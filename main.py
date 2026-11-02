"""
Извлечение триплетов + финальная сборка таблицы силами LLM (Арбитра).
"""
from __future__ import annotations

import asyncio
import json
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openai import AsyncOpenAI
from openpyxl import Workbook


# ============================================================
#                    ЧТЕНИЕ ФАЙЛОВ
# ============================================================

def pick_file() -> str | None:
    if len(sys.argv) > 1:
        p = sys.argv[1]
        if Path(p).is_file():
            return p
    try:
        from tkinter import Tk, filedialog
        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            path = filedialog.askopenfilename(title="Выберите файл", filetypes=[("Все файлы", "*.*")])
        finally:
            root.destroy()
        return path if path else None
    except Exception:
        return None


def _read_html(p: Path) -> str:
    from bs4 import BeautifulSoup
    raw = p.read_bytes()
    for enc in ("utf-8", "cp1251"):
        try:
            html = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        html = raw.decode("utf-8", errors="replace")
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style"]):
        tag.decompose()
    return soup.get_text("\n", strip=True)


def _read_docx(p: Path) -> str:
    from docx import Document
    doc = Document(p)
    parts = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def _read_xlsx(p: Path) -> list[tuple[str, str]]:
    from openpyxl import load_workbook
    wb = load_workbook(p, read_only=True, data_only=True)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]

    text_idx = id_idx = name_idx = None
    for i, h in enumerate(headers):
        if not h: continue
        hl = str(h).strip().lower()
        if hl in ("propertyvalue", "text", "текст", "value", "content"):
            text_idx = i
        elif hl in ("propertyid_ref", "id", "record_id"):
            id_idx = i
        elif hl in ("propertyname", "field", "поле", "type"):
            name_idx = i

    if text_idx is None:
        raise ValueError("Не найдена колонка с текстом")

    by_id: dict[str, list[str]] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[text_idx] is None: continue
        text = str(row[text_idx]).strip()
        if not text: continue
        rid = str(row[id_idx]) if id_idx is not None and row[id_idx] is not None else "1"
        prefix = f"=== {row[name_idx]} ===\n" if name_idx is not None and row[name_idx] is not None else ""
        by_id.setdefault(rid, []).append(prefix + text)
    wb.close()
    return [(rid, "\n\n".join(chunks)) for rid, chunks in by_id.items()]


def load_file(path: str) -> list[tuple[str, str]]:
    p = Path(path)
    ext = p.suffix.lower()
    if ext == ".txt": return [(p.stem, p.read_text(encoding="utf-8", errors="replace"))]
    if ext in (".html", ".htm"): return [(p.stem, _read_html(p))]
    if ext == ".docx": return [(p.stem, _read_docx(p))]
    if ext == ".xlsx": return _read_xlsx(p)
    raise ValueError(f"Неподдерживаемый формат: {ext}")


# ============================================================
#                   КОНФИГУРАЦИЯ И МОДЕЛИ
# ============================================================

def load_keys(path: str = "keys.txt") -> dict:
    cfg = {}
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"): continue
        k, v = line.split("=", 1)
        cfg[k.strip()] = v.strip()
    return cfg


def load_prompts(folder: str = "prompts") -> dict:
    base = Path(folder)
    def r(n): return (base / f"{n}.txt").read_text(encoding="utf-8").strip()
    return {
        "s1_sys": r("stage1_system"), "s1_usr": r("stage1_user"),
        "s2_sys": r("stage2_system"), "s2_usr": r("stage2_user"),
        "align_sys": r("align_system"), "align_usr": r("align_user"),
        "table_sys": r("table_system"), "table_usr": r("table_user"), # Новые промпты
    }


@dataclass
class Call:
    record_id: str
    model: str
    stage: str
    prompt: str
    raw: str
    parsed: list[dict]
    error: str | None = None
    timestamp: datetime = field(default_factory=lambda: datetime.now(timezone.utc))


class LLMClient:
    def __init__(self, name: str, api_key: str, base_url: str, max_tokens: int, temperature: float):
        self.name = name
        self.max_tokens = max_tokens
        self.temperature = temperature
        self.client = AsyncOpenAI(api_key=api_key, base_url=base_url, timeout=240)

    async def ask(self, prompt: str, system: str) -> tuple[str, str | None]:
        try:
            r = await self.client.chat.completions.create(
                model=self.name,
                messages=[{"role": "system", "content": system}, {"role": "user", "content": prompt}],
                max_tokens=self.max_tokens,
                temperature=self.temperature,
                response_format={"type": "json_object"},
            )
            return (r.choices[0].message.content or ""), None
        except Exception as e:
            return "", str(e)


# ============================================================
#                    ПАРСИНГ JSON
# ============================================================

_FENCE = re.compile(r"```(?:json)?\s*(.*?)\s*```", re.DOTALL)

def _load_json(raw: str):
    if not raw: return None
    try: return json.loads(raw)
    except json.JSONDecodeError: pass
    m = _FENCE.search(raw)
    if m:
        try: return json.loads(m.group(1))
        except json.JSONDecodeError: pass
    start = min((i for i in (raw.find("{"), raw.find("[")) if i >= 0), default=-1)
    end = max(raw.rfind("}"), raw.rfind("]"))
    if start >= 0 and end > start:
        try: return json.loads(raw[start:end + 1])
        except json.JSONDecodeError: pass
    return None

def _items(data) -> list[dict]:
    if isinstance(data, list): return [d for d in data if isinstance(d, dict)]
    if isinstance(data, dict):
        for v in data.values():
            if isinstance(v, list): return [d for d in v if isinstance(d, dict)]
    return []

def parse_pairs(raw: str) -> list[dict]:
    return [{"term": str(d.get("term", "")).strip(), "value": str(d.get("value", "")).strip()}
            for d in _items(_load_json(raw)) if str(d.get("term", "")).strip()]

def parse_triplets(raw: str) -> list[dict]:
    return [{"term": str(d.get("term", "")).strip(),
             "characteristic": str(d.get("characteristic", "")).strip(),
             "value": str(d.get("value", "")).strip()}
            for d in _items(_load_json(raw))
            if str(d.get("term", "")).strip() and str(d.get("characteristic", "")).strip()]


# ============================================================
#                    СРАВНЕНИЕ И АРБИТР
# ============================================================

def _norm(s: str) -> str:
    return " ".join(str(s).lower().strip().split())

def compare(items_a, items_b, stage, key_fields, value_field, model_a, model_b):
    idx_b = defaultdict(list)
    for it in items_b:
        idx_b[tuple(_norm(it[f]) for f in key_fields)].append(it)
    
    consensus, discrepancies, used_b = [], [], set()

    for a in items_a:
        k = tuple(_norm(a[f]) for f in key_fields)
        matches = idx_b.get(k, [])
        if not matches:
            discrepancies.append({"stage": stage, "kind": "presence_conflict", "model_a": model_a, "model_b": model_b,
                                  "item_a": json.dumps(a, ensure_ascii=False), "item_b": "", "key": " | ".join(k)})
            continue
        exact = next((b for b in matches if _norm(a[value_field]) == _norm(b[value_field])), None)
        if exact:
            consensus.append(a)
            used_b.add(id(exact))
        else:
            discrepancies.append({"stage": stage, "kind": "value_conflict", "model_a": model_a, "model_b": model_b,
                                  "item_a": json.dumps(a, ensure_ascii=False), "item_b": json.dumps(matches[0], ensure_ascii=False),
                                  "key": " | ".join(k)})
            used_b.add(id(matches[0]))

    for b in items_b:
        if id(b) not in used_b:
            k = tuple(_norm(b[f]) for f in key_fields)
            discrepancies.append({"stage": stage, "kind": "presence_conflict", "model_a": model_a, "model_b": model_b,
                                  "item_a": "", "item_b": json.dumps(b, ensure_ascii=False), "key": " | ".join(k)})
    return consensus, discrepancies

def _validate_items(items, expected_fields):
    result = []
    for d in _items(items):
        clean = {}
        ok = True
        for f in expected_fields:
            v = str(d.get(f, "")).strip()
            if not v:
                ok = False
                break
            clean[f] = v
        if ok: result.append(clean)
    return result

async def align_with_arbiter(record_id, items_a, items_b, arbiter, prompts, stage, expected_fields):
    prompt = prompts["align_usr"].format(
        stage=stage, expected_fields=", ".join(expected_fields),
        list_a=json.dumps({"items": items_a}, ensure_ascii=False, indent=2),
        list_b=json.dumps({"items": items_b}, ensure_ascii=False, indent=2),
    )
    raw, err = await arbiter.ask(prompt, prompts["align_sys"])
    call = Call(record_id, arbiter.name, f"align_{stage}", prompt, raw, [], err)

    if err: return [], items_a, items_b, call

    data = _load_json(raw) or {}
    agreed = _validate_items(data.get("agreed"), expected_fields) if isinstance(data, dict) else []
    only_a = _validate_items(data.get("only_a"), expected_fields) if isinstance(data, dict) else []
    only_b = _validate_items(data.get("only_b"), expected_fields) if isinstance(data, dict) else []

    key_fields = expected_fields[:2] if len(expected_fields) >= 2 else expected_fields
    seen_keys = {tuple(_norm(it.get(f, "")) for f in key_fields) for it in agreed + only_a + only_b}
    
    for it in items_a:
        k = tuple(_norm(it.get(f, "")) for f in key_fields)
        if k not in seen_keys:
            only_a.append(it); seen_keys.add(k)
    for it in items_b:
        k = tuple(_norm(it.get(f, "")) for f in key_fields)
        if k not in seen_keys:
            only_b.append(it); seen_keys.add(k)

    call.parsed = agreed
    return agreed, only_a, only_b, call


# ============================================================
#                    ПАЙПЛАЙН ИЗВЛЕЧЕНИЯ
# ============================================================

async def process_text(record_id, text, model_a, model_b, arbiter, prompts):
    audit_calls = []
    all_discrepancies = []

    p1 = prompts["s1_usr"].format(text=text)
    raw_a, err_a = await model_a.ask(p1, prompts["s1_sys"])
    raw_b, err_b = await model_b.ask(p1, prompts["s1_sys"])
    parsed_a, parsed_b = parse_pairs(raw_a), parse_pairs(raw_b)
    audit_calls += [
        Call(record_id, model_a.name, "stage1", p1, raw_a, parsed_a, err_a),
        Call(record_id, model_b.name, "stage1", p1, raw_b, parsed_b, err_b),
    ]

    cons_1_strict, disc_1 = compare(parsed_a, parsed_b, "stage1", ["term"], "value", model_a.name, model_b.name)

    if arbiter and (disc_1 or not cons_1_strict):
        agreed, only_a, only_b, call = await align_with_arbiter(record_id, parsed_a, parsed_b, arbiter, prompts, "stage1", ["term", "value"])
        audit_calls.append(call)
        cons_1 = agreed + only_a + only_b
        for it in only_a:
            all_discrepancies.append({"stage": "stage1", "kind": "only_a", "model_a": model_a.name, "model_b": model_b.name,
                                      "item_a": json.dumps(it, ensure_ascii=False), "item_b": "", "key": it.get("term", "")})
        for it in only_b:
            all_discrepancies.append({"stage": "stage1", "kind": "only_b", "model_a": model_a.name, "model_b": model_b.name,
                                      "item_a": "", "item_b": json.dumps(it, ensure_ascii=False), "key": it.get("term", "")})
    else:
        cons_1 = cons_1_strict
        all_discrepancies += disc_1

    cons_2 = []
    if cons_1:
        pairs_json = json.dumps({"items": cons_1}, ensure_ascii=False, indent=2)
        p2 = prompts["s2_usr"].format(pairs_json=pairs_json)
        raw_a_2, err_a_2 = await model_a.ask(p2, prompts["s2_sys"])
        raw_b_2, err_b_2 = await model_b.ask(p2, prompts["s2_sys"])
        parsed_a_2, parsed_b_2 = parse_triplets(raw_a_2), parse_triplets(raw_b_2)
        audit_calls += [
            Call(record_id, model_a.name, "stage2", p2, raw_a_2, parsed_a_2, err_a_2),
            Call(record_id, model_b.name, "stage2", p2, raw_b_2, parsed_b_2, err_b_2),
        ]

        cons_2_strict, disc_2 = compare(parsed_a_2, parsed_b_2, "stage2", ["term", "characteristic"], "value", model_a.name, model_b.name)

        if arbiter and (disc_2 or not cons_2_strict):
            agreed, only_a, only_b, call = await align_with_arbiter(record_id, parsed_a_2, parsed_b_2, arbiter, prompts, "stage2", ["term", "characteristic", "value"])
            audit_calls.append(call)
            cons_2 = agreed + only_a + only_b
            for it in only_a:
                all_discrepancies.append({"stage": "stage2", "kind": "only_a", "model_a": model_a.name, "model_b": model_b.name,
                                          "item_a": json.dumps(it, ensure_ascii=False), "item_b": "", "key": f"{it.get('term','')} | {it.get('characteristic','')}"})
            for it in only_b:
                all_discrepancies.append({"stage": "stage2", "kind": "only_b", "model_a": model_a.name, "model_b": model_b.name,
                                          "item_a": "", "item_b": json.dumps(it, ensure_ascii=False), "key": f"{it.get('term','')} | {it.get('characteristic','')}"})
        else:
            cons_2 = cons_2_strict
            all_discrepancies += disc_2

    return {"triplets": cons_2, "discrepancies": all_discrepancies, "calls": audit_calls}


# ============================================================
#        ШАГ 3: СИНТЕЗ ТАБЛИЦЫ СИЛАМИ LLM
# ============================================================

async def synthesize_table(record_id: str, triplets: list[dict], model: LLMClient, prompts: dict) -> tuple[dict, Call]:
    """Просит LLM собрать чистое резюме по пациенту."""
    if not triplets:
        return {}, Call(record_id, model.name, "table_synth", "", "", [])

    triplets_json = json.dumps(triplets, ensure_ascii=False, indent=2)
    prompt = prompts["table_usr"].format(triplets=triplets_json)
    
    raw, err = await model.ask(prompt, prompts["table_sys"])
    call = Call(record_id, model.name, "table_synth", prompt, raw, [], err)
    
    if err: return {}, call
    
    data = _load_json(raw)
    if isinstance(data, dict):
        # Ожидаем, что LLM вернет {"feature1": "value", "feature2": "value"}
        # или {"data": {...}}
        if "data" in data and isinstance(data["data"], dict):
            return data["data"], call
        
        # Если это просто словарь, фильтруем системные ключи
        clean = {k: str(v) for k, v in data.items() if not k.startswith("_")}
        return clean, call
        
    return {}, call


# ============================================================
#                    СОХРАНЕНИЕ В EXCEL
# ============================================================

def save_excel(path: str, results: list[dict]):
    wb = Workbook()
    
    # 1. Собираем все уникальные колонки из всех синтезированных словарей
    all_columns = set()
    for r in results:
        table_data = r.get("table_data", {})
        all_columns.update(table_data.keys())
    
    # Сортируем для красоты
    columns = sorted(list(all_columns))
    
    # 2. Лист "Датасет"
    ws1 = wb.active
    ws1.title = "Датасет"
    ws1.append(["record_id"] + columns)
    
    for r in results:
        row = [r["record_id"]]
        table_data = r.get("table_data", {})
        for col in columns:
            row.append(table_data.get(col, ""))
        ws1.append(row)

    # 3. Лист "Все_триплеты" (исходники)
    ws2 = wb.create_sheet("Все_триплеты")
    ws2.append(["record_id", "Термин", "Характеристика", "Значение"])
    for r in results:
        for t in r["result"]["triplets"]:
            ws2.append([r["record_id"], t["term"], t["characteristic"], t["value"]])

    # 4. Лист "Расхождения"
    ws3 = wb.create_sheet("Расхождения")
    ws3.append(["record_id", "Этап", "Тип", "Модель A", "Модель B", "A", "B", "Ключ"])
    for r in results:
        for d in r["result"]["discrepancies"]:
            ws3.append([r["record_id"], d.get("stage"), d.get("kind"), d.get("model_a"), d.get("model_b"),
                        d.get("item_a"), d.get("item_b"), d.get("key")])

    wb.save(path)


# ============================================================
#                        MAIN
# ============================================================

async def main():
    print("=" * 60)
    print("  Извлечение и финальная сборка таблицы (LLM Synthesis)")
    print("=" * 60)

    path = pick_file()
    if not path:
        print("Файл не выбран.")
        return
    print(f"Файл: {path}")

    cfg = load_keys("keys.txt")
    prompts = load_prompts("prompts")

    def make_client(name):
        return LLMClient(name=name, api_key=cfg["API_KEY"], base_url=cfg["BASE_URL"],
                         max_tokens=int(cfg.get("MAX_TOKENS", 8192)),
                         temperature=float(cfg.get("TEMPERATURE", 0.1)))

    model_a = make_client(cfg["MODEL_A"])
    model_b = make_client(cfg["MODEL_B"])
    
    # Арбитр обязателен для финального синтеза
    arb_name = cfg.get("ARBITER_MODEL", cfg["MODEL_A"]).strip()
    arbiter = make_client(arb_name)

    print(f"Модель A: {model_a.name}")
    print(f"Модель B: {model_b.name}")
    print(f"Сборщик:  {arbiter.name}\n")

    records = load_file(path)
    print(f"Записей: {len(records)}\n")

    all_results = []
    
    # Обрабатываем записи параллельно для скорости
    sem = asyncio.Semaphore(3) # Не более 3 записей одновременно

    async def process_record(i, rid, text):
        async with sem:
            print(f"[{i}/{len(records)}] Извлечение для {rid}...")
            # Шаг 1 и 2 (Триплеты)
            res = await process_text(rid, text, model_a, model_b, arbiter, prompts)
            
            print(f"[{i}/{len(records)}] Синтез таблицы для {rid}...")
            # Шаг 3 (Синтез)
            table_data, synth_call = await synthesize_table(rid, res["triplets"], arbiter, prompts)
            res["calls"].append(synth_call)
            
            return {"record_id": rid, "result": res, "table_data": table_data}

    tasks = [process_record(i, rid, text) for i, (rid, text) in enumerate(records, 1)]
    all_results = await asyncio.gather(*tasks)

    # --- СОХРАНЕНИЕ ---
    src = Path(path)
    out = src.parent / f"{src.stem}_synthesized.xlsx"
    save_excel(str(out), all_results)
    print(f"\nСохранено: {out}")

    try:
        import os
        os.startfile(str(out))
    except Exception:
        pass


if __name__ == "__main__":
    asyncio.run(main())