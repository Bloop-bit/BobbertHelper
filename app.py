# =============================================
# 0. УСТАНОВКА КОДИРОВКИ UTF-8
# =============================================
import os
import sys

os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONUTF8"] = "1"

try:
    sys.stdout.reconfigure(encoding='utf-8')
except (AttributeError, ValueError, TypeError):
    pass

# =============================================
# 1. ИМПОРТЫ
# =============================================
import re
import json
from typing import List, Dict, Optional
from io import BytesIO

import streamlit as st
import pandas as pd
from docx import Document
from openpyxl import load_workbook
from openai import OpenAI

# =============================================
# 2. ПОЛУЧЕНИЕ API-КЛЮЧА
# =============================================
CLOUD_RU_API_KEY = os.getenv("CLOUD_RU_API_KEY")

# =============================================
# 3. АВТООПРЕДЕЛЕНИЕ КОЛОНОК
# =============================================
def detect_text_column(headers: List[str], df_sample: pd.DataFrame) -> Optional[str]:
    keywords = [
        "propertyvalue", "описание", "жалобы", "анамнез", "заключение",
        "текст", "anamnesis", "complaints", "diagnosis", "text",
        "описание случая", "жалобы при поступлении", "жалобы пациента",
        "анамнез заболевания", "notes", "history", "description",
        "симптомы", "диагноз", "пациент", "жалобы", "медицинские записи"
    ]
    for col in headers:
        col_lower = col.lower().strip()
        if any(kw in col_lower for kw in keywords):
            return col
    if df_sample is not None and not df_sample.empty:
        avg_lengths = {}
        for col in df_sample.columns:
            sample = df_sample[col].dropna()
            if len(sample) == 0:
                continue
            try:
                numeric_ratio = pd.to_numeric(sample, errors='coerce').notna().mean()
                if numeric_ratio > 0.5:
                    continue
            except:
                pass
            avg_len = sample.astype(str).str.len().mean()
            avg_lengths[col] = avg_len
        if avg_lengths:
            return max(avg_lengths, key=avg_lengths.get)
    for col in headers:
        if df_sample is not None and df_sample[col].dtype == 'object':
            return col
    return None

def detect_id_column(headers: List[str], df_sample: pd.DataFrame, text_col: Optional[str] = None) -> Optional[str]:
    keywords = ["id", "номер", "код", "patient", "propertyid", "case", "идентификатор"]
    for col in headers:
        col_lower = col.lower().strip()
        if any(kw in col_lower for kw in keywords):
            return col
    if df_sample is not None and not df_sample.empty:
        unique_counts = {}
        for col in df_sample.columns:
            if col == text_col:
                continue
            n_unique = df_sample[col].nunique()
            if n_unique > 2 and n_unique / len(df_sample) > 0.05:
                unique_counts[col] = n_unique
        if unique_counts:
            return max(unique_counts, key=unique_counts.get)
    return None

# =============================================
# 4. ФУНКЦИИ ЧТЕНИЯ ФАЙЛОВ
# =============================================
def read_xlsx_full(file_path: str, text_col: Optional[str] = None, id_col: Optional[str] = None) -> tuple:
    wb = load_workbook(file_path, data_only=True)
    all_records = []
    full_dfs = []
    text_col_name = text_col
    id_col_name = id_col
    
    for sheet in wb.worksheets:
        data = sheet.values
        headers = None
        rows_data = []
        for i, row in enumerate(data):
            if i == 0:
                headers = [str(cell) if cell else "" for cell in row]
            else:
                rows_data.append(row)
        if not headers:
            continue
        
        df = pd.DataFrame(rows_data, columns=headers)
        
        if text_col_name is None:
            text_col_name = detect_text_column(headers, df.head(100))
            id_col_name = detect_id_column(headers, df.head(100), text_col_name)
        
        if text_col_name is None or text_col_name not in df.columns:
            continue
        
        df['_source'] = f"{os.path.basename(file_path)}, лист {sheet.title}, строка "
        df['_row_num'] = range(1, len(df) + 1)
        df['_full_source'] = df['_source'] + df['_row_num'].astype(str)
        
        for idx, row in df.iterrows():
            text = str(row[text_col_name]) if pd.notna(row[text_col_name]) else ""
            if text.strip():
                patient_id = str(row[id_col_name]) if id_col_name and id_col_name in df.columns and pd.notna(row[id_col_name]) else None
                all_records.append({
                    'source': row['_full_source'],
                    'text': text.strip(),
                    'patient_id': patient_id,
                    'row_data': row.to_dict()
                })
        
        full_dfs.append(df)
    
    full_df = pd.concat(full_dfs, ignore_index=True) if full_dfs else pd.DataFrame()
    return all_records, full_df, text_col_name, id_col_name

def read_docx_full(file_path: str, text_col: Optional[str] = None, id_col: Optional[str] = None) -> tuple:
    doc = Document(file_path)
    records = []
    row_counter = 0
    full_dfs = []
    text_col_name = text_col
    id_col_name = id_col
    
    for table in doc.tables:
        if not table.rows:
            continue
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        
        if text_col_name is None:
            sample_rows = []
            for row in table.rows[1:11]:
                sample_rows.append([cell.text.strip() for cell in row.cells])
            if headers and sample_rows:
                df_sample = pd.DataFrame(sample_rows, columns=headers)
                text_col_name = detect_text_column(headers, df_sample)
                id_col_name = detect_id_column(headers, df_sample, text_col_name)
            else:
                continue
        
        if text_col_name is None or text_col_name not in headers:
            continue
        
        rows_data = []
        for row in table.rows[1:]:
            row_counter += 1
            cells = row.cells
            row_data = {headers[i]: cells[i].text.strip() if i < len(cells) else "" for i in range(len(headers))}
            rows_data.append(row_data)
            
            text = row_data.get(text_col_name, "")
            if text.strip():
                source = f"{os.path.basename(file_path)}, строка {row_counter}"
                patient_id = row_data.get(id_col_name) if id_col_name else None
                records.append({
                    'source': source,
                    'text': text.strip(),
                    'patient_id': patient_id,
                    'row_data': row_data
                })
        
        if rows_data:
            df = pd.DataFrame(rows_data)
            df['_source'] = f"{os.path.basename(file_path)}, строка "
            df['_row_num'] = range(1, len(df) + 1)
            df['_full_source'] = df['_source'] + df['_row_num'].astype(str)
            full_dfs.append(df)
    
    full_df = pd.concat(full_dfs, ignore_index=True) if full_dfs else pd.DataFrame()
    return records, full_df, text_col_name, id_col_name

def read_txt_full(file_path: str) -> tuple:
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    records = []
    for i, line in enumerate(lines, 1):
        text = line.strip()
        if text:
            source = f"{os.path.basename(file_path)}, строка {i}"
            records.append({
                'source': source,
                'text': text,
                'patient_id': None,
                'row_data': {'text': text}
            })
    full_df = pd.DataFrame([r['row_data'] for r in records]) if records else pd.DataFrame()
    return records, full_df, None, None

def detect_file_type(file_path: str) -> Optional[str]:
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.docx': return 'docx'
    elif ext == '.xlsx': return 'xlsx'
    elif ext == '.txt': return 'txt'
    return None

def load_folder_full(folder_path: str, text_col: Optional[str] = None, id_col: Optional[str] = None) -> tuple:
    os.makedirs(folder_path, exist_ok=True)
    all_records = []
    all_dfs = []
    text_col_name = text_col
    id_col_name = id_col
    
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        ft = detect_file_type(file_path)
        records = []
        df = pd.DataFrame()
        
        if ft == 'docx':
            records, df, text_col_name, id_col_name = read_docx_full(file_path, text_col_name, id_col_name)
        elif ft == 'xlsx':
            records, df, text_col_name, id_col_name = read_xlsx_full(file_path, text_col_name, id_col_name)
        elif ft == 'txt':
            records, df, _, _ = read_txt_full(file_path)
        
        all_records.extend(records)
        if not df.empty:
            all_dfs.append(df)
    
    full_df = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
    return all_records, full_df, text_col_name, id_col_name

# =============================================
# 5. АДАПТЕР ДЛЯ CLOUD.RU
# =============================================
class CloudRuAdapter:
    def __init__(self, model: str = "Qwen/Qwen3-30B-A3B", api_key: str = None, timeout: int = 300):
        self.model = model
        self.api_key = api_key
        if not self.api_key:
            raise ValueError("API-ключ Cloud.ru не указан")
        self.client = OpenAI(
            base_url="https://foundation-models.api.cloud.ru/v1",
            api_key=self.api_key,
            timeout=timeout
        )

    def generate(self, system_prompt: str, user_prompt: str, 
                 temperature: float = 0.5, max_tokens: int = 2500,
                 top_p: float = 0.95, presence_penalty: float = 0,
                 enable_thinking: bool = False, **kwargs) -> str:
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": user_prompt})
        
        extra_body = {
            "top_k": 20,
            "min_p": 0.0,
            "repetition_penalty": 1.0,
        }
        if not enable_thinking:
            extra_body["chat_template_kwargs"] = {"enable_thinking": False}
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            max_tokens=max_tokens,
            temperature=temperature,
            top_p=top_p,
            presence_penalty=presence_penalty,
            extra_body=extra_body,
            timeout=self.client.timeout
        )
        return response.choices[0].message.content

# =============================================
# 6. ИЗВЛЕЧЕНИЕ ПАР
# =============================================
def extract_pairs(text: str, llm, enable_thinking: bool = False) -> list:
    if len(text) > 1500:
        text = text[:1500] + "..."
    prompt = f"""
Извлеки из текста все медицинские термины (симптомы, диагнозы, состояния) и их наличие/значение.
Верни JSON-массив массивов, каждый подмассив из двух строк: ["термин", "значение"].
Возможные значения: присутствует, отсутствует, усиливается, ослабевает, постоянный, периодический и т.п.
Только JSON, без пояснений.

Текст:
{text}
"""
    content = llm.generate(
        system_prompt="Ты медицинский эксперт. Отвечай только JSON.",
        user_prompt=prompt,
        temperature=0,
        enable_thinking=enable_thinking
    )
    content = content.strip().removeprefix("```json").removesuffix("```").strip()
    try:
        data = json.loads(content)
    except:
        match = re.search(r'\[.*\]', content, re.DOTALL)
        if match:
            try:
                data = json.loads(match.group())
            except:
                data = []
        else:
            data = []
    if isinstance(data, dict) and "terms" in data:
        data = data["terms"]
    if not isinstance(data, list):
        data = []
    filtered = []
    for item in data:
        if isinstance(item, (list, tuple)) and len(item) >= 2:
            filtered.append([item[0], item[1]])
    return filtered

# =============================================
# 7. ДОБАВЛЕНИЕ ХАРАКТЕРИСТИК
# =============================================
def enrich_with_characteristics(pairs: list, llm, enable_thinking: bool = False) -> list:
    unique_values = set(val for _, val in pairs)
    value_to_char = {}

    for value in unique_values:
        prompt = f"""
Определи категорию (характеристику) для медицинского значения "{value}".
Категория должна быть существительным или короткой фразой, обобщающей тип значения.
Примеры:
- "присутствует" → "присутствие"
- "отсутствует" → "отсутствие"
- "усиливается" → "динамика"
- "постоянный" → "характер"
- "ноющая" → "интенсивность" (или "характер боли")
- "при нагрузке" → "условие возникновения"
- "5 мг" → "дозировка"

Верни ТОЛЬКО категорию, без пояснений. Одно слово или короткая фраза.
"""
        try:
            response = llm.generate(
                system_prompt="Ты эксперт по медицинской терминологии.",
                user_prompt=prompt,
                temperature=0,
                enable_thinking=enable_thinking
            )
            char = response.strip().strip('"').strip("'")
            if not char or len(char) > 100:
                char = "характеристика"
            value_to_char[value] = char
        except Exception as e:
            value_to_char[value] = "характеристика"

    triples = []
    for term, value in pairs:
        char = value_to_char.get(value, "характеристика")
        triples.append((term, char, value))
    return triples

# =============================================
# 8. ОСНОВНОЙ ПАЙПЛАЙН
# =============================================
def run_pipeline(folder: str, model: str, api_key: str,
                 text_col: Optional[str] = None, id_col: Optional[str] = None,
                 enable_thinking: bool = False, timeout: int = 300):
    llm = CloudRuAdapter(api_key=api_key, model=model, timeout=timeout)
    os.makedirs(folder, exist_ok=True)

    records, full_df, text_col_name, id_col_name = load_folder_full(
        folder_path=folder, text_col=text_col, id_col=id_col
    )
    
    if not records:
        return None, None, None, None, None

    all_pairs = []
    for rec in records:
        pairs = extract_pairs(rec['text'], llm, enable_thinking=enable_thinking)
        if not pairs:
            continue
        for item in pairs:
            t, v = item[0], item[1]
            all_pairs.append({
                'term': t,
                'value': v,
                'source': rec['source'],
                'source_text': rec['text'],
                'patient_id': rec.get('patient_id')
            })

    if not all_pairs:
        return full_df, [], [], None, None

    # Собираем уникальные пары для получения характеристик
    unique_pairs = list(set((p['term'], p['value']) for p in all_pairs))
    pair_list = [(term, value) for term, value in unique_pairs]
    triples = enrich_with_characteristics(pair_list, llm, enable_thinking=enable_thinking)
    
    # Создаём словарь для быстрого поиска характеристики
    char_dict = {}
    for term, char, value in triples:
        char_dict[(term, value)] = char

    # Обогащаем all_pairs характеристиками
    for p in all_pairs:
        p['characteristic'] = char_dict.get((p['term'], p['value']), 'характеристика')

    # Создаём признаковую матрицу
    term_features = {}
    for p in all_pairs:
        feature_name = f"{p['term']}_{p['characteristic']}"
        if feature_name not in term_features:
            term_features[feature_name] = {}
        term_features[feature_name][p['source']] = p['value']

    # Создаём DataFrame признаков
    feature_df = pd.DataFrame.from_dict(term_features, orient='index').T
    feature_df = feature_df.fillna("")
    
    # Добавляем признаки в full_df
    result_df = full_df.copy()
    
    # Удаляем колонку с текстом
    if text_col_name and text_col_name in result_df.columns:
        result_df = result_df.drop(columns=[text_col_name])
    
    # Добавляем признаки по source
    for feature_name in feature_df.columns:
        result_df[feature_name] = result_df['_full_source'].map(feature_df[feature_name]).fillna("")
    
    # Удаляем служебные колонки
    for col in ['_source', '_row_num', '_full_source']:
        if col in result_df.columns:
            result_df = result_df.drop(columns=[col])

    return result_df, records, all_pairs, triples, feature_df

# =============================================
# 9. ФУНКЦИЯ ДЛЯ ПРЕОБРАЗОВАНИЯ В БИНАРНЫЙ ДАТАСЕТ
# =============================================
def create_binary_dataset(result_df: pd.DataFrame, feature_cols: List[str]) -> pd.DataFrame:
    """
    Преобразует текстовые значения (присутствует/отсутствует) в бинарные (1/0)
    """
    binary_df = result_df.copy()
    
    # Список слов, означающих "присутствует"
    present_keywords = ['присутствует', 'имеется', 'есть', 'выявлен', 'обнаружен', 'положительный', 'да', 'yes']
    # Список слов, означающих "отсутствует"
    absent_keywords = ['отсутствует', 'нет', 'не выявлен', 'не обнаружен', 'отрицательный', 'no']
    
    # Словарь для маппинга значений
    value_map = {}
    
    # Добавляем все возможные варианты
    for word in present_keywords:
        value_map[word] = 1
    for word in absent_keywords:
        value_map[word] = 0
    
    # Также обрабатываем числовые значения и другие текстовые
    for col in feature_cols:
        if col in binary_df.columns:
            # Применяем маппинг
            binary_df[col] = binary_df[col].map(value_map)
            
            # Если остались NaN или пустые строки, заполняем 0
            binary_df[col] = binary_df[col].fillna(0)
            
            # Если остались текстовые значения, пытаемся преобразовать
            # Если значение содержит 'присутствует' или похожее - 1, иначе 0
            mask = binary_df[col].apply(lambda x: isinstance(x, str) and any(kw in x.lower() for kw in present_keywords))
            binary_df.loc[mask, col] = 1
            
            # Если значение содержит 'отсутствует' или похожее - 0
            mask = binary_df[col].apply(lambda x: isinstance(x, str) and any(kw in x.lower() for kw in absent_keywords))
            binary_df.loc[mask, col] = 0
            
            # Преобразуем в int
            binary_df[col] = pd.to_numeric(binary_df[col], errors='coerce').fillna(0).astype(int)
    
    return binary_df

# =============================================
# 10. STREAMLIT ИНТЕРФЕЙС
# =============================================
st.set_page_config(page_title="Медицинский парсер", layout="centered")
st.title("🏥 Преобразование медицинских записей в датасет")
st.markdown("Загрузите файлы (Excel, Word, TXT). Исходная колонка с текстом будет заменена на извлечённые признаки.")

if CLOUD_RU_API_KEY:
    api_key_input = CLOUD_RU_API_KEY
    st.success("🔑 API-ключ загружен из секретов")
else:
    api_key_input = st.text_input("Введите ваш API-ключ Cloud.ru", type="password")

uploaded_files = st.file_uploader(
    "Выберите файлы",
    accept_multiple_files=True,
    type=['xlsx', 'docx', 'txt']
)

text_col = None
id_col = None

if uploaded_files:
    xlsx_files = [f for f in uploaded_files if f.name.endswith('.xlsx')]
    if xlsx_files:
        try:
            df_sample = pd.read_excel(xlsx_files[0], nrows=1)
            headers = df_sample.columns.tolist()
            if headers:
                temp_text = detect_text_column(headers, df_sample)
                temp_id = detect_id_column(headers, df_sample, temp_text)
                st.write("Выберите колонки для обработки (или оставьте автоопределение):")
                col1, col2 = st.columns(2)
                with col1:
                    text_col = st.selectbox(
                        "Колонка с текстом (описаниями):",
                        options=["(автоопределение)"] + headers,
                        index=0 if temp_text is None else headers.index(temp_text) + 1
                    )
                    if text_col == "(автоопределение)":
                        text_col = None
                with col2:
                    id_col = st.selectbox(
                        "Колонка с ID пациента (если есть):",
                        options=["(автоопределение)"] + headers,
                        index=0 if temp_id is None else headers.index(temp_id) + 1
                    )
                    if id_col == "(автоопределение)":
                        id_col = None
        except Exception as e:
            st.warning(f"Не удалось прочитать заголовки: {e}")

enable_thinking = st.checkbox(
    "🧠 Включить режим размышлений (thinking mode)",
    value=False,
    help="Увеличивает точность на сложных задачах, но замедляет работу"
)

if st.button("Обработать"):
    if not api_key_input:
        st.error("Пожалуйста, введите API-ключ.")
    elif not uploaded_files:
        st.error("Загрузите хотя бы один файл.")
    else:
        input_dir = "input"
        os.makedirs(input_dir, exist_ok=True)
        for f in os.listdir(input_dir):
            os.remove(os.path.join(input_dir, f))
        for uploaded_file in uploaded_files:
            with open(os.path.join(input_dir, uploaded_file.name), "wb") as f:
                f.write(uploaded_file.getbuffer())

        with st.spinner("Идёт обработка... Это может занять несколько минут."):
            try:
                result_df, records, all_pairs, triples, feature_df = run_pipeline(
                    folder="input",
                    model="Qwen/Qwen3-30B-A3B",
                    api_key=api_key_input,
                    text_col=text_col,
                    id_col=id_col,
                    enable_thinking=enable_thinking
                )

                if result_df is None or result_df.empty:
                    st.error("Не удалось загрузить данные. Проверьте формат файлов.")
                else:
                    # Определяем колонки-признаки
                    feature_cols = [col for col in result_df.columns if col not in ['PersonID_Ref', 'Sex', 'AGE', 'Death', 'TARGET', 'ServiceID', 'ServiceCode', 'ServiceName', 'StartDate', 'EndDate', 'CardMKB', 'MCardMKB', 'MKBCode_Ref', 'PropertyID_Ref', 'PropertyName', 'NormDescription', 'MinNormPropertyValue', 'MaxNormPropertyValue', 'MeasureID_Ref']]
                    
                    # Создаём бинарный датасет
                    binary_df = create_binary_dataset(result_df, feature_cols)
                    
                    # Переименовываем колонки-признаки (убираем суффиксы, добавляем пометку)
                    new_columns = {}
                    for col in binary_df.columns:
                        if col in feature_cols:
                            # Убираем суффикс после последнего подчёркивания
                            parts = col.rsplit('_', 1)
                            if len(parts) == 2:
                                # Если часть после подчёркивания — характеристика, убираем её
                                # Оставляем только название термина
                                term_name = parts[0]
                                new_columns[col] = f"{term_name} (0 - отсутствует, 1 - присутствует)"
                            else:
                                new_columns[col] = f"{col} (0 - отсутствует, 1 - присутствует)"
                        else:
                            new_columns[col] = col
                    
                    binary_df = binary_df.rename(columns=new_columns)
                    
                    # Переставляем колонки: сначала исходные, потом признаки
                    original_cols = [col for col in binary_df.columns if ' (0 - отсутствует, 1 - присутствует)' not in col]
                    feature_cols_renamed = [col for col in binary_df.columns if ' (0 - отсутствует, 1 - присутствует)' in col]
                    binary_df = binary_df[original_cols + feature_cols_renamed]
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # Страница 1: Исходные записи
                        if records:
                            pd.DataFrame(records).to_excel(writer, sheet_name='Исходные записи', index=False)
                        
                        # Страница 2: Пары
                        if all_pairs:
                            pd.DataFrame(all_pairs).to_excel(writer, sheet_name='Пары', index=False)
                        
                        # Страница 3: Тройки
                        if triples:
                            pd.DataFrame(triples, columns=['Термин', 'Характеристика', 'Значение']).to_excel(writer, sheet_name='Тройки', index=False)
                        
                        # Страница 4: Словарь характеристик
                        if all_pairs:
                            char_dict = {}
                            for p in all_pairs:
                                char_dict[p['value']] = p.get('characteristic', 'характеристика')
                            df_char = pd.DataFrame(list(char_dict.items()), columns=['Значение', 'Характеристика'])
                            df_char.to_excel(writer, sheet_name='Словарь характеристик', index=False)
                        
                        # Страница 5: Конечный датасет (бинарный)
                        binary_df.to_excel(writer, sheet_name='Конечный датасет', index=False)
                    
                    output.seek(0)

                    st.success(f"Обработка завершена! Создано {len(binary_df.columns)} колонок в конечном датасете.")
                    st.download_button(
                        label="📥 Скачать результат.xlsx",
                        data=output,
                        file_name="результат.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Ошибка: {e}")
                st.code(str(e))

if __name__ == "__main__":
    pass
