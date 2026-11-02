import streamlit as st
import asyncio
import os
from pathlib import Path
from io import BytesIO
import pandas as pd

# Импортируем наши функции из main.py
from main import process_text, synthesize_table, load_prompts, LLMClient

# Настройка страницы
st.set_page_config(page_title="Извлечение данных из ЭМК", page_icon="🏥", layout="wide")

st.title("🏥 Извлечение структурированных данных из клинических текстов")
st.markdown("""
**Как пользоваться:**
1. Загрузите файл (Excel, Word или HTML).
2. Нажмите «Запустить обработку».
3. Дождитесь окончания (это займет время, так как работают 3 нейросети).
4. Скачайте готовый Excel-файл.
""")

# Загружаем ключи из настроек Streamlit
# (Об этом ниже, в инструкции по деплою)
try:
    API_KEY = st.secrets["API_KEY"]
    BASE_URL = st.secrets["BASE_URL"]
    MODEL_A = st.secrets["MODEL_A"]
    MODEL_B = st.secrets["MODEL_B"]
    ARBITER_MODEL = st.secrets["ARBITER_MODEL"]
except KeyError:
    st.error("Ошибка: не настроены API-ключи. Добавьте их в настройки Streamlit.")
    st.stop()

# Загрузка файла
uploaded_file = st.file_uploader("Выберите файл с данными", type=["xlsx", "docx", "html", "txt"])

if uploaded_file is not None:
    # Сохраняем загруженный файл во временную папку
    temp_dir = Path("temp_uploads")
    temp_dir.mkdir(exist_ok=True)
    file_path = temp_dir / uploaded_file.name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    if st.button("🚀 Запустить обработку"):
        st.info("Обработка запущена. Пожалуйста, не закрывайте страницу.")
        
        # Прогресс-бар
        progress_bar = st.progress(0)
        status_text = st.empty()

        # Инициализация моделей
        def make_client(name):
            return LLMClient(name=name, api_key=API_KEY, base_url=BASE_URL,
                             max_tokens=8192, temperature=0.1)
        
        model_a = make_client(MODEL_A)
        model_b = make_client(MODEL_B)
        arbiter = make_client(ARBITER_MODEL)
        
        prompts = load_prompts("prompts")
        
        # Запуск асинхронного пайплайна
        async def run_pipeline():
            # В main.py у вас есть функция load_file, которая возвращает список (id, text)
            from main import load_file
            records = load_file(str(file_path))
            
            results = []
            total = len(records)
            
            for i, (rid, text) in enumerate(records, 1):
                status_text.text(f"Обработка записи {rid} ({i} из {total})...")
                
                # Шаг 1 и 2: Извлечение
                res = await process_text(rid, text, model_a, model_b, arbiter, prompts)
                
                # Шаг 3: Синтез
                table_data, synth_call = await synthesize_table(rid, res["triplets"], arbiter, prompts)
                res["calls"].append(synth_call)
                
                results.append({"record_id": rid, "result": res, "table_data": table_data})
                
                progress_bar.progress(i / total)
                
            return results

        # Запускаем асинхронный код в Streamlit
        try:
            all_results = asyncio.run(run_pipeline())
            
            # Формируем итоговый Excel в памяти
            from main import save_excel
            out_path = temp_dir / f"processed_{uploaded_file.name}"
            save_excel(str(out_path), all_results)
            
            st.success("✅ Обработка завершена!")
            
            # Кнопка скачивания
            with open(out_path, "rb") as f:
                st.download_button(
                    label="📥 Скачать готовый Excel",
                    data=f,
                    file_name=f"result_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # Предпросмотр таблицы
            st.subheader("Предпросмотр результата")
            # Собираем все колонки
            all_cols = set()
            for r in all_results:
                all_cols.update(r.get("table_data", {}).keys())
            
            df_data = []
            for r in all_results:
                row = {"record_id": r["record_id"]}
                row.update(r.get("table_data", {}))
                df_data.append(row)
            
            df = pd.DataFrame(df_data)
            st.dataframe(df)
            
        except Exception as e:
            st.error(f"Произошла ошибка: {e}")