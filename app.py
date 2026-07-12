Ты же этот адаптер используешь?
# %%
from openai import OpenAI
import time

def print_model_list(client: OpenAI):
    try:
        # Получаем список моделей
        models = client.models.list()
        
        print("Доступные модели:")
        print("-" * 30)
        
        # Проходимся по объектам моделей и выводим их ID
        for model in models.data:
            print(f"ID: {model.id} | Тип: {model.object}")
            
        print("-" * 30)
        print(f"Всего моделей: {len(models.data)}")

        print("-" * 30)
        
    except Exception as e:
        print(f"Произошла ошибка: {e}")

if __name__ == "__main__":

    client = OpenAI(
        base_url="BASE_URL",
        api_key="API_KEY",
    )

    print_model_list(client)

    messages = [
        {"role": "system", "content": "Ты — эксперт по Python."},
        {"role": "user", "content": "Напиши функцию для сортировки списка."}
    ]

    extra_body = {
        "top_k": 20,
        "min_p": 0.0,
        "repetition_penalty": 1.0,

        # Необходимые параметры для отключения режима размышлений в Qwen 3.5
        # Вариант 1
        "chat_template_kwargs": {"enable_thinking": False},
        # Вариант 2
        # "reasoning_effort": "none"
    }

    start_time = time.perf_counter()

    try:
        chat_response = client.chat.completions.create(
            model="MODEL_NAME",
            messages=messages,
            max_tokens=3000,
            temperature=0.7,
            top_p=0.8,
            presence_penalty=1.5,
            extra_body=extra_body,
            timeout=60
        )
        response_text = chat_response.choices[0].message.content

    except Exception as e:
        response_text = f"Ошибка при запросе к API: {str(e)}"

    time_spent = time.perf_counter() - start_time

    print("Ответ модели:")
    print(response_text)
    print("-" * 30)
    print(f"Время, затраченное на ответ: {time_spent:.4f} секунд")
 # %%
