"""Патч для parse_all_llm.py: добавляет retry-логику в call_llm."""
import re

with open("parse_all_llm.py", "r", encoding="utf-8") as f:
    content = f.read()

old = '''def call_llm(system_prompt: str, user_text: str, model: str) -> str:
    url = f"{API_BASE}/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_text},
        ],
        "temperature": 0.0,
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=600)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"]'''

new = '''def call_llm(system_prompt: str, user_text: str, model: str,
             max_retries: int = 3, timeout: int = 900) -> str:
    url = f"{API_BASE}/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_text},
        ],
        "temperature": 0.0,
    }
    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=timeout)
            resp.raise_for_status()
            data = resp.json()
            return data["choices"][0]["message"]["content"]
        except (requests.exceptions.Timeout, requests.exceptions.HTTPError) as e:
            last_error = e
            print(f"      [!] Попытка {attempt}/{max_retries} не удалась: {e}")
            if attempt < max_retries:
                wait = 15 * attempt
                print(f"      Жду {wait} сек перед повтором...")
                time.sleep(wait)
    raise last_error'''

if old not in content:
    print("ОШИБКА: старая функция не найдена")
    raise SystemExit(1)

content = content.replace(old, new)

with open("parse_all_llm.py", "w", encoding="utf-8") as f:
    f.write(content)

print("OK: патч применён")
