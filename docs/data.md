# Архитектура данных (data.md)

## 1. Файл конфигурации (Set.ini)
Конфигурация разделена на секции для удобного парсинга через `configparser`.

```ini
[AI]
provider = deepseek # Активный провайдер (влияет на маршрутизацию API в будущем)

[GOOGLE]
api_key = AIzaSy... # Ключ для Gemini
model = gemini-3.7-flash
temperature = 0.2

[DEEPSEEK]
api_key = sk-... # Ключ для DeepSeek
model = deepseek-chat
base_url = https://api.deepseek.com
temperature = 0.2

[TAGS]
e_top = ... # Список ФИО или тегов высшего руководства (запятая как разделитель)
e_dir = ... # Список ФИО или тегов директоров

[NOTES]
default_days = 2 # Глубина поиска писем по умолчанию
max_emails = 50 # Жесткий лимит выборки
max_body_chars = 1500 # Лимит символов тела письма для отправки в LLM (экономия токенов)
```

## 2. Схема базы данных (Dashboard_data.json)
Структура данных представляет собой кэшированный слепок для мгновенной отрисовки UI без повторного запроса к COM-интерфейсу или LLM.

### JSON Schema
```json
{
  "type": "object",
  "properties": {
    "emails": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "id": { "type": "string", "description": "Локальный ID для фронтенда" },
          "unid": { "type": "string", "description": "Lotus Document Universal ID" },
          "lotus_url": { "type": "string", "description": "Прямая ссылка notes:///" },
          "date": { "type": "string", "description": "Дата и время доставки (YYYY-MM-DD HH:mm)" },
          "senderName": { "type": "string", "description": "Очищенное имя отправителя" },
          "subject": { "type": "string", "description": "Тема письма" },
          "isRead": { "type": "boolean", "description": "Флаг прочтения" },
          "isReplied": { "type": "boolean", "description": "Флаг наличия ответа" },
          "needsReply": { "type": "boolean", "description": "AI-флаг: требуется ли ответ" },
          "priority": { "type": "integer", "description": "AI-приоритет (1=P1/Urgent, 2=P2/Dir, 3=P3/Info)" },
          "body": { "type": "string", "description": "Оригинальный текст (обрезанный)" },
          "threadMessages": {
            "type": "array",
            "items": {
              "type": "object",
              "properties": {
                "text": { "type": "string", "description": "Здесь хранится AI-резюме (summary)" }
              }
            }
          }
        },
        "required": ["id", "unid", "senderName", "priority", "isRead"]
      }
    }
  }
}
```

## 3. Маппинг Промпта (Prompt_template.txt)
Скрипт `Notes_gemini.py` производит следующие подстановки (string replace) перед отправкой в LLM:
* `[[E_TOP_LIST]]` -> Значение `[TAGS] e_top`
* `[[E_DIR_LIST]]` -> Значение `[TAGS] e_dir`
* `[[EMAIL_COUNT]]` -> Количество элементов в массиве `emails`
* `[[EMAILS_PAYLOAD]]` -> Сериализованный мини-JSON с полями `id`, `sender`, `subject`, `body`.
