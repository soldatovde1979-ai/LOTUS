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
e_top_emails = ... # Извлечённые email-адреса высшего руководства (автозаполняется после сбора писем)
e_dir_emails = ... # Извлечённые email-адреса директоров (автозаполняется после сбора писем)

[NOTES]
default_days = 2 # Глубина поиска писем по умолчанию
max_emails = 50 # Жесткий лимит выборки
max_body_chars = 1500 # Лимит символов тела письма для отправки в LLM (экономия токенов)
```

### Lifecycle полей e_top_emails / e_dir_emails
1. При вызове `GET /api/settings` значения считываются из `Set.ini` и возвращаются на фронтенд.
2. При выполнении полного сбора (`POST /api/fetch`): функция `resolve_names_to_emails()` сопоставляет список ФИО из `e_top` / `e_dir` с email-адресами, найденными в прочитанных письмах (по полям `senderName` / `senderEmail`). Найденные почты добавляются в `e_top_emails` / `e_dir_emails` и автоматически сохраняются обратно в `Set.ini` (секция `[TAGS]`).
3. При вызове `POST /api/save-settings` (сохранение настроек из модального окна): пересчёт почт выполняется по кэшу `dashboard_data.json` (если он существует), после чего обновлённые списки также записываются в `Set.ini`.
4. Ненайденные ФИО не блокируют работу — они будут повторно проверены при следующем сборе.

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
          "senderEmail": { "type": "string", "description": "Email отправителя (извлекается из полей Notes: INetFrom, InternetFrom, InternetAddr, InternetAddress, Originator, From)" },
          "subject": { "type": "string", "description": "Тема письма" },
          "isRead": { "type": "boolean", "description": "Флаг прочтения" },
          "isReplied": { "type": "boolean", "description": "Флаг наличия ответа" },
          "needsReply": { "type": "boolean", "description": "AI-флаг: требуется ли ответ" },
          "priority": { "type": "integer", "description": "AI-приоритет (1=P1/Urgent, 2=P2/Dir, 3=P3/Info)" },
          "body": { "type": "string", "description": "Оригинальный текст (обрезанный)" },
          "suggestedReply": { "type": "string", "description": "AI-сгенерированный проект ответа (из suggested_reply в ответе LLM)" },
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

### 2.1. Объединение писем с одинаковой темой
Письма с одинаковой темой **от одного отправителя** (обычно автописьма и напоминания) свёртываются в одно при сборе ([`merge_duplicate_subjects()`](../notes_gemini.py)). В дашборде остаётся **последнее** письмо группы, а в его тело (`body`) добавляется блок «когда приходили» остальных.

Дополнительные поля объединённого письма:
- `merged: true` — признак того, что запись агрегирует несколько писем;
- `mergedCount: N` — сколько писем с этой темой пришло всего;
- `mergedTimes: ["ДД.ММ.ГГГГ ЧЧ:ММ", ...]` — когда приходили все письма группы.

Темы короче 3 символов и пустые в объединение не берутся. Логика «новое / прочитано» считается до объединения и не ломается.

**Авто-отметка свёрнутых писем:** старые письма группы (все, кроме показанного последнего) при сборе автоматически помечаются в `work_state.json` как `done` («Отработано») — [`auto_mark_done()`](../notes_gemini.py). Активным остаётся только последнее напоминание; уже существующая метка `done` не перетирается.

## 3. Маппинг Промпта (Prompt_template.txt)
Скрипт `Notes_gemini.py` производит следующие подстановки (string replace) перед отправкой в LLM:
* `[[E_TOP_LIST]]` -> Значение `[TAGS] e_top`
* `[[E_DIR_LIST]]` -> Значение `[TAGS] e_dir`
* `[[EMAIL_COUNT]]` -> Количество элементов в массиве `emails`
* `[[EMAILS_PAYLOAD]]` -> Сериализованный мини-JSON с полями `id`, `sender`, `email`, `subject`, `body`.
