# Архитектура данных (data.md)

## 1. Файл конфигурации (set.ini)
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

[NETWORK]
ssl_verify = False

[NOTES]
default_days = 2 # Глубина поиска писем по умолчанию
max_emails = 500 # Жесткий лимит выборки
max_body_chars = 1500 # Лимит символов тела письма для отправки в LLM (экономия токенов)
batch_size = 15 # Сколько писем уходит в один запрос к модели

[UI]
mark_read_on_done = True # При отработке помечать письмо прочитанным в Notes
skip_weekends = True # Суббота и воскресенье не расходуют «дни» выборки
skip_ai_for_system = True # Автоматику не гонять через AI
subject_block = принято, отклонено # Фильтр тем (служебные ответы)
auto_done_days = 1 # «Отработать все»: закрывать письма старше этого срока (в днях)
```

### Lifecycle полей e_top_emails / e_dir_emails
1. При вызове `GET /api/settings` значения считываются из `set.ini` и возвращаются на фронтенд.
2. При выполнении полного сбора (`POST /api/analyze`): функция `resolve_names_to_emails()` сопоставляет список ФИО из `e_top` / `e_dir` с email-адресами, найденными в прочитанных письмах (по полям `senderName` / `senderEmail`). Найденные почты добавляются в `e_top_emails` / `e_dir_emails` и автоматически сохраняются обратно в `set.ini` (секция `[TAGS]`).
3. При вызове `POST /api/save-settings` (сохранение настроек из модального окна): пересчёт почт выполняется по кэшу `dashboard_data.json` (если он существует), после чего обновлённые списки также записываются в `set.ini`.
4. Ненайденные ФИО не блокируют работу — они будут повторно проверены при следующем сборе.

## 2. Схема базы данных (dashboard_data.json)
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
          "dateIso": { "type": "string", "description": "Дата доставки в ISO (YYYY-MM-DDTHH:MM:SS)" },
          "senderName": { "type": "string", "description": "Очищенное имя отправителя" },
          "senderEmail": { "type": "string", "description": "Email отправителя (извлекается из полей Notes: INetFrom, InternetFrom, InternetAddr, InternetAddress, Originator, From)" },
          "subject": { "type": "string", "description": "Тема письма" },
          "isRead": { "type": "boolean", "description": "Флаг прочтения" },
          "isReplied": { "type": "boolean", "description": "Флаг наличия ответа" },
          "needsReply": { "type": "boolean", "description": "AI-флаг: требуется ли ответ" },
          "priority": { "type": "integer", "description": "AI-приоритет (1=P1/Urgent, 2=P2/Dir, 3=P3/Info)" },
          "body": { "type": "string", "description": "Оригинальный текст (обрезанный)" },
          "suggestedReply": { "type": "string", "description": "AI-сгенерированный проект ответа (из suggested_reply в ответе LLM)" },
          "workState": { "type": "string", "description": "Отработка: new / wip / done (из work_state.json)" },
          "stateAt": { "type": "string", "description": "Момент последней отметки отработки (ISO)" },
          "threadKey": { "type": "string", "description": "Нормализованный ключ темы — цепочка писем" },
          "threadSize": { "type": "integer", "description": "Сколько писем в цепочке" },
          "threadRoot": { "type": "boolean", "description": "Первый элемент цепочки (голова); её пометка наследуется всей ветке" },
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

### 2.2. Цепочки писем (threadKey)
Письма одной темы объединяются в **цепочку** по нормализованному ключу ([`thread_key()`](../notes_gemini.py)): срезаются префиксы переписки (`RE:`/`FW:`/`HA:`/`>>:`/…) и нормализуются регистр, пробелы и кавычки. Пример: `HA: >>: [#2026-8607411]: Обращение в поддержку HRlink` → ключ `[#2026-8607411]: обращение в поддержку hrlink`.

- `threadKey` — ключ ветки; `threadSize` — размер цепочки; `threadRoot: true` — голова (первый элемент, самое свежее письмо).
- В UI цепочка **свёрнута по умолчанию**: показывается только голова; разворачивание показывает вложенные письма.
- **Наследование пометок:** пометка головы (отработка и прочтение) применяется ко **всем** письмам ветки; пометка вложенного — **только ему**. Запись выполняется пакетно (`/api/set-state`, `/api/mark-read` с массивом `unids`).

### 2.3. Отметки отработки и журнал показанных
- `work_state.json` — отдельная от Notes ось `new / wip / done` (ключ — UNID). Переживает пересбор: состояние восстанавливается через [`apply_work_state()`](../notes_gemini.py) при каждом чтении данных.
- `seen_mail.json` — журнал показанных писем для режима «новое с прошлого сбора» (когда Notes не отдаёт метки прочтения).

### 2.4. «Отработать все» (bulk done)
Кнопка в шапке дашборда ([`doneAllOld()`](../index.html)) помечает отработанными все письма **старше** `[UI] auto_done_days` (по умолчанию 1 день), **кроме писем «В работе»** (`wip`) — они остаются не отработанными.

## 3. Маппинг промптов

### 3.1. Разбор входящей почты (promt-get-post.txt)
Скрипт `notes_gemini.py` перед отправкой в LLM подставляет в шаблон:
* `[[E_TOP_LIST]]` -> Значение `[TAGS] e_top`
* `[[E_DIR_LIST]]` -> Значение `[TAGS] e_dir`
* `[[EMAIL_COUNT]]` -> Количество элементов в массиве `emails`
* `[[EMAILS_PAYLOAD]]` -> Сериализованный мини-JSON с полями `id`, `sender`, `email`, `subject`, `body`.

### 3.2. Чат по почте (promt-chat-post.txt)
System-промпт из `promt-chat-post.txt`; единственный плейсхолдер `[[EMAILS_PAYLOAD]]` заменяется на сжатую выжимку от [`compact_index()`](../notes_gemini.py): до 200 писем по важности, тела — у 60 важнейших (180 символов). Модель обязана вернуть `{"answer": "...", "refs": ["id", ...]}`.

### 3.3. Сводка (digest_template.txt)
Шаблон сводки подставляет `[[TODAY]]`, `[[PERIOD]]`, `[[EMAIL_COUNT]]`, `[[E_TOP_LIST]]`, `[[E_DIR_LIST]]`, `[[EMAILS_PAYLOAD]]`. Отработанные (`done`) письма из сводки **исключаются** — сводка строится только по активным.

## 4. Баланс DeepSeek
`GET /api/balance` → запрос `{base_url}/user/balance` с тем же API-ключом. Возвращает `is_available` и `balance_infos[]`:
- `currency` — валюта (обычно USD/CNY);
- `total_balance` — остаток всего;
- `granted_balance` — бесплатная часть;
- `topped_up_balance` — пополненная часть.
