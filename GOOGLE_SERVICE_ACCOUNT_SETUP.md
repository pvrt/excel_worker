# Google Service Account — подробная настройка для раздачи

Этот вариант **рекомендуется для раздачи** `ExcelWorker.exe` другим людям: не требует браузера, не истекает через 7 дней (как Test OAuth), работает без `token.json`.

---

## 1. Создайте проект Google Cloud

1. Откройте https://console.cloud.google.com
2. Вверху выберите проект → **New Project** → имя `excel-worker` → **Create**
3. Дождитесь создания и переключитесь на него.

## 2. Включите Google Drive API

1. **APIs & Services → Library**
2. Найдите **Google Drive API** → **Enable**
3. (Опционально) **Google Sheets API** → **Enable** (для конвертации xlsx → Google Sheets)

## 3. Создайте Service Account

1. **IAM & Admin → Service Accounts** → **Create Service Account**
2. Имя: `excel-worker-converter` → **Create and Continue**
3. Роль можно пропустить (или `Editor` не нужен) → **Continue** → **Done**
4. Откройте созданный аккаунт (клик по email вида `excel-worker-converter@hdv-monster-....iam.gserviceaccount.com`)

## 4. Создайте ключ JSON

1. Внутри Service Account → вкладка **Keys** → **Add Key → Create new key**
2. Тип **JSON** → **Create** → скачается файл вида `hdv-monster-...-a1b2c3.json`
3. **Переименуйте** его в `credentials.json` и положите **рядом с `app.py`** (при разработке) или **рядом с `ExcelWorker.exe`** (при раздаче). Путь можно выбрать в последней вкладке → `credentials.json:` → Обзор.

> **Важно:** `credentials.json` типа `service_account` содержит приватный ключ. В `.gitignore` — не коммитьте его и `google_sa_embedded.py`! Ключ **вшивается** в exe на этапе сборки из секрета `GOOGLE_SERVICE_ACCOUNT_JSON`, рядом с exe ничего не нужно.

## 5. (Опционально) Расшарьте папку — не обязательно

Для конвертации через наш код **расшаривать не нужно**: файл создаётся во временном `My Drive` сервис-аккаунта (`drive.files.create` → `export` → `delete`). Но если хотите, чтобы сервис-аккаунт видел ваши файлы:
- Создайте папку на своём Drive → **Share** → добавьте email сервис-аккаунта → **Editor**.

## 6. Проверьте подключение

1. В `ExcelWorker` → вкладка **Конвертер XLSX → PDF** → выберите движок **Только Google Sheets** — ключ уже вшит, настроек нет, просто конвертируйте.
2. Никаких `credentials.json`/`token.json`, кнопок проверки и браузера в программе нет — движок либо работает, либо показывает ошибку в логе.

## 7. Конвертация

- Выберите папку с `xlsx`, выходную папку, схему именования, **Движок: Google Sheets** → **Конвертировать в PDF**
- Параметры экспорта: `A4, альбом (portrait=false), fitw=true, gridlines=true, fzr=true` (`app.py:236`).

## 8. Сборка для раздачи (ключ вшивается в exe, рядом ничего не нужно)

1. GitHub → Settings → Secrets → Actions → секрет `GOOGLE_SERVICE_ACCOUNT_JSON` = содержимое JSON ключа целиком.
2. `git push` → workflow сам делает `python tools/gen_embedded_sa.py` (генерит `google_sa_embedded.py`, в `.gitignore`) и собирает `pyinstaller --onefile --noconsole --name ExcelWorker --hidden-import google_sa_embedded app.py`.
3. Раздайте один `dist/ExcelWorker.exe` — внутри уже ключ (`_get_embedded_sa_info` → `SACredentials.from_service_account_info`, без браузера и файлов).

Локально: `python tools/gen_embedded_sa.py` (читает `credentials.json`, НЕ коммитить) → `pyinstaller ExcelWorker.spec`.

При `PyInstaller --onefile` путь к `credentials.json` определяется как `Path(sys.executable).parent / "credentials.json"` (`app.py:15` `_get_app_dir()`), поэтому рядом с exe.

## 9. Частые проблемы

- **403 accessNotConfigured / Drive API not enabled** → включите Drive API в том же проекте, где создан Service Account.
- **401 Unauthorized** → проверьте, что скачан именно Service Account JSON (поле `"type": "service_account"`), а не OAuth `installed`.
- **SSL CERTIFICATE_VERIFY_FAILED** (корпоративный прокси) → код уже делает `ssl._create_unverified_context` и `verify=False` fallback (`app.py:203`/`app.py:257`), просто повторите.
- **Quota exceeded** → Service Account имеет свой Drive лимит 15 ГБ, файлы удаляются сразу после конвертации.

## 10. Безопасность

- Не публикуйте `credentials.json` на GitHub — он уже в `.gitignore:49`.
- Если скомпрометирован — **IAM → Service Accounts → Keys → Delete** старый ключ → создайте новый.
- Для личного использования можно оставить OAuth (`client_id` Desktop) — тогда нужен `token.json` и раз в 7 дней (Testing) перелогин, пока не переведёте в `Production`.
