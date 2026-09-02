# ExcelWorker

Десктопное приложение на Python (Tkinter) для работы с Excel: массовое сопоставление, поиск, хеши, diff, конвертер XLSX → PDF.

## Локальная сборка (Windows)

```bash
pip install -r requirements.txt
pyinstaller --onefile --noconsole --name "ExcelWorker" app.py
# или
pyinstaller ExcelWorker.spec
# exe -> dist/ExcelWorker.exe
```

## Git + релизы (Windows сборка в GitHub Actions)

В репозитории уже настроен workflow `.github/workflows/build.yml`:
- на каждый `push` в `main`/`master` и `pull_request` — собирает `ExcelWorker.exe` и загружает как артефакт (14 дней);
- при пуше тега `v*` (например `v1.0.0`) — собирает и **создаёт GitHub Release** с файлом `ExcelWorker.exe`.

### Первый пуш (создание репозитория)

1. Создайте **пустой** репозиторий на GitHub (без README, без .gitignore), скопируйте его URL, например `https://github.com/<user>/excel_worker.git`.

2. В папке проекта:

```bash
git init
git branch -M main
git add .
git commit -m "feat: initial commit with 5 tabs and Windows build workflow"
git remote add origin https://github.com/<user>/excel_worker.git
git push -u origin main
```

### Как выпустить релиз

```bash
# после изменений
git add .
git commit -m "feat: ..."
git push

# создать релиз (сборка запустится автоматически)
git tag v1.0.0
git push origin v1.0.0
# -> на вкладке Releases появится ExcelWorker.exe

# посмотреть сборку
# GitHub -> Actions -> Build Windows EXE and Release
```

### Авто-релиз на каждый пуш (опционально)

В `build.yml` есть закомментированная job `auto-release`. Раскомментируйте, чтобы каждый `push` в `main` сам создавал релиз `v0.0.<run_number>` без ручных тегов.

## Требования

- Python 3.10+
- `openpyxl`, `pyinstaller` (см. `requirements.txt`)
- Для точной конвертации XLSX→PDF — LibreOffice (иначе fallback `reportlab`)
