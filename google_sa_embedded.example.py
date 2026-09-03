# Пример вшитого ключа. Не коммитьте настоящий ключ!
# Локально: python tools/gen_embedded_sa.py  (прочитает credentials.json)
# В CI: секрет GOOGLE_SERVICE_ACCOUNT_JSON -> tools/gen_embedded_sa.py сгенерит этот файл.
GOOGLE_SA_JSON = None
