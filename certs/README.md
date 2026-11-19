# Сертификат MAX API

В этой папке находится публичный корневой сертификат Russian Trusted Root CA, который используется при необходимости для TLS-подключения к MAX API.

В Bothost укажи `MAX_CA_BUNDLE=/app/certs/russian_trusted_root_ca_pem.crt`. Не отключай проверку TLS и не используй `ssl=False`.
