Для MAX API может понадобиться корневой сертификат Минцифры.
Скачай его с https://gu-st.ru/content/lending/russian_trusted_root_ca_pem.crt
и положи сюда под именем russian_trusted_root_ca_pem.crt.
Затем установи MAX_CA_BUNDLE=/app/certs/russian_trusted_root_ca_pem.crt
в переменных окружения Bothost.
