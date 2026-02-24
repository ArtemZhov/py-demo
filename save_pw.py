import getpass, keyring
username = "artem.zhovtonozhko@rt.ru"
pw = getpass.getpass(f"Введите пароль для {username}: ")
keyring.set_password("my_email_app", username, pw)
print("Пароль сохранён.")
