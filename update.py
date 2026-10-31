# -*- coding: utf-8 -*-
"""
Автообновление бота из GitHub. Запускается из start.bat перед ботом.
Сверяет version.txt; если на GitHub версия новее — скачивает файлы из manifest.txt.
config.py, settings.json, кэш и журнал никогда не трогаются.
Если интернета/GitHub нет — просто запускается текущая версия.
"""
import os
import sys
import urllib.request

REPO = "alexbarny22/blagomed_tools"
BRANCH = "main"
FOLDER = "move_bot"          # папка проекта внутри репозитория
HERE = os.path.dirname(os.path.abspath(__file__))
PROTECTED = {"config.py", "settings.json", "cache_1c.json", "move_bot.log", "start.bat"}


def token():
    try:
        sys.path.insert(0, HERE)
        import config
        return getattr(config, "GITHUB_TOKEN", "") or ""
    except Exception:
        return ""


def fetch(name, tok):
    if tok:   # приватный репозиторий — через API с ключом
        url = f"https://api.github.com/repos/{REPO}/contents/{FOLDER}/{name}?ref={BRANCH}"
        req = urllib.request.Request(url, headers={
            "Authorization": "Bearer " + tok,
            "Accept": "application/vnd.github.raw",
            "User-Agent": "move_bot-updater"})
    else:     # публичный репозиторий
        url = f"https://raw.githubusercontent.com/{REPO}/{BRANCH}/{FOLDER}/{name}"
        req = urllib.request.Request(url, headers={"User-Agent": "move_bot-updater",
                                                   "Cache-Control": "no-cache"})
    with urllib.request.urlopen(req, timeout=20) as r:
        return r.read()


def local_version():
    try:
        with open(os.path.join(HERE, "version.txt"), encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return ""


def main():
    tok = token()
    try:
        remote = fetch("version.txt", tok).decode("utf-8").strip()
    except Exception as e:
        print(f"Проверка обновлений пропущена (нет связи с GitHub): {e}")
        return
    if remote == local_version():
        print(f"Версия {remote} — актуальная.")
        return
    print(f"Найдено обновление: {local_version() or '—'} -> {remote}")
    try:
        files = [x.strip() for x in fetch("manifest.txt", tok).decode("utf-8").splitlines()
                 if x.strip() and not x.startswith("#")]
        data = {}
        for name in files:
            if name in PROTECTED or "/" in name or "\\" in name or ".." in name:
                continue
            data[name] = fetch(name, tok)
    except Exception as e:
        print(f"Обновление не скачалось, запускаю текущую версию: {e}")
        return
    # записываем только когда всё скачалось — чтобы не получить половину новой версии
    for name, body in data.items():
        tmp = os.path.join(HERE, name + ".new")
        with open(tmp, "wb") as f:
            f.write(body)
        os.replace(tmp, os.path.join(HERE, name))
        print(f"  обновлён {name}")
    with open(os.path.join(HERE, "version.txt"), "w", encoding="utf-8") as f:
        f.write(remote)
    print("Обновление установлено.")


if __name__ == "__main__":
    main()
