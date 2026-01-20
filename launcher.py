import os
import requests
import subprocess

VERSION_URL = "https://raw.githubusercontent.com/sauravku07/invoice-automation/main/version.txt"
EXE_URL = "https://raw.githubusercontent.com/sauravku07/invoice-automation/main/invoice_once.exe"

LOCAL_VERSION_FILE = "version.txt"
APP_EXE = "invoice_once.exe"

def get_local_version():
    if os.path.exists(LOCAL_VERSION_FILE):
        return open(LOCAL_VERSION_FILE).read().strip()
    return "0"

def get_remote_version():
    return requests.get(VERSION_URL, timeout=5).text.strip()

def update_app():
    exe = requests.get(EXE_URL, timeout=10).content
    with open(APP_EXE, "wb") as f:
        f.write(exe)

def main():
    try:
        local = get_local_version()
        remote = get_remote_version()

        if remote != local:
            update_app()
            open(LOCAL_VERSION_FILE, "w").write(remote)
    except:
        pass

    subprocess.Popen([APP_EXE], shell=True)

if __name__ == "__main__":
    main()
