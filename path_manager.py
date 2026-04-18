import os
import configparser

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.ini")

config = configparser.ConfigParser()

if not os.path.exists(CONFIG_FILE):
    config["PATHS"] = {
        "jmsawb": "",
        "billcode": "",
        "dws1_8": "",
        "dws9_11": "",
        "auto_filter": "",
        "dws_filter": "",
        "last_dir": "",
        "output_auto": "",
        "output_dws": "",
        "output_filter": "",
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)
else:
    config.read(CONFIG_FILE, encoding="utf-8")


def load_path(key):
    return config.get("PATHS", key, fallback="")


def save_path(key, path):
    config.set("PATHS", key, path)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)


def get_initial_dir(*keys):
    for key in keys:
        path = load_path(key)
        if path and os.path.exists(path):
            if os.path.isfile(path):
                return os.path.dirname(path)
            return path

    return os.path.expanduser("~/Desktop")
