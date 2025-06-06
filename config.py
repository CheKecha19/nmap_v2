# config.py
# Цвета для состояний портов в формате HEX
COLORS = {
    'open': '90EE90',       # Зеленый
    'closed': 'FFCCCB',     # Красный
    'filtered': 'FFFFE0',   # Желтый
    'undefined': 'D3D3D3',  # Светло-серый
    'default': 'FFFFFF'     # Белый
}

# Настройки по умолчанию
DEFAULT_OUTPUT_DIR = r"C:\Users\cu-nazarov-na\Desktop\zenmap\files\таблицы"
OUTPUT_SUFFIX = '_filtered.xlsx'
SCAN_DIR = r"C:\Users\cu-nazarov-na\Desktop\zenmap\files\сканы"
NMAP_PATH = r"D:\nmap\nmap.exe"                                                 # Путь к Nmap (измените при необходимости)

# Профили сканирования
SCAN_PROFILES = {
    "ping":                         "Быстрое сканирование активных хостов",
    "intense":                      "Интенсивное сканирование с определением ОС и служб",
    "Intense scan plus UDP":        "Интенсивное сканирование + UDP порты",
    "Intense scan, all TCP ports":  "",
    "Intense scan, no ping":        "",
    "Quick scan":                   "",
    "Quick scan plus":              "",
    "Quick traceroute":             "",
    "Regular scan":                 "",
    "Slow comprehensive scan":      ""
}
