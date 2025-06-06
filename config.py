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
SCAN_DIR = r"C:\Users\cu-nazarov-na\Desktop\zenmap\files\сканы"
OUTPUT_SUFFIX = '_filtered.xlsx'

# Путь к Nmap
NMAP_PATH = r"D:\nmap\nmap.exe"

# Профили сканирования с полными командами
SCAN_PROFILES = {
    "ping":                         [NMAP_PATH, "-sn"],
    "intense":                      [NMAP_PATH, "-T4", "-A", "-v"],
    "Intense scan plus UDP":        [NMAP_PATH, "-sS", "-sU", "-T4", "-A", "-v"],
    "Intense scan, all TCP ports":  [NMAP_PATH, "-p", "1-65535", "-T4", "-A", "-v"],
    "Intense scan, no ping":        [NMAP_PATH, "-T4", "-A", "-v", "-Pn"],
    "Quick scan":                   [NMAP_PATH, "-T4", "-F"],
    "Quick scan plus":              [NMAP_PATH, "-sV", "-T4", "-O", "-F", "--version-light"],
    "Quick traceroute":             [NMAP_PATH, "-sn", "--traceroute"],
    "Regular scan":                 [NMAP_PATH],
    "Slow comprehensive scan":      [NMAP_PATH, "-sS", "-sU", "-T4", "-A", "-v", "-PE", "-PP", "-PS80,443", "-PA3389", "-PU40125", "-PY", "-g", "53", "--script", "default or (discovery and safe)"]
}
