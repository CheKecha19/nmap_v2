import subprocess
import socket
import os
import config

def get_local_ip():
    """Получаем локальный IP-адрес"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "N/A"

def run_nmap_scan(target, profile, output_file):
    """Запускает Nmap сканирование с выбранным профилем"""
    # Получаем команду для профиля
    if profile not in config.SCAN_PROFILES:
        raise ValueError(f"Неизвестный профиль: {profile}")
    
    command = config.SCAN_PROFILES[profile].copy()
    command.append("-oN")
    command.append(output_file)
    
    # Добавляем цель
    command.append(target)
    
    print(f"Выполняем команду: {' '.join(command)}")
    
    try:
        # Запускаем процесс
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        # Ждем завершения
        stdout, stderr = process.communicate()
        
        if process.returncode != 0:
            print(f"Ошибка при выполнении сканирования (код {process.returncode}):")
            print(stderr)
            return False
        
        return True
    except Exception as e:
        print(f"Неожиданная ошибка: {e}")
        return False
