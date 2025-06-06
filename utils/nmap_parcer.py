import re

def parse_nmap_txt(file_path):
    hosts = []
    scan_info = {
        'start_time': None,
        'command': None,
        'source_host': None,
        'total_ips': 0,
        'hosts_up': 0
    }
    
    current_host = None
    host_pattern = re.compile(r'Nmap scan report for (?:([\w\-. ]+)\s)?\(?([\d.]+)\)?')
    ports_section = False
    
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            
            # Парсинг общей информации
            if not scan_info['start_time'] and line.startswith('Starting Nmap'):
                scan_info['start_time'] = line.split(' at ')[-1]
            
            elif 'Nmap done:' in line:
                match = re.search(r'Nmap done: (\d+) IP addresses \((\d+) hosts? up\)', line)
                if match:
                    scan_info['total_ips'] = int(match.group(1))
                    scan_info['hosts_up'] = int(match.group(2))
            
            # Парсинг хостов
            host_match = host_pattern.match(line)
            if host_match:
                if current_host:
                    hosts.append(current_host)
                
                hostname, ip = host_match.groups()
                current_host = {
                    'ip': ip,
                    'hostname': hostname.strip() if hostname else None,
                    'ports': {}
                }
                ports_section = False
                continue
            
            if current_host is None:
                continue
            
            # Начало секции с портами
            if line.startswith('PORT') and 'STATE' in line and 'SERVICE' in line:
                ports_section = True
                continue
            
            # Конец секции с портами
            if ports_section and (not line or line.startswith('Nmap scan')):
                ports_section = False
            
            # Парсинг портов
            if ports_section and line:
                # Пропускаем информационные строки
                if line.startswith('Not shown:') or line.startswith('All ') or 'filtered' in line:
                    continue
                
                # Исправленный парсинг строк портов
                parts = re.split(r'\s+', line, maxsplit=2)
                if len(parts) < 2:
                    continue
                
                port_service = parts[0].strip()
                state = parts[1].strip().lower()
                service = parts[2].strip().split()[0] if len(parts) > 2 else 'unknown'
                
                if '/' not in port_service:
                    continue
                    
                port, protocol = port_service.split('/', 1)
                
                # Сохраняем как отдельные компоненты
                port_key = (port, protocol, service)
                current_host['ports'][port_key] = state

    # Добавляем последний хост
    if current_host:
        hosts.append(current_host)
    
    return hosts, scan_info
