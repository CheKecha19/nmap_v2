import os
import sys
import argparse
from utils import nmap_parser, excel_report, nmap_runner
import config

def main():
    parser = argparse.ArgumentParser(description='Обработчик результатов сканирования Nmap')
    parser.add_argument('input_file', nargs='?', help='Путь к файлу с результатами сканирования Nmap')
    parser.add_argument('--output-dir', help='Папка для сохранения отчета')
    parser.add_argument('--target', help='Цель для сканирования (например, 192.168.1.0/24)')
    parser.add_argument('--profile', choices=list(config.SCAN_PROFILES.keys()), 
                        help=f'Профиль сканирования: {", ".join(config.SCAN_PROFILES.keys())}')
    
    args = parser.parse_args()
    
    # Обработка аргументов
    input_file = None
    scan_info = {}
    
    # Если указаны target и profile, запускаем сканирование
    if args.target and args.profile:
        # Создаем имя файла для сканирования
        os.makedirs(config.SCAN_DIR, exist_ok=True)
        scan_file = os.path.join(config.SCAN_DIR, f"nmap_scan_{args.profile}.txt")
        
        if nmap_runner.run_nmap_scan(args.target, args.profile, scan_file):
            print(f"Сканирование завершено. Результаты сохранены в: {scan_file}")
            input_file = scan_file
            scan_info = {
                'command': f"nmap {args.profile} {args.target}",
                'source_host': nmap_runner.get_local_ip()
            }
        else:
            print("Ошибка при выполнении сканирования")
            sys.exit(3)
    
    # Если указан input_file
    if args.input_file:
        input_file = args.input_file
    
    # Проверка наличия входного файла
    if not input_file:
        print("Ошибка: необходимо указать либо файл с результатами, либо цель и профиль сканирования")
        parser.print_help()
        sys.exit(1)
    
    # Определяем папку для вывода
    output_dir = args.output_dir or config.DEFAULT_OUTPUT_DIR
    
    # Проверка существования файла
    if not os.path.exists(input_file):
        print(f"Ошибка: файл '{input_file}' не найден")
        sys.exit(1)
    
    # Определение пути для выходного файла
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_file_name = f"{base_name}{config.OUTPUT_SUFFIX}"
    
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, output_file_name)
    else:
        output_path = os.path.join(os.path.dirname(__file__), output_file_name)
    
    try:
        # Если scan_info не был создан при сканировании, парсим файл
        if not scan_info:
            hosts, scan_info = nmap_parser.parse_nmap_txt(input_file)
        else:
            hosts, scan_info_from_file = nmap_parser.parse_nmap_txt(input_file)
            # Объединяем информацию из сканирования и из файла
            scan_info.update({
                'start_time': scan_info_from_file.get('start_time'),
                'total_ips': scan_info_from_file.get('total_ips'),
                'hosts_up': scan_info_from_file.get('hosts_up')
            })
        
        if not hosts:
            print("Предупреждение: не найдено данных для обработки")
            wb = Workbook()
            wb.save(output_path)
            print(f"Создан пустой отчет: {output_path}")
        else:
            excel_report.create_excel_report(hosts, scan_info, output_path)
            
    except Exception as e:
        print(f"Критическая ошибка при обработке файла: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(2)

if __name__ == "__main__":
    main()
