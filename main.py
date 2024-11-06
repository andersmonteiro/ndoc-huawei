import os
import re
import gspread
from google.oauth2.service_account import Credentials

# Configuração da API do Google Sheets
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

sheet_id = "1z6PI5mYlywGCsWDG5FHMPkG7rzfZMBx42tUI8jt-Hos"  # Substitua pelo ID da sua planilha
workbook = client.open_by_key(sheet_id)

# Função para ler o arquivo de configuração
def read_config_file(file_path):
    try:
        with open(file_path, 'r', encoding="utf-8") as file:
            return file.readlines()
    except Exception as e:
        print(f"Erro ao ler o arquivo de configuração: {e}")
        return None

# Função para extração de interfaces, descrições, VLANs, IPs e tipo de porta
def extract_interfaces(config_lines):
    interfaces = []
    current_interface = None
    current_description = None
    current_vlan = None
    current_ipv4 = None
    current_ipv6 = None
    current_port_type = None
    allowed_vlans = []
    trunk_lines = []  # Acumula linhas de 'port trunk allow-pass vlan'

    for line in config_lines:
        interface_match = re.match(r"^interface\s+(\S+)", line.strip())
        trunk_match = re.match(r"port trunk allow-pass vlan\s+(.+)", line.strip())  # Detecta linhas de VLANs permitidas

        if interface_match:
            if current_interface:
                # Processa as VLANs acumuladas antes de iniciar uma nova interface
                if trunk_lines:
                    allowed_vlans = extract_allowed_vlans(' '.join(trunk_lines))
                    trunk_lines = []  # Reinicia a lista após processar

                interfaces.append((current_interface, current_description, current_vlan, current_ipv4, current_ipv6, current_port_type, ', '.join(map(str, sorted(set(allowed_vlans))))))

            current_interface = interface_match.group(1)
            current_description = None
            current_vlan = None
            current_ipv4 = None
            current_ipv6 = None
            current_port_type = None
            allowed_vlans = []

            vlan_if_match = re.match(r"Vlanif(\d+)", current_interface)
            if vlan_if_match:
                current_vlan = vlan_if_match.group(1)

        elif current_interface:
            description_match = re.match(r"^description\s+(.+)", line.strip())
            if description_match:
                current_description = description_match.group(1)

            vlan_match = re.search(r"vlan-type dot1q\s+(\d+)", line.strip())
            if vlan_match:
                current_vlan = vlan_match.group(1)

            ipv4_match = re.search(r"ip address (\d+\.\d+\.\d+\.\d+ \d+\.\d+\.\d+\.\d+)", line.strip())
            if ipv4_match:
                current_ipv4 = ipv4_match.group(1)

            ipv6_match = re.search(r"ipv6 address ([a-fA-F0-9:]+/\d+)", line.strip())
            if ipv6_match:
                current_ipv6 = ipv6_match.group(1)

            port_type_match = re.search(r"port link-type\s+(\S+)", line.strip())
            if port_type_match:
                current_port_type = port_type_match.group(1)

            # Acumula as linhas de 'port trunk allow-pass vlan'
            if trunk_match:
                trunk_lines.append(trunk_match.group(1))
            else:
                # Processa as VLANs acumuladas se a linha atual não for de VLANs permitidas
                if trunk_lines:
                    allowed_vlans = extract_allowed_vlans(' '.join(trunk_lines))
                    trunk_lines = []  # Reinicia a lista após processar

    if current_interface:
        # Processa as VLANs acumuladas da última interface
        if trunk_lines:
            allowed_vlans = extract_allowed_vlans(' '.join(trunk_lines))

        interfaces.append((current_interface, current_description, current_vlan, current_ipv4, current_ipv6, current_port_type, ', '.join(map(str, sorted(set(allowed_vlans))))))

    return interfaces

# Função auxiliar para extrair as VLANs permitidas de uma string concatenada
def extract_allowed_vlans(trunk_line):
    allowed_vlans = []
    # Remove caracteres inválidos e espaços extras
    trunk_line = re.sub(r'[^0-9to,\s]', '', trunk_line).strip()
    # Divide a string em itens, considerando os ranges de VLANs
    items = re.split(r'(?:,\s*|\s+)', trunk_line)  # Expressão regular modificada
    for i in range(len(items)):
        # Ignora itens vazios
        if not items[i]:
            continue

        if items[i] == 'to':
            # Concatena os números ao redor de "to" com a palavra "to"
            allowed_vlans.append(f"{items[i-1]} to {items[i+1]}")
        elif items[i].isdigit() and (i == 0 or items[i-1] != 'to'):
            allowed_vlans.append(items[i])
    return allowed_vlans

# Preparação dos dados para o Google Sheets
def prepare_data_for_sheet(interfaces):
    # Cabeçalho das colunas
    values = [["Interface", "Descrição", "VLAN", "IP", "IPv6", "Tipo de Porta", "VLANs Permitidas"]]
    for interface, description, vlan, ipv4, ipv6, port_type, allowed_vlans in interfaces:
        values.append([interface, description or "", vlan or "", ipv4 or "", ipv6 or "", port_type or "", allowed_vlans])
    return values

# Caminho relativo do arquivo de configuração
file_path = os.path.join(os.path.dirname(__file__), "huawei-cfg.txt")
config_lines = read_config_file(file_path)

if config_lines is not None:
    print("Arquivo lido com sucesso.")

    # Extração dos dados
    interfaces = extract_interfaces(config_lines)

    # Preparando os dados para o Google Sheets
    values = prepare_data_for_sheet(interfaces)

    # Atualizando o Google Sheets
    worksheet_list = map(lambda x: x.title, workbook.worksheets())
    new_worksheet_name = "ConfigData"

    if new_worksheet_name in worksheet_list:
        sheet = workbook.worksheet(new_worksheet_name)
    else:
        sheet = workbook.add_worksheet(new_worksheet_name, rows=1000, cols=10)

    sheet.clear()
    sheet.update(range_name=f"A1:G{len(values)}", values=values)
    sheet.format("A1:G1", {"textFormat": {"bold": True}})
else:
    print("Erro ao ler o arquivo de configuração. O script não continuará.")