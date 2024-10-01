import requests
import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox

# função para obter o token de autenticação
def get_access_token():
    url = "link oAuth da API"
    headers = {"Content-Type": "application/json"}
    data = {
        "client_id": "",
        "client_secret": "",
        "username": "",
        "password": "",
        "grant_type": "password"
    }
    
    response = requests.post(url, json=data, headers=headers)
    response_data = response.json()
    
    if response.status_code == 200:
        return response_data['access_token']
    else:
        print("Erro ao obter o token:", response_data)
        return None

# formatando datas
def format_date(date_str):
    if date_str:
        return datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y")
    return None

# obter ordens de serviço conforme datas
def get_service_orders(token, start_date, end_date):
    url = "endpoint da API"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    params = {
        "pagina": 0,
        "itens_por_pagina": 500,
        "data_inicio": start_date,
        "data_fim": end_date,
        "status": "finalizado",
        "relacoes": "tecnicos,ordem_servico_mensagem"
    }
    
    try:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        data = response.json()
    except requests.exceptions.RequestException as e:
        print(f"Erro ao consultar dados: {e}")
        return [], 0
    except requests.exceptions.JSONDecodeError:
        print("Erro ao decodificar a resposta da API.")
        return [], 0
    
    if 'status' in data and data['status'] == "success":
        return data.get('ordens_servico', []), data.get('paginacao', {}).get('total_registros', 0)
    else:
        print("Erro ao consultar dados:", data.get('msg', 'Erro desconhecido'))
        return [], 0

# processar e filtrar os dados obtidos
def process_data(orders):
    filtered_data = []
    excluded_types = {"COLETA DE EQUIPAMENTOS", "COLETA PÓS-CANCELAMENTO", "MUDANÇA", "COLETA POR INADIMPLÊNCIA"}
    
    for order in orders:
        tipo_os = order.get("tipo", "Não informado")
        
        if tipo_os in excluded_types:
            continue

        cliente_info = order.get("cliente", "").split(')', 1)
        cliente_codigo = cliente_info[0].strip('(') if len(cliente_info) > 1 else ""
        cliente_nome = cliente_info[1].strip() if len(cliente_info) > 1 else cliente_info[0].strip()

        try:
            cliente_codigo = int(cliente_codigo)
        except ValueError:
            pass

        # verificar se há técnicos
        if order.get("tecnicos") and len(order["tecnicos"]) > 0:
            tecnico_nome = order["tecnicos"][0].get("name", "Não informado")
        else:
            tecnico_nome = "Não informado"

        filtered_data.append({
            "TIPO DA O.S": tipo_os,
            "CÓDIGO DO CLIENTE": cliente_codigo,
            "NOME DO CLIENTE": cliente_nome,
            "DATA DE ABERTURA": format_date(order.get("data_cadastro")),
            "DATA DE TERMINO EXECUTADO": format_date(order.get("data_termino_executado")),
            "DESCRIÇÃO DE ABERTURA": order.get("descricao_abertura"),
            "DESCRIÇÃO DE FECHAMENTO": order.get("descricao_fechamento"),
            "TÉCNICO": tecnico_nome,
        })
    
    return filtered_data

# filtar clientes que aparecem mais de uma vez
def filter_retrabalho(data):
    df = pd.DataFrame(data)
    client_counts = df['CÓDIGO DO CLIENTE'].value_counts()
    repeated_clients = client_counts[client_counts >= 2].index
    filtered_df = df[df['CÓDIGO DO CLIENTE'].isin(repeated_clients)]
    return filtered_df

# exportar os dados para o excel
def export_to_excel(data, save_path):
    df = filter_retrabalho(data)
    df = df.sort_values(by=["CÓDIGO DO CLIENTE", "NOME DO CLIENTE"], ascending=[False, True])
    df.to_excel(save_path, index=False)

# consultar em períodos divididos, contornando o limite de 500 requisições da API
def query_in_chunks(token, start_date, end_date):
    all_data = []
    current_start = datetime.strptime(start_date, "%Y-%m-%d")
    current_end = current_start + timedelta(days=30)  # período de 30 em 30 dias

    end_date_obj = datetime.strptime(end_date, "%Y-%m-%d")

    while current_start < end_date_obj:
        # garantindo que a última consulta não exceda a data final
        if current_end > end_date_obj:
            current_end = end_date_obj

        orders, _ = get_service_orders(token, current_start.strftime("%Y-%m-%d"), current_end.strftime("%Y-%m-%d"))
        if orders:
            all_data.extend(process_data(orders))

        # avançar para o próximo intervalo de tempo
        current_start = current_end + timedelta(days=1)
        current_end = current_start + timedelta(days=30)

    return all_data

# função para solicitar as datas e consultar os dados
def ask_for_dates():
    def submit_dates():
        start_date = start_date_entry.get()
        end_date = end_date_entry.get()

        try:
            start_date_converted = datetime.strptime(start_date, "%d/%m/%Y").strftime("%Y-%m-%d")
            end_date_converted = datetime.strptime(end_date, "%d/%m/%Y").strftime("%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/YYYY.")
            return
        
        token = get_access_token()
        if not token:
            messagebox.showerror("Erro", "Não foi possível obter o token.")
            return

        # dividindo a consulta em períodos menores
        all_data = query_in_chunks(token, start_date_converted, end_date_converted)
        
        if not all_data:
            messagebox.showinfo("Informação", "Nenhuma ordem de serviço encontrada.")
            return
        
        # exportando para o excel
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx")],
            initialfile=f"retrabalhos_{datetime.now().strftime('%d_%m_%Y')}.xlsx"
        )
        if save_path:
            export_to_excel(all_data, save_path)
            messagebox.showinfo("Sucesso", "Dados exportados com sucesso.")
        
        if not messagebox.askyesno("Continuar", "Deseja realizar uma nova consulta?"):
            root.quit()

    global root
    root = tk.Tk()
    root.title("Consulta de Ordens de Serviço")

    tk.Label(root, text="Data de Início (DD/MM/YYYY):").pack(pady=5)
    start_date_entry = tk.Entry(root)
    start_date_entry.pack(pady=5)

    tk.Label(root, text="Data de Fim (DD/MM/YYYY):").pack(pady=5)
    end_date_entry = tk.Entry(root)
    end_date_entry.pack(pady=5)

    tk.Button(root, text="Consultar e Exportar", command=lambda: [root.withdraw(), submit_dates()]).pack(pady=20)

    root.mainloop()

def main():
    ask_for_dates()

if __name__ == "__main__":
    main()
