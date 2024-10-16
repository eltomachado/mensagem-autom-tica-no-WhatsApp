import pandas as pd
import pywhatkit as kit
from datetime import datetime, timedelta
import os
import tkinter as tk
from tkinter import messagebox

# Função para enviar mensagem no WhatsApp
def enviar_mensagem_whatsapp(numero_cliente, mensagem):
    try:
        kit.sendwhatmsg_instantly(f"+{numero_cliente}", mensagem)
        print(f"Mensagem enviada para {numero_cliente}: {mensagem}")
    except Exception as e:
        print(f"Erro ao enviar mensagem para {numero_cliente}: {e}")

# Função para gerar a mensagem personalizada
def gerar_mensagem(nome, plano, data_vencimento_formatada, chave_pix, nome_empresa):
    return (f"Olá, {nome}, tudo bem?\n\n"
            f"Lembrando que o seu plano de *R$ {plano},00* vence em *{data_vencimento_formatada}.*\n\n"
            f"Para evitar o *bloqueio automático*, efetue o pagamento via Pix: CPF: {chave_pix}. "
            f"Por favor, envie o comprovante para atualização no sistema.\n\n"
            f"Qualquer dúvida, estamos à disposição!\n\n"
            f"Abraços,\n{nome_empresa}")

# Função para carregar e processar a planilha
def processar_planilha(caminho_planilha, chave_pix, nome_empresa):
    # Ler a planilha de clientes
    planilha = pd.read_excel(caminho_planilha)

    # Definir quantos dias antes do vencimento deseja notificar
    dias_para_notificar = 2
    hoje = datetime.now().date()

    # Loop para verificar cada cliente
    for index, cliente in planilha.iterrows():
        try:
            # Converter a data corretamente para o formato datetime
            data_vencimento = pd.to_datetime(cliente['Data de Vencimento'], format='%d/%m/%Y').date()
            data_vencimento_formatada = data_vencimento.strftime('%d/%m/%Y')  # Formata a data
        except Exception as e:
            print(f"Erro ao converter a data para o cliente {cliente['Nome']}: {e}")
            continue

        # Verifica se a data de vencimento está próxima
        if hoje >= data_vencimento - timedelta(days=dias_para_notificar):
            nome = cliente['Nome']
            telefone = cliente['Telefone']
            plano = cliente['Plano']  # Assume que o valor do plano está na coluna 'Plano'

            # Mensagem personalizada
            mensagem = gerar_mensagem(nome, plano, data_vencimento_formatada, chave_pix, nome_empresa)

            # Enviar a mensagem
            enviar_mensagem_whatsapp(telefone, mensagem)
            print(f"Mensagem enviada para {nome} no número {telefone}")
        else:
            print(f"Cliente {cliente['Nome']} não está dentro do prazo de notificação.")

# Função para iniciar o processo com os dados da interface
def iniciar_processamento():
    chave_pix = chave_pix_entry.get()
    nome_empresa = nome_empresa_entry.get()

    if chave_pix and nome_empresa:
        caminho_planilha = os.path.join(os.path.dirname(__file__), 'clientes.xlsx')
        processar_planilha(caminho_planilha, chave_pix, nome_empresa)
        messagebox.showinfo("Info", "Mensagens enviadas com sucesso!")
    else:
        messagebox.showwarning("Aviso", "Por favor, preencha todos os campos.")

# Configurar a interface gráfica
root = tk.Tk()
root.title("Configuração de Notificação")

# Configurar a largura e a altura da janela
root.geometry("400x200")

# Label e entrada para chave Pix
chave_pix_label = tk.Label(root, text="Chave Pix:")
chave_pix_label.pack(pady=10)
chave_pix_entry = tk.Entry(root, width=50)
chave_pix_entry.pack(pady=5)

# Label e entrada para nome da empresa
nome_empresa_label = tk.Label(root, text="Nome da Empresa:")
nome_empresa_label.pack(pady=10)
nome_empresa_entry = tk.Entry(root, width=50)
nome_empresa_entry.pack(pady=5)

# Botão para iniciar o processamento
processar_button = tk.Button(root, text="Enviar Mensagens", command=iniciar_processamento)
processar_button.pack(pady=20)

# Iniciar o loop da interface gráfica
root.mainloop()
