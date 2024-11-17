import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import datetime
import mysql.connector


# Função para ler as configurações de conexão do arquivo rede.txt
def ler_configuracoes_conexao():
    try:
        # Caminho do arquivo
        caminho_arquivo = r'C:\LC sistemas - Softhouse\rede.txt'

        # Lê o arquivo de configuração
        with open(caminho_arquivo, 'r') as arquivo:
            config = {}
            for linha in arquivo:
                if ':' in linha:
                    chave, valor = linha.strip().split(':', 1)
                    config[chave.strip()] = valor.strip()

            # Verifica se todas as configurações necessárias estão presentes
            obrigatorias = ['IP', 'DB', 'USER', 'KEY', 'PORT', 'TERMINAL_TIPO', 'ID_EMPRESA_PADRAO']
            for chave in obrigatorias:
                if chave not in config:
                    raise ValueError(f"Configuração faltando: {chave}")

            return config
    except FileNotFoundError:
        messagebox.showerror("Erro", "Arquivo de configuração não encontrado!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo de configuração: {e}")
    return None


# Função para conectar ao banco de dados
def conectar_banco():
    config = ler_configuracoes_conexao()
    if not config:
        return None

    try:
        # Conecta ao banco de dados com base nas configurações lidas
        conn = mysql.connector.connect(
            host=config['IP'],  # Host
            user=config['USER'],  # Usuário
            password=config['KEY'],  # Senha
            database=config['DB'],  # Banco de dados
            port=int(config['PORT'])  # Porta
        )
        return conn
    except mysql.connector.Error as err:
        messagebox.showerror("Erro de Conexão", f"Erro ao conectar ao banco de dados: {err}")
        return None

# Função para consultar nosso_numero_banco onde o nosso_numero do boleto é igual ao nosso_numerotitulo no banco
def consultar_nosso_numero_banco(nosso_numero_boleto):
    conn = conectar_banco()
    if not conn:
        return None

    cursor = conn.cursor()
    query = """
        SELECT nosso_numero, c.nome
        FROM boletoremessadet b
        INNER JOIN cliente c on c.id = b.id_cliente
        WHERE nosso_numero = %s
    """
    try:
        cursor.execute(query, (nosso_numero_boleto,))
        resultado = cursor.fetchone()
        return resultado  # Retorna tanto o 'nosso_numero' quanto o 'nome' do cliente
    except mysql.connector.Error as err:
        messagebox.showerror("Erro", f"Erro ao executar a consulta: {err}")
        return None
    finally:
        cursor.close()
        conn.close()

# Função para processar os registros e extrair os dados
def processar_retorno_cnab400(caminho_arquivo):
    boletos = []
    try:
        with open(caminho_arquivo, 'r') as arquivo:
            for linha in arquivo:
                tipo_registro = linha[0:1]

                # Verifica se o tipo de registro é 7, que é o tipo de detalhe do boleto
                if tipo_registro == '7':
                    boleto = {}
                    boleto['nosso_numero'] = linha[62:80].strip()  # Extrai o nosso_numero
                    boleto['nosso_numero_consulta_bd'] = linha[73:80].strip()  # Extrai o nosso_numero para que eu possa consultar no banco

                    # Data de pagamento (posição 110 a 116, formato DDMMAA)
                    data_pagamento_str = linha[110:116].strip()
                    try:
                        data_pagamento = datetime.datetime.strptime(data_pagamento_str, "%d%m%y").date()
                    except ValueError:
                        data_pagamento = None  # Caso a data não seja válida
                    boleto['data_pagamento'] = data_pagamento.strftime("%d/%m/%Y") if data_pagamento else "Data inválida"

                    # Data de vencimento (posição 146 a 152, formato DDMMAA)
                    data_vencimento_str = linha[146:152].strip()
                    try:
                        data_vencimento = datetime.datetime.strptime(data_vencimento_str, "%d%m%y").date()
                    except ValueError:
                        data_vencimento = None  # Caso a data não seja válida
                    boleto['data_vencimento'] = data_vencimento.strftime("%d/%m/%Y") if data_vencimento else "Data inválida"

                    # Tenta converter o valor do título, lidando com possíveis erros de conversão
                    try:
                        valor_titulo = int(linha[153:165].strip()) / 100  # valor do título em centavos
                    except ValueError:
                        valor_titulo = 0.0  # Valor inválido, define como 0.00

                    # Tenta converter o valor pago, lidando com possíveis erros de conversão
                    try:
                        valor_pago = int(linha[254:266].strip()) / 100  # valor pago em centavos
                    except ValueError:
                        valor_pago = 0.0  # Valor inválido, define como 0.00

                    boleto['valor_pago'] = valor_pago
                    boleto['valor_titulo'] = valor_titulo

                    boletos.append(boleto)
        return boletos
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o arquivo: {e}")
        return []


# Função para abrir o arquivo e carregar os boletos
def carregar_boletos():
    caminho_arquivo = filedialog.askopenfilename(title="Selecionar Arquivo CNAB400",
                                                 filetypes=[("Arquivo CNAB400", "*.ret")])

    if caminho_arquivo:
        boletos = processar_retorno_cnab400(caminho_arquivo)
        if boletos:
            listar_boletos(boletos)
        else:
            messagebox.showinfo("Aviso", "Nenhum boleto encontrado no arquivo.")


# Função para listar os boletos na árvore de visualização
def listar_boletos(boletos):
    for item in treeview.get_children():
        treeview.delete(item)

    for boleto in boletos:
        # Consultar o 'nosso_numero' no banco
        resultado_banco = consultar_nosso_numero_banco(boleto['nosso_numero_consulta_bd'])

        if resultado_banco:
            nosso_numero_banco_value, nome_cliente = resultado_banco
            treeview.insert('', 'end', values=(
                boleto['nosso_numero'],
                nome_cliente,
                boleto['data_pagamento'],
                boleto['data_vencimento'],
                f"R$ {boleto['valor_pago']:.2f}".replace('.', ','),
                f"R$ {boleto['valor_titulo']:.2f}".replace('.', ',')
            ))
        else:
            # Caso não encontre o nosso_numero no banco, exibe 'Não encontrado'
            treeview.insert('', 'end', values=(
                boleto['nosso_numero'],
                f"Não encontrado",
                boleto['data_pagamento'],
                boleto['data_vencimento'],
                f"R$ {boleto['valor_pago']:.2f}".replace('.', ','),
                f"R$ {boleto['valor_titulo']:.2f}".replace('.', ',')
            ))


# Função para exibir detalhes do boleto selecionado
def exibir_detalhes():
    selecionado = treeview.selection()
    if selecionado:
        index = selecionado[0]
        boleto = treeview.item(index, 'values')
        messagebox.showinfo("Detalhes do Boleto",
                            f"Nosso Número: {boleto[0]}\nNome do Cliente: {boleto[1]}\nData de Vencimento: {boleto[3]}\nValor Pago: {boleto[4]}\nValor: {boleto[5]}")
    else:
        messagebox.showwarning("Aviso", "Selecione um boleto para visualizar os detalhes.")


# Criando a janela principal
root = tk.Tk()
root.title("Sistema de Retorno CNAB400")
root.geometry("1000x600")

# Botões e funcionalidades
btn_carregar = tk.Button(root, text="Carregar Arquivo CNAB400", command=carregar_boletos, width=30)
btn_carregar.pack(pady=10)

btn_detalhes = tk.Button(root, text="Exibir Detalhes", command=exibir_detalhes, width=30)
btn_detalhes.pack(pady=10)

# Configuração da árvore (treeview) para exibir os boletos
columns = ("Nosso Número", "Nome Cliente", "Data de Pagamento", "Data de Vencimento", "Valor Pago", "Valor")
treeview = ttk.Treeview(root, columns=columns, show="headings", height=30)

treeview.heading("Nosso Número", text="Nosso Número")
treeview.heading("Nome Cliente", text="Nome Cliente")
treeview.heading("Data de Pagamento", text="Data de Pagamento")
treeview.heading("Data de Vencimento", text="Data de Vencimento")
treeview.heading("Valor Pago", text="Valor Pago")
treeview.heading("Valor", text="Valor")

# Ajustando a largura das colunas
treeview.column("Nosso Número", width=150, anchor="center")
treeview.column("Nome Cliente", width=200, anchor="center")
treeview.column("Data de Pagamento", width=120, anchor="center")
treeview.column("Data de Vencimento", width=120, anchor="center")
treeview.column("Valor Pago", width=100, anchor="center")
treeview.column("Valor", width=100, anchor="center")

treeview.pack(pady=20, padx=20)

# Iniciando a interface gráfica
root.mainloop()
