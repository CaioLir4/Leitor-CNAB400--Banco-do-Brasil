import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import datetime
import webbrowser
import mysql.connector

# Variável global para armazenar a conexão
conn = None
boletos = []
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

            # Verifca se todas as configurações necessárias estão presentes
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
    global conn
    if not conn:
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
        except mysql.connector.Error as err:
            messagebox.showerror("Erro de Conexão", f"Erro ao conectar ao banco de dados: {err}")
            return None
    return conn

# Função para desconectar do banco de dados
def desconectar_banco():
    global conn
    if conn:
        conn.close()
        conn = None

# Função para consultar nosso_numero_banco onde o nosso_numero do boleto é igual ao nosso_numerotitulo no banco
def consultar_nosso_numero_banco(nosso_numero_boleto):
    conn = conectar_banco()
    if not conn:
        return None

    cursor = conn.cursor()
    query = """
        SELECT c.nome
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

# Função para processar os registros e extrair os dados
def processar_retorno_cnab400(caminho_arquivo):
    try:
        with open(caminho_arquivo, 'r') as arquivo:
            for linha in arquivo:
                tipo_registro = linha[0:1]


                # Verifica se o tipo de registro é 7, que é o tipo de detalhe do boleto
                if tipo_registro == '7':
                    boleto = {}

                    boleto['prefixo_da_agencia'] = linha[17:21].strip()  # Extrai o Prefixo da Agência

                    boleto['digito_verificador_dv_agencia'] = linha[21:22].strip() # Extrai o Dígito Verificador - D.V. - do Prefixo da Agência

                    boleto['numero_conta_corrente'] = linha[23:30].strip()  # Extrai o Número da Conta Corrente do Cedente

                    boleto['digito_verificador_dv_conta_corrente_do_cedente'] = linha[
                                                      30:31].strip()  # Extrai o Dígito Verificador - D.V. - do Número da Conta conta corrente

                    boleto['numero_do_convenio_de_cobranca_cedente'] = linha[31:38].strip()  # Extrai o Número do Convênio de Cobrança do Cedente

                    boleto['numero_controle_do_participante'] = linha[38:63].strip()  # Extrai o Número de Controle do Participante

                    boleto['nosso_numero'] = linha[63:80].strip()  # Extrai o nosso_numero

                    boleto['nosso_numero_consulta_bd'] = linha[73:80].strip()  # Extrai o nosso_numero para que eu possa consultar no banco

                    boleto['tipo_cobranca_especifico_para_comando_72'] = linha[80:82].strip()  # Extrai o Tipo de cobrança específico para comando 72

                    boleto['dias_para_calculo'] = linha[82:86].strip()  # Extrai dias para calculo

                    boleto['natureza_recebimento'] = linha[86:88].strip()  # Extrai a Natureza do recebimento

                    boleto['prefixo_boleto'] = linha[88:91].strip()  # Extrai o Prefixo do boleto

                    boleto['variacao_da_carteira'] = linha[91:94].strip()  # Extrai a Variação da Carteira

                    boleto['conta_caucao'] = linha[94:95].strip()  # Extrai a conta caução

                    boleto['taxa_para_desconto'] = linha[95:100].strip()  # Extrai a Taxa para desconto

                    boleto['taxa_IOF'] = linha[100:105].strip()  # Extrai a Taxa para IOF

                    boleto['Branco'] = linha[105:106].strip()  # BRANCO

                    boleto['Carteira'] = linha[106:108].strip()  # Extrai a Carteira

                    boleto['Comando'] = linha[108:110].strip()  # Extrai o Comando

                    resultado_banco = consultar_nosso_numero_banco(boleto['nosso_numero_consulta_bd'])

                    boleto['nome_cliente'] = resultado_banco

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

                    '''for boleto in boletos:
                        print(boleto)'''
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
    # Limpar todas as entradas anteriores
    for item in treeview.get_children():
        treeview.delete(item)

    # Definir as tags para zebramento
    treeview.tag_configure("even", background="#f2f2f2")  # Cor para linhas pares
    treeview.tag_configure("odd", background="#ffffff")   # Cor para linhas ímpares

    for idx, boleto in enumerate(boletos):
        # Consultar o 'nosso_numero' no banco
        resultado_banco = consultar_nosso_numero_banco(boleto['nosso_numero_consulta_bd'])

        if resultado_banco:
            nome_cliente = resultado_banco
            # Definir a tag para a linha com base no índice (par ou ímpar)
            tag = "even" if idx % 2 == 0 else "odd"
            treeview.insert('', 'end', values=(
                boleto['nosso_numero'],
                nome_cliente[0],
                boleto['data_pagamento'],
                boleto['data_vencimento'],
                f"R$ {boleto['valor_pago']:.2f}".replace('.', ','),
                f"R$ {boleto['valor_titulo']:.2f}".replace('.', ',')
            ), tags=(tag,), iid=idx)  # Usar 'iid' para associar o boleto completo ao item
        else:
            # Caso não encontre o nosso_numero no banco, exibe 'Não encontrado'
            tag = "even" if idx % 2 == 0 else "odd"
            treeview.insert('', 'end', values=(
                boleto['nosso_numero'],
                f"Não encontrado",
                boleto['data_pagamento'],
                boleto['data_vencimento'],
                f"R$ {boleto['valor_pago']:.2f}".replace('.', ','),
                f"R$ {boleto['valor_titulo']:.2f}".replace('.', ',')
            ), tags=(tag,), iid=idx)  # Usar 'iid' para associar o boleto completo ao item


def exibir_detalhes():
    selecionado = treeview.selection()

    if selecionado:
        # Acessa o índice do item selecionado
        index = selecionado[0]
        boleto_completo = boletos[int(index)]  # Acessa o boleto completo da lista 'boletos'

        # Exibir todos os detalhes do boleto
        try:
            nome_cliente = boleto_completo.get('nome_cliente', ('Não encontrado',))[0]
        except TypeError:
            nome_cliente = 'Não encontrado'

        # Identificar o comando com base nas condições
        if boleto_completo['Comando'] in ['05', '06', '07', '08', '15', '46']:
            comando_descricao = f"{boleto_completo['Comando']} (Liquidado)"
        elif boleto_completo['Comando'] == '02':
            comando_descricao = f"{boleto_completo['Comando']} (Entrada)"
        elif boleto_completo['Comando'] in ['09', '10', '20']:
            comando_descricao = f"{boleto_completo['Comando']} (Baixa)"
        elif boleto_completo['Comando'] == '03':
            comando_descricao = f"{boleto_completo['Comando']} (Recusa)"
        else:
            comando_descricao = boleto_completo['Comando']

        # Criar os detalhes do boleto
        detalhes = (
            f"Nosso Número: {boleto_completo['nosso_numero']}\n"
            f"Nome Cliente: {nome_cliente}\n"
            f"Data de Pagamento: {boleto_completo['data_pagamento']}\n"
            f"Data de Vencimento: {boleto_completo['data_vencimento']}\n"
            f"Valor Pago: R$ {boleto_completo['valor_pago']:.2f}\n"
            f"Valor do Título: R$ {boleto_completo['valor_titulo']:.2f}\n"
            f"Prefixo da Agência: {boleto_completo['prefixo_da_agencia']}\n"
            f"Dígito Verificador da Agência: {boleto_completo['digito_verificador_dv_agencia']}\n"
            f"Número Conta Corrente: {boleto_completo['numero_conta_corrente']}\n"
            f"Número do Convênio de Cobrança: {boleto_completo['numero_do_convenio_de_cobranca_cedente']}\n"
            f"Nosso Número (consulta BD): {boleto_completo['nosso_numero_consulta_bd']}\n"
            f"Tipo de Cobrança: {boleto_completo['tipo_cobranca_especifico_para_comando_72']}\n"
            f"Dias para Cálculo: {boleto_completo['dias_para_calculo']}\n"
            f"Natureza de Recebimento: {boleto_completo['natureza_recebimento']}\n"
            f"Prefixo Boleto: {boleto_completo['prefixo_boleto']}\n"
            f"Variação da Carteira: {boleto_completo['variacao_da_carteira']}\n"
            f"Conta Caução: {boleto_completo['conta_caucao']}\n"
            f"Taxa para Desconto: {boleto_completo['taxa_para_desconto']}\n"
            f"Taxa IOF: {boleto_completo['taxa_IOF']}\n"
            f"Carteira: {boleto_completo['Carteira']}\n"
            f"Comando: {comando_descricao}\n"
        )
        messagebox.showinfo("Detalhes do Boleto", detalhes)
    else:
        messagebox.showwarning("Aviso", "Selecione um boleto para visualizar os detalhes.")


# Função para abrir sites
def abrir_SITES():
    webbrowser.open("https://github.com/CaioLir4")
    webbrowser.open("https://acsistemasnet.com.br/")

# Criando a janela principal
root = tk.Tk()
root.title("Sistema de Retorno CNAB400")
root.geometry("900x550")

# Botões e funcionalidades
btn_carregar = tk.Button(root, text="Carregar Arquivo CNAB400", command=carregar_boletos, width=30)
btn_carregar.pack(pady=10)

btn_detalhes = tk.Button(root, text="Exibir Detalhes", command=exibir_detalhes, width=30)
btn_detalhes.pack(pady=10)

# Configuração da árvore (treeview) para exibir os boletos
columns = ("Nosso Número", "Nome Cliente", "Data de Pagamento", "Data de Vencimento", "Valor Pago", "Valor")
treeview = ttk.Treeview(root, columns=columns, show="headings", selectmode="browse")

# Cabeçalhos
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

# Exibindo a árvore de visualização
treeview.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# Adicionando o Label de crédito com link clicável
link_credito = tk.Label(root, text="Created By Caio Lira for ACSISTEMAS",
                        font=("Helvetica", 8, "bold"), fg="black", bg="#f4f4f4")
link_credito.pack(side="bottom", pady=10)
link_credito.bind("<Button-1>", lambda e: abrir_SITES())  # Abre ao clicar

# Inicia a interface gráfica
root.mainloop()

# Fechar a conexão ao banco de dados ao sair
desconectar_banco()
