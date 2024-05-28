import openpyxl

# Criar um novo arquivo de planilha
wb = openpyxl.Workbook()
ws = wb.active

# Definir o título da planilha
ws.title = "ITENS"
nova_aba = wb.create_sheet("Pedido de Compra")

# Cabeçalho com os itens desejados
cabecalho = [
    "ANO", "MÊS (RFQ)", "REQUESTOR", "Quotation no.", "Request Quotation Date",
    "Request Vendor Date", "Receive Vendor Quotation Date", "Send Quotation Date to dept",
    "GP APPROVAL DATE", "EMISSÃO", "ORDEM DE COMPRA", "TYPE 2", "TYPE 3", "P/N",
    "DESCRIÇÃO", "QUANTIDADE", "UND MEDIDA", "VALOR UNITÁRIO 1", "VALOR UNITÁRIO 2",
    "VALOR TOTAL", "FORNECEDOR", "REQUISITANTE"
]

# Adicionar o cabeçalho na primeira linha
ws.append(cabecalho)

# Selecionar a aba recém-criada
nova_aba = wb["Pedido de Compra"]

# Cabeçalho da tabela
cabecalho_tabela = ["", "Valores", "", "", ""]
sub_cabecalho = ["DESCRIPTION", "QTY", "US$ UNIT", "US$ TOTAL"]
info_adicional = ["NCM/ HS CODE", "PART Nº", "WON AUTOMATION KOREA CO., LTD", "", ""]

# Dados da tabela
dados = [
    ["85444200#000", "GGLZ241029405", "VACUUM HOLDER  KSLD100S H420 B850P", 1, "$2.072,00", "$2.072,00"],
    ["84719012#000", "GGLZ241029406", "Amount", 1, "2072", "$2.072,00"]
]

# Adicionar o cabeçalho e sub-cabeçalho à tabela
nova_aba.append(cabecalho_tabela)
nova_aba.append(sub_cabecalho)
nova_aba.append(info_adicional)

# Adicionar os dados à tabela
for linha in dados:
    nova_aba.append(linha)

# Salvar o arquivo de planilha
wb.save("Cotacoes.xlsx")

# Salvar o arquivo de planilha
wb.save("Controle de Compras.xlsx")
