import pandas as pd
import os
import pyautogui as py

# Pedir ao usuário para inserir o caminho da pasta
py.alert('Sempre escreva se baseando nesse modelo : \n\nC:\\Users\\usuário\\pasta do arquivo\\nome do arquivo  \n\n AVISO: Caso voce tenha copiado o caminho diretamente da pasta, não se esqueça de retirar as aspas "" iniciais e finais')

caminho_pasta = py.prompt(text='Coloque o caminho da pasta',
                                  title='Separador de Arquivos', default="")

# Definir o caminho do arquivo original
arquivo_original = f"{caminho_pasta}"

# Extrair o diretório do arquivo original
diretorio_original = os.path.dirname(arquivo_original)

# Ler a planilha "VM" do arquivo original
planilha_base = py.prompt(text="Coloque o nome da planilha(aba) de Excel que será lida: \n\n EX: 'Linha9' ou 'VM'",
                                  title='Separador de Arquivos')

planilha_vm = pd.read_excel(arquivo_original, sheet_name = f"{planilha_base}")

# Filtrar as linhas que têm uma quantidade diferente de 28 caracteres na coluna "Classificação"
linhas_erradas = planilha_vm[planilha_vm["Classificação"].str.len() != 28]

if not linhas_erradas.empty:
    # Criar um novo arquivo com as linhas erradas
    arquivo_errado = os.path.join(diretorio_original, "Linhas_erradas.xlsx")
    linhas_erradas.to_excel(arquivo_errado, index=False)

    # Remover as linhas erradas da planilha original
    planilha_vm = planilha_vm[planilha_vm["Classificação"].str.len() == 28]

    # Verificar e adicionar colunas ausentes
    colunas_necessarias = ["Categoria", "Identificador", "Título", "Revisão", "Data", "Sistema (Subsistema)",
                           "Linha", "Tr/Subtr", "Etapa", "Classe", "Classificação", "Projeto", "Projetista",
                           "Nº Contrato", "Nome do arquivo", "Caminho arquivo", "Segurança"]

    for coluna in colunas_necessarias:
        if coluna not in planilha_vm.columns:
            planilha_vm[coluna] = ""

        # Criar variáveis para cada coluna extraída da "Classificação"
        tipo = "DT_"
        planilha_vm['Categoria'] = tipo + planilha_vm['Classificação'].str[0:2]
        sistema = planilha_vm['Classificação'].str[3]
        planilha_vm['Sistema (Subsistema)'] = sistema + planilha_vm['Classificação'].str[14:18]
        planilha_vm['Linha'] = planilha_vm['Classificação'].str[5:7]
        tr = planilha_vm['Classificação'].str[8:10]
        planilha_vm['Tr/Subtr'] = tr + planilha_vm['Classificação'].str[11:13]
        planilha_vm['Etapa'] = planilha_vm['Classificação'].str[19]
        planilha_vm['Classe'] = planilha_vm['Classificação'].str[21:24]

        # Criar um novo arquivo com as linhas corretas
        arquivo_novo = os.path.join(diretorio_original, "Novo_documento.xlsx")
        planilha_vm.to_excel(arquivo_novo, index=False)

        # Exibir informações em uma caixa de diálogo
    mensagem = (
        f"Documento original lido: {os.path.basename(arquivo_original)}\n\n"
        f"\nNovo documento criado em: {arquivo_novo}\n"
        f"\nProcesso concluído com sucesso!\n"
        f"\nLinhas erradas salvas em: {arquivo_errado}\n"
        f"\nNúmero de linhas fora do padrão correto: {len(linhas_erradas)}\n"
    )
    py.alert(mensagem, "Informações do Processo")

else:
    # Exibir aviso se não houver linhas erradas
    py.alert(
        f"Todas as linhas na coluna 'Classificação' têm 28 caracteres. Nenhum arquivo foi criado.",
        "Aviso"
    )