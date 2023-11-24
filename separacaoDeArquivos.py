import pandas as pd
import os

# Definir o caminho do arquivo original
arquivo_original = "C:\\Users\\garot\\Downloads\\Documentos na fila para cadastro (3) (1).xlsx"

# Ler a planilha "VM" do arquivo original
planilha_vm = pd.read_excel(arquivo_original, sheet_name="VM")

# Filtrar as linhas que têm uma quantidade diferente de 28 caracteres na coluna "Classificação"
linhas_erradas = planilha_vm[planilha_vm["Classificação"].str.len() != 28]

if not linhas_erradas.empty:
    # Criar um novo arquivo com as linhas erradas
    arquivo_errado = "C:\\Users\\garot\\Downloads\\Linhas erradas.xlsx"
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
    arquivo_novo = "C:\\Users\\garot\\Downloads\\Novo documento.xlsx"
    planilha_vm.to_excel(arquivo_novo, index=False)

    # Imprimir um aviso com as informações do processo
    print(f"\nDocumento original lido: {os.path.basename(arquivo_original)}")
    print(f"Novo documento criado em: {arquivo_novo}")
    print(f"Processo concluído com sucesso!")
    print(f"Linhas erradas salvas em: {arquivo_errado}")
    print(f"Número de linhas erradas: {len(linhas_erradas)}\n")
else:
    # Imprimir aviso se não houver linhas erradas
    print(f"Todas as linhas na coluna 'Classificação' têm 28 caracteres. Nenhum arquivo foi criado.")
