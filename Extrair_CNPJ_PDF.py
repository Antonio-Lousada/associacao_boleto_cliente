import os
import pandas as pd
import pdfplumber
import re

# Caminhos
pasta_boletos = "D:\\AntonioLocal\\Python\\Extrair_CNPJ_PDF\\boletos"  # Substitua pelo caminho correto
planilha_clientes = "D:\\AntonioLocal\\Python\\Extrair_CNPJ_PDF\\clientes_boletos.xlsx"

# CNPJ do emissor (deve ser ignorado)
CNPJ_EMISSOR = "44107573000194"

# Carregar a planilha
clientes_df = pd.read_excel(planilha_clientes, dtype={"CNPJ": str})
clientes_df["Caminho Boleto"] = ""

# Função para extrair o CNPJ do cliente de um PDF
def extrair_cnpj(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        # Como cada boleto tem apenas uma página, pegamos a única disponível
        pagina = pdf.pages[0]
        texto = pagina.extract_text() or ""
        
        # Encontrar todos os CNPJs (sequências de 14 dígitos)
        cnpjs = re.findall(r"\d{14}", texto)
        
        # Remover o CNPJ do emissor e pegar o outro (do cliente)
        cnpjs_filtrados = [cnpj for cnpj in cnpjs if cnpj != CNPJ_EMISSOR]
        
        if cnpjs_filtrados:
            return cnpjs_filtrados[0].zfill(14)  # Garantir que tenha 14 dígitos
        return None

# Associar boletos aos clientes
for arquivo in os.listdir(pasta_boletos):
    if arquivo.lower().endswith(".pdf"):
        caminho_pdf = os.path.join(pasta_boletos, arquivo)
        cnpj_extraido = extrair_cnpj(caminho_pdf)
        
        if cnpj_extraido:
            # Buscar o cliente pelo CNPJ (corrigindo possíveis perdas de zeros à esquerda)
            mask = clientes_df["CNPJ"].astype(str).str.zfill(14) == cnpj_extraido
            if mask.any():
                clientes_df.loc[mask, "Caminho Boleto"] = caminho_pdf

# Salvar a planilha atualizada
clientes_df.to_excel("clientes_atualizado.xlsx", index=False)

print("Processo concluído! Os boletos foram identificados e associados aos clientes.")
