import os
import comtypes.client
from tqdm import tqdm

def docx_para_pdf(word, caminho_docx, caminho_pdf):
    doc = word.Documents.Open(caminho_docx)
    doc.SaveAs(caminho_pdf, FileFormat=17)  # 17 = PDF
    doc.Close()

pasta_origem = r"C:\pasta\concluido"
pasta_destino = r"C:\pasta\concluido_pdf"

# Inicializa o Word (visível = False)
word = comtypes.client.CreateObject('Word.Application')
word.Visible = False

# Lista todos os arquivos DOCX
arquivos_docx = []
for raiz, _, arquivos in os.walk(pasta_origem):
    for arquivo in arquivos:
        if arquivo.lower().endswith(".docx"):
            arquivos_docx.append((raiz, arquivo))

try:
    for raiz, arquivo in tqdm(arquivos_docx, desc="Convertendo arquivos", unit="arquivo"):
        caminho_word = os.path.join(raiz, arquivo)
        
        caminho_relativo = os.path.relpath(raiz, pasta_origem)
        pasta_pdf = os.path.join(pasta_destino, caminho_relativo)
        os.makedirs(pasta_pdf, exist_ok=True)
        
        caminho_pdf = os.path.join(pasta_pdf, os.path.splitext(arquivo)[0] + ".pdf")
        
        # Se já existe, pula
        if os.path.exists(caminho_pdf):
            continue
        
        # Converte
        docx_para_pdf(word, caminho_word, caminho_pdf)

finally:
    # Fecha o Word quando terminar ou der erro
    word.Quit()

print("✅ Conversão concluída com estrutura de pastas preservada!")
