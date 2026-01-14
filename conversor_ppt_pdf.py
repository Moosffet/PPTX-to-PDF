import os
import platform
import subprocess

def pptx_para_pdf_universal(caminho_pptx, caminho_pdf):
    sistema = platform.system()
    entrada = os.path.abspath(caminho_pptx)
    saida = os.path.abspath(caminho_pdf)
    diretorio_saida = os.path.dirname(saida)

    print(f"Sistema detectado: {sistema}")

    if sistema == "Windows":
        try:
            import comtypes.client
            # Inicia o PowerPoint no Windows
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            powerpoint.Visible = 1
            deck = powerpoint.Presentations.Open(entrada)
            # 32 é o código para formato PDF
            deck.SaveAs(saida, 32)
            deck.Close()
            powerpoint.Quit()
            print("Conversão concluída via PowerPoint (Windows).")
        except ImportError:
            print("Erro: Instale 'pip install comtypes' no Windows.")
            
    elif sistema == "Linux":
        try:
            # No Linux, usamos o comando 'soffice' ou 'libreoffice'
            comando = [
                "libreoffice", 
                "--headless", 
                "--convert-to", "pdf", 
                "--outdir", diretorio_saida, 
                entrada
            ]
            subprocess.run(comando, check=True)
            print("Conversão concluída via LibreOffice (Linux).")
        except Exception as e:
            print(f"Erro no Linux: Certifique-se de que o LibreOffice está instalado. {e}")

    else:
        print(f"Sistema {sistema} não suportado nativamente por este script.")


arquivo_entrada = "/home/matheusvasconcelos/Documentos/conversor/apresentacao.pptx"
arquivo_saida = "resultado.pdf"

if os.path.exists(arquivo_entrada):
    pptx_para_pdf_universal(arquivo_entrada, arquivo_saida)
else:
    print("Arquivo de entrada não encontrado.")
