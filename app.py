import os
import fitz  # PyMuPDF
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread

def extrair_dados_por_posicao(texto_completo):
    """
    Extrai dados gerais do RDO (cabeçalho) usando a lógica de posições.
    """
    dados = {
        "Data RDO": "Não encontrada",
    }
    linhas = texto_completo.splitlines()
    try:
        dados["Data RDO"] = linhas[10].strip()
    except IndexError:
        pass
    return dados

def extrair_mao_de_obra(caminho_pdf):
    try:
        doc = fitz.open(caminho_pdf)
        texto = "\n".join(pagina.get_text() for pagina in doc)
        doc.close()

        dados_cabecalho = extrair_dados_por_posicao(texto)
        data_rdo = dados_cabecalho.get("Data RDO", "Não encontrada")

        inicio = texto.find("RECURSOS EM OPERAÇÃO EQUIPAMENTO")
        fim = texto.find("ASSINATURAS")

        if inicio == -1 or fim == -1 or fim <= inicio:
            return []

        bloco = texto[inicio:fim]
        linhas_brutas = [l.strip() for l in bloco.splitlines()]

        # --- CORREÇÃO DO FILTRO AQUI ---
        ignorar = {
            "Classificação", "Função",
            "Manhã", "Tarde", "Noite", "Em Operação", "Fiscalizado", "Geral", "Contratado"
        }
        
        linhas = []
        for l in linhas_brutas:
            if not l: continue # Pula linhas vazias
            if l in ignorar: continue # Pula palavras da lista ignorar
            
            # A CORREÇÃO:
            # Antes você usava: if "TOTAL" in l.upper(): continue
            # Isso apagava "ESTAÇÃO TOTAL". Agora só apagamos se a linha for EXATAMENTE "TOTAL"
            if l.upper() == "TOTAL": 
                continue
                
            linhas.append(l)
        # -------------------------------

        dados = []
        i = 0
        while i < len(linhas) - 6:
            bloco_numeros = []
            j = i
            # Identifica o bloco de números (quantidades)
            while j < len(linhas) and re.fullmatch(r'\d+', linhas[j]):
                bloco_numeros.append(int(linhas[j]))
                j += 1

            if len(bloco_numeros) >= 6:
                classificacao = ""
                frente = ""
                funcao_linhas = []
                
                # Varre para trás para achar a Classificação
                for k in range(i - 1, -1, -1):
                    linha = linhas[k]
                    if linha in ("Direto", "Indireto"):
                        classificacao = linha
                        
                        # Tenta pegar a Frente de Obra
                        if k - 1 >= 0:
                            idx_frente = k - 1
                            # Pula números perdidos acima da classificação
                            while idx_frente >= 0 and re.fullmatch(r'\d+', linhas[idx_frente]):
                                idx_frente -= 1
                            frente = linhas[idx_frente] if idx_frente >= 0 else ""
                        
                        # Captura o nome do equipamento
                        # Tenta pegar entre a classificação e os números
                        funcao_linhas = linhas[k + 1:i]
                        texto_teste = " ".join(funcao_linhas).strip()
                        
                        # Se estiver vazio (ou só números), faz o backtracking (recurso de segurança)
                        if not texto_teste or re.fullmatch(r'\d+', texto_teste):
                            for m in range(k - 1, -1, -1):
                                linha_candidata = linhas[m]
                                if linha_candidata in ("Direto", "Indireto"): break
                                if "FRENTE DE OBRA" in linha_candidata.upper(): continue
                                if re.fullmatch(r'\d+', linha_candidata): continue # Pula o '0' isolado
                                
                                funcao_linhas = [linha_candidata]
                                break
                        
                        break

                funcao = " ".join(funcao_linhas).strip()

                while len(bloco_numeros) < 7:
                    bloco_numeros.append(0)

                dados.append({
                    "Nome do Arquivo": os.path.basename(caminho_pdf),
                    "Data da RDO": data_rdo,
                    "Função": funcao, # Nome do Equipamento
                    "Frente de Obra": frente,
                    "Classificação": classificacao,
                    "Contratado Geral": bloco_numeros[0],
                    "Em operação (manhã)": bloco_numeros[5],
                    "Fiscalizado (manhã)": bloco_numeros[6],
                    "Em operação (tarde)": bloco_numeros[3],
                    "Fiscalizado (tarde)": bloco_numeros[4],
                    "Em operação (noite)": bloco_numeros[1],
                    "Fiscalizado (noite)": bloco_numeros[2],
                })

                i = j
            else:
                i += 1

        return dados

    except Exception as e:
        print(f"Erro ao processar {caminho_pdf}: {e}")
        return []

def processar_pasta(pasta_origem, pasta_destino, nome_arquivo, barra_progresso, botao):
    todos_dados = []
    arquivos_pdf = [
        os.path.join(raiz, arquivo)
        for raiz, _, arquivos in os.walk(pasta_origem)
        for arquivo in arquivos if arquivo.lower().endswith('.pdf')
    ]

    total = len(arquivos_pdf)
    for idx, caminho_pdf in enumerate(arquivos_pdf, 1):
        dados = extrair_mao_de_obra(caminho_pdf)
        todos_dados.extend(dados)
        barra_progresso["value"] = (idx / total) * 100
        barra_progresso.update()

    if todos_dados:
        df = pd.DataFrame(todos_dados)
        colunas = [
            "Nome do Arquivo", "Data da RDO", "Função", "Frente de Obra", "Classificação",
            "Contratado Geral", "Em operação (manhã)", "Fiscalizado (manhã)",
            "Em operação (tarde)", "Fiscalizado (tarde)",
            "Em operação (noite)", "Fiscalizado (noite)"
        ]
        df = df[colunas]
        df.to_excel(os.path.join(pasta_destino, f"{nome_arquivo}.xlsx"), index=False)
        messagebox.showinfo("Concluído", "Planilha gerada com sucesso!")
    else:
        messagebox.showwarning("Aviso", "Nenhum dado encontrado.")

    botao["state"] = "normal"
    barra_progresso["value"] = 0

def iniciar_interface():
    def selecionar_origem():
        pasta = filedialog.askdirectory()
        entrada_origem.delete(0, tk.END)
        entrada_origem.insert(0, pasta)

    def selecionar_destino():
        pasta = filedialog.askdirectory()
        entrada_destino.delete(0, tk.END)
        entrada_destino.insert(0, pasta)

    def iniciar():
        origem = entrada_origem.get()
        destino = entrada_destino.get()
        nome = entrada_nome.get().strip()
        if not origem or not destino or not nome:
            messagebox.showerror("Erro", "Todos os campos devem ser preenchidos.")
            return
        botao_executar["state"] = "disabled"
        Thread(target=processar_pasta, args=(origem, destino, nome, barra, botao_executar)).start()

    root = tk.Tk()
    root.title("Extrator de Mão de Obra e Equipamentos - RDO")

    tk.Label(root, text="📁 Pasta com PDFs:").grid(row=0, column=0, sticky="w")
    entrada_origem = tk.Entry(root, width=50)
    entrada_origem.grid(row=0, column=1)
    tk.Button(root, text="Selecionar", command=selecionar_origem).grid(row=0, column=2)

    tk.Label(root, text="📂 Pasta para salvar:").grid(row=1, column=0, sticky="w")
    entrada_destino = tk.Entry(root, width=50)
    entrada_destino.grid(row=1, column=1)
    tk.Button(root, text="Selecionar", command=selecionar_destino).grid(row=1, column=2)

    tk.Label(root, text="📝 Nome do Excel:").grid(row=2, column=0, sticky="w")
    entrada_nome = tk.Entry(root, width=50)
    entrada_nome.grid(row=2, column=1, columnspan=2)

    barra = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
    barra.grid(row=3, column=0, columnspan=3, pady=10)

    botao_executar = tk.Button(root, text="Iniciar Extração", command=iniciar)
    botao_executar.grid(row=4, column=0, columnspan=3, pady=5)

    root.mainloop()

if __name__ == "__main__":
    iniciar_interface()
