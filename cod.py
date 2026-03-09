# AUTOMAÇÃO LISTA MESTRA PQ 10
import re
import time
import logging
from datetime import datetime
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
from dateutil.parser import parse

# --- CONFIGURAÇÃO ---
# Coloque aqui o nome do seu arquivo Excel de entrada
CAMINHO_ARQUIVO_ENTRADA = "PQ 10 Anexo I.xlsx" 
NOME_ARQUIVO_SAIDA = "resultado_verificacao.xlsx"

CONFIG = {
    'timeout_requisicao': 20,
    'user_agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36",
    'prefixo_excluir': "https://sgp.madrix.app/",
    'max_tentativas': 3
}

LINK_PATTERN = re.compile(r"https?://[^\s)'\"]+")

# Padrões para Datas
PATTERNS_ULTIMA_MOD = [
    re.compile(r"Última modificação:.*?(\d{1,2}\s+de\s+[a-zA-Zçãõé]+\s+de\s+\d{4})", re.I),
    re.compile(r"Última modificação:.*?(\d{2}/\d{2}/\d{4})", re.I),
]
PATTERNS_ATUALIZACAO = [
    re.compile(r"Atualizad[oa]\s+(?:em|:)?\s*(\d{2}/\d{2}/\d{4})", re.I),
    re.compile(r"Atualizad[oa]\s+(?:em|:)?\s*(\d{1,2}\s+de\s+[a-zA-Zçãõé]+\s+de\s+\d{4})", re.I),
]
PATTERNS_PUBLICACAO = [
    re.compile(r"Publicado\s+(?:em|:)?\s*(\d{2}/\d{2}/\d{4})", re.I),
    re.compile(r"Publicado\s+(?:em|:)?\s*(\d{1,2}\s+de\s+[a-zA-Zçãõé]+\s+de\s+\d{4})", re.I),
]
GENERIC_PATTERNS = [
    re.compile(r"(\d{1,2}\s+de\s+[a-zA-Zçãõé]+\s+de\s+\d{4})", re.I),
    re.compile(r"(\d{2}/\d{2}/\d{4})", re.I),
    re.compile(r"(\d{4}-\d{2}-\d{2})", re.I),
]

def criar_sessao():
    session = requests.Session()
    session.headers.update({"User-Agent": CONFIG['user_agent']})
    retries = Retry(total=CONFIG['max_tentativas'], backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session

def buscar_data(sessao, link):
    try:
        if not link.startswith('http'): return None, "Link inválido"
        
        response = sessao.get(link, timeout=CONFIG['timeout_requisicao'], verify=False)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, "html.parser")
        for script in soup(["script", "style"]): script.decompose()
        texto = soup.get_text(" ", strip=True)

        # Tenta todos os padrões de data
        for lista_patterns in [PATTERNS_ULTIMA_MOD, PATTERNS_ATUALIZACAO, PATTERNS_PUBLICACAO, GENERIC_PATTERNS]:
            for pattern in lista_patterns:
                if match := pattern.search(texto):
 
                    if lista_patterns == GENERIC_PATTERNS:
                        all_matches = re.compile('|'.join(p.pattern for p in GENERIC_PATTERNS), re.I).findall(texto)
                        if len(all_matches) > 3: return None, None
                        return match.group(0).strip(), None
                    return match.group(1).strip(), None
        return None, None
    except Exception as e:
        return None, f"Erro: {type(e).__name__}"

def main():
    print("--- INICIANDO VERIFICAÇÃO (MODO LOCAL) ---")
    requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)
    
    # 1. Leitura do Excel com detecção automática do cabeçalho
    try:
        print(f"Lendo arquivo: {CAMINHO_ARQUIVO_ENTRADA}...")
        df_raw = pd.read_excel(CAMINHO_ARQUIVO_ENTRADA, header=None, dtype=str)
        
        header_idx = -1
        for idx, row in df_raw.iterrows():
            txt = " ".join([str(v).lower() for v in row.values if pd.notna(v)])
            if "código" in txt and "título" in txt:
                header_idx = idx
                break
        
        if header_idx == -1:
            print("ERRO: Não foi possível encontrar a linha de cabeçalho (Código/Título).")
            return

        df = pd.read_excel(CAMINHO_ARQUIVO_ENTRADA, header=header_idx, dtype=str)
        print(f"Tabela encontrada na linha {header_idx + 1}. Total de registros: {len(df)}")
    except FileNotFoundError:
        print(f"ERRO: O arquivo '{CAMINHO_ARQUIVO_ENTRADA}' não foi encontrado.")
        return
    except Exception as e:
        print(f"ERRO CRÍTICO AO LER ARQUIVO: {e}")
        return

    sessao = criar_sessao()
    resultados = []
    links_processados = set()

    for index, row in df.iterrows():
        # Extração de Código e Título
        codigo = str(row['Código']).strip() if 'Código' in df.columns else str(row.iloc[0]).strip()
        titulo = str(row['Título']).strip() if 'Título' in df.columns else str(row.iloc[1]).strip()
        
        if pd.isna(codigo) or codigo.lower() == 'nan' or 'código' in codigo.lower():
            continue

        # Busca link na linha inteira
        linha_txt = " ".join([str(val) for val in row.values if pd.notna(val)])
        matches = LINK_PATTERN.finditer(linha_txt)
        
        for match in matches:
            link = match.group(0).strip()
            
            if link in links_processados or link.startswith(CONFIG['prefixo_excluir']):
                continue
                
            links_processados.add(link)
            print(f"Verificando: {codigo} - {link[:40]}...", end=" ")
            
            data_encontrada, erro = buscar_data(sessao, link)
            situacao = ""
            
            if erro:
                situacao = erro
                print(f"[ERRO: {erro}]")
            elif data_encontrada:
                situacao = "Não atualizado"
                agora = datetime.now()
                try:
                    dt = parse(data_encontrada, dayfirst=True)
                    if dt.year == agora.year and dt.month == agora.month:
                        situacao = "Atualizado"
                except: pass
                print(f"[{data_encontrada}] -> {situacao}")
            else:
                situacao = "Verificar manualmente"
                print("[DATA NÃO ENCONTRADA]")

            resultados.append({
                "Código": codigo,
                "Título": titulo,
                "Link": link,
                "Data Encontrada": data_encontrada,
                "Situação": situacao
            })

    # Salvar
    if resultados:
        df_res = pd.DataFrame(resultados)
        df_res.to_excel(NOME_ARQUIVO_SAIDA, index=False)
        print(f"\nConcluído! Arquivo salvo como: {NOME_ARQUIVO_SAIDA}")
    else:
        print("\nNenhum link foi encontrado ou processado.")

if __name__ == "__main__":
    main()