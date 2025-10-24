import pandas as pd
import re
from tkinter import Tk, filedialog
import os
import sys
import logging
import io

def limpar_telefones(df_telefones):
    """Função auxiliar para limpar uma coluna/Series de telefones."""
    # Garante que é string e remove caracteres indesejados
    return df_telefones.astype(str).str.replace(r'[()\s-]', '', regex=True)

def carregar_planilha(titulo, tipos_arquivo):
    """Abre uma caixa de diálogo para selecionar um arquivo."""
    logging.info(f"Por favor, selecione o arquivo: {titulo}...")
    file_path = filedialog.askopenfilename(
        title=titulo,
        filetypes=tipos_arquivo
    )
    
    if not file_path:
        logging.warning("Nenhum arquivo selecionado. Encerrando.")
        sys.exit()

    logging.info(f"Arquivo '{os.path.basename(file_path)}' carregado.")
    return file_path

def carregar_dataframe(file_path):
    """Tenta carregar um arquivo Excel ou CSV para um DataFrame."""
    try:
        if file_path.endswith('.csv'):
            try:
                df = pd.read_csv(file_path, sep=';', on_bad_lines='skip')
                if df.shape[1] == 1: 
                   df = pd.read_csv(file_path, sep=',', on_bad_lines='skip')
            except Exception:
                 df = pd.read_csv(file_path, sep=',', on_bad_lines='skip')
        else:
            df = pd.read_excel(file_path)
        return df
    except Exception as e:
        logging.error(f"Erro ao ler o arquivo {file_path}: {e}")
        sys.exit()

def main():
    # --- INÍCIO DA CONFIGURAÇÃO DO LOG ---
    log_stream = io.StringIO()
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.StreamHandler(log_stream)
        ]
    )
    # --- FIM DA CONFIGURAÇÃO DO LOG ---

    # 1. Configurar o Tkinter
    root = Tk()
    root.withdraw() 
    logging.info("Iniciando o processo...")

    # 2. Carregar planilha principal
    path_principal = carregar_planilha(
        "Selecione a planilha principal (Excel/CSV)",
        [("Planilhas", "*.xlsx *.xls *.csv"), ("Todos os arquivos", "*.*")]
    )
    df = carregar_dataframe(path_principal)

    # 3. Filtrar colunas
    colunas_desejadas = ['cnpj', 'razao_social', 'email', 'telefones']
    colunas_existentes = [col for col in colunas_desejadas if col in df.columns]
    
    if 'telefones' not in colunas_desejadas:
        logging.error("Erro: A coluna 'telefones' é essencial e não foi encontrada. Encerrando.")
        sys.exit()
        
    df = df[colunas_existentes]
    logging.info("Colunas filtradas.")

    # 4. Limpar caracteres especiais da coluna 'telefones'
    df['telefones'] = limpar_telefones(df['telefones'])
    
    df['telefones'] = df['telefones'].replace(r'^\s*$', pd.NA, regex=True)
    df = df.dropna(subset=['telefones'])

    # 5. Separar telefones em colunas dinâmicas (telefone_1, telefone_2, ...)
    logging.info("Separando telefones em colunas dinâmicas...")

    split_telefones = df['telefones'].str.split(',', expand=True)
    
    max_phones = split_telefones.shape[1]
    
    # Esta é a lista de colunas de telefone, ex: ['telefone_1', 'telefone_2']
    phone_columns = [f'telefone_{i+1}' for i in range(max_phones)]
    
    split_telefones.columns = phone_columns
    
    for col in phone_columns:
        split_telefones[col] = split_telefones[col].str.strip()

    df = df.drop(columns='telefones').join(split_telefones)
    
    df[phone_columns] = df[phone_columns].fillna('')
    
    logging.info(f"Telefones divididos em {max_phones} colunas (de 'telefone_1' a 'telefone_{max_phones}').")
    logging.info(f"Total de leads para verificar: {len(df)}")

    # 6. Carregar e aplicar a Blocklist
    path_blocklist = carregar_planilha(
        "Selecione o arquivo 'Blocklist' (CSV)",
        [("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")]
    )
    df_blocklist = carregar_dataframe(path_blocklist)
    
    logging.info("Extraindo telefones da Blocklist (assumindo ser a 2ª coluna)...")
    
    telefones_brutos_blocklist = df_blocklist.iloc[:, 1]
    telefones_blocklist = set(limpar_telefones(telefones_brutos_blocklist))
    
    if '' in telefones_blocklist:
        telefones_blocklist.remove('')
    
    logging.info(f"Carregados {len(telefones_blocklist)} telefones únicos da Blocklist.")

    total_antes = len(df)
    is_blocked = df[phone_columns].isin(telefones_blocklist).any(axis=1)
    df = df[~is_blocked]
    logging.info(f"Filtro 'Blocklist' aplicado. {total_antes - len(df)} linhas removidas.")

    # 7. Carregar e aplicar 'Não Perturbe'
    path_nao_perturbe = carregar_planilha(
        "Selecione o arquivo 'Não Perturbe' (CSV)",
        [("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")]
    )
    df_nao_perturbe = carregar_dataframe(path_nao_perturbe)

    logging.info("Extraindo telefones do 'Não Perturbe' (assumindo ser a 2ª coluna)...")
    
    telefones_brutos_nao_perturbe = df_nao_perturbe.iloc[:, 1]
    telefones_nao_perturbe = set(limpar_telefones(telefones_brutos_nao_perturbe))

    if '' in telefones_nao_perturbe:
        telefones_nao_perturbe.remove('')
    
    logging.info(f"Carregados {len(telefones_nao_perturbe)} telefones únicos do 'Não Perturbe'.")

    total_antes = len(df)
    is_not_disturb = df[phone_columns].isin(telefones_nao_perturbe).any(axis=1)
    df = df[~is_not_disturb]
    logging.info(f"Filtro 'Não Perturbe' aplicado. {total_antes - len(df)} linhas removidas.")


    # --- INÍCIO DA MODIFICAÇÃO (Seção 8) ---

    # 8. Reordenar colunas
    logging.info("\nProcessamento de filtros concluído.")
    try:
        # Define a ordem ideal, começando com 'razao_social', 
        # depois TODAS as colunas de telefone, e por fim 'cnpj' e 'email'
        ordem_desejada = ['razao_social'] + phone_columns + ['cnpj', 'email']
        
        # Pega todas as colunas atuais do DataFrame
        colunas_atuais = list(df.columns)
        
        # Cria a nova lista de colunas, na ordem desejada,
        # mas apenas com as colunas que realmente existem no DataFrame
        nova_ordem = [col for col in ordem_desejada if col in colunas_atuais]
        
        # Adiciona ao final quaisquer colunas que estavam no DataFrame 
        # mas não foram mencionadas na 'ordem_desejada' (garante que não haja perda de dados)
        colunas_restantes = [col for col in colunas_atuais if col not in nova_ordem]
        nova_ordem.extend(colunas_restantes)
        
        # Aplica a nova ordem
        df = df[nova_ordem]
        logging.info(f"Colunas reordenadas para: {', '.join(nova_ordem)}")
        
    except Exception as e:
        logging.error(f"Não foi possível reordenar as colunas: {e}")

    # --- FIM DA MODIFICAÇÃO (Seção 8) ---


    # 9. Salvar o resultado e o log
    logging.info(f"Total de linhas restantes: {len(df)}")
    
    log_file_path = "" # Inicializa a variável
    
    if len(df) == 0:
        logging.warning("Nenhum dado restou após os filtros. Nenhum arquivo será salvo.")
        log_file_path = "automacao_log.txt" # Salva um log padrão no local do script
    else:
        save_path = filedialog.asksaveasfilename(
            title="Salvar arquivo final",
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx"), ("Arquivo CSV", ".csv")]
        )

        if not save_path:
            logging.warning("Nenhum local selecionado. O arquivo não foi salvo.")
            log_file_path = "automacao_log.txt" # Salva um log padrão no local do script
        else:
            try:
                if save_path.endswith('.csv'):
                    df.to_csv(save_path, index=False, encoding='utf-8-sig', sep=';')
                else:
                    df.to_excel(save_path, index=False, engine='openpyxl')
                
                logging.info(f"Arquivo final salvo com sucesso em: {save_path}")
                # Define o caminho do log para ser o mesmo do arquivo salvo, mas com .txt
                log_file_path = os.path.splitext(save_path)[0] + "_log.txt"
                
            except Exception as e:
                logging.error(f"Erro ao salvar o arquivo: {e}")
                log_file_path = "automacao_log_ERRO.txt" # Salva log de erro

    # --- SALVAR O ARQUIVO DE LOG ---
    if log_file_path:
        try:
            with open(log_file_path, 'w', encoding='utf-8') as f:
                f.write(log_stream.getvalue())
            logging.info(f"Log de execução salvo em: {log_file_path}")
        except Exception as e:
            logging.error(f"Falha ao salvar o arquivo de log: {e}")
            
    logging.info("--- Execução Finalizada ---")
    log_stream.close() # Libera a memória

if __name__ == "__main__":
    main()