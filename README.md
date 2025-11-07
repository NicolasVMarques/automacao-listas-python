# Automação para Limpeza de Listas de Contatos

Este projeto contém um script Python desenvolvido para automatizar o processo de limpeza e filtragem de planilhas de contatos (leads), preparando-as para campanhas de marketing ou prospecção.

Com esta automação, consegui otimizar tempo em uma tarefa repetitiva e demorada que eu fazia todo dia manualmente no trabalho.

Com o código executável, foi possível disponibilizar a automação em outras máquinas caso precisasse.

O script utiliza `pandas` para manipulação de dados e `tkinter` para criar interfaces gráficas de seleção de arquivos.

## Funcionalidades Principais

* **Seleção de Arquivos:** Solicita ao usuário que selecione uma planilha principal (Excel ou CSV).
* **Filtragem de Colunas:** Mantém apenas as colunas essenciais: `razao_social`, `cnpj`, `email` e os `telefones`.
* **Limpeza de Telefones:** Remove caracteres especiais como `(`, `)`, `-` e espaços.
* **Divisão Dinâmica:** Separa múltiplos telefones (separados por vírgula) em colunas dinâmicas (ex: `telefone_1`, `telefone_2`, ...).
* **Filtragem Dupla (Blocklist):**
    1.  Solicita um arquivo `.csv` de "Blocklist".
    2.  Solicita um arquivo `.csv` de "Não Perturbe".
    3.  Remove qualquer linha da planilha principal se *qualquer um* dos seus telefones for encontrado em *qualquer uma* dessas listas.
* **Reordenação:** Organiza as colunas finais na ordem: `razao_social`, `telefone_1` (e outros), `cnpj`, `email`.
* **Geração de Log:** Salva um arquivo `_log.txt` detalhado junto com o arquivo final, registrando todos os passos da execução.
* **Salvar Resultado:** Pede ao usuário para escolher onde salvar a planilha limpa (Excel ou CSV).

## Como Usar

Para executar este projeto em sua máquina local, siga os passos abaixo.

1. **Clone o repositório:**
```bash
git clone https://github.com/NicolasVMarques/automacao-listas-python
```

2. **Entre na pasta do projeto:**
```bash
cd Automacao-CNPJ-Biz
```

3. **Criar um ambiete virtual:**
```bash
python -m venv venv_limpo
```

4. **Ativar o ambiete virtual:**
```bash
# No Windows (Prompt de Comando ou Anaconda Prompt)
.\venv_limpo\Scripts\activate

# No macOS / Linux (ou Git Bash)
source venv_limpo/bin/activate
```

5. **Instalar as depndências:**
```bash
pip install pandas openpyxl pyinstaller
```

6. **Executar o Script:**
```bash
python script.py
```

7. **(Opcional) Como Criar o Arquivo Executável (.exe):**
```bash
python -m PyInstaller --onefile script.py
```

O seu arquivo script.exe estará pronto dentro da pasta dist que será criada.
