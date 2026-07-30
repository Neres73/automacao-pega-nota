# Automação Pega Nota

Scripts de automação para localizar, baixar e conferir notas fiscais de serviço (NFS-e) via SAP/NIX, e para dar baixa em recebimentos de requisições no Coupa.

## Scripts

- **`main.py`** — Lê uma planilha de ocorrências, entra no SAP (transação `/nNIXNFSE`) e salva o PDF de cada NFS-e encontrada em uma pasta de destino. Ao final, gera um relatório em Excel com o status de cada nota (OK, erro ou não encontrada).
- **`verificanota.py`** — Confere quais notas da planilha original ainda não foram baixadas e gera uma planilha só com as faltantes, no mesmo formato esperado por `main.py`.
- **`nix_material.py`** — Automação equivalente via navegador (Playwright), usada para o cockpit web do NIX.
- **`recebimento.py`** — Lê números de requisição de uma planilha e realiza o recebimento correspondente no Coupa, via navegador (Playwright).

## Requisitos

- Python 3.10+
- Dependências: `openpyxl`, `pywin32` (`win32com`, `win32gui`), `pyautogui`, `pyperclip`, `pywinauto`, `playwright`, `python-dotenv`
- Para os scripts baseados em Playwright (`nix_material.py`, `recebimento.py`), é necessário ter o navegador instalado via `playwright install`.

## Configuração

Os caminhos de planilha e pasta de destino estão definidos no topo de cada script (seção `CONFIGURAÇÕES`). Ajuste conforme o seu ambiente antes de rodar.

Para `recebimento.py`, configure as variáveis de ambiente em um arquivo `.env` (veja o cabeçalho do script para as chaves esperadas, como `COUPA_URL` e `COLUNA_REQ`).

## Uso

```bash
python main.py
python verificanota.py
python nix_material.py
python recebimento.py --login   # primeiro uso: login manual no Coupa
python recebimento.py           # execuções seguintes, já logado
```
