# Importador de Vendas (CSV)

## Como gerar o executável (Windows)

Use o PyInstaller para criar um `.exe` único que roda sem instalar Python.

### 1) Criar ambiente virtual e instalar dependências

```bash
python -m venv .venv
.venv\Scripts\activate
pip install pyinstaller pyautogui
```

### 2) Gerar o executável

```bash
pyinstaller --onefile --windowed main.py
```

### 3) Onde está o arquivo final?

O executável será criado em:

```
dist\main.exe
```

Você pode renomear o arquivo para algo como `ImportadorVendas.exe` e distribuir.

## Observações

- Gere o `.exe` no Windows onde você quer usar o arquivo final.
- O `pyautogui` pode pedir permissões de automação do teclado/mouse.
- Os arquivos `config.json` e `logs/app.log` são criados automaticamente na primeira execução.
