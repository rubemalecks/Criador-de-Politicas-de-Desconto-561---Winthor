# Criador de Políticas de Desconto (561) — WinThor

Programa desktop (Python + tkinter) que gera políticas de desconto da rotina 561 do WinThor.
Tudo fica em um arquivo: `gerador_561.py`.

## Rodar pelo código

```powershell
git clone https://github.com/rubemalecks/Criador-de-Politicas-de-Desconto-561---Winthor.git
cd Criador-de-Politicas-de-Desconto-561---Winthor
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe gerador_561.py
```

Ambiente usado: Python 3.14 x64 no Windows. O tkinter já vem com o Python.

## Chamada pelo WinThor

Pela ação da rotina 529 o WinThor passa 5 argumentos, nesta ordem:
`USUARIOWT SENHABD ALIASBD USUARIOBD CODROTINA`. A senha nunca é exibida, salva ou registrada.

## Gerar o .exe

O `.exe` leva o Oracle Client x64 embutido. A pasta `oracle_client\` (com `oci.dll` direto nela) **não vai para o Git**
(passa de 100 MB): copie do pendrive ou baixe o Instant Client x64 da Oracle antes de compilar.

```powershell
.\.venv\Scripts\python.exe -m PyInstaller --clean --noconfirm RB561.spec    # gera dist\RB561.exe
.\.venv\Scripts\python.exe -m PyInstaller --clean --noconfirm RB9750.spec   # mesmo programa com o nome RB9750.exe
```

`build\`, `dist\` e o `.exe` ficam fora do Git; o `.exe` é distribuído à parte.
