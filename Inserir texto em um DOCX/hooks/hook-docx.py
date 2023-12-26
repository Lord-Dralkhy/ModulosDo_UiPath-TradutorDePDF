# hooks/hook-docx.py
from PyInstaller.utils.hooks import collect_all

# Adicione todas as dependências necessárias do docx e docx.shared
datas, binaries, hiddenimports = collect_all('docx')
datas_shared, binaries_shared, hiddenimports_shared = collect_all('docx.shared')

# Combine as listas de dados, binários e importações ocultas
datas += datas_shared
binaries += binaries_shared
hiddenimports += hiddenimports_shared
