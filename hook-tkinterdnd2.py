from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = collect_submodules('tkinterdnd2')
datas = collect_data_files('tkinterdnd2')
