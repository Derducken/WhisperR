from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Collect all py7zr Python modules
hiddenimports = collect_submodules('py7zr')

# Collect all data files (including 7z DLLs if they exist)
datas = collect_data_files('py7zr')

# Explicitly include py7zr's required binaries
binaries = []
