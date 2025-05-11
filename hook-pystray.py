from PyInstaller.utils.hooks import collect_submodules, copy_metadata, collect_data_files

# Collect submodules for pystray
hiddenimports = collect_submodules('pystray')

# pystray might have platform-specific backends, ensure they are included.
# For example, on Windows, it might use a specific module.
# Add known ones if issues persist, e.g., hiddenimports.append('pystray._win32')
# However, collect_submodules should generally cover this.

# Collect metadata for pystray
datas = copy_metadata('pystray')

# pystray uses Pillow (PIL) for icons. While PIL has its own hooks,
# ensuring its data files are collected can sometimes be helpful if icon issues arise.
# datas += collect_data_files('PIL', include_py_files=True) # Usually not needed as PIL has good hooks
