from PyInstaller.utils.hooks import collect_submodules, copy_metadata

# Collect all submodules of the 'keyboard' package
hiddenimports = collect_submodules('keyboard')

# Explicitly add known backends, though collect_submodules should get them.
# This is a "just in case" measure, especially for Windows.
hiddenimports.append('keyboard._winkeyboard')

# Collect metadata associated with the 'keyboard' package
datas = copy_metadata('keyboard')
