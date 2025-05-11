from PyInstaller.utils.hooks import collect_submodules, copy_metadata, collect_dynamic_libs

hiddenimports = collect_submodules('numpy')
datas = copy_metadata('numpy')
# NumPy has many .dlls, especially for MKL/OpenBLAS integration
binaries = collect_dynamic_libs('numpy', destdir='numpy')
