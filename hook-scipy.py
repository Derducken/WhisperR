from PyInstaller.utils.hooks import collect_submodules, copy_metadata, collect_dynamic_libs

hiddenimports = collect_submodules('scipy')
datas = copy_metadata('scipy')
# SciPy also has many .dlls and depends on NumPy's BLAS/LAPACK
binaries = collect_dynamic_libs('scipy', destdir='scipy')

# Ensure scipy._lib.messagestream is included, sometimes missed.
hiddenimports.append('scipy._lib.messagestream')
