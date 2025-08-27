import importlib.util
import sys
import os

# Garante que estamos rodando Python 3.12
if not (sys.version_info.major == 3 and sys.version_info.minor == 12):
    print("❌ Este programa requer Python 3.12.")
    sys.exit(1)

# Adiciona a pasta atual ao sys.path para imports locais funcionarem
sys.path.insert(0, os.path.dirname(__file__))

# Caminho do main.pyc
pyc_path = os.path.join(os.path.dirname(__file__), "__pycache__", "main.cpython-312.pyc")

# Carrega o módulo main a partir do .pyc
spec = importlib.util.spec_from_file_location("main", pyc_path)
main = importlib.util.module_from_spec(spec)
sys.modules["main"] = main
spec.loader.exec_module(main)
