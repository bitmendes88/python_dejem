import importlib.util
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

def load_pyc(name, filename):
    spec = importlib.util.spec_from_file_location(name, filename)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module

# Carrega módulos
cadastrar = load_pyc("cadastrar", os.path.join(os.path.dirname(__file__), "cadastrar.pyc"))
main = load_pyc("main", os.path.join(os.path.dirname(__file__), "main.pyc"))
