import json

def load_json(path):
    try:
        with open(path, "r", encoding="utf-8") as arquivo:
            return json.load(arquivo)  # Converte JSON para dicionário
    except FileNotFoundError:
        print(f"Erro: Arquivo '{path}' não encontrado.")
        return {}
    except json.JSONDecodeError:
        print(f"Erro: Arquivo '{path}' contém um JSON inválido.")
        return {}