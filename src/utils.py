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
    
def format_data(data):
    """
    Formata os dados para o formato desejado.
    
    Args:
        data (dict): Dados a serem formatados.
        
    Returns:
        dict: Dados formatados.
    """

    formatted_data = {}
    
    for key, value in data.items():
        if isinstance(value, list):
            formatted_data[key] = "\n".join(value)
        else:
            formatted_data[key] = value
            
    return formatted_data