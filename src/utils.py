import json
from unidecode import unidecode
from pathlib import Path

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

# def _clean_string(string):
#     # Lista de preposições comuns em português
#     preposicoes = {
#         'a', 'ante', 'após', 'até', 'com', 'contra', 'de', 'desde',
#         'em', 'entre', 'para', 'per', 'perante', 'por', 'sem',
#         'sob', 'sobre', 'trás', 'o', 'os', 'a', 'as', 'um', 'uma',
#         'uns', 'umas', 'do', 'da', 'dos', 'das', 'no', 'na', 'nos', 'nas',
#         'ao', 'aos', 'à', 'às', 'pelo', 'pela', 'pelos', 'pelas'
#     }
    
#     texto = string.lower()

#     # Remove acentuação
#     texto = unidecode(texto)

#     # Divide em palavras
#     palavras = texto.split()

#     # Remove preposições
#     palavras_filtradas = [p for p in palavras if p not in preposicoes]

#     # Junta com underline
#     return '_'.join(palavras_filtradas)

def _clean_secoes(secoes):
    def clean(secao):
        level_title = secao.get("title")
        data = []

        for d in secao.get("data", []):
            heading_title = d.get("title")
            subtitles = []
            
            for sub in secao.get("subtitles", []):
                # Ignora subtítulos irrelevantes
                if "Digite" in sub.get("title", ""):
                    continue

                # Aplica a função recursivamente para pegar títulos aninhados
                clean_subtitle = clean(sub)
                if clean_subtitle:  # Se não for None
                    subtitles.append(clean_subtitle)
                    
            data.append({
                "title": heading_title,
                "subtitles": subtitles,
            })
            subtitles = []

        # Retorna o dicionário apenas se o título for válido
        if "Digite" not in level_title:
            return {
                "title": level_title,
                "data": data
            }
            
        return None

    # Aplica a função de limpeza para todas as seções de nível 1
    return [secao_limpa for secao in secoes if (secao_limpa := clean(secao))]

def format_data(data: dict):
    """
    Formata os dados para o formato desejado no relatório.
    
    Args:
        data: Dados a serem formatados.
        
    Returns:
        dict: Dados formatados.
    """

    formatted_data = {}
    
    for key, value in data.items():
        if isinstance(value, list) and len(value) > 0:
            if isinstance(value[0], str): # Lista do sharepoint
                formatted_data[key] = "\n".join(value)
            elif isinstance(value[0], dict): # Lista do front (seções)
                formatted_data[key] = _clean_secoes(value)
        else:
            formatted_data[key] = value

    return formatted_data

def get_status_processo(tipo_relatorio: str, processo_tipo: str):
    """
    Retorna o status do processo com base no tipo de relatório
    e tipo de processo mapeados.
    
    Args:
        tipo_relatorio: Tipo de relatório
        processo_tipo: Tipo de processo
    
    Returns:
        str: Status do processo.
    """

    mapping_path = Path("src/mappings/status_processo.json")
    mapping = load_json(mapping_path)
    
    first_layer = mapping.get(tipo_relatorio.lower(), {})
    
    if not first_layer:
        return ""
    
    # Pega lista de opções de tipos de processo e remove o último elemento, que é o "default"
    options = list(first_layer.keys())
    options.pop(-1)
    
    # Verifica se o tipo de processo está na lista de opções, capturando o índice
    bool_list = [option in processo_tipo.lower() for option in options]
    true_index = bool_list.index(True) if any(bool_list) else -1
    
    if true_index != -1:
        return first_layer.get(options[true_index], "")
    else:
        return first_layer.get("default", "")
