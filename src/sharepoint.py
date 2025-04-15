from shareplum import Site
from shareplum import Office365
import os
from dotenv import load_dotenv
from babel.numbers import format_currency
from pathlib import Path
import re

try:
    from .utils import load_json
except ImportError:
    from utils import load_json

class Sharepoint():
    def __init__(self) -> None:
        site_url_base = "https://tcepi365.sharepoint.com"
        site_url = "https://tcepi365.sharepoint.com/sites/SecretariadeControleExterno"
        username = os.getenv("USUARIO")
        password = os.getenv("SENHA")
        self.site = Site(site_url, authcookie=Office365(site_url_base, username=username, password=password).GetCookies())
        self.diretorias_mapping = load_json(Path('src/mappings/diretorias.json'))
        self.divisoes_mapping = load_json(Path('src/mappings/divisoes.json'))
        
        load_dotenv()

    def get_all_lists(self):
        lists = self.site.GetListCollection()
        list_names = [list_info['Title'] for list_info in lists]
    
        return list_names

    def _transform_data(self, data):
        def safe_split(value):
            """Realiza split simples.
            
            Exemplo:
            "item1;#valor1"
            
            Resultado:
            "valor1"
            """
            if not value or not isinstance(value, str):
                return ''
            parts = value.split(";#")
            return parts[1] if len(parts) > 1 else list(filter(None, parts))

        def safe_date_format(value):
            """Formata data de datetime para string no formato "%d/%m/%Y"."""
            if not value:
                return ''
            try:
                return value.strftime("%d/%m/%Y")
            except (AttributeError, TypeError):
                return ''

        def safe_int(value, default=0):
            """Converte para inteiro de forma segura, retornando valor default se houver erro."""
            try:
                if isinstance(value, str) and ";#" in value:
                    value = safe_split(value)
                if '.' in value:
                    value = float(value)
                return int(value)
            except (ValueError, TypeError):
                return default

        def safe_list_split(value):
            """Processa lista de valores com split."""
            if not value:
                return []
            if isinstance(value, list):
                return [safe_split(x) for x in value]
            return value

        def safe_multiple_split(value):
            """Processa múltiplos valores com split, removendo elementos vazios do início e do fim.
            
            Exemplo:
            ";#valor1;#valor2;#", retorna "valor1, valor2"
            
            Resultado:
            ["valor1", "valor2"]
            """
            if not value or not isinstance(value, str):
                return []
            parts = value.split(";#")
            if len(parts) <= 1:
                return value
            return list(filter(None, parts[1:-1]))
        
        def safe_alternate_split(value):
            """
            Realiza split e retorna elementos alternados (índices ímpares).
            Exemplo:
            "item1;#valor1;#item2;#valor2"
            
            Resultado:
            ["valor1, "valor2"]
            """
            if not value or not isinstance(value, str):
                return []
            try:
                parts = value.split(";#")
                # Pega elementos em índices ímpares (1, 3, 5, etc)
                values = parts[1::2]
                return values
            except (IndexError, TypeError):
                return []
            
        def format_name(nome):
            """
            Converte um nome para title case (primeira letra de cada palavra maiúscula),
            mas mantém preposições e artigos em minúsculas.
            
            Args:
                nome (str): O nome a ser formatado
                
            Returns:
                str: O nome formatado
            """

            preposicoes = ['a', 'o', 'as', 'os', 'de', 'da', 'do', 'das', 'dos', 'em', 'na', 'no', 
                        'nas', 'nos', 'por', 'para', 'com', 'e']
            

            palavras = nome.lower().split()
            resultado = []
            
            for i, palavra in enumerate(palavras):
                # A primeira palavra e palavras que não são preposições recebem title case
                if i == 0 or palavra not in preposicoes:
                    resultado.append(palavra.capitalize())
                else:
                    # Preposições ficam em minúsculas
                    resultado.append(palavra)
            
            return ' '.join(resultado)

        return [
            {
                'acao__controle_ativa': i.get('Ação de controle ativa?', ''),
                'acoes_controle_PAI_objeto': i.get('Ações de controle PAI: Objeto ', ''),
                'anexos': safe_int(i.get('Anexos', 0)),
                'beneficios_efetivos': i.get('Benefícios efetivos:', ''),
                'beneficios_qualitativos': safe_multiple_split(i.get('Benefícios Qualitativos', [])),
                'classe': safe_split(i.get('Nº Processo: Classe', '')),
                'criado_data': safe_date_format(i.get('Criado', '')),
                'data_conclusao_relatorio_preliminar': safe_date_format(i.get('Data de conclusão do Relatório Preliminar', '')),
                'data_conclusão_acao_de_controle': safe_date_format(i.get('Data de conclusão da Ação de Controle', '')),
                'data_inicio_acao': safe_date_format(i.get('Data de Início da Ação:', '')),
                'dias_em_atividade': safe_int(i.get('Dias em atividade', 0)),
                'divisao_origem_ajustada': i.get('Divisão de Origem Ajustada', ''),
                'divisao_origem_ajustada_diretoria': (
                    self.diretorias_mapping.get(i.get('Divisão de Origem Ajustada', '').split('/')[0].strip(), '')
                    ),
                'divisao_origem_ajustada_divisao': (
                    self.divisoes_mapping.get(i.get('Divisão de Origem Ajustada', '').split('/')[1].strip(), '')
                    ),
                'equipe_fiscalizacao': safe_list_split(i.get('Equipe de Fiscalização', [])),
                'exercicios': safe_alternate_split(i.get('Exercícios', [])),
                'finalidade_acao_de_controle': i.get('Finalidade da ação de controle', ''),
                'id': i.get('ID', ''),
                'informe_metodologia_VRF': i.get('Informe a metodologia do VRF:', ''),
                'linha_atuacao_descrição_tema': safe_alternate_split(i.get('Linha de Atuação: Descrição Tema', [])),
                'modificado_data': safe_date_format(i.get('Modificado', '')),
                'modificado_por': i.get('Modificado por', ''),
                'motivo_encerramento_acao': i.get('Motivo do Encerramento da ação', ''),
                'municipios_visitados_in_loco': safe_alternate_split(i.get('Municípios visitados in loco', [])),
                'n_processo_eTCE': safe_split(i.get('Nº Processo e-TCE', '')),
                'processo_tipo': safe_split(i.get('Nº Processo e-TCE: processoTipo', '')).title(),
                'procurador': format_name(safe_split(i.get('Nº Processo: procurador', ''))),
                'proposta_beneficios_potenciais': i.get('Proposta de benefícios potenciais ', ''),
                'quantidade_medidas_cautelares_solicitadas': i.get('Quantidade de medidas cautelares solicitadas;', ''),
                'relator': format_name(safe_split(i.get('Nº Processo: relator', ''))),
                'situacao_acao_de_controle': safe_split(i.get('Situação da Ação de Controle', '')),
                'subclasse': safe_split(i.get('Nº Processo: Subclasse', '')),
                'tecnicas_aplicadas': safe_multiple_split(i.get('Técnicas Aplicadas', [])),
                'temas_PACEX': safe_alternate_split(i.get('Tema(s) do PACEX', [])),
                'tempestividade_acao_de_controle': i.get('Tempestividade da Ação de Controle', ''),
                'tipo_acao': i.get('Tipo de ação', ''),
                'trimestre_conclusao': safe_int(i.get('Trimestre de conclusão', 0)),
                'unidades_fiscalizadas': safe_alternate_split(i.get('Unidades Fiscalizadas', [])),
                'utilizou_matriz_risco_NUGEI': i.get('Utilizou matriz de Risco da NUGEI?', ''),
                'VRF': format_currency(i.get('Volume de Recursos Fiscalizados (VRF):', ''), 'BRL', locale='pt_BR'),
            } for i in data
        ]

    def _get_data(self, list_name=None, query=None):
        if not list_name: return None
        
        sp_list = self.site.List(list_name)
        
        if query:
            fields = None
            data = sp_list.GetListItems(fields=fields, query=query)
        else:
            data = sp_list.GetListItems()
            
        return data
    
    def get_acao_controle_data(self, item_id=None):
        query=None
        if item_id:
            query = {'Where': [('Eq', 'ID', str(item_id))]}
            
        data = self._get_data(list_name='Cadastro de Ação de Controle', query=query)
        
        return self._transform_data(data)

if __name__ == '__main__':
    
    result = Sharepoint().get_acao_controle_data(item_id=3868)
    
    print(result)
