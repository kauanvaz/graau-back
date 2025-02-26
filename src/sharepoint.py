from shareplum import Site
from shareplum import Office365
import os
from dotenv import load_dotenv
from babel.numbers import format_currency

load_dotenv()

# Configurações do SharePoint
class Sharepoint():
    def __init__(self) -> None:
        site_url_base = "https://tcepi365.sharepoint.com"
        site_url = "https://tcepi365.sharepoint.com/sites/SecretariadeControleExterno"
        username = os.getenv("USUARIO")
        password = os.getenv("SENHA")
        self.site = Site(site_url, authcookie=Office365(site_url_base, username=username, password=password).GetCookies())

    def get_all_lists(self):
        lists = self.site.GetListCollection()
        list_names = [list_info['Title'] for list_info in lists]
    
        return list_names

    def _transform_data(self, data):
        def safe_split(value, index=1):
            """Realiza split de forma segura, retornando valor vazio se houver erro."""
            if not value or not isinstance(value, str):
                return ''
            parts = value.split(";#")
            return parts[index] if len(parts) > index else value

        def safe_date_format(value):
            """Formata data de forma segura, retornando string vazia se houver erro."""
            if not value:
                return ''
            try:
                return value.strftime("%d/%m/%Y")
            except (AttributeError, TypeError):
                return ''

        def safe_int(value, default=0):
            """Converte para inteiro de forma segura, retornando valor default se houver erro."""
            try:
                if isinstance(value, str):
                    value = safe_split(value)
                return int(float(value))
            except (ValueError, TypeError):
                return default

        def safe_list_split(value):
            """Processa lista de valores com split, retornando string vazia se houver erro."""
            if not value:
                return ''
            if isinstance(value, list):
                return ', '.join([safe_split(x) for x in value])
            return value

        def safe_multiple_split(value):
            """Processa múltiplos valores com split, removendo elementos vazios."""
            if not value or not isinstance(value, str):
                return ''
            parts = value.split(";#")
            if len(parts) <= 1:
                return value
            return ', '.join(filter(None, parts[1:-1]))
        
        def safe_alternate_split(value):
            """
            Realiza split e retorna elementos alternados de forma segura.
            Para dados no formato "item1;#valor1;#item2;#valor2", retorna "valor1, valor2"
            """
            if not value or not isinstance(value, str):
                return ''
            try:
                parts = value.split(";#")
                # Pega elementos em índices ímpares (1, 3, 5, etc)
                values = parts[1::2]
                return ', '.join(filter(None, values))
            except (IndexError, TypeError):
                return ''

        return [
            {
                'id': i.get('ID', ''),
                'unidades_fiscalizadas': safe_split(i.get('Unidades Fiscalizadas', '')),
                'criado_por': safe_split(i.get('Criado por', '')[0]) if isinstance(i.get('Criado por', ''), list) else '',
                'divisao_origem_ajustada': i.get('Divisão de Origem Ajustada', ''),
                'finalidade_acao_de_controle': i.get('Finalidade da ação de controle', ''),
                'tipo_acao': i.get('Tipo de ação', ''),
                'n_processo_eTCE': safe_split(i.get('Nº Processo e-TCE', '')),
                'n_processo_eTCE_processo_tipo': safe_split(i.get('Nº Processo e-TCE: processoTipo', '')),
                'situacao_acao_de_controle': safe_split(i.get('Situação da Ação de Controle', '')),
                'data_inicio_acao': safe_date_format(i.get('Data de Início da Ação:', '')),
                'exercicios': ', '.join(filter(None, i.get('Exercícios', '').split(";#"))),
                'data_conclusão_acao_de_controle': safe_date_format(i.get('Data de conclusão da Ação de Controle', '')),
                'tempestividade_acao_de_controle': i.get('Tempestividade da Ação de Controle', ''),
                'informe_metodologia_VRF': i.get('Informe a metodologia do VRF:', ''),
                'VRF': format_currency(i.get('Volume de Recursos Fiscalizados (VRF):', ''), 'BRL', locale='pt_BR'),
                'temas_PACEX': safe_split(i.get('Tema(s) do PACEX', '')),
                'linha_atuacao_descrição_tema': safe_split(i.get('Linha de Atuação: Descrição Tema', '')),
                'equipe_fiscalizacao': safe_list_split(i.get('Equipe de Fiscalização', '')),
                'beneficios_qualitativos': safe_multiple_split(i.get('Benefícios Qualitativos', '')),
                'tecnicas_aplicadas': safe_multiple_split(i.get('Técnicas Aplicadas', '')),
                'beneficios_efetivos': i.get('Benefícios efetivos:', ''),
                'proposta_beneficios_potenciais': i.get('Proposta de benefícios potenciais ', ''),
                'beneficio_nao_financeiro_proposto': i.get('Benefício não Financeiro Proposto', ''),
                'beneficio_nao_financeiro_efetivo': i.get('Benefício não financeiro efetivo', ''),
                'quantidade_medidas_cautelares_solicitadas': i.get('Quantidade de medidas cautelares solicitadas;', ''),
                'municipios_visitados_in_loco': safe_alternate_split(i.get('Municípios visitados in loco', '')),
                'quantidade_visitas_realizadas': safe_int(i.get('Quantidade de visitas realizadas', 0)),
                'canva_fiscalizacao': i.get('Canva de Fiscalização ', ''),
                'DVR': i.get('DVR', ''),
                'matriz_planejamento': i.get('Matriz de Planejamento ', ''),
                'matriz_achados': i.get('Matriz de achados', ''),
                'data_conclusao_relatorio_preliminar': safe_date_format(i.get('Data de conclusão do Relatório Preliminar', '')),
                'motivo_encerramento_acao': i.get('Motivo do Encerramento da ação', ''),
                'encaminhamentos': safe_multiple_split(i.get('Encaminhamentos', '')),
                'acao__controle_ativa': i.get('Ação de controle ativa?', ''),
                'dias_em_atividade': safe_int(i.get('Dias em atividade', '0.0')),
                'anexos': i.get('Anexos', ''),
                'acoes_controle_relacionadas': safe_split(i.get('Ações de controle relacionadas', '')),
                'acoes_controle_PAI_objeto': i.get('Ações de controle PAI: Objeto ', ''),
                'acoes_controle_PAI_VRF': safe_split(i.get('Ações de controle PAI: VRF (R$)', '')),
                'modificado_por': i.get('Modificado por', ''),
                'modificado_data': safe_date_format(i.get('Modificado', '')),
                'valor_licitacoes_analisadas': i.get('Valor das Licitações Analisadas', 0.0),
                'utilizou_matriz_risco_NUGEI': i.get('Utilizou matriz de Risco da NUGEI?', ''),
                'trimestre_conclusao': safe_int(i.get('Trimestre de conclusão', '0.0')),
                'criado_data': safe_date_format(i.get('Criado', '')),
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
    
    def get_processos_ETCE_data(self, item_processo_codigo=None):
        query=None
        if item_processo_codigo:
            query = {'Where': [('Eq', 'processoCodigo', str(item_processo_codigo))]}
            
        data = self._get_data(list_name='Processos E-TCE', query=query)
        
        return data

if __name__ == '__main__':
    names_list = Sharepoint().get_all_lists()
    
    print(names_list, end="\n\n")
    print("#"*80, end="\n\n")
    
    # result = Sharepoint().get_acao_controle_data(item_id=3868)
    result = Sharepoint().get_processos_ETCE_data()
    
    print(result)
