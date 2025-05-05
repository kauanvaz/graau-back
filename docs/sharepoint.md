# Documentação do módulo Sharepoint

## Visão Geral

O módulo `sharepoint.py` fornece funcionalidades para interagir com o Microsoft SharePoint, permitindo recuperar e transformar dados que serão utilizados na geração de relatórios. O módulo se especializa na conexão com um site específico do SharePoint, autenticação e formatação de dados obtidos para uso no sistema de relatórios.

## Componentes Principais

### Classes

#### `Sharepoint`

A classe principal para interação com o SharePoint.

**Construtor:**

```python
def __init__(self)
```

- Inicializa a conexão com o SharePoint usando credenciais do arquivo `.env`
- Carrega mapeamentos de diretorias a partir do arquivo `src/mappings/diretorias.json`
- Carrega mapeamentos de divisões a partir do arquivo `src/mappings/divisoes.json`

**Métodos:**

##### `get_all_lists`

```python
def get_all_lists(self)
```

Obtém a lista de todas as listas disponíveis no site do SharePoint.

**Retorna:**
- Lista com os nomes de todas as listas no site

##### `_transform_data`

```python
def _transform_data(self, data)
```

Método interno que transforma os dados brutos do SharePoint em um formato estruturado e padronizado.

**Parâmetros:**
- `data`: Dados brutos obtidos do SharePoint

**Retorna:**
- Lista de dicionários com dados transformados e formatados

**Funcionalidades:**
- Realiza o parsing de valores separados por ";#"
- Formata datas para o padrão brasileiro (DD/MM/AAAA)
- Converte strings para inteiros
- Formata nomes próprios seguindo regras de capitalização
- Mapeia códigos de diretoria para seus nomes completos
- Mapeia códigos de divisão para seus nomes completos
- Formata valores monetários para o padrão brasileiro
- Trata campos de múltiplos valores

##### `_get_data`

```python
def _get_data(self, list_name=None, query=None)
```

Método interno para obter dados de uma lista específica do SharePoint.

**Parâmetros:**
- `list_name`: Nome da lista a ser consultada
- `query`: Consulta opcional para filtrar os dados

**Retorna:**
- Dados brutos obtidos da lista do SharePoint

##### `get_acao_controle_data`

```python
def get_acao_controle_data(self, item_id=None)
```

Obtém dados de ações de controle do SharePoint, opcionalmente filtrados por ID.

**Parâmetros:**
- `item_id`: ID opcional do item específico a ser obtido

**Retorna:**
- Lista de dicionários com dados de ações de controle transformados

## Funções Auxiliares Internas

O método `_transform_data` utiliza várias funções auxiliares internas:

- `safe_split`: Realiza split de valores separados por ";#" e retorna o segundo elemento
- `safe_date_format`: Formata datas para o padrão brasileiro (DD/MM/AAAA)
- `safe_int`: Converte valores para inteiros, tratando casos especiais
- `safe_list_split`: Processa listas de valores com split
- `safe_multiple_split`: Processa múltiplos valores com split separados por ";#", removendo elementos vazios do início e do fim
- `safe_alternate_split`: Realiza split e retorna elementos alternados (índices ímpares)
- `format_name`: Formata nomes próprios com capitalização adequada, mantendo preposições e artigos em minúsculas

## Transformação de Dados

A transformação de dados inclui:

- Extração de valores de diretoria e divisão a partir do campo 'Divisão de Origem Ajustada'
- Mapeamento desses códigos para nomes completos utilizando os arquivos de mapeamento
- Formatação de nomes próprios (relatores, procuradores)
- Formatação de valores monetários (VRF) utilizando localização brasileira
- Processamento de listas de valores em diversos formatos

## Dependências

- `shareplum`: Para interação com a API do SharePoint
- `dotenv`: Para carregar variáveis de ambiente
- `babel.numbers`: Para formatação de valores monetários
- `pathlib`: Para manipulação de caminhos de arquivo
- `utils`: Módulo interno contendo função `load_json` para carregar arquivos JSON

## Configuração

O módulo depende das seguintes configurações:

- Arquivo `.env` com as variáveis:
  - `USUARIO`: Nome de usuário para autenticação no SharePoint. Formato: user.name@tce.pi.gov.br
  - `SENHA`: Senha para autenticação no SharePoint

- Arquivos de mapeamento:
  - `src/mappings/diretorias.json`: Mapeamento de códigos de diretoria para nomes completos
  - `src/mappings/divisoes.json`: Mapeamento de códigos de divisão para nomes completos

## Exemplo de Uso

```python
# Importar o módulo
from src.sharepoint import Sharepoint

# Criar instância do cliente SharePoint
sp_client = Sharepoint()

# Obter dados de uma ação de controle específica
acao_controle = sp_client.get_acao_controle_data(item_id=3868)

print(acao_controle)
```

## Observações

- O módulo realiza conexão automática com o site do SharePoint "https://tcepi365.sharepoint.com/sites/SecretariadeControleExterno"
- As consultas são realizadas na lista "Cadastro de Ação de Controle"
- O processo de transformação trata diversos casos especiais e formatos de dados
- Os mapeamentos de diretorias e divisões permitem a exibição de nomes completos e formatados nos relatórios
