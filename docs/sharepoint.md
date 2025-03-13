# Documentação do módulo Sharepoint

## Visão Geral

O módulo `sharepoint.py` fornece funcionalidades para interagir com o Microsoft SharePoint, permitindo recuperar e transformar dados que serão utilizados na geração de relatórios. O módulo se especializa na conexão com um site específico do SharePoint, autenticação e formatação de dados obtidos para uso no sistema de relatórios.

## Componentes Principais

### Classes

#### `Sharepoint`

A classe principal para interação com o SharePoint.

**Construtor:**

```python
def __init__(self) -> None
```

- Inicializa a conexão com o SharePoint usando credenciais do arquivo `.env`
- Carrega mapeamentos de diretorias a partir de um arquivo JSON

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
- Realiza o parsing seguro de valores separados por ";#"
- Formata datas para o padrão brasileiro (DD/MM/AAAA)
- Converte strings para inteiros de forma segura
- Formata nomes próprios seguindo regras de capitalização
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

- `safe_split`: Realiza split de forma segura em valores separados por ";#"
- `safe_date_format`: Formata datas de forma segura
- `safe_int`: Converte valores para inteiros de forma segura
- `safe_list_split`: Processa listas de valores com split
- `safe_multiple_split`: Processa múltiplos valores com split
- `safe_alternate_split`: Realiza split em elementos alternados
- `format_name`: Formata nomes próprios com capitalização adequada
- `mapear_divisao`: Mapeia códigos de divisão para nomes completos

## Dependências

- `shareplum`: Para interação com a API do SharePoint
- `dotenv`: Para carregar variáveis de ambiente
- `babel.numbers`: Para formatação de valores monetários
- `pathlib`: Para manipulação de caminhos de arquivo
- `re`: Para operações com expressões regulares

## Configuração

O módulo depende das seguintes configurações:

- Arquivo `.env` com as variáveis:
  - `USUARIO`: Nome de usuário para autenticação no SharePoint
  - `SENHA`: Senha para autenticação no SharePoint

- Arquivo de mapeamento:
  - `src/mappings/diretorias.json`: Mapeamento de códigos de diretoria para nomes completos

## Exemplo de Uso

```python
# Importar o módulo
from src.sharepoint import Sharepoint

# Criar instância do cliente SharePoint
sp_client = Sharepoint()

# Obter dados de uma ação de controle específica
acao_controle = sp_client.get_acao_controle_data(item_id=3868)

# Processar os dados
if acao_controle:
    item = acao_controle[0]
    print(f"Unidades Fiscalizadas: {item['unidades_fiscalizadas']}")
    print(f"Processo: {item['n_processo_eTCE']}")
    print(f"VRF: {item['VRF']}")
    print(f"Exercícios: {item['exercicios']}")
```

## Observações

- O módulo realiza conexão automática com o site do SharePoint "https://tcepi365.sharepoint.com/sites/SecretariadeControleExterno"
- As consultas são realizadas principalmente na lista "Cadastro de Ação de Controle"
- O processo de transformação de dados é robusto, tratando diversos casos especiais e formatos de dados
- Métodos com prefixo `_` são considerados internos e não devem ser chamados diretamente pelos consumidores do módulo