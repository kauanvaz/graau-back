# Documentação do módulo ReportGenerator

## Visão Geral

O módulo `report_generator.py` fornece funcionalidade para gerar relatórios DOCX usando templates. Ele é especializado na criação de documentos, com imagens de capa opcionais, títulos e subtítulos organizados hierarquicamente e conteúdo personalizado com base nos dados de contexto fornecidos.

## Componentes Principais

### Classes

#### `ReportGenerator`

A classe principal responsável por gerar relatórios DOCX a partir de templates.

**Construtor:**

```python
def __init__(self, template_path: str)
```
- `template_path`: Caminho para o arquivo DOCX template

**Métodos:**

##### `_insert_headings_recursively`

```python
def _insert_headings_recursively(self, doc, headings: list, index: int, level: int=1)
```

Insere os títulos e subtítulos no documento de forma recursiva.

**Parâmetros:**
- `doc`: O documento DOCX onde os títulos serão inseridos
- `headings`: Lista de dicionários com os títulos e subtítulos
- `index`: Índice onde o título será inserido
- `level`: Nível do título (1 para Heading 1, 2 para Heading 2, etc.)

**Retorna:**
- `int`: O índice atualizado após a inserção

**Funcionalidades:**
- Adiciona quebras de página antes de cada título de nível 1
- Insere os títulos com o estilo apropriado (Heading 1, Heading 2, etc.)
- Chama recursivamente para inserir subtítulos

##### `_get_signing_area_name`

```python
def _get_signing_area_name(self, headings: list) -> str
```

Determina qual seção deve conter a área de assinaturas.

**Parâmetros:**
- `headings`: Lista de dicionários com os títulos e subtítulos

**Retorna:**
- `str`: Nome da seção onde a área de assinaturas deve ser inserida, ou None se não for encontrada

**Funcionalidades:**
- Verifica se existe uma seção "proposta de encaminhamentos" ou "conclusão"
- Retorna o nome da primeira seção encontrada

##### `_add_content`

```python
def _add_content(self, doc, text=None, bold=False, color=None, alignment=None, font='Segoe UI', space_after=0)
```

Adiciona conteúdo de texto ao documento com formatação específica.

**Parâmetros:**
- `doc`: O documento DOCX onde o conteúdo será adicionado
- `text`: Texto a ser adicionado
- `bold`: Se o texto deve estar em negrito
- `color`: Cor do texto (RGBColor)
- `alignment`: Alinhamento do parágrafo
- `font`: Fonte do texto
- `space_after`: Espaço após o parágrafo em Pt

**Funcionalidades:**
- Adiciona um parágrafo com o texto e formatação especificados

##### `_add_signing_content`

```python
def _add_signing_content(self, doc)
```

Adiciona a área de assinaturas ao documento.

**Parâmetros:**
- `doc`: O documento DOCX onde a área de assinaturas será adicionada

**Funcionalidades:**
- Adiciona espaços em branco
- Adiciona seção de instrução com campo para auditores signatários
- Adiciona seção de supervisão com campos para nome e cargo
- Adiciona seção de visto com campos para nome e cargo

##### `generate_headings_from_structure`

```python
def generate_headings_from_structure(self, doc, headings: list)
```

Gera os tópicos a partir da estrutura fornecida.

**Parâmetros:**
- `doc`: O documento DOCX onde os tópicos serão gerados
- `headings`: Lista de dicionários com os títulos e subtítulos

**Funcionalidades:**
- Localiza o marcador `<CONTEUDO>` no documento
- Substitui o marcador pelo primeiro título
- Insere os subtítulos e títulos subsequentes
- Adiciona a área de assinaturas após a seção apropriada

##### `replace_existing_image`

```python
def replace_existing_image(self, docx_path: str, target_image_filename: str, new_image_path: Union[str, Path]) -> bool
```

Substitui uma imagem de capa existente.

**Parâmetros:**
- `docx_path`: Caminho para o arquivo DOCX
- `target_image_filename`: Nome do arquivo de imagem a ser substituído
- `new_image_path`: Caminho do novo arquivo de imagem

**Retorna:**
- `bool`: True se a imagem tiver sido substituída com sucesso, False se não

**Funcionalidades:**
- Transforma o arquivo DOCX em arquivo ZIP para acesso aos arquivos internos.
- Localiza a imagem a ser substituída.
- Substitui a imagem original pela nova.
- Reempacota o DOCX

##### `generate_report`

```python
def generate_report(self, context: dict, output_path: str, cover_image_path: Optional[Union[str, Path]] = None, target_image_filename: Optional[str] = "image1.png") -> bool
```

Gera o relatório utilizando o template e o contexto fornecidos.

**Parâmetros:**
- `context`: Dicionário com as variáveis de contexto do template
- `output_path`: Caminho para salvar o arquivo de saída
- `cover_image_path`: Caminho opcional para a imagem de capa (padrão: None)
- `target_image_filename`: Nome do arquivo de imagem a ser substituído (padrão: "image1.png")

**Retorna:**
- `bool`: True se o relatório tiver sido gerado com sucesso, False se não

**Funcionalidades:**
- Cria uma nova instância de DocxTemplate para cada relatório
- Extrai dados hierárquicos (seções e elementos textuais) do contexto
- Gera a estrutura de títulos e subtítulos
- Adiciona a imagem de capa se um caminho de imagem for fornecido
- Renderiza o modelo com o contexto fornecido
- Salva o documento no caminho de saída especificado
- Registra o resultado da operação

## Dependências

- `typing`: Para anotações de tipo
- `docx.shared`: Para medidas de documentos (Pt, RGBColor)
- `docx.enum.text`: Para constantes de texto (WD_BREAK, WD_ALIGN_PARAGRAPH)
- `docxtpl`:  Para geração de documentos baseada em templates
- `pathlib`: Para manipulação de caminhos de arquivos
- `logging`: Para operações de logging
- `zipfile`: Para manipulação de arquivos ZIP (DOCX internamente)
- `os`: Para operações de sistema de arquivos
- `shutil`: Para operações de cópia de arquivos
- `tempfile`: Para criação de diretórios temporários

## Estrutura do Contexto

A estrutura esperada do contexto para o relatório é:

```python
{
    "seccoes": [
        {
            "title": "Elementos pré-textuais",
            "data": [
                {
                    "title": "...",
                    "subtitles": [
                        {
                            "title": "...",
                            "subtitles": [...]
                        }
                    ]
                },
                ...
            ]
        },
        {
            "title": "Elementos textuais",
            "data": [
                {
                    "title": "EXEMPLO 1",
                    "subtitles": [
                        {
                            "title": "Exemplo 1.1",
                            "subtitles": [...]
                        }
                    ]
                },
                ...
            ]
        },
        {
            "title": "Elementos pós-textuais",
            "data": [
                {
                    "title": "...",
                    "subtitles": [
                        {
                            "title": "...",
                            "subtitles": [...]
                        }
                    ]
                },
                ...
            ]
        }
    ],
    # Outras variáveis de contexto para o template
    "unidades_fiscalizadas": "P. M. CIDADE",
    "n_processo_eTCE": "TC/XXXXXX/20XX",
    ...
}
```

## Exemplo de Uso

```python
from src.report_generator import ReportGenerator
from utils import load_json

context = {
    "unidades_fiscalizadas": "P. M. CIDADE",
    "n_processo_eTCE": "TC/XXXXXX/20XX",
    "n_processo_eTCE_processo_tipo": "CONTAS-TOMADA DE CONTAS ESPECIAL",
    "exercicios": "20XX, 20YY",
    "VRF": "R$ 100.000,00"
}
# Carregar dados de seções do arquivo JSON de exemplo
context["seccoes"] = load_json("examples/sections.json")

generator = ReportGenerator("src/templates/Relatório Padrão - GRAAU.docx")

success = generator.generate_report(
    context=context,
    output_path="src/reports/report_example.docx",
    cover_image_path="src/cover_images/cover_page_2.jpg",
)
```

## Observações

- O módulo usa a lib `docxtpl` que permite templates estilo Jinja2-like em arquivos DOCX
- O template deve ter um marcador `<CONTEUDO>` a partir de onde a estrutura de títulos será inserida
- O template deve ter variáveis placeholders que correspondam às chaves no dicionário de contexto
- Imagens de capa são opcionais
- A área de assinaturas é inserida automaticamente após a seção "proposta de encaminhamentos" ou "conclusão"