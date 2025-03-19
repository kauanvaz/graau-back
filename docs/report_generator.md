# Documentação do módulo Report Generator

## Visão Geral

O módulo `report_generator.py` fornece funcionalidade para gerar relatórios DOCX usando templates. Ele é especializado na criação de documentos, com imagens de capa opcionais e conteúdo personalizado com base nos dados de contexto fornecidos.

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

* *Funcionalidades:**
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
- Adiciona a página de capa se um caminho de imagem for fornecido
- Renderiza o modelo com o contexto fornecido
- Salva o documento no caminho de saída especificado
- Registra o resultado da operação

## Dependências

- `typing`: Para anotações de tipo
- `docx.shared`: Para medidas de documentos (Pt, Inches)
- `docxtpl`:  Para geração de documentos baseada em templates
- `pathlib`: Para manipulação de caminhos de arquivos
- `logging`: Para operações de logging

## Exemplo de Uso

```python
from src.report_generator import ReportGenerator

sharepoint_data = [
    {
        "unidades_fiscalizadas": "P. M. CIDADE",
        "n_processo_eTCE": "TC/XXXXXX/20XX",
        "n_processo_eTCE_processo_tipo": "CONTAS-TOMADA DE CONTAS ESPECIAL",
        "exercicios": "20XX, 20YY",
        "VRF": "R$ 100.000,00"
    }
]

generator = ReportGenerator("src/templates/Relatório Padrão - GRAAU.docx")

success = generator.generate_report(
    context=sharepoint_data[0],
    output_path="src/reports/report_example.docx",
    cover_image_path="src/cover_images/cover_page_2.jpg",
)
```

## Observações

- O módulo usa a lib `docxtpl` que permite templates estilo Jinja2-like em arquivos DOCX
- O template deve ter variáveis placeholders que correspondam às chaves no dicionário de contexto
- Imagens de capa são opcionais