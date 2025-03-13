# Documentação do módulo Report Generator

## Visão Geral

O módulo `report_generator.py` fornece funcionalidade para gerar relatórios DOCX usando modelos. Ele é especializado na criação de documentos, com imagens de capa opcionais e conteúdo personalizado com base nos dados de contexto fornecidos.

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

##### `add_cover_page`

```python
def add_cover_page(self, doc: DocxTemplate, image_path: Union[str, Path]) -> bool
```

Adiciona uma imagem de capa em página inteira ao documento.

**Parâmetros:**
- `doc`: instância DocxTemplate
- `image_path`: Caminho para a imagem de capa

**Retorna:**
- `bool`: True se a imagem tiver sido adicionada com sucesso, False se não

* *Funcionalidades:**
- Valida se o arquivo de imagem existe.
- Configura a primeira seção do documento para remover margens.
- Adiciona a imagem no início do documento em tamanho de página inteira.
- Cria uma nova seção com margens normais para o conteúdo do template.
- Registra erros caso o processo falhe.

##### `generate_report`

```python
def generate_report(self, context: dict, output_path: str, cover_image_path: Optional[Union[str, Path]] = None, graau_params: dict = {}) -> bool
```

Gera o relatório utilizando o template e o contexto fornecidos.

**Parâmetros:**
- `context`: Dicionário com as variáveis de contexto do template
- `output_path`: Caminho para salvar o arquivo de saída
- `cover_image_path`: Caminho opcional para a imagem de capa (padrão: None)
- `graau_params`: Dicionário com parâmetros adicionais do relatório (padrão: dicionário vazio)

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
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)

# Sample data that would come from SharePoint
sharepoint_data = {
    "unidades_fiscalizadas": "P. M. CIDADE",
    "n_processo_eTCE": "TC/XXXXXX/20XX",
    "n_processo_eTCE_processo_tipo": "CONTAS-TOMADA DE CONTAS ESPECIAL",
    "exercicios": "20XX, 20YY",
    "VRF": "R$ 100.000,00"
}

# Create the report generator with the template path
generator = ReportGenerator("templates/Relatório Padrão - GRAAU.docx")

# Generate the report
success = generator.generate_report(
    context=sharepoint_data,
    output_path="reports/final_report.docx",
    cover_image_path="cover_images/cover_page.jpg"
)

if success:
    print("Report generated successfully!")
else:
    print("Failed to generate report")
```

## Notes

- O módulo usa a lib `docxtpl` que permite templates estilo Jinja2-like em arquivos DOCX
- O template deve ter variáveis placeholders que correspondam às chaves no dicionário de contexto
- Imagens de capa são opcionais