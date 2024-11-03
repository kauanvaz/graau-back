# Report Generator

Uma biblioteca Python para geração automatizada de relatórios em formato DOCX com suporte a estilos personalizados, imagens de capa em página inteira e estrutura hierárquica de conteúdo.

## Características

- Geração de relatórios em formato DOCX
- Sistema de estilos personalizáveis
- Suporte a imagem de capa em página inteira
- Estrutura hierárquica de seções e subseções
- Suporte a tabelas
- Sistema de logging integrado
- Tipagem estática com Type Hints

## Requisitos

- Python 3.6+
- python-docx
- Typing (para suporte a type hints)

## Instalação

1. Instale as dependências necessárias:

```bash
pip install python-docx
```

2. Copie os arquivos do projeto para seu ambiente de desenvolvimento.

## Uso Básico

### 1. Importando a classe

```python
from report_generator import ReportGenerator
```

### 2. Criando um template básico

```python
template = {
    "title": "Título do Relatório",
    "sections": [
        {
            "title": "Primeira Seção",
            "text": "Conteúdo da primeira seção...",
            "subsections": [
                {
                    "title": "Subseção",
                    "text": "Conteúdo da subseção..."
                }
            ]
        }
    ]
}
```

### 3. Gerando o relatório

```python
generator = ReportGenerator()
success = generator.generate_report(
    template=template,
    output_path="relatorio.docx",
    cover_image_path="caminho/para/imagem.jpg"  # Opcional
)
```

## Estrutura do Template

O template é um dicionário Python com a seguinte estrutura:

```python
{
    "title": str,                    # Título do relatório
    "sections": [                    # Lista de seções
        {
            "title": str,            # Título da seção
            "text": str,             # Texto da seção (opcional)
            "table": List[List],     # Tabela (opcional)
            "subsections": [         # Lista de subseções (opcional)
                {
                    "title": str,
                    "text": str,
                    # ... pode conter os mesmos elementos que uma seção
                }
            ]
        }
    ]
}
```

## Personalização de Estilos

### Estilos Padrão

O sistema possui quatro estilos predefinidos:

- `title`: Título principal (16pt, negrito, centralizado)
- `heading1`: Títulos de seção (14pt, negrito, azul)
- `heading2`: Títulos de subseção (12pt, negrito, azul)
- `normal`: Texto normal (11pt)

### Personalizando Estilos

Você pode criar estilos personalizados ao instanciar o ReportGenerator:

```python
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

custom_styles = {
    'custom_style': Style(
        size=13,
        bold=True,
        color=RGBColor(255, 0, 0),
        alignment=WD_ALIGN_PARAGRAPH.LEFT
    )
}

generator = ReportGenerator(custom_styles=custom_styles)
```

## Exemplos

### Relatório com Tabela

```python
template = {
    "title": "Relatório de Vendas",
    "sections": [
        {
            "title": "Resumo de Vendas",
            "text": "Análise do período...",
            "table": [
                ["Produto", "Quantidade", "Valor"],
                ["Produto A", "100", "R$ 1.000"],
                ["Produto B", "150", "R$ 2.250"]
            ]
        }
    ]
}
```

### Relatório com Imagem de Capa

```python
generator = ReportGenerator()
success = generator.generate_report(
    template=template,
    output_path="relatorio_vendas.docx",
    cover_image_path="logo_empresa.jpg"
)
```

## Tratamento de Erros

O sistema inclui logging integrado para facilitar a depuração:

```python
import logging

logging.basicConfig(level=logging.INFO)
generator = ReportGenerator()
```

Todos os erros são registrados e o método `generate_report` retorna `False` em caso de falha.

## Limitações

- Suporta apenas formato DOCX
- Imagens de capa devem ter proporções compatíveis com o tamanho da página
- Tabelas seguem um estilo fixo ('Table Grid')

## Contribuindo

Sinta-se à vontade para abrir issues ou enviar pull requests com melhorias.

## Licença

Este projeto está sob a licença MIT.