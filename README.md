# Report Generator

Uma biblioteca Python para geração automatizada de relatórios em formato DOCX com suporte a imagens de capa.

## Características

- Geração de relatórios em formato DOCX
- Suporte a imagem de capa em página inteira
- Estrutura hierárquica de seções e subseções
- Sistema de logging integrado
- Tipagem estática com Type Hints

## Instalação

1. Instale as dependências necessárias:

```bash
pip install python-docx
```

2. Copie os arquivos do projeto para seu ambiente de desenvolvimento.

## Uso Básico

### Relatório com Imagem de Capa

```python
generator = ReportGenerator()
success = generator.generate_report(
    template=template,
    output_path="relatorio_vendas.docx",
    cover_image_path="logo_empresa.jpg"
)
```

## Limitações

- Suporta apenas formato DOCX
- Imagens de capa devem ter proporções compatíveis com o tamanho da página

## Estrutura do arquivo .env

```
- USUARIO="user.name@tce.pi.gov.br"
- SENHA="suaSenha"
```