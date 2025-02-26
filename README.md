# GRAAU - backend

Uma biblioteca Python para geração automatizada de relatórios em formato DOCX com suporte a imagens de capa.

## Características

- Geração de relatórios em formato DOCX
- Suporte a imagem de capa em página inteira
- Estrutura hierárquica de seções e subseções

## Instalação

1. Crie um ambiente virtual e o ative

```bash
virtualenv venv
source venv/bin/activate
```

2. Instale as dependências necessárias:

```bash
pip install -r requirements.txt
```

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