# GRAAU - backend

API Flask para geração automatizada de relatórios em formato DOCX com suporte a imagens de capa.

## Características

- Geração de relatórios em formato DOCX
- Suporte a imagem de capa em página inteira

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

## Limitações

- Suporta apenas formato DOCX
- Imagens de capa devem ter proporções compatíveis com o tamanho da página

## Estrutura do arquivo .env

```
- USUARIO="user.name@tce.pi.gov.br"
- SENHA="suaSenha"
```

## Documentação

Para a documentação completa, incluindo todos os endpoints, parâmetros e exemplos de uso, consulte a [Documentação do projeto](./docs).