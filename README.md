# FastPlac

> Gerador automatizado de placas técnicas integrado ao Google Drive e Google Slides.

---

## Sobre o projeto

O **FastPlac** é uma aplicação web interna desenvolvida em Streamlit para automatizar a geração de placas técnicas de estruturas de armazenagem. A ferramenta substitui o processo manual em planilhas, oferecendo uma interface limpa onde o usuário seleciona o tipo de placa, preenche apenas os campos relevantes e gera um PDF consolidado — que é salvo automaticamente na pasta de concluídos no Google Drive.

---

## Funcionalidades

- Formulário dinâmico por tipo de placa — só aparecem os campos do modelo selecionado
- Suporte a múltiplos tipos de placa em uma única sessão
- Geração de PDF consolidado com todas as placas do pedido
- Salvamento automático na pasta de concluídos no Google Drive
- Link direto para o arquivo gerado
- Nome do arquivo padronizado automaticamente (`Placas - CLIENTE (N° PEDIDO)`)

---

## Tipos de placa suportados

| Tipo | Descrição |
|------|-----------|
| Placa Porta Paletes | Estrutura porta paletes convencional |
| Placa Flow Rack | Estrutura de fluxo por gravidade |
| Placa Drive In | Estrutura drive-in |
| Placa Mezanino | Mezanino estrutural |
| Placa Mezanino Carga Piso | Mezanino com carga por piso |
| Placa Dinâmico | Estrutura dinâmica |
| Placa Mercado Livre | Padrão Mercado Livre |
| Placa PP Palete e Plano | Push-back palete e plano |
| Placa PP Peso Plano | Push-back peso por plano |
| Placa PP Peso por Nível | Push-back peso por nível |
| Placa Dinâmico Peso Máximo | Dinâmico com peso máximo |

---

## Estrutura do projeto

```
FASTPLAC/
├── app.py                  # Interface principal (Streamlit)
├── requirements.txt        # Dependências Python
├── README.md
├── .gitignore
│
├── config/
│   └── field_config.py     # Campos por tipo de placa e IDs dos templates
│
├── services/
│   └── google_api.py       # Integração com Google Drive e Slides API
│
├── images/
│   └── logo_aguia1.png     # Logo da aplicação
│
└── secrets/                # ⚠️ Ignorado pelo Git — nunca versionar
    ├── credentials.json    # Service account (legado)
    ├── client_secret.json  # OAuth2 client secret
    └── token.json          # Token OAuth2 gerado no primeiro login
```

---

## Pré-requisitos

- Python 3.10+
- Conta Google com acesso ao Google Drive e Google Slides
- Projeto no [Google Cloud Console](https://console.cloud.google.com) com as APIs ativadas:
  - Google Drive API
  - Google Slides API

---

## Instalação local

**1. Clone o repositório**
```bash
git clone https://github.com/taynapuppo/fastplac.git
cd fastplac
```

**2. Instale as dependências**
```bash
pip install -r requirements.txt
```

**3. Configure as credenciais**

Crie a pasta `secrets/` na raiz do projeto e adicione os arquivos:
- `client_secret.json` — baixado do Google Cloud Console (OAuth2 → Aplicativo de computador)
- `token.json` — gerado automaticamente no primeiro login

Na primeira execução, o navegador abrirá para autenticação com sua conta Google. O `token.json` será salvo automaticamente em `secrets/`.

**4. Rode a aplicação**
```bash
streamlit run app.py
```

---

## Deploy no Streamlit Cloud

**1.** Faça o push do repositório para o GitHub (o `.gitignore` já protege os arquivos sensíveis)

**2.** Acesse [share.streamlit.io](https://share.streamlit.io) → **New app** → selecione o repositório

**3.** Em **Settings → Secrets**, adicione:

```toml
[google]
credentials    = '{ conteúdo do credentials.json em uma linha }'
client_secret  = '{ conteúdo do client_secret.json em uma linha }'
token          = '{ conteúdo do token.json em uma linha }'
```

> ⚠️ O `token.json` deve ser gerado localmente antes do deploy. Rode a aplicação uma vez na sua máquina para criá-lo.

---

## Configuração dos templates

Os templates de Slides e os campos de cada tipo de placa são configurados em `config/field_config.py`.

Para adicionar um novo tipo de placa:

1. Adicione o ID do template no dicionário `TEMPLATE_IDS`
2. Adicione os campos específicos no dicionário `CAMPOS_ESPECIFICOS`
3. Os campos comuns (Cliente, N° do Pedido, N° do Projeto, Qtd. de Placas) são herdados automaticamente

---

## Tecnologias utilizadas

| Tecnologia | Uso |
|------------|-----|
| [Streamlit](https://streamlit.io) | Interface web |
| [Google Slides API](https://developers.google.com/slides) | Preenchimento dos templates |
| [Google Drive API](https://developers.google.com/drive) | Upload e gestão dos arquivos |
| [pypdf](https://pypdf.readthedocs.io) | Merge dos PDFs gerados |
| [google-auth-oauthlib](https://google-auth-oauthlib.readthedocs.io) | Autenticação OAuth2 |

---

## Observações

- A API do Google Slides tem limite de **60 requisições de escrita por minuto**. O sistema aplica retry automático com backoff exponencial em caso de erro 429.
- Os arquivos temporários criados durante a geração são **deletados permanentemente** do Drive logo após a exportação do PDF.
- O nome do cliente é sempre salvo em **letras maiúsculas** para manter o padrão.

---

*Desenvolvido para uso interno — Águia Sistemas de Armazenagem*