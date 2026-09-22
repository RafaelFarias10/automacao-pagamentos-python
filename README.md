# 📊 Automação de Registro de Pagamentos

Projeto em Python desenvolvido para automatizar o preenchimento de uma planilha de registro de pagamentos a partir de dados exportados de um sistema ERP.

## 🚀 Funcionalidades

- Leitura de relatórios exportados do ERP
- Tratamento e limpeza de arquivos XML
- Identificação de pedidos de compra (PO)
- Remoção de registros duplicados
- Preenchimento automático de informações na planilha
- Preservação de dados já existentes
- Identificação de POs com múltiplos registros para revisão manual
- Geração de uma nova planilha com os dados processados

## 🛡️ Regras de segurança

A automação utiliza uma abordagem conservadora:

- Nunca sobrescreve células já preenchidas
- Apenas POs com um único registro válido são preenchidos automaticamente
- POs não encontrados são mantidos para revisão
- POs com múltiplos registros não são preenchidos automaticamente

## 🛠️ Tecnologias

- Python
- OpenPyXL
- XML (ElementTree)
- Regex

## 📁 Estrutura

```text
automacao-protheus-pagamentos/
│
├── automacao_completa.py
├── requirements.txt
├── README.md
├── .gitignore
│
├── entrada/   # arquivos locais (não incluídos no repositório)
└── saida/     # arquivos gerados (não incluídos no repositório)
```

## ⚙️ Instalação

Instale a dependência:

```bash
pip install -r requirements.txt
```

## ▶️ Execução

Crie as pastas `entrada` e `saida`.

Adicione na pasta `entrada`:

- o relatório exportado do ERP;
- a planilha que será preenchida.

Depois execute:

```bash
python automacao_completa.py
```

O arquivo processado será criado na pasta `saida`.

## 🔒 Privacidade

Arquivos e dados corporativos utilizados durante o desenvolvimento não estão incluídos neste repositório.

O projeto publicado contém apenas o código-fonte da automação.