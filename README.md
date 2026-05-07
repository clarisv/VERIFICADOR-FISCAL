# Robô de Conferência Fiscal v1.0

![Python](https://img.shields.io/badge/python-3.10+-blue.svg)
![DevSecOps](https://img.shields.io/badge/Security-Audited-green.svg)
![RPA](https://img.shields.io/badge/RPA-Automated-red.svg)

## Visão Geral
Solução de automação robótica (RPA) desenvolvida para a reconciliação e validação de documentos fiscais. O sistema realiza o processamento de guias tributárias em formato PDF (ISS, ICMS e Importação), executando o cruzamento de dados com livros de apuração em Excel para identificar divergências de valores, competências e Unidades Federativas (UF) de destino.

## Funcionalidades Principais
* **OCR & Extração de Dados:** Processamento de PDFs complexos para captura de CNPJ, IE, valores monetários e códigos fiscais.
* **Validação por Estilo:** Motor de busca inteligente que adapta a extração conforme o layout da guia (Estilos 1, 2 ou 3).
* **Reconciliação Inteligente:** Algoritmo de busca que prioriza células marcadas com formatação condicional (pendências) na planilha de destino.
* **Agrupamento Lógico:** Consolidação automática de guias parceladas ou matriciais para conferência de valores totais.
* **Dashboard Visual:** Interface moderna desenvolvida em `CustomTkinter` para operação simplificada e acompanhamento em tempo real via logs.

## Stack Tecnológica
* **Linguagem:** Python 3.10+
* **Extração de PDF:** `pdfplumber` (análise de texto estruturado)
* **Manipulação de Dados:** `openpyxl` (integração avançada com Excel/XML)
* **Interface Gráfica (GUI):** `CustomTkinter` (UI/UX moderna e responsiva)
* **Segurança:** Implementação de auditoria interna para prevenção de vazamento de credenciais e limpeza de dados temporários.

## Práticas de DevSecOps Aplicadas
1.  **Sanitização de Caminhos:** Utilização de diretórios temporários (`tempfile`) para processamento de dados sensíveis, garantindo a deleção completa após a execução.
2.  **Segurança de Variáveis:** Arquitetura livre de *hardcoded credentials*, utilizando passagem de parâmetros e variáveis de ambiente.
3.  **Modularização:** Separação estrita entre o motor de processamento (`robo.py`) e a interface de usuário (`visual.py`) para facilitar a manutenção e auditoria de código.

## 📂 Estrutura do Projeto
```text
├── robo.py          # core: motor rpa e lógica de conferência fiscal
├── visual.py        # ui: dashboard e gestão de estados da interface
├── requirements.txt # dependências do projeto
└── README.md        # documentação técnica
