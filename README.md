# ⏱️ Automação de Ponto e Gestão de RH (Microsoft 365 Integration)

Este projeto é um ecossistema completo de automação para Recursos Humanos. Ele extrai batidas de ponto do **Microsoft Teams**, cruza com regras de contratos dinâmicas, calcula horas extras, detecta faltas, cruza dados com calendários de feriados e aprovações de férias, e gera um relatório final em Excel que é enviado automaticamente para o **SharePoint**.

## 🚀 Arquitetura do Sistema

O fluxo funciona através de uma arquitetura modularizada utilizando Python e a **Microsoft Graph API**:

1. **Extração:** Coleta batidas de ponto reais via Microsoft Teams (Shifts/Turnos).
2. **Validação de Regras:** Consulta listas no SharePoint para descobrir a carga horária de cada funcionário e as tolerâncias de ponto.
3. **Fluxos de Aprovação (Power Automate):** - Solicitações de Horas Extras e Férias são feitas via **Microsoft Forms**.
   - Fluxos do Automate direcionam aprovações (via Teams/Outlook) para gestores ou RH.
   - Decisões são salvas em listas do SharePoint.
4. **Motor de Cálculo (Python):** - Calcula Horas Brutas, Horas Reais, Descontos de Almoço e Horas Extras (com 10 minutos de tolerância).
   - Detecta dias sem ponto e classifica inteligentemente como: `FÉRIAS` (se aprovado pelo RH), `RECESSO` (se houver feriado na Intranet) ou `FALTA`.
5. **Carga e Upload:** Gera um arquivo `.xlsx` limpo e faz o upload automático via Graph API direto para a pasta do RH no SharePoint.

---

## 📋 Pré-requisitos e Configuração

Para rodar este projeto localmente, você precisará de:
* Python 3.10+
* Conta de Desenvolvedor Microsoft / Acesso de Admin (Azure AD) para registrar uma Aplicação.
* Permissões na Graph API (Teams, SharePoint, Files).

### 1. Variáveis de Ambiente (`.env`)
Crie um arquivo `.env` na raiz do projeto com as credenciais do seu app no Azure:
```env
CLIENT_ID=seu_client_id_aqui
CLIENT_SECRET=seu_client_secret_aqui
TENANT_ID=seu_tenant_id_aqui
TEAM_ID=id_da_equipe_do_teams_aqui
SHAREPOINT_HOSTNAME=nome_empresa.sharepoint.com
SHAREPOINT_SITE_PATH=/sites/RH-Intranet
```
### 2. Instalação das Bibliotecas
```Bash
pip install requests pandas msal python-dotenv openpyxl
```
## 🗄️ Estrutura do Banco de Dados (Listas do SharePoint)
Para que o motor funcione perfeitamente, o ambiente Microsoft 365 deve possuir as seguintes listas configuradas no SharePoint do RH:

1. Contratos Colaboradores (Regras Base)
Title: E-mail do funcionário

Carga: Texto (Ex: 08:00)

Almoco: Texto (Ex: 01:00)

Setor: Texto (Ex: TI)

2. Gestores Setores (Hierarquia de Aprovação)
Setor: Texto (Ex: TI)

Email do Gestor: Texto (Ex: chefe.ti@empresa.com.br)

3. Aprovacoes Hora Extra (Alimentada via Automate)
Title: E-mail do funcionário

Data: Data (Ex: 2026-03-12)

Status: Texto (Aprovado ou Rejeitado)

4. Aprovações Férias (Alimentada via Automate)
Email: Texto (E-mail do funcionário)

Data de Início: Data

Data de Fim: Data

Status: Opção / Choice (Aprovado ou Rejeitado)

5. Eventos (Calendário da Intranet Principal)
A busca de feriados varre automaticamente o calendário raiz da empresa. O robô concede "RECESSO" se encontrar as palavras Feriado, Recesso ou Folga no Título ou na Categoria do evento.

## ⚙️ Power Automate: Resumo dos Fluxos
Este projeto depende de dois fluxos de Nuvem Automatizados (Forms -> Approvals -> SharePoint):

Solicitação de Hora Extra: O funcionário preenche o dia e o motivo. O Automate lê a lista Contratos para descobrir o setor, lê a lista Gestores para descobrir o chefe, envia o card de aprovação para o chefe correto e salva a decisão no SharePoint.

Solicitação de Férias: O funcionário preenche as datas de início e fim. O Automate envia o card de aprovação para o RH (sobrescrevendo o remetente nativo) e salva o período na lista de férias.

```
📂 Estrutura de Arquivos do Projeto
Plaintext
├── main.py               # Arquivo principal (Maestro do fluxo ETL e Upload)
├── integracoes.py        # Todas as chamadas para a Microsoft Graph API
├── motor_calculo.py      # Lógica matemática (horas, tolerância de 10 min, detecção de faltas)
├── .env                  # Variáveis de ambiente (oculto no Git)
├── .gitignore            # Ignora o .env, cache e arquivos locais do Excel
└── README.md             # Esta documentação
```

## ▶️ Como Executar
Com todas as dependências instaladas e o .env configurado, basta rodar o comando:

```Bash
python main.py
```
Resultado esperado:
O terminal informará os relatórios lidos, a classificação das datas (BINGO! de Férias/Feriados) e finalizará com a mensagem de SUCESSO ABSOLUTO confirmando que a planilha Relatorio_Ponto_ValeShop.xlsx foi gerada e enviada ao SharePoint.

***
