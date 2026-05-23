# 📊 Aging Automation

> Pipeline de automação completo para extração, tratamento e geração de relatórios de Aging a partir do Power BI — reduzindo um processo de 30 a 60 minutos diários para execução com um único comando.

---

## 🧩 Contexto

Em ambientes hospitalares, o relatório de **Aging de Faturamento** é um instrumento crítico de gestão: ele mapeia o tempo que as contas médicas estão em cada etapa do ciclo de faturamento, permitindo identificar gargalos, cobranças em atraso e prioridades de ação.

O processo manual envolvia:
1. Acessar o dashboard no Power BI
2. Aplicar filtros de data e convênio
3. Exportar os dados manualmente
4. Abrir o modelo Excel e inserir os dados
5. Reaplicar fórmulas, formatações e atualizar tabelas dinâmicas

**Esse fluxo consumia entre 30 minutos e 1 hora por dia.** Este projeto automatiza todas essas etapas em um pipeline executável com um único comando.

---

## ⚙️ Funcionalidades

- 🌐 **Acesso automatizado ao Power BI** via Playwright (Microsoft Edge com perfil persistente)
- 📅 **Aplicação automática de filtros** de data (D-1) e convênio
- 📥 **Download programático** do arquivo exportado
- 🧹 **Tratamento de dados** com Pandas e Openpyxl (inserção de dados, fórmulas, formatação)
- 📊 **Atualização automática de Tabelas Dinâmicas** via COM Automation (pywin32)
- 📝 **Registro de execução** completo com logging estruturado
- 💾 **Arquivo final nomeado por data** e salvo automaticamente na pasta de resultados
- 🔒 **Configurações sensíveis** isoladas em variáveis de ambiente (.env)

---

## 🏗️ Arquitetura do Pipeline

```
main.py
  │
  ├── src/browser.py       # Automação do navegador (Playwright)
  ├── src/processing.py    # Tratamento de dados (Pandas + Openpyxl)
  ├── src/pivot.py         # Atualização de Tabelas Dinâmicas (pywin32)
  └── utils/dates.py       # Utilitários de data
```

**Fluxo de execução:**

```
[Power BI Dashboard]
        │
        ▼
  src/browser.py    ──►  data/modelo/export.xlsx
        │
        ▼
  src/processing.py ──►  data/resultado/{data}_Aging.xlsx
        │
        ▼
  src/pivot.py      ──►  Tabelas Dinâmicas atualizadas
```

---

## 🛠️ Tecnologias Utilizadas

| Tecnologia | Uso |
|---|---|
| `playwright` | Automação do navegador Edge |
| `pandas` | Leitura e manipulação dos dados exportados |
| `openpyxl` | Escrita, formatação e fórmulas no Excel |
| `pywin32` | Atualização de Tabelas Dinâmicas via COM |
| `python-dotenv` | Gerenciamento de variáveis de ambiente |
| `logging` | Registro estruturado de execução |
| `datetime` | Cálculo e formatação de datas |

---

## 📁 Estrutura do Projeto

```
aging-automation/
│
├── main.py                        # Orquestrador do pipeline
│
├── src/
│   ├── browser.py                 # Automação do navegador
│   ├── processing.py              # Tratamento de dados
│   └── pivot.py                   # Atualização de tabelas dinâmicas
│
├── utils/
│   └── dates.py                   # Utilitários de data
│
├── data/
│   ├── modelo/                    # Arquivos de entrada (não versionados)
│   │   ├── export.xlsx            # Exportado do BI automaticamente
│   │   └── Modelo Aging.xlsx      # Modelo base do relatório
│   ├── resultado/                 # Relatórios finais por data (não versionados)
│   └── logs/
│       └── historico.log          # Log de todas as execuções
│
├── scripts/
│   └── run.bat                    # Atalho para execução no Windows
│
├── .env                           # Variáveis de ambiente (não versionado)
├── .env.example                   # Modelo de configuração
├── .gitignore
├── requirements.txt
└── README.md
```

> ⚠️ Os arquivos em `data/modelo/` e `data/resultado/` não são versionados por conterem dados sensíveis de faturamento hospitalar.

---

## 🚀 Como Executar

### Pré-requisitos

- Python 3.10+
- Microsoft Edge instalado
- Acesso ao Power BI com as permissões necessárias

### Instalação

```bash
# 1. Clone o repositório
git clone https://github.com/seu-usuario/aging-automation.git
cd aging-automation

# 2. Crie e ative o ambiente virtual
python -m venv .venv
.venv\Scripts\activate

# 3. Instale as dependências
pip install -r requirements.txt

# 4. Instale o browser do Playwright
playwright install msedge

# 5. Configure as variáveis de ambiente
cp .env.example .env
# Edite o .env com suas configurações
```

### Configuração

Crie um arquivo `.env` baseado no `.env.example`:

```env
# URL do relatório no Power BI
PBI_REPORT_URL=https://app.powerbi.com/groups/SEU_GROUP_ID/reports/SEU_REPORT_ID/...

# Caminho para o perfil do Edge usado pelo Playwright
EDGE_PROFILE_PATH=C:\playwright-edge-profile
```

### Execução

```bash
python main.py
```

Ou via arquivo `.bat`:

```bash
scripts\run.bat
```

---

## 📋 Exemplo de Log

```
2026-05-17 08:00:01 - root - INFO - Extraindo dados do Power BI...
2026-05-17 08:01:43 - root - INFO - Extração concluída!
2026-05-17 08:01:43 - root - INFO - Começando o tratamento de dados..
2026-05-17 08:01:51 - root - INFO - Os arquivos foram tratados com sucesso!
2026-05-17 08:01:52 - root - INFO - Atualizando as tabelas dinâmicas..
2026-05-17 08:02:10 - root - INFO - Pipeline finalizado com sucesso!
```

---

## 🔒 LGPD e Dados Sensíveis

Este projeto foi desenvolvido em ambiente hospitalar e lida com dados de faturamento médico. As seguintes medidas foram adotadas para conformidade com a **Lei Geral de Proteção de Dados (LGPD — Lei nº 13.709/2018)**:

- ✅ Nenhum dado de paciente é armazenado no repositório
- ✅ Arquivos Excel com dados reais estão no `.gitignore`
- ✅ URLs e IDs do ambiente corporativo são carregados via variáveis de ambiente
- ✅ Caminhos de rede e identificadores de usuário removidos do código
- ✅ O repositório contém apenas lógica de automação, sem dados pessoais ou sensíveis

---

## 📈 Resultados

| Métrica | Antes | Depois |
|---|---|---|
| Tempo médio de geração | 30–60 min | ~2 min |
| Intervenção manual necessária | Total | Zero |
| Consistência de formatação | Variável | 100% padronizada |
| Rastreabilidade | Nenhuma | Log completo por execução |

---

## 🧠 Decisões Técnicas

**Por que Playwright ao invés de Selenium?**
O Power BI renderiza componentes complexos com estados assíncronos. O Playwright oferece melhor suporte a `expect()` e esperas automáticas, tornando a automação mais resiliente a variações de carregamento.

**Por que perfil persistente no Edge?**
O Power BI exige autenticação corporativa (SSO). Usar `launch_persistent_context` com um perfil já autenticado evita gerenciar login automatizado, que seria frágil e incompatível com políticas de segurança.

**Por que pywin32 para as Tabelas Dinâmicas?**
O Openpyxl não executa macros nem recalcula pivôs. O `win32com` despacha a chamada diretamente para o Excel instalado na máquina, garantindo resultado idêntico ao processo manual.

**Por que dicionário no retorno de `ler_arquivos()`?**
Retornar múltiplos valores como tupla é frágil — qualquer mudança na ordem quebra quem consome a função. O dicionário torna o contrato explícito e resistente a refatorações.

---

## 👨‍💻 Autor

**Pedro** — Assistente de Dados | Setor Hospitalar

Projeto desenvolvido como solução interna para otimização de processos de faturamento.

---

*Desenvolvido com Python • Playwright • Pandas • Openpyxl • pywin32*