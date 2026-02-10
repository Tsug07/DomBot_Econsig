<p align="center">
  <img src="assets/DomBot_New.png" alt="DomBot Econsig Logo" width="150">
</p>

<h1 align="center">DomBot - Empréstimo Consignado</h1>

<p align="center">
  Automação inteligente para geração e publicação de relatórios de Empréstimo Consignado no sistema Domínio Folha
</p>

<p align="center">
  <img src="https://img.shields.io/badge/python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white" alt="Windows">
  <img src="https://img.shields.io/badge/GUI-CustomTkinter-1ABC9C?style=for-the-badge" alt="CustomTkinter">
  <img src="https://img.shields.io/badge/automation-PyWinAuto-E74C3C?style=for-the-badge" alt="PyWinAuto">
</p>

<p align="center">
  <img src="https://img.shields.io/github/last-commit/Tsug07/DomBot_Econsig?style=flat-square&color=2ECC71" alt="Last Commit">
  <img src="https://img.shields.io/github/repo-size/Tsug07/DomBot_Econsig?style=flat-square&color=3498DB" alt="Repo Size">
  <img src="https://img.shields.io/badge/version-2.0-1ABC9C?style=flat-square" alt="Version">
  <img src="https://img.shields.io/badge/status-em%20desenvolvimento-F39C12?style=flat-square" alt="Status">
</p>

---

## Sobre

O **DomBot Econsig** automatiza o processo completo de geração de relatórios de **Empréstimo Consignado - Por Mês** no sistema **Domínio Folha**, eliminando o trabalho manual repetitivo de:

- Trocar entre empresas via F8
- Navegar até o Gerenciador de Relatórios
- Preencher os 5 parâmetros do relatório
- Publicar o documento na categoria correta
- Gerar e salvar PDFs com nomes padronizados

Tudo controlado por uma interface gráfica moderna com logs em tempo real, estatísticas e controle total da execução.

## Funcionalidades

| Funcionalidade | Descrição |
|---|---|
| **Processamento em lote** | Processa múltiplas empresas a partir de planilha Excel |
| **Troca automática de empresa** | Alterna entre empresas via F8 antes de cada relatório |
| **Publicação de documentos** | Publica automaticamente na categoria `Pessoal/E-Consignado` |
| **Geração de PDF** | Salva relatórios em PDF com nome e pasta padronizados |
| **Interface moderna** | GUI dark theme com CustomTkinter e logo personalizada |
| **Logs coloridos** | Logs em tempo real com cores por tipo (sucesso, erro, aviso, info) |
| **Preview do Excel** | Visualização dos dados carregados antes de iniciar |
| **Controle de execução** | Iniciar, pausar, retomar e parar a qualquer momento |
| **Estatísticas em tempo real** | Cards com total, sucesso, erros, empresa atual e cronômetro |
| **Exportação de logs** | Salvar logs da sessão em arquivo `.txt` |
| **Tratamento de erros** | Detecção automática de diálogos de erro via Win32 API |
| **Performance otimizada** | Waits condicionais e `EnumWindows` em vez de sleeps fixos |

## Screenshot

```
┌──────────────────────────────────────────────────────────┐
│  🤖 DomBot - Econsig                  ● Aguardando...   │
├──────────────────────────────────────────────────────────┤
│  📁 [planilha.xlsx]  [Procurar]  Linha: [2]             │
│  [▶ Iniciar]  [⏸ Pausar]  [⏹ Parar]                    │
├──────────────────────────────────────────────────────────┤
│  📊 Total  ✅ Sucesso  ❌ Erros  🏢 Empresa  ⏱ Tempo    │
│  ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓░░░░░░░░░░░░░░░░░░░░░  42.8%          │
├──────────────────────────────────────────────────────────┤
│  📋 Logs  │  📊 Preview                                  │
│  [14:28:37] ℹ️ Nomeando PDF...                           │
│  [14:28:41] ℹ️ Salvando PDF                              │
│  [14:28:53] ✅ Linha 8 processada com sucesso            │
│  [14:28:59] ⏳ Processando linha 9 - Empresa 123         │
└──────────────────────────────────────────────────────────┘
```

## Pré-requisitos

- **Windows** (obrigatório — utiliza Win32 API)
- **Python 3.8+**
- **Domínio Folha** instalado e aberto

## Instalação

```bash
# Clonar o repositório
git clone https://github.com/Tsug07/DomBot_Econsig.git
cd DomBot_Econsig

# Instalar dependências
pip install customtkinter pandas pywinauto pywin32 pillow openpyxl
```

## Uso

### 1. Preparar a planilha Excel

A planilha deve conter as seguintes colunas obrigatórias:

| Coluna | Descrição | Exemplo |
|---|---|---|
| `Nº` | Número da empresa no Domínio | `103` |
| `Data Inicial` | Data inicial do período | `01/01/2026` |
| `Data Final` | Data final do período | `31/01/2026` |
| `Salvar Como` | Nome do arquivo PDF gerado | `103-Empresa LTDA-012026` |
| `EMPRESAS` | Nome da empresa (exibição) | `Empresa LTDA` |

### 2. Executar

```bash
python DomBot_Econsig.py
```

### 3. Na interface

1. Clique em **Procurar** e selecione a planilha Excel
2. Verifique os dados na aba **📊 Preview**
3. Ajuste a **linha inicial** se necessário
4. Certifique-se que o **Domínio Folha** está aberto
5. Clique em **▶ Iniciar**

## Fluxo da Automação

```
Início
  │
  ├─ Carregar planilha Excel
  ├─ Conectar ao Domínio Folha (UIA backend)
  │
  └─ Para cada linha:
       │
       ├─ Fechar janelas abertas (se não for a primeira)
       ├─ Trocar empresa via F8
       ├─ Fechar avisos de vencimento
       │
       ├─ Abrir Relatórios Integrados (ALT+R → I → I)
       ├─ Navegar até "Empréstimo Consignado - Por Mês"
       ├─ Preencher parâmetros:
       │   ├─ Arg 1: Empresa (Excel)
       │   ├─ Arg 2: Código Empregados (*)
       │   ├─ Arg 3: Data Inicial (Excel)
       │   ├─ Arg 4: Data Final (Excel)
       │   └─ Arg 5: Somente Valor Aberto (0)
       ├─ Executar relatório
       │
       ├─ Publicar documento
       │   ├─ Categoria: Pessoal/E-Consignado
       │   └─ Nome: coluna "Salvar Como"
       │
       ├─ Gerar PDF
       │   ├─ Navegar até pasta de destino
       │   └─ Salvar com nome padronizado
       │
       └─ Próxima linha
  │
  Fim → Resumo da execução
```

## Estrutura do Projeto

```
DomBot_Econsig/
├── DomBot_Econsig.py          # Aplicação principal
├── assets/
│   ├── DomBot_New.png         # Logo do aplicativo
│   ├── favicon.ico            # Ícone da janela
│   └── ...
├── logs/                      # Logs de execução (gerado automaticamente)
│   ├── success_YYYY-MM-DD.log
│   └── error_YYYY-MM-DD.log
└── README.md
```

## Dependências

| Pacote | Uso |
|---|---|
| `customtkinter` | Interface gráfica moderna (dark theme) |
| `pandas` | Leitura e manipulação da planilha Excel |
| `pywinauto` | Automação da interface do Domínio Folha |
| `pywin32` | Interação com janelas do Windows (Win32 API) |
| `Pillow` | Processamento de imagens (logo/ícones) |
| `openpyxl` | Engine para leitura de arquivos `.xlsx` |

## Otimizações de Performance

| Técnica | Descrição |
|---|---|
| `smart_sleep()` | Sleep interruptível com polling de 0.15s (respeita pause/stop) |
| `EnumWindows` | Detecção de diálogos de erro em passagem única (~1ms vs ~500ms por busca UIA) |
| `wait_for_condition()` | Waits condicionais com timeout em vez de sleeps fixos |
| `_is_connection_alive()` | Validação de conexão via `win32gui.IsWindow` antes de reconectar |
| Batch keystrokes | Teclas TAB agrupadas em vez de enviar uma a uma com sleep |

---

<p align="center">
  Desenvolvido por <a href="https://github.com/Tsug07">Tsug07</a>
</p>
