# Excel Organizer Agent

Suplemento para Excel que organiza planilhas automaticamente via painel lateral. Descreva em linguagem natural o que quer fazer e a IA interpreta e executa as transformações.

## Requisitos

| Ferramenta | Versão mínima | Download |
|---|---|---|
| Python | 3.10+ | https://python.org/downloads |
| Node.js | 18+ | https://nodejs.org |
| Excel | Microsoft 365 / Office 2019+ | — |

> **Chave de API:** é necessária uma conta Anthropic com `ANTHROPIC_API_KEY` para usar o modo de linguagem natural. Sem ela, as ferramentas manuais funcionam normalmente.

---

## Instalação

### 1. Clonar o repositório

```bash
git clone https://github.com/seu-usuario/excel-organizer-agent.git
cd excel-organizer-agent
```

### 2. Configurar variáveis de ambiente

```bash
cd backend
copy ..\\.env.example .env
```

Edite o arquivo `.env` e preencha sua `ANTHROPIC_API_KEY`:

```
ANTHROPIC_API_KEY=sk-ant-...
```

### 3. Instalar dependências

**Backend:**
```bash
cd backend
pip install -r requirements.txt
```

**Add-in:**
```bash
cd excel-addin
npm install
```

---

## Iniciar

Dê dois cliques no arquivo **`iniciar.bat`** na raiz do projeto.

Ele abre automaticamente:
- **Backend API** em `http://localhost:8000`
- **Servidor do Add-in** em `http://localhost:3001`

---

## Instalar o suplemento no Excel (feito uma vez)

1. Abra o **Excel**
2. Vá em **Inserir → Suplementos → Meus Suplementos**
3. Clique em **Carregar Suplemento**
4. Selecione o arquivo `excel-addin\manifest.xml`
5. O painel **Organizador Excel** aparece na lateral

> O sideload precisa ser refeito ao reabrir o Excel. Para instalação permanente, consulte a documentação oficial da Microsoft sobre catálogos de suplementos.

---

## Funcionalidades

| Função | Descrição |
|---|---|
| **Organizar com IA** | Descreva em português o que quer fazer |
| **Ordenar** | Ordene por uma ou mais colunas |
| **Remover duplicatas** | Com controle de quais colunas verificar |
| **Padronizar texto** | Maiúsculas, minúsculas, capitalizar, remover espaços |
| **Colorir por regra** | Destaque linhas com base em condições |
| **Dividir em abas** | Separe os dados por categoria automaticamente |
| **Perfis** | Salve e reutilize configurações frequentes |
| **Proteção de formatação** | Bloqueia alterações em planilhas com cores ou fórmulas já definidas |

### Proteção contra alterações não autorizadas

Antes de aplicar qualquer transformação, o agente verifica automaticamente se a planilha ativa contém formatação pré-existente:

- **Fórmulas** — células com fórmulas Excel (`=SOMA`, `=PROCV`, etc.)
- **Cores de preenchimento** — células com cor de fundo definida manualmente

Se algum desses elementos for detectado, um **modal de confirmação** é exibido listando o que foi encontrado. O agente só prossegue com a autorização explícita do usuário.

> Esta proteção se aplica a todos os modos de operação: instrução por IA, ferramentas manuais e perfis salvos.

### Exemplos de instruções para a IA

- "Ordene pelo nome do cliente de A a Z"
- "Remova linhas duplicadas baseadas na coluna Email"
- "Coloque todos os nomes em maiúsculas"
- "Colorir de vermelho as linhas onde Status é igual a Inativo"
- "Crie uma aba separada para cada departamento"
- "Ordene do maior para o menor valor e destaque em verde quem tem valor acima de 1000"

---

## Exemplos de uso da API

### Verificar saúde

```bash
curl http://localhost:8000/health
```

### Organizar com parâmetros

```bash
curl -X POST http://localhost:8000/api/organize \
  -F "file=@planilha.xlsx" \
  -F 'parameters={"sort_by":[{"column":"Nome","direction":"asc"}],"remove_duplicates":true}'
```

### Organizar com instrução em linguagem natural

```bash
curl -X POST http://localhost:8000/api/organize \
  -F "file=@planilha.xlsx" \
  -F 'parameters={"natural_language_instruction":"Ordene por nome em ordem alfabetica e remova duplicatas"}'
```

### Prosseguir mesmo com formatação existente

Se o arquivo contiver cores ou fórmulas, o backend retorna `requires_confirmation: true`. Para confirmar e prosseguir, envie novamente com `force_override: true`:

```bash
curl -X POST http://localhost:8000/api/organize \
  -F "file=@planilha.xlsx" \
  -F 'parameters={"sort_by":[{"column":"Nome","direction":"asc"}],"force_override":true}'
```

### Criar perfil

```bash
curl -X POST http://localhost:8000/api/profiles \
  -H "Content-Type: application/json" \
  -d '{"name":"Organizacao Mensal","parameters":{"sort_by":[{"column":"Nome","direction":"asc"}],"remove_duplicates":true}}'
```

---

## Variáveis de Ambiente

| Variável | Descrição | Padrão |
|---|---|---|
| `ANTHROPIC_API_KEY` | Chave da API Anthropic | obrigatório para IA |
| `DATABASE_URL` | URL do banco SQLite | `sqlite:///./storage/organizer.db` |
| `MAX_FILE_SIZE_MB` | Limite de tamanho de arquivo | `50` |
| `CORS_ORIGINS` | Origens permitidas pelo CORS | localhost 3000 e 3001 |
| `ENVIRONMENT` | Ambiente de execução | `development` |

---

## Deploy no Railway

1. Faça push do repositório para o GitHub
2. No Railway, crie um novo projeto e conecte o repositório
3. Configure a variável de ambiente `ANTHROPIC_API_KEY`
4. O `railway.json` já configura o build automaticamente
5. Atualize a `API_BASE_URL` em `excel-addin/src/taskpane/taskpane.js` com a URL pública gerada pelo Railway

---

## Segurança

- O arquivo `.env` **nunca é commitado** (incluído no `.gitignore`)
- Uploads são deletados automaticamente após o processamento
- Arquivos de saída expiram em 1 hora
- Limite de tamanho de arquivo configurável via `MAX_FILE_SIZE_MB`
