# Como instalar em uma nova máquina

## 1. Instalar os programas necessários

Antes de tudo, instale:

- **Python 3.10 ou superior** → https://python.org/downloads  
  Durante a instalação, marque a opção **"Add Python to PATH"**

- **Node.js 18 ou superior** → https://nodejs.org

- **Excel Microsoft 365 ou Office 2019/2021**  
  Versões mais antigas não suportam suplementos modernos.

---

## 2. Baixar o projeto

Abra o **Prompt de Comando** ou **PowerShell** e execute:

```bash
git clone https://github.com/Lucas-Barbosa-Vicente/excel-organizer-agent.git
cd excel-organizer-agent
```

> Se não tiver o Git instalado: https://git-scm.com/download/win

---

## 3. Configurar a chave de API

O programa usa a IA da Anthropic (Claude). Você precisa de uma chave de API.

1. Copie o arquivo de exemplo:
```bash
copy backend\.env.example backend\.env
```

2. Abra o arquivo `backend\.env` em qualquer editor de texto e preencha:
```
ANTHROPIC_API_KEY=sk-ant-SUA_CHAVE_AQUI
```

> Sem a chave, o modo "Organizar com IA" fica desativado.  
> As ferramentas manuais (ordenar, duplicatas, etc.) funcionam normalmente.

---

## 4. Instalar as dependências

**Backend (Python):**
```bash
cd backend
pip install -r requirements.txt
cd ..
```

**Add-in (Node.js):**
```bash
cd excel-addin
npm install
cd ..
```

---

## 5. Iniciar o programa

Dê **dois cliques** no arquivo **`iniciar.bat`** na pasta do projeto.

Duas janelas vão abrir:
- **Backend API** — servidor rodando em `http://localhost:8000`
- **Add-in Server** — arquivos do suplemento em `http://localhost:3001`

Mantenha as duas janelas abertas enquanto usar o programa.

---

## 6. Instalar o suplemento no Excel

> Este passo precisa ser feito uma vez por máquina.

1. Abra o **Excel**
2. Clique em **Inserir** no menu superior
3. Clique em **Suplementos** → **Meus Suplementos**
4. Clique em **Carregar Suplemento**
5. Navegue até a pasta do projeto e selecione o arquivo:
   ```
   excel-addin\manifest.xml
   ```
6. O painel **Organizador Excel** aparece na lateral direita

---

## 7. Usar o programa

Com o painel aberto no Excel:

| Função | Como usar |
|---|---|
| **Organizar com IA** | Digite em português o que quer fazer e clique em "Organizar com IA" |
| **Ordenar** | Escolha a coluna e a direção (crescente/decrescente) |
| **Remover duplicatas** | Marque a opção e escolha as colunas a verificar |
| **Padronizar texto** | Escolha maiúsculas, minúsculas, capitalizar ou remover espaços |
| **Colorir por regra** | Defina uma condição e uma cor para destacar linhas |
| **Dividir em abas** | Separa os dados em abas por categoria automaticamente |
| **Perfis** | Salve configurações para reutilizar depois |

### Exemplos de instruções para a IA

- "Ordene pelo nome do cliente de A a Z"
- "Remova linhas duplicadas pela coluna Email"
- "Coloque todos os nomes em maiúsculas"
- "Pinte de vermelho as linhas onde Status é Inativo"
- "Crie uma aba separada para cada departamento"

---

## Solução de problemas

**O painel não aparece no Excel**  
→ Verifique se o `iniciar.bat` está rodando e repita o passo 6.

**Erro "Python não reconhecido"**  
→ Reinstale o Python marcando "Add Python to PATH".

**Erro "npm não reconhecido"**  
→ Reinstale o Node.js e reinicie o computador.

**O modo IA não funciona**  
→ Verifique se a `ANTHROPIC_API_KEY` está preenchida corretamente no arquivo `backend\.env`.
