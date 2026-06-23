# Pipeline Automático de Holerites — Design Spec
**Data:** 2026-06-23
**Status:** Aprovado

## Problema

Hoje o processo de renomear e enviar holerites exige intervenção manual em cada etapa. O objetivo é automatizar o fluxo completo: baixar o ZIP do Outlook, renomear os PDFs e enviar cada holerite ao funcionário correspondente — tudo com um duplo clique.

## Decisões de Design

- **Sem Azure / sem TI:** autenticação via Outlook COM (win32com), usando a sessão do Outlook já aberta no computador. Não requer cadastro externo, tokens ou permissões especiais.
- **Sem IA:** o processo é determinístico e previsível. A lógica de matching por matrícula já existe e funciona. IA pode ser adicionada futuramente se o processo se tornar imprevisível.
- **Reutiliza `renomear_holerites.py`:** a lógica de renomeação existente não é alterada.

## Premissas

1. O Outlook está instalado e aberto no computador onde o script roda.
2. O e-mail com o ZIP chega na caixa de entrada com a palavra "holerite" no assunto (case-insensitive).
3. O ZIP contém apenas os PDFs dos holerites, com nomes no formato `000003 EMILCE.pdf`.
4. A planilha `Emails funcionarios.xlsx` tem colunas `Matricula` e `Email`.
5. A chave de ligação entre as planilhas e os PDFs é o número de 6 dígitos da matrícula.

## Fluxo Completo

```
duplo clique em "RENOMEAR E ENVIAR HOLERITES.bat"
  ↓
[Passo 1 — email_reader.py]
  Abre Outlook via win32com
  Busca e-mail com "holerite" no assunto na caixa de entrada
  Se mais de um encontrado: lista e pede para o usuário escolher
  Baixa o ZIP, extrai PDFs em HOLERITES\

  ↓
[Passo 2 — renomear_holerites.py (existente, sem alteração)]
  Lê Holerites renomear.xlsx
  Renomeia os PDFs em HOLERITES\

  ↓
[Passo 3 — conferencia (dentro de pipeline.py)]
  Lê Emails funcionarios.xlsx (colunas: Matricula, Email)
  Para cada PDF renomeado, extrai matrícula e busca e-mail
  Exibe tabela de conferência:
    FUNCIONÁRIO              | E-MAIL                   | ARQUIVO
    Emilce Gomes             | emilce@empresa.com       | A000003_...pdf
    Marcia Vicente           | marcia@empresa.com       | A000007_...pdf
    [SEM E-MAIL]             | —                        | A000099_...pdf
  Mostra contagem: "X prontos para envio, Y sem e-mail cadastrado"
  Pergunta: "Confirmar envio? (s/n)"

  ↓
[Passo 4 — email_sender.py]
  Para cada PDF com e-mail encontrado:
    Compõe e-mail via win32com
    Assunto: configurável em config.json
    Corpo: configurável em config.json
    Anexo: PDF do funcionário
    Envia
    Exibe progresso: "Enviando 1/51... 2/51..."
  Relatório final:
    ✓ 49 e-mails enviados
    ✗ 2 sem e-mail cadastrado: [lista de nomes]
```

## Arquivos

| Ação | Arquivo | Responsabilidade |
|---|---|---|
| Criar | `pipeline.py` | Orquestra os 4 passos, tela de conferência |
| Criar | `email_reader.py` | Busca e-mail no Outlook, baixa e extrai ZIP |
| Criar | `email_sender.py` | Envia PDFs via Outlook COM |
| Criar | `config.json` | Assunto e corpo do e-mail configuráveis |
| Criar | `RENOMEAR E ENVIAR HOLERITES.bat` | Novo ponto de entrada único |
| Manter | `renomear_holerites.py` | Sem alteração |
| Manter | `RENOMEAR HOLERITES.bat` | Mantido para uso avulso |

## config.json

```json
{
  "assunto_email": "Holerite {mes}/{ano}",
  "corpo_email": "Olá {nome},\n\nSegue em anexo seu holerite referente a {mes}/{ano}.\n\nAtenciosamente,\nRH"
}
```

Os placeholders `{mes}`, `{ano}` e `{nome}` são preenchidos automaticamente a partir do nome do arquivo PDF e do nome do funcionário.

## Tratamento de Erros

| Situação | Comportamento |
|---|---|
| Outlook fechado | Mensagem: "Abra o Outlook antes de executar." |
| Nenhum e-mail com holerite encontrado | Mensagem: "Nenhum e-mail com holerite encontrado na caixa de entrada." |
| Mais de um e-mail candidato | Lista os e-mails e pede ao usuário que escolha (digita número) |
| PDF sem correspondência na planilha de nomes | Listado como "sem correspondência" antes da conferência |
| Funcionário sem e-mail cadastrado | Pulado no envio, listado no relatório final |
| Falha no envio de um e-mail | Continua os demais, lista falhas no relatório |

## Dependências Novas

| Biblioteca | Uso | Instalação |
|---|---|---|
| `pywin32` | Outlook COM automation | `pip install pywin32` |

Todas as outras dependências (`pandas`, `openpyxl`) já estão instaladas.

## Fora do Escopo

- Monitoramento contínuo da caixa de entrada (o processo é iniciado manualmente)
- Envio automático sem confirmação humana
- Suporte a outros clientes de e-mail além do Outlook instalado
- Uso de IA para tomada de decisões
