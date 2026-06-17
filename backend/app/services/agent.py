import json
import asyncio
from typing import List
import anthropic
from app.config import settings
from app.schemas.organize import OrganizeParameters

SYSTEM_PROMPT = """Você é um especialista em organização de planilhas Excel.
Sua função é converter instruções em linguagem natural para
parâmetros estruturados de organização.

Você receberá:
1. Uma instrução do usuário em português
2. A lista de colunas disponíveis na planilha

Você deve retornar APENAS um JSON válido (sem markdown, sem explicações)
seguindo exatamente este schema:

{
  "sort_by": [{"column": "nome_coluna", "direction": "asc|desc"}],
  "remove_duplicates": true|false,
  "duplicate_columns": ["col1"] ou null,
  "keep_duplicate": "first|last",
  "standardize_text": {"nome_coluna": "upper|lower|capitalize|strip"},
  "color_rules": [{
    "column": "nome_coluna",
    "operator": "equals|not_equals|starts_with|ends_with|contains|greater_than|less_than|is_empty|is_not_empty",
    "value": "valor ou null",
    "color": "yellow|green|red|blue|orange|gray|pink"
  }],
  "split_by_category": "nome_coluna ou null",
  "keep_original_sheet": true|false,
  "create_summary_sheet": true|false
}

REGRAS:
- Use apenas colunas que existem na lista fornecida
- Se o usuário mencionar uma coluna que não existe, use a mais similar
- Se uma funcionalidade não foi mencionada, omita do JSON ou use null/false
- Interprete intenções: "do maior para o menor" = direction "desc"
- Cores: vermelho=problemas/negativo, verde=ok/positivo, amarelo=atenção"""


async def interpret_instruction(
    instruction: str, columns: List[str]
) -> OrganizeParameters:
    client = anthropic.Anthropic(api_key=settings.anthropic_api_key)
    user_message = (
        f"Instrução: {instruction}\n\nColunas disponíveis: {', '.join(columns)}"
    )

    last_error = None
    for attempt in range(2):
        try:
            response = await asyncio.to_thread(
                client.messages.create,
                model="claude-sonnet-4-6",
                max_tokens=1024,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_message}],
            )
            raw = response.content[0].text.strip()
            data = json.loads(raw)
            return OrganizeParameters(**data)
        except (json.JSONDecodeError, Exception) as e:
            last_error = e
            if attempt == 0:
                await asyncio.sleep(1)

    raise RuntimeError(
        f"Falha ao interpretar instrução após 2 tentativas: {last_error}"
    )
