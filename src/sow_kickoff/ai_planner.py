from __future__ import annotations

import json

from openai import OpenAI

from .models import DeckPlan, SourceDocument, TemplateInventory


SYSTEM_PROMPT = """You are a meticulous enterprise account manager preparing a client kickoff presentation.

Your task is to map an approved statement of work into an existing PowerPoint template inventory.

NON-NEGOTIABLE RULES:
1. Use only facts supported by the supplied SoW or explicit user metadata.
2. Never invent dates, deliverables, responsibilities, quantities, names, risks, governance, scope, exclusions, or commitments.
3. Turn contractual prose into concise, presentation-ready language without changing its meaning.
4. Keep assumptions separate from commitments and out-of-scope items separate from scope.
5. If an editable placeholder expects information that is absent, return an edit with an empty text value.
6. Do not edit logos, footers, confidentiality labels, page numbers, navigation labels, or decorative copy.
7. Use only slide numbers, shape IDs, table coordinates, and target kinds present in the inventory.
8. Prefer short bullets. Separate bullets with newline characters and do not prefix them with decorative symbols.
9. Preserve exact product names, dates, milestones, quantities, and party responsibilities.
10. Source references must use the supplied labels, for example "Paragraph 12" or "Table 2, row 3". Metadata-only edits may use "User metadata".
11. Do not add a title edit merely to restate an existing good slide title.
12. Return each target at most once.

Cover objectives, scope, deliverables, timeline, roles, governance, risks, assumptions, and next steps when the template contains appropriate destinations. Leave missing items blank rather than adding TBD.
"""


def build_deck_plan(
    *,
    source: SourceDocument,
    inventory: TemplateInventory,
    api_key: str,
    model: str,
    client_name: str,
    project_name: str,
    kickoff_date: str,
    generation_notes: str,
) -> DeckPlan:
    if not api_key:
        raise ValueError("An OpenAI API key is required.")
    if not model:
        raise ValueError("An OpenAI model name is required.")

    client = OpenAI(api_key=api_key)
    user_payload = {
        "user_metadata": {
            "client_name": client_name,
            "project_name": project_name,
            "kickoff_date": kickoff_date,
            "generation_notes": generation_notes,
        },
        "template_inventory": inventory.model_dump(),
        "statement_of_work": source.text,
        "source_was_truncated": source.truncated,
    }

    try:
        completion = client.chat.completions.parse(
            model=model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {
                    "role": "user",
                    "content": "Create the evidence-backed slide edit plan from this JSON input:\n"
                    + json.dumps(user_payload, ensure_ascii=False),
                },
            ],
            response_format=DeckPlan,
        )
    except Exception as exc:
        raise RuntimeError(f"The AI planning request failed: {exc}") from exc

    message = completion.choices[0].message
    if getattr(message, "refusal", None):
        raise RuntimeError(f"The AI could not process this document: {message.refusal}")
    if message.parsed is None:
        raise RuntimeError("The AI did not return a valid slide plan.")
    return message.parsed

