"""
ai_service.py
Generates structured interview answers using OpenAI GPT-4o-mini.
Returns { points: list[str], illustrated: str }
"""

import json
import re
from openai import AsyncOpenAI
from config import settings

client = AsyncOpenAI(api_key=settings.openai_api_key)

ROLE_CONTEXTS: dict[str, str] = {
    "Test Manager": (
        "You are an expert Test Manager with 15+ years in software testing, QA strategy, "
        "test planning, team management, defect tracking, and test metrics."
    ),
    "Scrum Master": (
        "You are an expert Scrum Master with deep knowledge of Agile/Scrum ceremonies, "
        "sprint planning, backlog refinement, impediment removal, team facilitation, and Agile metrics."
    ),
    "Defect Manager": (
        "You are an expert Defect Manager specialising in defect lifecycle management, "
        "root cause analysis, defect triage, metrics, and defect prevention strategies."
    ),
    "SAP Tester": (
        "You are an expert SAP Tester with hands-on experience testing SAP ERP modules "
        "(FI, CO, MM, SD, HR), regression testing, integration testing, and SAP-specific test tools."
    ),
    "SAP Test Manager": (
        "You are an expert SAP Test Manager managing end-to-end SAP testing programmes, "
        "cutover testing, UAT coordination, SAP Solution Manager, and test governance."
    ),
    "Project Manager": (
        "You are an expert Project Manager with PMP/PRINCE2 knowledge, skilled in project planning, "
        "risk management, stakeholder communication, budgeting, and delivery management."
    ),
}

SYSTEM_TEMPLATE = """{role_context}

You are helping a candidate prepare for a job interview for the role of {role}.
When answering interview questions, respond in the following strict JSON format:

{{
  "points": ["Point 1", "Point 2", "Point 3", ...],
  "illustrated": "A detailed, well-explained answer with real-world examples, context, and depth."
}}

Rules:
- "points": 5 to 8 concise bullet points. Each point must be a complete short sentence.
- "illustrated": A thorough 150-250 word explanation with practical examples and specifics.
- Keep the tone professional and interview-appropriate.
- Do NOT include markdown formatting inside JSON strings.
- Respond ONLY with valid JSON — no extra text outside the JSON object."""


async def generate_answer(role: str, question: str) -> dict:
    """Call OpenAI and return {'points': [...], 'illustrated': '...'}."""
    role_context = ROLE_CONTEXTS.get(role, "You are an experienced IT professional.")
    system_prompt = SYSTEM_TEMPLATE.format(role_context=role_context, role=role)

    response = await client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f'Interview Question: "{question}"'},
        ],
        temperature=0.7,
        max_tokens=900,
    )

    raw = response.choices[0].message.content.strip()

    # Strip markdown code fences if present
    raw = re.sub(r"^```json\s*", "", raw, flags=re.IGNORECASE)
    raw = re.sub(r"```$", "", raw).strip()

    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        # Graceful fallback
        parsed = {"points": [raw[:200]], "illustrated": raw}

    return {
        "points": parsed.get("points", []) if isinstance(parsed.get("points"), list) else [],
        "illustrated": parsed.get("illustrated", "") if isinstance(parsed.get("illustrated"), str) else "",
    }
