

import json
import logging
from pathlib import Path

from src.gemini_client import send_text_to_gemini, parse_json_response

logger = logging.getLogger(__name__)


def _build_ddr_prompt(merged_data: list) -> str:
    return f"""
You are a professional property consultant writing a Detailed Diagnostic Report (DDR) for a client.

Source data (merged inspection + thermal findings):
{json.dumps(merged_data, indent=2)}

Write a complete, professional DDR with the 7 sections below.
Language: simple, client-friendly. No technical jargon.
Rule: Do NOT invent facts. Write "Not Available" for anything missing.

Return ONLY the following JSON structure. No explanation. No markdown fences.

{{
  "property_issue_summary": "A single clear paragraph summarising all major issues found across the property.",

  "area_wise_observations": [
    {{
      "area_name": "Hall - Ceiling",
      "observation_text": "A paragraph combining inspection and thermal findings for this area in plain English.",
      "thermal_data": {{
        "hotspot": "27.5°C",
        "coldspot": "22.3°C",
        "emissivity": "0.95",
        "date": "12-Jan-2024"
      }}
    }}
  ],

  "probable_root_cause": [
    {{
      "area_name": "Hall - Ceiling",
      "root_cause": "Plain-language explanation of the likely cause."
    }}
  ],

  "severity_assessment": [
    {{
      "area_name": "Hall - Ceiling",
      "severity": "High",
      "reasoning": "Brief explanation of why this severity level was assigned."
    }}
  ],

  "recommended_actions": [
    {{
      "area_name": "Hall - Ceiling",
      "actions": ["Specific action 1", "Specific action 2"],
      "priority": "Immediate"
    }}
  ],

  "additional_notes": "A paragraph covering recurring patterns, general observations, or important context.",

  "missing_or_unclear_information": [
    {{
      "item": "Description of the missing or conflicting information.",
      "source": "Which section or area this gap affects."
    }}
  ]
}}

Severity levels to use: Critical | High | Medium | Low
Priority levels to use: Immediate | Short-term | Long-term
"""


def generate_ddr_content(
    model,
    merged_data: list,
    max_retries: int,
    checkpoint_dir: Path
) -> dict:
    """
    Generate the full DDR report content as a structured dict.
    """
    checkpoint_dir.mkdir(parents=True, exist_ok=True)
    checkpoint_path = checkpoint_dir / "ddr_content.json"

    logger.info("Generating DDR content from merged data...")

    prompt = _build_ddr_prompt(merged_data)

    try:
        raw = send_text_to_gemini(model, prompt, max_retries)
        ddr = parse_json_response(raw)

        if not isinstance(ddr, dict):
            raise ValueError(f"Expected dict, got {type(ddr)}")

        with open(checkpoint_path, "w", encoding="utf-8") as f:
            json.dump(ddr, f, indent=2)
        logger.info(f"DDR content checkpoint saved → {checkpoint_path}")

        return ddr

    except Exception as e:
        logger.error(f"DDR content generation failed: {e}")
        raise
