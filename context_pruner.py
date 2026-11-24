"""Context Pruner & Surgical Skill Router (adapted from RealmPlay Context Router).

Evaluates user intent and drawing sheet scope to inject only relevant
architectural discipline rules, avoiding context bloat and prompt drift.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Set

# Skill keyword triggers mapped to project skill directory names
SKILL_ROUTING_MAP: Dict[str, Dict[str, Any]] = {
    "hospitality-bar-kitchen-equipment": {
        "keywords": ["bar", "kitchen", "counter", "cocktail", "sink", "bottle", "ice bin", "บาร์", "ครัว"],
        "sheet_prefix": ["ID-203", "B-02", "C-02", "F-01"],
    },
    "wet-area-plumbing": {
        "keywords": ["toilet", "restroom", "wc", "lavatory", "shower", "drain", "ห้องน้ำ", "สุขภัณฑ์"],
        "sheet_prefix": ["ID-206", "ID-207", "ID-208", "A-07", "A-08", "D-01W", "E-01W"],
    },
    "rcp-lighting-electrical": {
        "keywords": ["ceiling", "rcp", "downlight", "cove", "lighting", "diffuser", "ฝ้า", "ไฟ", "ดวงโคม"],
        "sheet_prefix": ["ID-201-2", "ID-202-2", "ID-203-2", "ID-204-2", "ID-205-2", "ID-206-2", "ID-207-2", "ID-208-2"],
    },
    "av-stage-ktv-integration": {
        "keywords": ["ktv", "stage", "dj", "speaker", "display", "karaoke", "จอ", "เวที", "ลำโพง"],
        "sheet_prefix": ["ID-204", "ID-205", "D-01", "E-01"],
    },
    "acoustic-soundproofing-detail": {
        "keywords": ["acoustic", "stc", "soundproof", "resilient", "กันเสียง", "ซับเสียง"],
        "sheet_prefix": ["ID-204", "ID-205", "D-", "E-"],
    },
    "millwork-joinery-detail": {
        "keywords": ["millwork", "joinery", "cabinet", "banquette", "wardrobe", "built-in", "บิลท์อิน", "ตู้"],
        "sheet_prefix": ["ID-201-3", "ID-202-3", "ID-203-3"],
    },
}

# Always active core drafting skills
CORE_SKILLS: Set[str] = {
    "autocad-professional-draftsman",
    "cad-layer-lineweight",
    "annotation-symbols",
}


def route_active_skills(user_prompt: str, active_sheet: str = "") -> List[str]:
    """Determine the optimal minimal set of skills for the active task."""
    selected_skills = set(CORE_SKILLS)
    prompt_lower = (user_prompt or "").lower()
    sheet_clean = (active_sheet or "").strip().upper()

    for skill_name, config in SKILL_ROUTING_MAP.items():
        # Check sheet prefix match
        if any(sheet_clean.startswith(prefix) for prefix in config["sheet_prefix"]):
            selected_skills.add(skill_name)
            continue

        # Check prompt keyword match
        if any(kw in prompt_lower for kw in config["keywords"]):
            selected_skills.add(skill_name)

    return sorted(selected_skills)


def get_routing_summary(user_prompt: str, active_sheet: str = "") -> Dict[str, Any]:
    """Generate explainable routing report for logging and debugging."""
    active = route_active_skills(user_prompt, active_sheet)
    pruned_count = len(SKILL_ROUTING_MAP) + len(CORE_SKILLS) - len(active)
    return {
        "active_sheet": active_sheet,
        "routed_skills": active,
        "pruned_skills_count": pruned_count,
        "context_savings_pct": round((pruned_count / (len(SKILL_ROUTING_MAP) + len(CORE_SKILLS))) * 100, 1),
    }
