"""Configuration globale de l'application (nom et code du cours)."""

import json
from pathlib import Path

CONFIG_PATH = Path(__file__).parent / "data" / "configuration.json"

PAR_DEFAUT = {"nom_cours": "", "code_cours": ""}


def charger_configuration():
    if not CONFIG_PATH.exists():
        return dict(PAR_DEFAUT)
    return {**PAR_DEFAUT, **json.loads(CONFIG_PATH.read_text(encoding="utf-8"))}


def sauver_configuration(configuration):
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(configuration, indent=2, ensure_ascii=False), encoding="utf-8")
