import json
from pathlib import Path
from typing import Dict

from .config import BASE_DIR

LINKS_FILE = BASE_DIR / 'saved_links.json'


def load_links() -> Dict[str, str]:
    """Load saved Google Sheet links.

    Returns:
        dict: Mapping of name -> url.
    """
    if LINKS_FILE.exists():
        try:
            with open(LINKS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return {str(k): str(v) for k, v in data.items()}
        except Exception:
            pass
    return {}


def save_link(name: str, url: str) -> None:
    """Save a Google Sheet link under a user-defined name."""
    links = load_links()
    links[name] = url
    LINKS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(LINKS_FILE, 'w', encoding='utf-8') as f:
        json.dump(links, f, ensure_ascii=False, indent=2)
