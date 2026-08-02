import json
import re
from urllib.parse import urlparse, urlunparse

from curl_cffi import requests


def download_challonge_page(url: str) -> str:
    """Scrapes raw HTML/JSON store state from a specified Challonge tournament link."""
    headers = {
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "accept-language": "en-US,en;q=0.9",
        "cache-control": "no-cache",
        "referer": "https://challonge.com/",
    }

    parsed = urlparse(url.strip())
    if not parsed.scheme:
        parsed = urlparse("https://" + url.strip())

    base = urlunparse((parsed.scheme or "https", parsed.netloc, parsed.path.rstrip("/"), "", "", ""))
    variants = [base, base + "/module?multiplier=1&match_width_multiplier=1&show_final_results=1"]

    for candidate_url in variants:
        for imp in ["chrome124", "chrome123", "chrome120"]:
            try:
                res = requests.get(candidate_url, headers=headers, impersonate=imp, timeout=15)
                if res.status_code == 200:
                    return res.text
            except Exception:
                continue

    raise RuntimeError("[!] Failed to fetch Challonge data: Blocked by Challonge")


def parse_challonge_display_leader(display_name: str) -> str:
    """Extracts team leader display names from raw Challonge match label strings."""
    player_text = (display_name or "").split("|", 1)[0]
    pattern = r"([^\s\[(|]+)(?:\s*\[(.*?)\])?(?:\s*\((-?\d+(?:\.\d+)?)\))?"
    ignored = {"total", "guesses", "average", "avg", "="}

    for name, _, _ in re.findall(pattern, player_text):
        if not name or name.casefold() in ignored or re.fullmatch(r"-?\d+(?:\.\d+)?", name):
            continue
        return name

    return "Unknown Team"