"""Where the tool reads and writes files, so it works from any folder once installed.

Bundled with the package: the peer lists and the example assumptions.
In the folder you run it from: .env, your own peers/ and assumptions/ files, and output/.
In your user cache folder: downloaded SEC and market data.
"""
import os
from pathlib import Path

PACKAGE = Path(__file__).resolve().parent


def cache_dir():
    override = os.environ.get("EDGAR_COMPS_CACHE")
    path = Path(override) if override else Path.home() / ".cache" / "edgar-comps"
    path.mkdir(parents=True, exist_ok=True)
    return path


def peers_file(name):
    """A peer list by name: your own ./peers/<name>.txt wins over the bundled one. A file path also works."""
    direct = Path(name)
    if direct.suffix == ".txt" and direct.exists():
        return direct
    for folder in (Path.cwd() / "peers", PACKAGE / "peers"):
        candidate = folder / f"{name}.txt"
        if candidate.exists():
            return candidate
    bundled = sorted(p.stem for p in (PACKAGE / "peers").glob("*.txt") if p.stem != "predecessors")
    raise FileNotFoundError(f"No peer list named '{name}'. Bundled lists: {', '.join(bundled)}. "
                            f"Or create ./peers/{name}.txt with one ticker per line.")


def assumptions_file(ticker):
    """DCF overrides for a ticker: ./assumptions/<TICKER>.csv wins over the bundled example. None if neither exists."""
    for folder in (Path.cwd() / "assumptions", PACKAGE / "assumptions"):
        candidate = folder / f"{ticker.upper()}.csv"
        if candidate.exists():
            return candidate
    return None


def predecessors_file():
    return PACKAGE / "peers" / "predecessors.txt"


def env_file():
    return Path.cwd() / ".env"


def default_output():
    return Path.cwd() / "output" / "comps.xlsx"
