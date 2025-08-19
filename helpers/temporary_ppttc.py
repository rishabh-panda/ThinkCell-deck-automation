from pathlib import Path

def write_ppttc(json_str: str, ppttc_path: Path):
    ppttc_path.write_text(json_str, encoding='utf-8')