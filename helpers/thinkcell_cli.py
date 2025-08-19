import subprocess
import os
from pathlib import Path

def run_thinkcell_cli(ppttc_path: Path, output_pptx: Path, exe_path=None):
    exe = exe_path or r"C:\Program Files (x86)\think-cell\ppttc.exe"
    if not os.path.exists(exe):
        raise FileNotFoundError(f"Could not find think-cell CLI at {exe}")
    cmd = [exe, str(ppttc_path), "-o", str(output_pptx)]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)