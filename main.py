"""
Script to generate combined ThinkCell charts from Excel files
and update a PowerPoint deck using a template.
"""

from pathlib import Path
from typing import List, Tuple

from helpers.json_array import generate_multi_series_bar_json
from helpers.temporary_ppttc import write_ppttc
from helpers.thinkcell_cli import run_thinkcell_cli


def main() -> None:
    """
    Generate combined ThinkCell charts and update the PowerPoint presentation.
    """
    root: Path = Path.cwd()
    template_dir: Path = root / "_inputs" / "_template_deck"
    refresh_dir: Path = root / "_inputs" / "_refresh_files"

    # Locate the single PowerPoint template
    templates: List[Path] = list(template_dir.glob("*.pptx"))
    if len(templates) != 1:
        raise RuntimeError(f"Expected one .pptx in {template_dir}, found {len(templates)}")
    template_pptx: str = str(templates[0])

    # Define mapping of data files to ThinkCell chart identifiers
    file_chart_pairs: List[Tuple[Path, str]] = [
        (refresh_dir / "Sales_YOY.xlsx"     , "Sales_YOY"),
        (refresh_dir / "SpendItem_YOY.xlsx" , "SpendItem_YOY"),
        (refresh_dir / "Slide_5.xlsx" , "Slide_5a"),
        (refresh_dir / "Slide_5.xlsx" , "Slide_5b"),
        (refresh_dir / "Slide_5.xlsx"  , "Slide_5a"),
        (refresh_dir / "Slide_5.xlsx"  , "Slide_5b"),
        (refresh_dir / "Slide_6.xlsx" , "Slide_6a"),
        (refresh_dir / "Slide_6.xlsx" , "Slide_6b"),
        (refresh_dir / "Slide_6.xlsx" , "Slide_6c"),
        (refresh_dir / "Slide_6.xlsx"  , "Slide_6a"),
        (refresh_dir / "Slide_6.xlsx"  , "Slide_6b"),
        (refresh_dir / "Slide_6.xlsx"  , "Slide_6c"),
    ]

    ppttc_path: Path = root / "_intermediate" / "combined_charts.ppttc"
    final_pptx: Path = root / "_output" / "Updated_PPT.pptx"

    # Generate ThinkCell JSON configuration
    json_config = generate_multi_series_bar_json(file_chart_pairs, template_pptx)

    # Write the .ppttc file and run ThinkCell CLI to produce the final PPTX
    write_ppttc(json_config, ppttc_path)
    run_thinkcell_cli(ppttc_path, final_pptx)


if __name__ == "__main__":
    main()