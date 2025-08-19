import json
from pathlib import Path
from collections import defaultdict
from helpers.data_ingestion import read_single_series_bar

def generate_multi_series_bar_json(
    file_chart_pairs: list[tuple[Path,str]],
    template_pptx: str
) -> str:
    """
    From a list of (file_path, chart_name), group all series per chart
    and emit one JSON array of chart objects.
    If the first file for a chart errors, we log & skip that chart.
    """
    # 1) group paths by chart_name
    groups = defaultdict(list)
    for path, cname in file_chart_pairs:
        groups[cname].append(path)

    all_charts = []
    for chart_name, paths in groups.items():
        first_path = paths[0]
        try:
            first_dict = read_single_series_bar(first_path)
        except ValueError as e:
            print(f"[ERROR] Could not read '{first_path.name}' for chart '{chart_name}': {e}")
            print("        Skipping chart:", chart_name)
            continue

        # If OK, build header row
        header = [{"string": ""}] + [{"string": k} for k in first_dict.keys()]

        # build each series row
        rows = []
        for path in paths:
            try:
                data = read_single_series_bar(path)
            except ValueError as e:
                print(f"[WARN]  Could not read '{path.name}' as a series in chart '{chart_name}': {e}")
                print("        Skipping that series.")
                continue

            series_label = path.stem
            row = [{"string": series_label}] + [{"number": data[k]} for k in first_dict.keys()]
            rows.append(row)

        # if after skipping you have at least one row, emit this chart
        if rows:
            chart_obj = {
                "template": template_pptx,
                "data": [
                    {
                        "name": chart_name,
                        "table": [header] + rows
                    }
                ]
            }
            all_charts.append(chart_obj)
        else:
            print(f"[ERROR] No valid series left for chart '{chart_name}', skipping it entirely.")

    return json.dumps(all_charts, indent=4)