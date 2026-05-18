from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any, Iterable

def write_json(path: Path, data:Any) -> None:
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

def write_csv(path: Path, rows:Iterable[dict], fieldnames:list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)
