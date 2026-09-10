#!/usr/bin/env python3
"""Reuse the committed percentile/Amdahl analyzer for the 0499 after controls."""
import importlib.util
import json
from pathlib import Path
HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("batch_analysis", HERE.parent / "change-0498/analyze-final.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
result = module.analyze(HERE / "after")
result["historical_baseline_note"] = (
    "Both 0499 executables use the unchanged harness with plain data() on the ordinary serial path. "
    "This table compares serial and batch routes only within the after executable; comparison.json "
    "contains the separate matched before/after results. The earlier 0498 historical accounting-path "
    "confound does not apply to these fresh 0499 controls."
)
(HERE / "after-scaling.json").write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
markdown = module.build_markdown(result).replace("# Change 0498 final benchmark analysis", "# Change 0499 after: serial versus batch scaling", 1)
markdown = markdown.replace("## Historical baseline boundary", "## Matched baseline boundary")
(HERE / "after-scaling.md").write_text(markdown)
print("0499 after scaling regenerated")
