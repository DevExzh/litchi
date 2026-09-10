#!/usr/bin/env python3
"""Seal retained evidence, excluding the inventory itself to avoid recursion."""
from pathlib import Path
import hashlib, json
HERE = Path(__file__).resolve().parent
files = [{"path": str(p.relative_to(HERE)), "bytes": p.stat().st_size,
          "sha256": hashlib.sha256(p.read_bytes()).hexdigest()}
         for p in sorted(HERE.rglob("*")) if p.is_file() and p.name != "inventory.json"]
(HERE / "inventory.json").write_text(json.dumps({"files": files, "file_count": len(files),
    "total_bytes": sum(p["bytes"] for p in files)}, indent=2) + "\n")
print(json.dumps({"files": len(files), "bytes": sum(p["bytes"] for p in files)}))
