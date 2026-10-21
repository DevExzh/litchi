"""Hash the isolated PPT measurement source, excluding build/evidence output."""
import argparse
import hashlib
import json
from pathlib import Path

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("source", type=Path)
parser.add_argument("output", type=Path)
args = parser.parse_args()
root = args.source.resolve()
paths = [path for path in (root / "crates").rglob("*") if path.is_file()]
paths.extend(root / name for name in (
    ".cargo/config.toml", "Cargo.lock", "Cargo.toml", "clippy.toml",
    "rust-toolchain.toml", "rustfmt.toml",
))
manifest = {
    str(path.relative_to(root)): hashlib.sha256(path.read_bytes()).hexdigest()
    for path in sorted(paths)
}
args.output.write_text(json.dumps(manifest, indent=2) + "\n")
