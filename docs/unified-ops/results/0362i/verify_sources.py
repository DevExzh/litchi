"""Verify ODraw's naming-only source change from the repository root."""
from pathlib import Path
import hashlib
import json
import re
import subprocess


def image_kind(source):
    match = re.search(
        r"#\[derive\(Debug, Clone, Copy, PartialEq, Eq\)\]\n"
        r"pub enum (?:ImageLimit|LimitKind) \{\n    /// Maximum size of one BLIP.*?\n\}",
        source, re.S,
    )
    return match.group().replace("enum ImageLimit", "enum LimitKind")


def normalize(source, path, *, previous):
    # The image resource enum moves unchanged into image/model.rs. Imports
    # and public exports are checked by compilation and the rustdoc gate.
    source = re.sub(
        r"/// A bounded resource used while parsing `OfficeArt` image records\.\n"
        r"#\[derive\(Debug, Clone, Copy, PartialEq, Eq\)\]\n"
        r"pub enum (?:ImageLimit|LimitKind) \{.*?\n\}\n",
        "", source, flags=re.S,
    )
    source = re.sub(r"^\s*(?:pub(?:\([^)]*\))?\s+)?use\s+[^;]+;", "", source, flags=re.M)
    source = source.replace("pub enum Limit {", "pub enum LimitKind {")
    source = source.replace("limit: Limit,", "limit: LimitKind,")
    source = source.replace("limit: ImageLimit,", "limit: IMAGE_RESOURCE,")
    source = source.replace("limit: crate::image::LimitKind,", "limit: IMAGE_RESOURCE,")
    source = re.sub(r"\bImageLimit(?:Kind)?::", "IMAGE_RESOURCE::", source)
    if not previous and path.endswith(("image/codec.rs", "image/validation.rs", "image/tests.rs")):
        source = re.sub(r"\bLimitKind::", "IMAGE_RESOURCE::", source)
    source = re.sub(r"\bLimit::", "LimitKind::", source)
    return re.sub(r"\s+", "", source)


if __name__ == "__main__":
    proof = json.loads(Path(__file__).with_name("source-parity.json").read_text())
    old_error = subprocess.check_output(
        ["git", "show", proof["base_commit"] + ":crates/litchi-odraw/src/error.rs"], text=True
    )
    assert image_kind(old_error) == image_kind(Path("crates/litchi-odraw/src/image/model.rs").read_text())
    for row in proof["files"]:
        old = subprocess.check_output(
            ["git", "show", proof["base_commit"] + ":" + row["path"]], text=True
        )
        current = Path(row["path"]).read_text()
        before = normalize(old, row["path"], previous=True)
        after = normalize(current, row["path"], previous=False)
        assert before == after, row["path"]
        assert hashlib.sha256(after.encode()).hexdigest() == row["normalized_sha256"]
    print(f"All {len(proof['files'])} changed ODraw sources match after the declared naming normalization.")
