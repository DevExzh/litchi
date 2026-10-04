"""Export a source-bound selector census; this does not execute any workload."""
import argparse
import hashlib
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[4]


def census():
    path = ROOT / "tools/perf-baseline/src/lib.rs"
    raw = path.read_bytes()
    source = raw.decode()
    enum = source.split("enum Case {", 1)[1].split("\n}\n\nimpl Case", 1)[0]
    variants = re.findall(r"^\s*([A-Z]\w*),$", enum, re.M)
    implementation = source.split("impl Case {", 1)[1]
    default = implementation.split("const DEFAULT:", 1)[1].split("];", 1)[0]
    defaults = re.findall(r"Self::(\w+)", default)
    declared_default = int(re.search(r"\[Self; (\d+)\]", default).group(1))
    names_body = implementation.split("const fn name(self)", 1)[1].split(
        "const fn uses_synthetic_cfb", 1)[0]
    names = re.findall(r'Self::(\w+)\s*=>\s*(?:\{\s*)?"([^"]+)"', names_body)
    parse_body = source.split("fn parse_case(value: &str)", 1)[1].split("\nfn parse_shape", 1)[0]
    parsed = re.findall(r'"([^"]+)"\s*=>\s*(?:\{\s*)?Some\(Case::(\w+)\)', parse_body)
    if (len(variants) != len(set(variants)) or len(names) != len(variants)
            or len(parsed) != len(variants) or len(dict(names)) != len(variants)
            or len(dict(parsed)) != len(variants) or set(dict(names)) != set(variants)
            or {(variant, name) for name, variant in parsed} != set(names)
            or len(defaults) != declared_default or len(set(defaults)) != len(defaults)
            or not set(defaults) <= set(variants)):
        raise ValueError("selector grammar or enum/name/parser/default registry changed")
    rows = [{"variant": variant, "selector": name, "default": variant in defaults,
             "name_line": source[:source.index("Self::" + variant + " =>", source.index("const fn name(self)", source.index("impl Case {")))].count("\n") + 1}
            for variant, name in names]
    return {"schema": "litchi.selector-census-0906.v1", "status": "pass",
            "source": {"path": str(path.relative_to(ROOT)), "bytes": len(raw),
                       "sha256": hashlib.sha256(raw).hexdigest()},
            "selector_count": len(rows), "default_count": len(defaults),
            "family_prefix_counts": {family: sum(r["selector"].startswith(family + "_") for r in rows)
                                     for family in ["doc", "xls", "ppt", "docx", "xlsx", "pptx", "xlsb"]},
            "selectors": rows,
            "scope": "Lexical enum/name/CLI/default bijection under the checked source grammar. "
                     "Counts are selectors, not independently measured scenarios, supported CRUD directions, "
                     "corpora or result rows. Family prefixes do not include generic substrate cases. "
                     "No binary, workload, build, or performance measurement is executed."}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = census()
    with args.output.open("x") as out:
        json.dump(result, out, indent=2, sort_keys=True)
        out.write("\n")
