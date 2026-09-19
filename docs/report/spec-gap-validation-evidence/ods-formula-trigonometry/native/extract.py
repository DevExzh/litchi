#!/usr/bin/env python3
"""Extract scalar cached-result observations from LibreOffice's function corpus.

Usage: python3 extract.py /path/to/libreoffice-core [output-directory]
Inputs are never modified. The selected observations are not a native resave test.
"""
import hashlib
import json
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET

UPSTREAM_COMMIT = "d804d6aff49054bad1719ec3c2d136b545bbc7e7"
NAMES = "ACOS ACOSH ACOT ACOTH ASIN ASINH ATAN ATAN2 ATANH COS COSH COT COTH CSC CSCH SEC SECH SIN SINH TAN TANH DEGREES RADIANS PI".split()
TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"


def main():
    source = Path(sys.argv[1]).resolve()
    output = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).resolve().parent
    output.mkdir(parents=True, exist_ok=True)
    inputs, rows, excluded = {}, [], []
    for name in NAMES:
        relative = f"sc/qa/unit/data/functions/mathematical/fods/{name.lower()}.fods"
        data = (source / relative).read_bytes()
        inputs[relative] = hashlib.sha256(data).hexdigest()
        for table in ET.fromstring(data).iter(TABLE + "table"):
            row_index = 1
            for row in table.findall(TABLE + "table-row"):
                column = 1
                for cell in row:
                    formula = cell.get(TABLE + "formula", "")
                    calls = re.findall(r"([A-Za-z][A-Za-z0-9_.]*)\s*\(", formula)
                    if (calls and calls[0].upper() == name and "[" not in formula
                            and not set(c.upper() for c in calls) - set(NAMES) - {"TRUE", "FALSE"}
                            and cell.get(OFFICE + "value-type") == "float"):
                        observation = {"function": name, "source": relative,
                                       "sheet": table.get(TABLE + "name"),
                                       "row": row_index, "column": column,
                                       "formula": formula, "cached": cell.get(OFFICE + "value")}
                        if formula == "of:=ATAN2(0;0)":
                            observation["reason"] = "OpenFormula permits zero or an error; this evaluator selects an error."
                            excluded.append(observation)
                        else:
                            rows.append(observation)
                    column += int(cell.get(TABLE + "number-columns-repeated", "1"))
                row_index += int(row.get(TABLE + "number-rows-repeated", "1"))
    assert len(rows) == 115 and len(excluded) == 1
    assert {row["function"] for row in rows} == set(NAMES)
    expected = json.loads((Path(__file__).resolve().parent / "provenance.json").read_text())
    assert inputs == expected["inputs"], "inputs differ from the pinned upstream fixture hashes"
    receipt = {"upstream": "https://github.com/LibreOffice/core",
               "commit": UPSTREAM_COMMIT,
               "inputs": inputs, "excluded_profile_variances": excluded}
    (output / "cached-results.json").write_text(json.dumps(rows, indent=2) + "\n")
    (output / "provenance.json").write_text(json.dumps(receipt, indent=2) + "\n")


if __name__ == "__main__":
    main()
