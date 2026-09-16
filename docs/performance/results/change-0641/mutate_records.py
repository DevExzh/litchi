"""Apply one mutation to the merged tree, run the measure-only tests, revert."""
import subprocess, sys, shutil, os, json

ROOT = "/home/zhuhe/code/litchi-worktrees/0641"
MUT = [
 ("M1 measure_label drops the string refusal instead of reporting it",
  "crates/litchi-xls/src/records.rs",
  "        utils::measure_string_record(&data[6..], encoding)?;",
  "        let _outcome = utils::measure_string_record(&data[6..], encoding);"),
 ("M2 the refusal is reported with decode_utf16's wording, not from_utf16's",
  "crates/litchi-xls/src/utils.rs",
  "        return parse_string_record(data, encoding).map(|_text| ());",
  "        let _ = encoding;\n        return Err(Error::Encoding(\"UTF-16 decoding error: unpaired surrogate found\".to_string()));"),
 ("M3 measure_formula skips validate_formula_extra",
  "crates/litchi-xls/src/formula_metadata/codec.rs",
  "    validate_formula_extra(framed.tokens, framed.extra)?;",
  "    let _ = validate_formula_extra(framed.tokens, framed.extra);"),
 ("M4 measure_formula skips decode_flags",
  "crates/litchi-xls/src/formula_metadata/codec.rs",
  "    decode_flags(framed.flags, framed.tokens)?;",
  "    let _ = framed.flags;"),
 ("M5 measure_formula skips the empty-token checks",
  "crates/litchi-xls/src/formula_metadata/codec.rs",
  "    check_token_stream(&framed)?;\n    decode_flags(framed.flags, framed.tokens)?;",
  "    decode_flags(framed.flags, framed.tokens)?;"),
 ("M6 LabelSst measures against the wrong minimum length",
  "crates/litchi-xls/src/records.rs",
  "            0x00FD => Self::measure_fixed(data, 10),       // LabelSst",
  "            0x00FD => Self::measure_fixed(data, 6),        // LabelSst"),
 ("M7 wants_record answers false for a payload too short to carry a position",
  "crates/litchi-xls/src/workbook/source.rs",
  "    payload.get(..4).is_none_or(|head| {",
  "    payload.get(..4).is_some_and(|head| {"),
 ("M8 measure_label reports the wrong column",
  "crates/litchi-xls/src/records.rs",
  "        let (row, col, xf_index) = Self::cell_head(data, 8)?;\n        utils::measure_string_record(&data[6..], encoding)?;",
  "        let (row, col, xf_index) = Self::cell_head(data, 8)?;\n        let col = col.wrapping_add(1);\n        utils::measure_string_record(&data[6..], encoding)?;"),
]

results = []
for name, rel, old, new in MUT:
    path = os.path.join(ROOT, rel)
    original = open(path).read()
    if original.count(old) != 1:
        results.append((name, "NOT-APPLIED (anchor count %d)" % original.count(old)))
        continue
    open(path, "w").write(original.replace(old, new, 1))
    try:
        proc = subprocess.run(
            ["cargo", "test", "-p", "litchi-xls", "--all-features", "--lib",
             "cell_measure_tests", "--", "--nocapture"],
            cwd=ROOT, capture_output=True, text=True)
        out = proc.stdout + proc.stderr
        if "error[" in out or "error: could not compile" in out:
            verdict = "DID NOT COMPILE"
        else:
            failed = [l.split("...")[0].strip() for l in out.splitlines()
                      if l.startswith("test ") and "FAILED" in l and not l.startswith("test result:")]
            verdict = ("caught by " + ", ".join(sorted(f.replace("test records::cell_measure_tests::","") for f in failed))) if failed else "SURVIVED"
        results.append((name, verdict))
    finally:
        open(path, "w").write(original)

for name, verdict in results:
    print("%-70s %s" % (name, verdict))
