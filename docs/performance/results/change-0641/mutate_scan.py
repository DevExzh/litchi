"""Scan-level mutations, checked against the whole litchi-xls suite."""
import subprocess, os

ROOT = "/home/zhuhe/code/litchi-worktrees/0641"
MUT = [
 ("N1 the position peek reads the XF index where the column is",
  "crates/litchi-xls/src/workbook/source.rs",
  "            u16::from_le_bytes([head[2], head[3]]),",
  "            u16::from_le_bytes([head[0], head[1]]),"),
 ("N2 accept_measured skips the XF validation",
  "crates/litchi-xls/src/workbook/source.rs",
  """        scan.formatting
            .validate_cell_xf(measured.xf_index)
            .map_err(SourceBackedError::Parse)
    }""",
  """        let _ = (scan, measured);
        Ok(())
    }"""),
 ("N3 the measure path skips validation entirely",
  "crates/litchi-xls/src/workbook/source.rs",
  """                if !wants_record(sink, payload) {
                    let measured = CellRecord::measure(frame.kind, payload, &owner.encoding)
                        .map_err(SourceBackedError::Parse)?;
                    sink.accept_measured(&measured, &context)?;
                    continue;
                }
                let cell = CellRecord::parse(frame.kind, payload, &owner.encoding)
                    .map_err(SourceBackedError::Parse)?;
                if matches!(""",
  """                if !wants_record(sink, payload) {
                    let _ = CellRecord::measure(frame.kind, payload, &owner.encoding);
                    continue;
                }
                let cell = CellRecord::parse(frame.kind, payload, &owner.encoding)
                    .map_err(SourceBackedError::Parse)?;
                if matches!("""),
 ("N4 the worksheet chain hint is shared with the shared-string resolver",
  "crates/litchi-xls/src/workbook/source.rs",
  """            let mut sheet_chain = self.inner.cfb.chain_hint();
            for sheet in self""",
  """            let mut sheet_chain = strings.chain;
            for sheet in self"""),
]

for name, rel, old, new in MUT:
    path = os.path.join(ROOT, rel)
    original = open(path).read()
    if original.count(old) != 1:
        print("%-62s NOT-APPLIED (%d anchors)" % (name, original.count(old)))
        continue
    open(path, "w").write(original.replace(old, new, 1))
    try:
        proc = subprocess.run(
            ["cargo", "test", "-p", "litchi-xls", "--all-features"],
            cwd=ROOT, capture_output=True, text=True)
        out = proc.stdout + proc.stderr
        if "error[" in out or "error: could not compile" in out:
            verdict = "DID NOT COMPILE"
        else:
            failed = sorted({l.split("...")[0].replace("test ","").strip()
                             for l in out.splitlines() if l.startswith("test ") and "FAILED" in l and not l.startswith("test result:")})
            verdict = ("caught by %d test(s): " % len(failed)) + ", ".join(failed[:4]) if failed else "SURVIVED"
        print("%-62s %s" % (name, verdict))
    finally:
        open(path, "w").write(original)
