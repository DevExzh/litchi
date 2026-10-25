# 0362ul89anc — PPTX notes-memo process-boundary ABBA closure

This closure is based on the 89anb validator packet in
`uo2/pptx-notes-perf-89an`. It remains harness-only: no Cargo, build, native,
benchmark, public API, or production dependency change was made.

The dedicated validator now closes the cross-report evidence boundary before
writing its summary. For each of the three probe rows it requires all four
legs to have identical output byte/hash identity, identical semantic/raw
fidelity evidence, and identical source inventory (member/payload, relationship,
master/theme, and slide/notes-root counts). The source inventory is also
required to remain identical between all three operation rows. The summary
retains the output identity, source inventory, and complete first-leg fidelity
evidence instead of reducing these checks to a boolean claim.

`tools/run_pptx_notes_memo_abba.py` adds the process boundary. A strict
manifest admits exactly `A1, B1, B2, A2`, requires absolute unique report
paths under a fresh output directory, launches each command as a separate
process without a shell, records PID plus `/proc` start-time and executable
identity, bounds timeout/report/log sizes, requires unique report content, and
validates the complete finite matrix through the dedicated validator. Its
create-new receipt preserves process identities, commands, report paths and
hashes, source revisions, binary identities, and the validated ABBA summary.
The runner requires at least four four-sample rounds and rejects unknown modes,
manifest fields, reused paths, reused report content, and incomplete output
directories.

Allocator validation now checks exact mode-specific allocator identity and
counter revision, positive warmup, unsigned-64-bit numeric bounds, live-byte
arithmetic, allocation/reallocation count ordering, and peak/live/region peak
bounds. Binary paths must be absolute, mode bits must fit executable metadata,
and paired ABBA legs must use the same complete binary identity, including the
path, while control and candidate hashes must differ.

Validation performed:

```text
PYTHONWARNINGS=error python3 -m unittest tools/test_pptx_notes_memo_abba.py tools/test_run_pptx_notes_memo_abba.py  # 18 passed
python3 -m py_compile tools/pptx_notes_memo_abba.py tools/test_pptx_notes_memo_abba.py tools/run_pptx_notes_memo_abba.py tools/test_run_pptx_notes_memo_abba.py
rustfmt --edition 2024 --config-path rustfmt.toml --check tools/perf-baseline/src/pptx_notes_memo.rs
```

No performance or qualification result is claimed. Any later measurement
must use fresh endpoint binaries and this process runner under the approved
quiet ABBA protocol.
