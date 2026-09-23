# Change 0747 evidence packet

Record: [0747](../../0747-xlsx-publication-audit-reuse.md). Base `009d515bef`;
candidate commits `1ccfe6b354` (pair audit) and `6b49ce999a` (observer gate).
`performance_claim: none`.

Everything here was produced on the recorded host with every measured process
pinned to CPU 24. Binaries, perf data, raw callgrind profiles, captured payload
bytes and corpora are not kept; their identities are.

| path | what it is |
| --- | --- |
| `binaries.txt` | SHA-256 and provenance of every harness and probe binary used |
| `timing/final/` | the reported campaign: self-built base (`0578e443…`) vs final candidate (`5c63a831…`), 9 cases × 4 rounds × A B B A. `analysis.json` (per-process p50/p95/mean, medians, paired ratios, bootstrap CI, output digests, phase medians), `flags.json` (every paired comparison over 5%), `status.txt` (144 exits), `raw/` (every harness JSON report and log, gzip) |
| `timing/followup/` | the two noisiest controls again, 8 rounds (16 processes per leg): eager one-edit medium and pptx one-edit |
| `timing/preliminary/` | the first campaign, coordinator's prebuilt base (`fb535ebb…`) vs first candidate (`52528d5b…`): same headline, plus 2.7–3.4% control shifts that same-flag builds remove; retained unedited |
| `probe/first/`, `probe/final/` | unit probe CSVs (A = base-code probe, B = candidate probe), 4 rounds × 2 shapes × A B B A; `abba.summary.txt` / `abba2.summary.txt` are the per-measure medians |
| `callgrind/` | per-measured-iteration inclusive instructions and call counts from isolation pairs (`--samples 1` vs `3`): `report-oneedit.json` (self-built base vs final candidate, both shapes), `report-eager.json` (eager dense-sparse control), `report-prebuilt-base.json` (coordinator's base, for the planning-variation note) |
| `capture/` | the gdb capture of every `verify_with_policy` call in one run per shape: backtraces (`*-cap.log.gz`) and SHA-256 plus length of each audited payload (`*-payloads.txt`); hits 44–47 are the measured publication |
| `differential/` | the release-mode differential campaigns: `campaign/` 8 × 5,000,000 cases on the first candidate, `campaign2/` 8 × 3,000,000 on the final one; per-seed accepted / window / refused counts, all exits 0 |
| `scripts/` | every script used: timing runner and analysis, flag scan, probe source and runners, callgrind runners and parser, gdb capture scripts, differential campaign runner, leg build script, gates script |
| `gates.txt` | the gate commands, test totals and exit codes on the committed candidate |
| `cleanup.json` | what was deleted after the evidence was copied here, with sizes |
| `log-sections.md` | ready-to-paste sections for `HOTSPOTS.md`, `REPORT.md` and `GOAL_AUDIT.md` |

Replaying the analysis: `python3 scripts/analyze.py <dir>` over a directory whose
`raw/` holds the decompressed JSON reports reproduces `analysis.json`;
`python3 scripts/flags.py analysis.json flags.json` reproduces the flag list;
`python3 scripts/summarize_probe.py probe/final` reproduces the probe medians.
