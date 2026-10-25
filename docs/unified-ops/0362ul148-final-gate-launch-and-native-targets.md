# 0362ul148 — guarded final gate launched; native targets repaired

Root applies 135g on the separate protocol branch and commits
`53451cc8d22499c3532502b3c986339a7206afdd`. The change supplies the absolute
external target directory to the PPT build, all seven PPT test/Clippy
commands, and the legacy RTF build. Each command now matches its recorded
endpoint binary paths. All 29 focused tests pass; independent review confirms
the limited policy/test/record diff and preserved before/after separation.
Frozen candidate C remains `f8cd398a770f7363f628be40edb3ce4219566c16`.

The final gate's v5 driver is returned after mutation tests show that altered
commands and phase-state paths were accepted. The initial v6 remains an
unexecuted diagnostic because its runtime command paths reused compiler
state. Corrected v6-final independently regenerates all phases, requires the
approved plan SHA, binds every command/environment/isolation receipt and
exact phase-state path, and checks these before and after each step. Tool,
source, lock, target, free-space and process-ownership guards remain in place.
Five source-only tests and root/independent exact packet dry-checks pass.

The scoped final gate retains the approved Cargo systemd `MemoryMax=20G`
mechanism and 18/19/20 GB target thresholds. It does not claim newly measured
process-tree RSS telemetry. Native performance protocols retain their own
requirements. Documentation rows 34 and 36 remain isolated. Tooling row 48
contains all 42 modules. Row 37 must pass before its OfficeCLI TSV is copied
and verified externally; row 38 cannot start before that preservation.

The reviewed packet is external at
`.coord/final-gate-Cf8cd398a-20261025-v6-final/` under the worktree parent:

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `final_gate_v6.py` | 23,875 | `3e7aba2bdfc65f434c8730cb5ffa12db301583eb033b6f759da9976134d9ec01` |
| `test_final_gate_v6.py` | 4,682 | `f0b2495e93289daa9b27987408c7044645b60d46ab5fcd6fc566c508b62b60e8` |
| `packet/plan.json` | 183,171 | `8059ad69f2eacd9a5b59d2b738c8b6ea86944196b261e1d946dc5e4d7ea1103b` |
| `packet/changed-rust-paths.txt` | 26,703 | `7b3fc80f16b97a26da340c65cd269304104e0ed62c43bf49105f82c66c1c2546` |

Root starts the reviewed driver from the parent worktree directory with:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -B \
  .coord/final-gate-Cf8cd398a-20261025-v6-final/final_gate_v6.py --go \
  --expected-plan-sha256 8059ad69f2eacd9a5b59d2b738c8b6ea86944196b261e1d946dc5e4d7ea1103b \
  --root uo2-coordinator-metadata-integration \
  --packet .coord/final-gate-Cf8cd398a-20261025-v6-final/packet \
  --state .coord/final-gate-Cf8cd398a-20261025-v6-final/state \
  --target .coord/targets/final-gate-Cf8cd398a-20261025-v6-final/target
```

The driver runs seven compiler supplements, full53 and five runtime
supplements in that order. The changed-path formatting row passes; the
remaining rows are running or pending. Some internal result schema labels
retain the v5 spelling, but the exact driver/plan hashes above identify this
v6-final execution. This is a launch record, not a final gate pass.

The failed raw census remains preserved under 147. Its v10 inventory-cap
adapter has six source-only tests and a real shim-admission seam; independent
review and a fresh C-bound packet precede any new GO. The single build slot
belongs to the full gate. Native binding recipes, actual builds/fidelity and
quiet measurements, raw callable review/finalization, final trunk merge and
cleanup remain open. Nothing is pushed.
