# Evidence: change 0671, DOC admission residues

Change record: [`0671-doc-admission-residues.md`](../../0671-doc-admission-residues.md).

Disposition: retained correctness fix. `performance_claim: none`. The change
keeps the normative `FBKF.ibkl` uniqueness refusal, admits one narrowly proven
PAPX alignment byte at the PAPX consumer, and exposes existing DOC leniency
through the unified facade. It closes the DOC portion of row 16 in 0651 under
0652's standing trade-offs (small breaking APIs are acceptable; correctness and
safety take priority over performance).

## Contents

| Path | What it is |
| --- | --- |
| `decision.json` | The `litchi-perf-change-decision` record and the specification-based disposition of both residues. |
| `log-sections.md` | Four paragraphs for the coordinator to merge into `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`. |
| `gates.txt` | The deterministic commands and results run in this worktree. |
| `../../0671-doc-admission-residues.md` | The full byte reading, checked-in MS-DOC section citations, implementation rule and limitations. |

## Provenance

| | |
|---|---|
| Base | `5fa92d7ce` |
| Branch | `perf/0671-doc-admission-residues` |
| Worktree | `/home/zhuhe/code/litchi-worktrees/0671` |
| Toolchain | rustc/cargo 1.95.0 |
| Build settings | `CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_BUILD_JOBS=2` |
| Corpus | 57 `.doc` files under `test-data/`; the 0640 retained scan is the corpus-level baseline |

## Witnesses

| Fixture | Before 0671 | After 0671 | Evidence |
|---|---|---|---|
| `test-data/ole/doc/watermark.doc` | typed refusal: `bookmark ibkl values must be unique and in range` | same typed refusal | MS-DOC `FBKF.ibkl` explicitly requires uniqueness; the duplicate CP does not override it |
| `test-data/poi/test-data/document/test.doc` | refusal: trailing one-byte SPRM opcode | opens and returns nonempty text | 24 PAPX entries share `cb=0`, odd valid SPRM prefix, final zero alignment byte |
| `test-data/ole/doc/duplicate-style-names.doc` | default facade refusal | opens with `TolerateStylesheetDefects` through both new facade methods | existing stylesheet leniency contract; structural defects remain fatal |

The PAPX compatibility rule is tested with both a complete odd SPRM prefix plus
zero and a nonzero malformed tail. The checked-in MS-DOC reference supports the
PAPX length arithmetic but requires whole Prl elements and does not specify this
pad; the packet records the allowance as a bounded compatibility exception. A
separate `sprm` unit test proves that the shared parser still rejects a trailing
zero when called without PAPX context.
