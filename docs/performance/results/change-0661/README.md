# Evidence packet for change 0661

Record: [0661 — owned OPC packages defer part payload decoding](../../0661-opc-lazy-part-decode.md).

Authority: accepted [ADR 0030](../../../adr/0030-lazy-opc-part-decode.md),
change [0652](../../0652-owner-decisions-for-the-third-wave.md) decision 3,
and queue row 3 of [0651](../../0651-queue-refresh-after-the-second-wave.md).
The frozen design and sizing baseline are [0610](../../0610-opc-lazy-part-decode-design.md).

This packet reports deterministic correctness and read-set evidence. It makes
no latency, RSS, allocation, or registered performance claim.

## Contents

| path | purpose |
| --- | --- |
| `differ-eager-vs-lazy.txt` | 336-fixture eager-versus-lazy open, part catalog, exact no-op, and one-part reblob differential; 1,342 `MATCH` rows and zero `MISMATCH` rows |
| `crosscheckout/before-opc-members.txt`, `after-opc-members.txt` | metadata and payload-independent member observations from the before and after checkouts |
| `crosscheckout/before-opc-noop.txt`, `after-opc-noop.txt` | exact no-op output observations |
| `crosscheckout/before-opc-reblob.txt`, `after-opc-reblob.txt` | one-part reblob output observations |
| `crosscheckout/before-xlsx-hide.txt`, `after-xlsx-hide.txt` | XLSX hide operation outcomes |
| `readset-opc-noop.txt` | per-fixture lazy decode counters for exact no-op publication |
| `readset-opc-reblob.txt` | per-fixture lazy decode counters for one-part reblob publication |
| `readset-xlsx-hide.txt` | per-fixture lazy decode counters for the current XLSX semantic hide route |
| `probe/src/bin/lazy.rs` | source for after-leg differential and read-set probe |
| `probe/src/bin/publish.rs` | source for before/after cross-checkout probe |
| `probe/Cargo.toml.example` | replay manifest template using a checkout path dependency |
| `decision.json` | machine-readable disposition, authority, evidence, and limitations |
| `log-sections.md` | four coordinator-ready sections requested by change 0652 |

## Provenance

| | |
| --- | --- |
| base commit | `ab07e2a47` |
| branch | `perf/0661-opc-lazy-part-decode` |
| source corpus | repository's 336-fixture `test-data` selection used by the retained probe |
| build target | `/tmp/litchi-0661-target`, reused across checks; `-j2` |
| timing | not run |
| claim registry | no entry; `performance_claim: none` |

The before and after cross-checkout files are paired by operation. `cmp` exits
zero for all four pairs (`opc-members`, `opc-noop`, `opc-reblob`, and
`xlsx-hide`). The before files are retained as evidence and must remain until
the coordinator has completed the other agents' integration checks.

## Read-set totals

The three read-set files have 336 fixture rows each. Their common admitted
catalog is 5,077 parts and 45,562,463 inflated bytes across 334 successful
opens; two rows are refused while opening.

| operation | published | refused after open | open refused | decoded parts | decoded bytes |
| --- | ---: | ---: | ---: | ---: | ---: |
| `opc-noop` | 334 | 0 | 2 | 0 | 0 |
| `opc-reblob` | 325 | 9 signed-policy refusals | 2 | 334 | 389,581 |
| `xlsx-hide` | 33 | 301 | 2 | 1,050 | 17,069,415 |

For `xlsx-hide`, the all-row total includes refused non-XLSX and skipped
operations and is not a working-set percentage. Among its 33 published XLSX
fixtures, the semantic route decoded 216 of 550 parts and 6,044,263 of
8,274,037 inflated bytes. This is the observed route result; it does not
repeat 0610's earlier one-part design estimate.

## Replay

The focused contract tests are the primary replay gate:

```sh
CARGO_TARGET_DIR=/tmp/litchi-0661-target \
  cargo test -p litchi-opc --test lazy_part_decode -j2
CARGO_TARGET_DIR=/tmp/litchi-0661-target \
  cargo test -p litchi-opc -j2
```

The differential and read-set probe uses `probe/Cargo.toml.example` with the
checkout path substituted for its placeholder. `lazy` consumes the fixture
list and emits the eager/lazy comparison and read-set files. `publish` emits
the cross-checkout operation files. The exact command environment used for
the retained artifacts is intentionally kept out of the claim registry because
these artifacts are correctness evidence rather than timing measurements.

## Interpretation

An exact no-op keeps all deferred payloads cold and copies the retained source
archive. A fallible selected-part access inflates only the selected member.
The corruption fixture in `lazy_part_decode.rs` shows that a first-access ZIP
failure is stable, that exact no-op output preserves the corrupt source, and
that a targeted edit never turns the failed member into an empty generated
part. All three tests pass.

The packet does not claim that the old 0610 one-part XLSX route has been
reproduced, that lazy decoding improves wall-clock time, or that all refused
operations could be admitted. Those questions require separate measurements
or format-specific policy decisions.

The follow-up audit also has focused malformed-deferred coverage: PPTX
master/layout authoring leaves the source graph unchanged after a late decode
refusal and forces an orphan theme before linking it; XLSB threaded graph
validation propagates a corrupt root workbook and graph removal follows the
resolved root workbook when multiple XLSB parts exist.
