# 0556 XLSX sorted provenance merge preparation

Status: candidate correctness checks passed; production source restored.
Performance measurement and adoption remain pending. OLE2 and OOXML have priority; ODF is
deferred and iWork is excluded from this coordinator's work.

The isolated candidate replaces append-then-sort in `Store::merge_omitted_cells`
with a checked linear merge. It shares the existing index/extent assembler
with the ordinary sorted-and-validated constructor, moves parsed entries, and
clones selected source entries. Complete XML validation, reduced parsing,
optional refusal, complete-parse fallback, and parsed worksheet metadata remain
authoritative. The parsed iterator is explicitly dropped before rebuilding
indexes; overlap during the merge still requires peak-memory measurement.

The [source proof](source-proof.md), [independent candidate review](candidate-review.md),
and [prospective measurement plan](measurement-plan.md) are separate evidence.
Direct tests compare the original Store merge algorithm and every retained
field; existing snapshot tests compare complete XML parsing and preservation.
The retained 0550 merge attribution motivates the candidate but cannot prove
that it improves latency or memory. No native, allocation, or instruction
capture is made in this preparation batch.

## Artifacts and reproduction

`baseline-cell.rs`, `candidate/cell.rs`, and `candidate/candidate.patch` retain
the exact source alternatives. Source manifests bind workspace crates and tools;
`environment.json` identifies the toolchain and the byte-identical retained
workspace lock in change 0555. The thirty accepted-ADR/index files were rehashed
and match the ADR manifest read earlier in this session.

`quality-plan.json` and `quality.py` define the serial checks. For a new attempt,
create the owned target and its `tmp` child, restore the corresponding source
manifest exactly, and invoke:

```text
python3 -B docs/performance/results/change-0556/quality.py baseline NEW-LABEL targeted
python3 -B docs/performance/results/change-0556/quality.py baseline NEW-LABEL-2 commands
```

Use stage `candidate` only after checking and applying the retained patch to
the matching baseline. Each label must be new; the driver refuses to overwrite
an attempt and checks source and lock identity before and after each command.
Wait for each command to terminate before editing source or starting the next
one. Preserve the original source and restore it after temporary candidate
validation. Future performance captures belong in a new bundle with a newly
frozen measurement plan; this packet's prospective plan has explicit unresolved
noise, allocation-scope, process-isolation, and symbol-mapping decisions.

`candidate-01` passed tests and Clippy but failed formatting. Its original
source, patch, manifest, and review are retained under
`candidate-attempts/before-format`. `format-correction.json` binds the actual
formatter invocation; the production-code prefix is byte-identical. The
premature `candidate-02` overlapped formatting and failed its source guard;
`aborted.json` marks its outputs inadmissible. That attempt supplies no passing
quality evidence. A fresh attempt validates the exact corrected source.

The adapted quality driver's `ole2_0556_*` schema names are inherited labels;
its command list, source manifests, and receipts specify the actual XLSX scope.
No fuzz, sanitizer, native-producer capture, cold-cache, or scaling result is
claimed by these preparation checks.

## Correctness results

| Source | XLSX tests | Clippy and format | Workspace features, XLSX rustdoc, minimal features |
| --- | --- | --- | --- |
| Baseline | 1,313 passed, 59 groups, zero failed/ignored | Passed | Passed |
| Formatted candidate | 1,316 passed, 59 groups, zero failed/ignored | Passed | Passed |

The successful candidate attempts are `candidate-03` and `candidate-04`.
`restoration.json` binds their results and verifies every baseline source file
after restoration. The three added differential/refusal tests supplement the
existing public preservation, complete-parser, patch, and source-sharing tests.
`final-checks` records repository boundaries, claim validation, and formatting
against the restored source. `decision.json` keeps adoption false.

After terminal cleanup, verify this packet with:

```text
python3 -B docs/performance/results/change-0556/verify.py --cleaned --sealed
```

The recursive SHA-256 inventory includes failed attempts and reviews. The
verifier rechecks source identities, exact patch reconstruction, successful and
failed command receipts, test counts, the formatting correction, and the
inadmissible overlapping attempt. It verifies preparation evidence only; it
cannot substitute for future measured admission or completion of `docs/GOAL.md`.
