# 0554 harness audit

`status: read-only harness audit`

`scope: OLE2/OOXML active; ODF deferred; iWork excluded`

`performance_claim: none`

This audit records the smallest reproducible way to run the frozen 0554 N
name-handoff experiment.  It does not alter Rust, the driver, the plan, a
binary, or a capture.  The frozen driver is already the 0547/0548/0549 serial
driver with the 0554 target and lock bindings.  The old campaigns are useful
as parser and protocol references; their campaign-specific verifiers must not
be used as the 0554 verifier.

## Reusable pieces

The following existing pieces are sufficient for the campaign:

| Need | Reusable piece | Boundary |
| --- | --- | --- |
| Build, copy, and bind binaries | `change-0554/run.py` | 0554 owns the frozen copy; it is derived from the serial `change-0547/run.py` shape. |
| CFB/XLS corpus generation and timers | `tools/perf-baseline` selectors used by `run.py` | `cfb_open` takes `--shape tiny`, `many-small`, or `few-large` with `--payload incompressible`; the XLS row uses the fixed generated manifest. |
| Raw Callgrind function/edge parsing | Immutable parser imported by `change-0554/analyze_profiles.py` from `change-0536/analyze_profiles.py` | Only the raw parser and positive-edge logic are reused; 0536/0547 plan, target, cleanup, and verifier globals are not. |
| Profile protocol | `change-0547` and `change-0548` `analyze_profiles.py` / `protocol.md` | Owner ancestry and setup classification are retained; 0554 adds name-decoder and scalar-field rows. |
| Native and allocator command shape | `change-0548/run.py` and `change-0549/run.py` | The 0554 frozen `run.py` is authoritative for paths, source stage, lock, and environment. |
| Static instruction follow-up | `change-0547/inspect_assembly.py` and `instruction_analysis.py` as references | Their 0547 collector-sector assumptions do not prove N; inspect decoder, scalar, and drop symbols for 0554. |

The old `change-0548` and `change-0549` verifiers contain hard-coded plan
schemas, target paths, cleanup shapes, and admission fields.  Copying those
verifiers wholesale can silently validate the wrong campaign.  A 0554
checker must bind the current plan, stage manifests, receipts, binary hashes,
raw dump hashes, and the current analyzer explicitly.

## Exact serial command sequence

Run from the repository root.  Every command is a separate serial child.  Do
not combine commands with shell separators in a receipt, and use `python3 -B`
so an analyzer run cannot create a `__pycache__` in the evidence bundle.

The baseline source is frozen and built before the candidate patch is applied
to the private candidate source.  The baseline profile and native/allocation
captures must finish before candidate source is made live.  The candidate
stage uses the reviewed N patch from the frozen plan; the source transition is
owned by the coordinator and is not part of this audit.

```sh
python3 -B docs/performance/results/change-0554/run.py freeze
python3 -B docs/performance/results/change-0554/run.py build-normal \
  --stage baseline --execution-stage baseline
python3 -B docs/performance/results/change-0554/run.py build-alloc \
  --stage baseline --execution-stage baseline
```

Capture the normal baseline in the first ABBA lane.  The first and second
profile commands produce four jobs each.  In the historical 0547/0548/0549
shape, a CFB job has a part 1 corpus-construction setup open followed by
parts 2--6 for five positive timed opens, while the XLS job has five positive
timed parts and no CFB setup open.  0554 must classify each fresh dump from
its actual positive ancestry and retain any different shape as evidence; it
must not assign role from a numbered suffix.

```sh
python3 -B docs/performance/results/change-0554/run.py profile \
  --stage baseline --execution-stage baseline --repeat 1
python3 -B docs/performance/results/change-0554/run.py profile \
  --stage baseline --execution-stage baseline --repeat 2
python3 -B docs/performance/results/change-0554/run.py alloc \
  --stage baseline --execution-stage baseline --repeat 1
python3 -B docs/performance/results/change-0554/run.py alloc \
  --stage baseline --execution-stage baseline --repeat 2
```

The frozen native order is baseline r1, candidate r1, candidate r2, baseline
r2.  For baseline r2, `--stage baseline` selects the retained baseline
binary/output folder while `--execution-stage candidate` binds the live
candidate source manifest.  This is why stage and execution stage are both
recorded in every receipt; do not infer source identity from the folder name.

```sh
python3 -B docs/performance/results/change-0554/run.py native \
  --stage baseline --execution-stage baseline --repeat 1

# Apply the reviewed N source to the private candidate checkout, then build:
python3 -B docs/performance/results/change-0554/run.py build-normal \
  --stage candidate --execution-stage candidate
python3 -B docs/performance/results/change-0554/run.py build-alloc \
  --stage candidate --execution-stage candidate

python3 -B docs/performance/results/change-0554/run.py native \
  --stage candidate --execution-stage candidate --repeat 1
python3 -B docs/performance/results/change-0554/run.py native \
  --stage candidate --execution-stage candidate --repeat 2
python3 -B docs/performance/results/change-0554/run.py native \
  --stage baseline --execution-stage candidate --repeat 2
```

If profiles are captured as a separate lane, run them after the candidate
normal build with the same stage bindings:

```sh
python3 -B docs/performance/results/change-0554/run.py profile \
  --stage candidate --execution-stage candidate --repeat 1
python3 -B docs/performance/results/change-0554/run.py profile \
  --stage candidate --execution-stage candidate --repeat 2
```

The allocator executable is already copied by `build-alloc`; allocator
captures use the same group matrix as the native lane, with two repeats, three
warmups, and 30 samples.  Native captures use the same matrix with 20
warmups and 1,000 samples.  The profile lane uses zero warmups and five
samples per selected job, with Callgrind flags frozen as
`--dump-instr=yes --dump-line=no --compress-pos=no --collect-jumps=yes`.

After all captures terminate, run the standalone profile analyzer.  It is
read-only with respect to captures and refuses to overwrite a non-identical
report:

```sh
python3 -B docs/performance/results/change-0554/analyze_profiles.py \
  --stage baseline \
  --output docs/performance/results/change-0554/baseline/profile-analysis.json
python3 -B docs/performance/results/change-0554/analyze_profiles.py \
  --stage candidate \
  --output docs/performance/results/change-0554/candidate/profile-analysis.json
python3 -B docs/performance/results/change-0554/analyze_profiles.py \
  --compare \
  --output docs/performance/results/change-0554/profile-comparison.json
```

For a focused human inspection, the deterministic environment and existing
Callgrind tool are sufficient.  This command is diagnostic output only; save
its output as a separately receipted artifact if it is retained:

```sh
env PERL_HASH_SEED=0 PERL_PERTURB_KEYS=0 \
  callgrind_annotate --auto=no --threshold=100 --show-percs=no \
  --inclusive=yes --tree=both \
  docs/performance/results/change-0554/candidate/profile-r1-cfb-many-small.callgrind.2
```

## Owner and mechanism evidence

The analyzer must classify each numbered dump from its positive incoming edge
and runner ancestry.  Dump ordinal is only a consistency check.  The selected
owners and runners are:

| Workload | Owner | Positive timed runner | CFB setup callers |
| --- | --- | --- | --- |
| CFB tiny/many-small/few-large | `litchi_cfb::file::OleFile<R>::open` | `litchi_perf_baseline::run_cfb_open` | `litchi_perf_baseline::build_cfb_corpus`, or the inlined `litchi_perf_baseline::run::{{closure}}` |
| XLS owned one-cell | `litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits` | `litchi_perf_baseline::run_xls_owned_source_case` | none |

For every numbered dump, retain the owner caller, owner call count, incoming
Ir, selected owner self/direct Ir, caller ancestry, raw dump hash, trigger,
part number, and role.  The owner edge must have one positive call and its
inclusive Ir must equal the raw `summary` for the selected constructor.  An
XLS owner whose immediate caller is
`SourceBackedWorkbook::from_read_at` is valid only when that wrapper has a
positive ancestry path to the XLS benchmark runner.  A CFB setup dump may have
the inline `run::{{closure}}` caller, so it must not be discarded because the
named `build_cfb_corpus` symbol is absent.

The per-dump mechanism rows must include separate attributions for:

```text
validated_directory_entries
directory_name_data
parse_directory_entry
decode_utf16le
format_clsid
build_storage_tree_iterative
DirectoryNameData drop paths (when emitted)
```

For `parse_directory_entry -> decode_utf16le`, retain every positive direct
edge with caller, callee, calls, and inclusive Ir.  The standard-root expected
call counts are four for CFB tiny, 257 for CFB many-small, five for CFB
few-large, and 12 for the fixed XLS owner corpus in the retained baseline.
The N candidate prediction is exactly one positive standard-root decoder call
per timed open: the classic root remains the one public decode, while normal
validated names are moved.  The analyzer reports the prediction for every
timed dump and makes any mismatch visible rather than converting a missing
symbol into zero work.

`format_clsid` and the other scalar parse rows are controls.  A candidate
profile may move or inline them, but the report must say `out_of_line`,
`inlined_or_absent`, or `no_positive_incoming_edge` and retain the function IDs,
self Ir, direct Ir, and incoming edges.  A lower aggregate Ir value or a
missing `parse_directory_entry`, `decode_utf16le`, or `format_clsid` symbol is
not proof of eliminated work.  The source candidate's required
`directory_name_data` validation work and its `DirectoryNameData`/public name
drop paths remain separate from the public decoder edge.

The CFB setup dump and the final `Program termination` dump are retained but
excluded from timed aggregates.  Final dumps must have `events: Ir`,
`positions: instr`, and zero summary Ir.  Collection-off call labels are
context evidence only; operation-local dynamic counts come from the selected
positive owner and direct target edges.

### Symbol-bounded follow-up

Capture static bounds from the exact retained normal binary after its build,
before changing the source, and repeat for the candidate normal binary after
the candidate build.  The existing `change-0547/inspect_assembly.py` pattern
is the reusable command wrapper: it runs `nm -S --defined-only`, retains only
text symbols, and invokes `objdump -d --disassemble=<exact-symbol>` once per
selected symbol through the same source/binary-bound receipt helper.  For N,
the owner fragments to select are:

```text
OleFile$LT$R$GT$21parse_directory_entry
OleFile$LT$R$GT$27validated_directory_entries
OleFile$LT$R$GT$28build_storage_tree_iterative
parse_validated_directory_entry
decode_utf16le
format_clsid
directory_name_data
DirectoryNameData
drop_in_place$LT$litchi_cfb..directory_name..DirectoryNameData
drop_in_place$LT$alloc..vec..Vec$LT$core..option..Option$LT$litchi_cfb..directory_name..DirectoryNameData
```

The frozen 0554 wrapper applies that selection to each normal binary with
these serial commands (the candidate command follows candidate application
and build):

```sh
python3 -B docs/performance/results/change-0554/inspect_assembly.py baseline
python3 -B docs/performance/results/change-0554/inspect_assembly.py candidate
```

Once both assembly indexes and all numbered profile dumps exist, use the
standalone bounded mapper.  It validates the exact stage assembly receipts,
finds the single relocation bias from all selected target positions, and
retains each target's raw/static address, symbol offset, bytes, text, and
self instruction Ir.  The comparison reports per-dump decoder calls and
parser/`format_clsid` scalar control calls alongside the mapped instruction
totals:

```sh
python3 -B docs/performance/results/change-0554/instruction_analysis.py \
  --stage baseline \
  --output docs/performance/results/change-0554/baseline/instruction-analysis.json
python3 -B docs/performance/results/change-0554/instruction_analysis.py \
  --stage candidate \
  --output docs/performance/results/change-0554/candidate/instruction-analysis.json
python3 -B docs/performance/results/change-0554/instruction_analysis.py \
  --compare \
  --output docs/performance/results/change-0554/instruction-comparison.json
```

The mapper accepts baseline-r2's deliberate ABBA receipt binding: its output
folder and normal binary remain baseline while its receipt records
`execution_stage: candidate`.  That exception is checked explicitly by the
current plan/repeat, and does not permit a folder-name-only source or binary
substitution.  Its instruction totals add only exclusive function self Ir;
inclusive parent/callee Ir and edge call counts remain separate.  A missing
or inlined symbol is reported as `absent_or_inlined` or
`partial_static_mapping`, and the report cannot interpret either state as
eliminated work without owner/caller disassembly.

The Rust symbols are normally mangled, so filter the raw `nm -S` output by
these fragments (for example `decode_utf16le`), retain every matching `t` or
`T` symbol and its address/size, then disassemble each exact symbol.  A
minimal manual check has this shape, with the actual binary and output paths
bound to the stage receipt:

```sh
nm -S --defined-only /home/zhuhe/litchi-goal-0554-target/retained/baseline/normal \
  > docs/performance/results/change-0554/baseline/symbols.stdout
objdump -d --disassemble=_ZN10litchi_cfb4file14decode_utf16le17...E \
  /home/zhuhe/litchi-goal-0554-target/retained/baseline/normal \
  > docs/performance/results/change-0554/baseline/assembly-decode-utf16le.stdout
```

The ellipsis above is descriptive: use the complete symbol copied from the
`nm` row, including its hash, and issue one `objdump` invocation per matching
monomorphization.  The retained wrapper should record the exact argv, binary
hash, source manifest hash, stdout/stderr hashes, and symbol address/size.  Do
not use a wildcard as the `objdump --disassemble` value.

Callgrind's `positions: instr` records can then be relocated against the
matching static symbol address and size.  For each stage, retain instruction
work for the validator/name path (`validated_directory_entries`,
`directory_name_data`), public parse (`parse_directory_entry`), decoder,
`format_clsid`, and all emitted owner/drop code.  If a symbol is absent or a
helper is inlined into a caller, report that state and attribute the relevant
instructions in the caller's disassembly.  The absence of a symbol, a smaller
symbol, or a lower aggregate Ir total by itself proves none of code removal,
latency improvement, or allocation improvement.  `--dump-line=no` also means
that source line labels cannot substitute for instruction-address mapping.

## Allocation and native controls

The allocation report must compare, for every operation row, status and the
allocation/deallocation/reallocation/failed-call counts, allocated and
deallocated bytes, live bytes before/after, absolute peaks before/after, and
the operation region peak.  `Unavailable` and `Overflow` are explicit states,
not zeros.  Region peak is allocator evidence and does not establish process
RSS.

Native admission uses p50 and mean for CFB tiny, many-small, and few-large plus
the nine XLS rows from the frozen `groups.xls` matrix.  The required primary
gate is CFB many-small p50 and mean improving by at least three percent in
both repeats.  The frozen `admission-supplement.json` independently requires
each of `xls_source_backed_open`, `xls_source_backed_open_one_cell`,
`xls_owned_source_open`, and `xls_owned_source_open_one_cell` to improve p50 by
at least three percent in both paired repeats.  Every XLS/CFB p50, mean, and
process peak RSS increase must be at most five percent in each paired repeat.
The profile mechanism gate is separate: all candidate standard-root decoder
rows must show exactly one positive direct call, and many-small owner Ir must
decrease in both repeats.  No Callgrind, RSS, hardware, or allocation row
alone is a speedup claim.

## Custody and verifier pitfalls

* Bind each report to the current `plan.json`, `run.py`, stage source manifest,
  workspace `Cargo.lock` hash, selected binary metadata, and every raw output
  and dump hash.  The 0554 driver puts its owned temporary directory under
  `/home/zhuhe/litchi-goal-0554-target/tmp`; do not substitute `/tmp`.
* Keep baseline r2's source/binary distinction explicit: its output lives in
  `baseline`, its binary is the retained baseline normal executable, and its
  receipt must record `execution_stage: candidate`. A folder-name-only check
  can accept the wrong source or binary.
* The driver uses exclusive output creation. Preserve failed build/capture and
  analyzer attempts under an attempt directory; never rerun into a canonical
  path or replace a failed receipt with a later result.
* Use `python3 -B` for all Python helpers and scan for bytecode before sealing.
  A generated `__pycache__` is an owned artifact that must be recorded and
  removed before the final recursive hash.
* A host compiler-process observation covers only accessible `/proc` entries;
  it does not prove machine quiescence. Capture children remain serial, and
  unrelated host activity must not be killed or silently omitted.
* Callgrind's setup open can be the first numbered CFB dump, and inlining can
  remove a named helper. Classify by positive ancestry and report explicit
  absence. Never infer a timed operation from `.callgrind.2` alone or infer
  eliminated work from a vanished symbol.
* `decode_utf16le -> try_reserve` is an outgoing allocator edge. It is not a
  count of decoder calls. The decoder call count is the incoming
  `parse_directory_entry -> decode_utf16le` edge.
* Preserve the classic-Mac root two-view fixture and malformed/error
  differential separately. The synthetic CFB selectors do not exercise the
  compact `00 52` root form, so a normal profile cannot satisfy that proof.
* Do not relax the three-percent/five-percent gates, hide adverse rows, or
  use the 0553 rejected XLSX result as a candidate/baseline control. Failed N
  evidence remains useful diagnostic evidence and must be sealed with an
  explicit rejection decision.

This document is an audit of reusable harness boundaries and commands.  It
records no new measurement and makes no adoption or performance claim.
