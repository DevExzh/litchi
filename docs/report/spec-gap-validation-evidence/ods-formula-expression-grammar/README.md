# OpenFormula expression grammar

This batch adds `codec::formula::expression::Expression`, an immutable, bounded
inspection API for ODF 1.4 Part 4 chapter 5. The baseline is
`b609407cc672896f6954491ac9501cd1a131e941`. The existing `FormulaParser`
remains available as the compatibility tokenizer.

The new tree retains precedence and associativity, prefix/postfix operators,
parentheses, function calls (including host-defined names and missing argument
slots), arrays and their row shapes, labels, named expressions, constant errors,
and bracketed references. `source()` retains the exact input; borrowed node
views expose original lexemes, typed operators/scopes, direct children and inert
reference metadata. `is_force_recalculate()` distinguishes the optional `==`
marker. The parser accepts a bare expression or the optional `=`/`==` intro;
`of:=` and `of:==` are convenience wrappers, not XML namespace resolution.

Defaults admit 1 MiB of source, 65,536 nodes, 65,536 aggregate array cells, and
256 parser recursion levels. Callers can supply `Limits`; a hard 256-level
recursion ceiling also applies. This is parser recursion depth, not the height
of a left-associated tree. Node/edge/reference arenas and fallible growth keep
ownership flat, so destruction does not recurse through long chains. Child
lookup is constant-time and child iteration allocates nothing. Strings and
names borrow spans of the single owned source; references occupy a separate
arena instead of inflating every node. Ragged arrays retain their grammar shape
without claiming that an evaluator can compute them.

The batch also corrects reference and name whitespace admission. Section 5.14
allows the four formula whitespace characters between grammar components:
`$$ Name`, `'Sheet' . Name`, `[ . $A$1 ]`, and `[.A 1]` parse. Lexical splits such
as `$ $Name`, `SUM (1)`, `[.$ A1]`, and `[.A$ 1]` remain invalid. Only SPACE, TAB,
LF and CR are ignored; quote contents remain data. This supersedes the earlier
reference batch's blanket rejection of bracket-interior whitespace. Public
reference wrapper trimming now uses those four characters rather than Unicode
`trim`. XML embedding character validation remains a separate layer.

[Specification requirements](specification.json) and the independent
[final specification/code review](spec-review.md) record the normative decisions.
XML name classes use the exact historical XML 1.0 Appendix-B productions cited
by ODF 1.4, with an independently extracted range fixture and source receipt.

## Validation

All five scoped gates passed on the frozen source vector: 852 tests across 51
Cargo targets, Clippy with warnings denied, rustdoc with warnings denied, one
compiled API example, and formatting. The hard-depth test additionally replays
one case in a child process; that replay is not counted as an extra Cargo target.
The 14 expression integration groups cover precedence, complete consumption,
all grammar families, source/forced-marker retention, precise whitespace
boundaries, limits, long flat chains, malformed input, injected allocation
failures, and subprocess-isolated depth refusal. The existing ODS lifecycle,
reference, tokenizer and literal tests also passed. Workspace crate-boundary
checking passed with its existing 11 migration debt items.

The [gate receipts](gates/results.json) bind exact commands and before/after
source hashes. [candidate.patch](candidate.patch) was applied to the single
baseline checkout and its resulting source vector was independently checked.
`verify.py` replays the patch in memory against the baseline Git objects,
checks the frozen source, normative character ranges, gates, raw profile rows,
build/source/harness/binary receipts, and the complete artifact manifest.

## Measurements

The same legacy harness ran 12 reference/tokenizer and 9 literal lanes before
and after the change. Every lane preserved expected results, allocation counts,
requested bytes and peak live bytes. Initial median latency changes ranged
from approximately -5.0% to +2.8%; these small shared-host differences do not
support a speedup claim. All initial deltas stayed within ±5%, so no ABAB confirmation was run. Raw
profiles are retained under `performance/`.

Three additional successful `perf stat` captures retain process-level cycles,
instructions, branches, branch misses, and cache misses for the 4,096-size flat,
name and reference cases. They include harness startup and warmup work. Absolute
wall timestamps were not captured for the primary baseline/candidate lanes;
per-lane durations and provenance are retained.

The new AST operation has 24 separate lanes covering scaling and malformed/limit
refusal. It performs additional syntax validation and tree construction, so
its timings are not compared with tokenization as equivalent work. Representative
initial measurements (3 warmups, 15 measured samples, CPU 2; medians normalized
by each lane's repeat count) are:

| AST workload | Median µs/parse | Peak additional live bytes | Allocations/parse |
|---|---:|---:|---:|
| `expr-flat-64` | 1.434 | 13,440 | 13 |
| `expr-flat-4096` | 68.384 | 860,160 | 25 |
| `expr-array-64` | 1.208 | 13,474 | 15 |
| `expr-array-4096` | 49.518 | 860,194 | 27 |
| `expr-name-4096` | 13.955 | 4,481 | 2 |
| `expr-string-4096` | 1.232 | 4,483 | 2 |
| `expr-reference-4096` | 19.181 | 9,482 | 5 |

Flat-expression and array peak storage increases exactly fourfold at each
64→256→1024→4096 size step in this corpus. Long-name and string lanes retain
two allocations per parse, with payload growth proportional to input size.
Latency is noisy at individual points and is not an asymptotic guarantee.
[Environment](performance/environment.json), harnesses and raw per-process
outputs define the scope of these microbenchmarks; they do not measure workbook
recalculation or an end-to-end spreadsheet workflow.

The temporary checkout, build target, temporary directory, executable copies
and Python caches were removed after source and binary verification. The
[cleanup receipt](performance/cleanup.json) and root confirmation retain that
final state; captured text evidence and source harnesses remain.

## Replay and scope

Run `python3 docs/report/spec-gap-validation-evidence/ods-formula-expression-grammar/verify.py`
from the repository to verify retained evidence without rebuilding. Exact Cargo
commands are in `gates/results.json`; isolated release-build commands and raw
measurement commands are in each performance result directory. Reproduction
should use separate output paths to preserve these captured receipts.

No evaluation, arity/type validation, name/source/label resolution, external I/O,
array calculation/spilling, or recalculation is added. Source-preserving syntax
inspection does not establish the full chapter 6–8 semantics. The broader spec
gap audit and performance program remain active.
