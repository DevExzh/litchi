# 0182 — Indexed source materialization

## Contract

Both ordinary source-backed OPC materialization and prepared-topology
materialization now stage their effective parts and validate the complete
namespace through `PartNameIndex`. This replaces the repeated existing-name
scan in each individual insertion. Public single-part insertion and its
conflict checks are unchanged; the new helper is crate-private.

The batch preflights existing and incoming names, preserving exact duplicate,
ASCII-equivalent and ancestor/descendant conflict classes. Destination maps
are reserved fallibly before insertion. Untouched source XML retains its
payload provenance; prepared replacements and additions stay authored even
when their bytes originate from another source part. Package and part
relationships, content types, non-part members, tolerance, source attachment
and final signature-ingress policy retain their existing conversion paths.

The source/cache checks and per-read accounting remain in place. Indexed
preflight also checks freshness and execution context, with a final fence
after insertion. That final fence can reject a populated local candidate;
both callers own disposable conversion packages and return no partial result.
The helper must not be reused to mutate an externally visible package.

Temporary part staging and the namespace index require linear, fallibly
reserved auxiliary storage. This change does not claim constant-memory or
zero-copy conversion. There is no new public API, registry capability,
property vocabulary or OfficeCLI delta.

## Verification and measurements

Independent implementation, review, test and benchmark audits were used.
The signing-enabled OPC suite passed 1,162 tests/doctests, native XLSB passed
1,151, and the combined DOCX/XLSX/PPTX/XLSB/encryption facade suite passed
1,230. After adding the two integration provenance cases, all eight focused
materialization tests passed. They cover authored-versus-source XML captures,
effective graph ownership, managed sources/donors and changing source errors.
Batch unit tests cover all namespace conflict classes and source/signature
state remaining unchanged on preflight refusal.

Warnings-denied combined all-target Clippy and affected-crate rustdoc passed.
Repository lockfiles are unchanged; Cargo verification used an external copied
lockfile with cached cc 1.6.0. No base-only commits were pending. Temporary
implementation and test worktrees were removed after integration. This record
does not claim Python/CLI/parser-fuzz or full-goal completion.

The crate-boundary audit passed for 66 packages and 268 internal declarations
with 11 existing debt items; generated OfficeCLI parity, scoped rustfmt and
diff checks also passed. Temporary build, measurement and verification outputs
were removed after retaining the reduced evidence below.

## Paired measurements

The unchanged sheet-command harness ran 1, 10, 100, 1,000 and 10,000 commands
for eager/source owners, three ABBA rounds (six samples per leg), pinned to
CPU 0 with no concurrent build/test/audit jobs. Source promotion remains inside
the source leg; corpus construction, prepared selectors and final serialization
are outside timing. All cases preserve byte-exact native/facade equivalence,
one commit per leg and an empty failure matrix. Source/build/binary hashes
and samples are retained in [before](results/0182-xlsb-sheet-batch-before.json),
[after](results/0182-xlsb-sheet-batch-after.json) and
[repeat](results/0182-xlsb-sheet-batch-repeat.json).

| Owner / commands | Before native / facade ms | After native / facade ms | Repeat native / facade ms |
|---|---:|---:|---:|
| eager / 1,000 | 65.585 / 68.680 | 65.294 / 68.189 | 65.796 / 68.102 |
| eager / 10,000 | 914.424 / 944.651 | 906.559 / 932.592 | 929.332 / 937.182 |
| source / 1,000 | 101.529 / 103.528 | 90.459 / 92.969 | 89.719 / 91.691 |
| source / 10,000 | 2614.444 / 2633.248 | 1190.011 / 1226.304 | 1182.503 / 1213.290 |

Native source-backed 10k time improves about 55%, while the unchanged eager
control stays within 2% of its baseline. Source 1k-to-10k growth drops from
25.8x to about 13.2x, near the eager leg's 13.9–14.1x. The indexed namespace
algorithm removes the identified quadratic scan, but the remaining measured
slope does not by itself prove the goal's full linear-scaling requirement.
After/repeat facade overhead at 1k/10k is below 4.5%, within the scoped 2x
comparison bar. These numbers cover homogeneous non-active visibility runs;
they do not qualify every operation family or state-transition barrier.
General coalescing, remaining scaling qualification and the complete
all-format/Python/CLI goal remain open.
