# 0442: Shared auxiliary ODP staging traversal

ODP transaction staging now collects settings, declarations and page metadata
in one namespace-aware XML traversal. The individual public parsers remain;
shared scanning preserves their historical full-pass error priority, limits,
namespace semantics and owned outputs. Original parser bodies are retained
byte-for-byte as independent test-only references.

The frozen practical gate passes: normal medium/large p50 improves
9.830–13.307% across both repeats; tiny improves 6.923–7.020%.
The [decision](decision.json) keeps the change on that basis. Each shape saves
only 38 calls and 11,902 requested bytes; peak and retained memory are unchanged.
All five repeat flags remain disclosed in [measurements](measurements.md),
including a +7.091% candidate large p99 repeat change. No general tail,
material allocation, RSS or bounded-append benefit is claimed.

The [protocol](protocol.json) was frozen before builds and edits. Its
A1/B1/B2/A2 matrix retains all 24 reports and 720 samples across 64/4,096/8,192
slides, normal/allocator executables, CPU 2, one worker, 30 samples and three
warmups. Four fresh whole-process profiles retain setup, warmups, oracle work,
raw data and all symbolization warnings. No measurement was excluded or rerun.

The independent report oracle is byte-identical to 0439–0441. Rust fixture
gates inspect actual archives and verify slide semantics, exact tail append,
opaque preservation, no-op sharing, reversible patches and stale-source
refusal. Python independently regenerates expectations and checks report
contracts; reports do not embed full archives for Python to reopen.

Portable verification needs Python 3, without Cargo, perf, original binaries,
temporary directories or the original checkout:

```sh
python3 -B verify.py --sealed --cleanup
python3 -B derive.py --check
python3 -B profile-summary.py --check
python3 -B measurements.py --check
```

The verifier binds frozen protocol/oracle/drivers, exact source/build identities,
ordered captures, commands, report oracles, profile artifacts, chronology,
validation receipts, cleanup and the complete sealed inventory. It also compares
all three test-only reference bodies to retained baseline source. Mutation
probes refresh inventories and report hashes before testing rejection.
See [validation notes](validation-notes.md) for gates and limits.

Capture and evidence tooling is adapted from 0441. The optional historical
prior-attribution.py reads the sealed 0441 profile and is not needed for this
bundle's portable verification. The initial cloned-event prototype is retained
in draft-history; only the final borrowed-event candidate was measured.

The registry remains 436 selectors, the default matrix 36 cases, and coverage
15 categories/33 mappings (10 measured, 23 correctness-only). No promotion.
Remaining source-fragment scans, one-shot caching, bounded existing append,
Part addition, repackaging, native breadth, cold/range and scaling remain open.
The full non-iWork goal stays active.
