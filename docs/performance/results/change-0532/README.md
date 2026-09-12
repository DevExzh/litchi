# 0532: current CFB ownership and reconciliation attribution

This campaign measures the unchanged source at revision `70e04d901`. It follows the rejected 0531 OOXML candidate and prioritizes OLE2. It does not adopt an optimization or claim a before/after speedup.

The frozen matrix contains nine XLS eager, instrumented-source, and owned-source open/list/one-cell cases, plus direct CFB tiny, many-small-stream, and few-large-stream shapes. Native captures use two repeats with 20 warmups and 1,000 samples per case. Allocation captures use three warmups and 30 samples with the separate instrumented executable; their elapsed times are excluded from latency claims. Every child runs on CPU 2 with source and binary receipts, exact corpus identities, raw vectors, and runtime oracles.

Eight Callgrind children isolate parent constructors and retain five measured dumps per child. Direct CFB corpus construction contributes a separate setup dump. Positive incoming caller edges distinguish setup from timed calls. The three residual owners are reported separately by exclusive instructions; their inclusive counts overlap and must not be added. Parent toggles avoid producing a dump for each claimed sector. Generated-code inspection records six 363-byte claim variants with a 15-instruction success path and a 112-byte stack reservation. Whether changing their layout improves native latency requires a fresh candidate comparison.

Two hardware captures are whole-child diagnostics, including setup and correctness work outside the constructor clock. They cannot establish operation-local cycles or hardware speedups. Accessible compiler-process observations accompany each receipt; the host is not claimed to be quiescent.

All Rust and harness source hashes equal the verified 0531 restored source. Its eight quality gates are reused with exact source and receipt bindings. Five fresh CFB/XLS test, Clippy, and rustdoc checks supplement that evidence. No runtime source, dependency, unsafe code, validation policy, or resource policy is changed by this baseline.

The source review must preserve collect-then-claim order, bounds and fallible allocations, duplicate ownership errors, atomic publication, and physical reconciliation of unclaimed sectors. The rejected 0524 visited-bit fusion and 0279 freshness-session proposals are not reinstated.

OLE2/OOXML remains active. ODF is deferred until that optimization goal completes; iWork is excluded. These warm synthetic measurements do not establish cold I/O, remote-source, native-producer, scaling, or program completion claims.

The completed capture contains 24,000 native samples, 720 allocation samples, and 40 timed constructor profiles. All five fresh quality gates passed with 1,957 test executions, supplementing 4,757 explicitly reused executions. [Cleanup](cleanup.json) removed both owned scratch paths. The [post-cleanup verifier](verification.json) replays the analyzers and validates custody, quality, cleanup, and the exact [sealed inventory](SHA256SUMS).
