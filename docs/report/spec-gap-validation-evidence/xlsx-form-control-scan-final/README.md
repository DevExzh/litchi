# Form-control namespace/event scan fusion

The source owner previously parsed worksheet/drawing XML once to validate namespaces and again solely to count MCE events. The change obtains the same count during namespace validation. It retains the original second sequence of execution checks/work charges after successful admission, followed by the same object reservation. Namespace checks, XML error precedence, event limits, MCE selection/provenance, readsets, and guards remain unchanged.

Independent review approved the narrow patch after comparing both implementations: the removed subject-specific event-limit error was unreachable because the first namespace pass enforced the identical cap. Cooperative cancellation checks retain their sequence, although the second sequence no longer interleaves a redundant physical parse. There is no requirement to preserve wall-clock cancellation timing.

Root applied only this patch in an isolated checkout at `01c31e6e4`. The resulting owner bytes match the staged file and retained SHA-256. Validation passed 1,171 library tests, 39 public owner tests, strict library/owner-test Clippy, and warnings-denied rustdoc. Source lock, patch, and compressed logs are retained. Uncommitted scalar-lifecycle changes are excluded.

The separate matched profiling experiment and clean root replay are committed in `2a0572778`; see [the replay bundle](../xlsx-form-control-scan-performance/README.md). The 16-control source projection measured 57,796 → 57,788 allocations and 34,526,604 → 34,520,703 requested bytes with unchanged source reads and peak live bytes. Refusal-lane allocation totals differ across captures and are documented separately. This source gate makes no production latency, broad speedup, or scalar-lifecycle completion claim.
