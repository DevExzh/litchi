# Byte-text validation status

The candidate is based on commit `3844f235bac545ff0ae1580b97612883c1fd9f89`.
The 25 selected inputs in [freeze.json](gates/freeze.json) include production,
focused tests, the contract, independent oracle and raw native fixtures. Gate
validation uses retained Cargo.lock `58b4be6c…`; the ambient root lock is a
different file and is deliberately excluded from that source claim.

| Requirement | Current evidence |
| --- | --- |
| All seven section 6.7 functions | Six focused evaluation tests cover scalar/value/matrix calls, UTF-8 clipping, arity/defaults, Missing, Complex, conversions, numeric-text overflow, broadcasting and projected reducers. |
| Independent UTF-8 expectations | One Rust test checks all 1,376 observations reproduced by `byte_oracle.py --check`. The oracle uses integer/string arguments and selected stable full-fold characters; it does not claim exhaustive numeric conversion or Unicode coverage. |
| Native interoperability boundary | One Rust test checks all 19 selected-profile results; retained LibreOffice output and `native/reproduce.py` establish seven ASCII matches and ten non-ASCII divergences. |
| Resource and failure behavior | Five focused limits tests check read/type refusals, charged text work, output/storage failure, cancellation/source fences, typed precedence and budget refund. Independent [resource/cache review](resource-review.md) is PASS against the frozen sources. |
| Projected cache correctness | Permanent regression checks AVERAGE(LENB(reference)) and AVERAGE(LEN(reference)) per coordinate. SUM keeps complete matrix arguments. Conservative cache refusal can repeat invariant computed SUM work; that cost must remain disclosed. |
| Isolated crate verification | All seven gates pass: 1,605 tests, strict all-target clippy, rustdoc with warnings denied, both format checks, crate boundaries and diff whitespace. `gates/verify.py` independently verifies the receipts and unchanged source hashes; zero failed or ignored tests. |
| Performance | All 3,360 samples independently verify: 840 baseline and 2,520 candidate. Matched allocation/work/read/output medians are unchanged. Latency p50 shifts range from -5.33% to +3.30%; no causal speedup claim is made. See [performance interpretation](performance-review.md). |

Validation is complete for this batch; see [completion.md](completion.md). The saved [independent verification receipt](verification.json) covers gates,
oracle, native observations and complete performance source/measurement custody.
Independent [semantic review](semantic-review.md) is PASS, and owned temporary
files were removed. Raw evidence remains reproducible after checkout cleanup.
