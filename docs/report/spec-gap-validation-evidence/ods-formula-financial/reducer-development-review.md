# Financial reducer development review

Status: **HOLD**, pre-integration evidence only. This is not a production gate.

The root reviewer compiled a temporary standalone Rust test harness containing
unchanged snapshots of `financial/reducers.rs` and `numerics.rs`, with a minimal
`ScalarError` declaration. The harness ran the six reducer unit tests and two
independent probes. It does not exercise evaluator admission, resolver reads,
resource accounting, or the public API. Temporary sources and binaries were
removed after execution.

| Snapshot | SHA-256 |
| --- | --- |
| `financial/reducers.rs` | `18e3f8b393faeb999be199f23e5633ae62bf8cb5f01cdca81eee600b99754cb9` |
| `numerics.rs` | `627ec3d9c5f0f15d43f483fc64b28a861c9235877173a1db087a82c7697ee6e0` |

Compilation passed; execution returned 101: **5 passed, 3 failed**.

1. Independent NPV probe: rate `0.5`, followed by 2,000 zero cash flows and
   `1e300`, returned zero. The finite expected result is approximately
   `4.379158148872767e-53`, calculated as
   `exp(ln(1e300) - 2001 * ln1p(0.5))`. Rounding the discount factor to
   binary64 before multiplying by the cash flow silently underflows. The
   discount and cash-flow magnitude must be combined in scaled arithmetic.
2. Existing reversed-NPV fixture expected `9.09090909090908` for
   `[110, -100]` at rate `0.1`. The one-based equation gives approximately
   `17.35537190082645`; the fixture needs correction.
3. Existing MIRR fixture expected `0.21` for `[-100, 110]`, with both rates
   `0.1`. The contracted equation gives `0.1`. The independent probe for
   this value passed; the fixture needs correction.

Source review also found `ProductSumError::ExponentSpan` converted to a
catchable formula `#NUM!`. Its resource/profile failure mapping requires
review at the adapter boundary; a typed failure must not silently become a
formula error. No runtime claim is made for this source-only finding.

Separately, `cargo check --locked --offline -p litchi-ods` in the isolated
development checkout failed with 86 dead-code diagnostics while financial
value-dispatch hooks were absent. Integration and focused Cargo tests remain
pending. These findings were handed to the respective implementation owners.
