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

## Numerical follow-up

The same two independent probes plus seven reducer unit tests pass against
reducer snapshot
`b7a949e9ec3cc6ad7b33f68879c0a6904532c3f69d3026bebd2deb26e8580ec1`,
with the unchanged numeric helper snapshot above: **9 passed, 0 failed**.
The discounted cash-flow underflow and the two fixture errors are resolved
in that snapshot. The temporary standalone harness was removed.

This does not clear the resource hold: fixed accumulator work accounting and
the distinction between numerical-domain refusal and typed accumulator-span
failure still require adapter review and public-evaluator tests.

## Discount cancellation and signed MIRR follow-up

Root compiled reducer snapshot
`60618d78b87b178f5f7196cf458476c1f8286c00e7129629716eff93d4f1cf82`
with the unchanged numeric helper. Two independent probes passed:

- `NPV(0.125; [-1e16, 1, 12656250000000000]) = 64/81`, within `1e-12`.
- `MIRR([110, -100]; 0; -2) = -2.1`, within `1e-12`.

Including the nine embedded reducer tests, the run had **10 passed, 1 failed**.
The failure was a new fixture expecting `#NUM!` for
`MIRR([110, -100, 0]; 0; -2)`. Its intermediate ratio is positive `1.1`,
so the implementation's `sqrt(1.1) - 1` result is correct. A negative-ratio
nonintegral-power refusal can instead use investment `-2` and reinvestment
`-2` with the same values. The owner received this fixture correction;
no tolerance or domain rule was relaxed. Temporary probe files were removed.

## XNPV shared reducer follow-up

Root ran the retained `probe_reducers.py .codex-tmp/ods-financial-development`
harness against reducer SHA256
`7c0f57f24322dc244dec09763a1f4c057adb63e6e76bc9260ba25e3331127ebe`
and numeric helper SHA256
`627ec3d9c5f0f15d43f483fc64b28a861c9235877173a1db087a82c7697ee6e0`.
All **13 embedded reducer tests passed**, including exact rate-zero XNPV
cancellation, discounted integer-date cancellation, fractional-date powers,
and the corrected signed MIRR fixture. Temporary snapshot sources and binary
were removed automatically. This verifies numerical kernels only; the value
adapter's streaming, error precedence, charging, and public replay remain
separate validation requirements.
