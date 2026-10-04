# 0902 source review

Status: **PASS (source review)**. This review covers the candidate diff in
`crates/xml-minifier/src/audit.rs` and
`crates/xml-minifier/tests/source_replacement.rs`. It is a source and contract
review; the coordinator owns builds, tests, and any workload measurement.

## Candidate

`verify_source_replacement` already computes the common prefix and suffix and
sets `WindowSearch::stopped` when the replacement span containing all changed
bytes is strictly larger than half of the replacement. The candidate dispatches
from that same condition to two ordinary `verify_source` calls, returning
`ReplacementProof::Complete` after both pass. The existing window search and
window proof path are unchanged when the strict condition is false.

The final source binding is `audit.rs` SHA-256
`e1d73295becd05347dfa6c07910089f16e60b44a6eeb50d92938a3a50cb10fdb` and
`source_replacement.rs` SHA-256
`f7f05d3e2f733c1586c9402f79b29ba4087c8148b7afc88bc8a2c86c6e246d20`.

The condition is sound. If `d = replacement.len() - suffix - prefix`, every
replacement window must contain at least `d` bytes, because bytes outside it
must be byte-identical to the original. `Window::prove` only accepts a window
whose replacement span satisfies `2 * span <= replacement.len()`. Therefore
`2 * d > replacement.len()` makes a window proof impossible. The candidate
uses the existing saturating multiplication, preserving the overflow-safe
boundary behavior. Equality remains eligible for the existing search path.

## Verdict and error-order equivalence

The direct path preserves the public contract in the function documentation and
the `oracle` helper in the test module:

1. `verify_source(original, limits)` runs first. Any encoding, byte, token,
   structural, namespace, or well-formedness error remains
   `ReplacementError::Original`, with the original offset and detail.
2. Only after the original passes, `verify_source(replacement, limits)` runs.
   Its failure remains `ReplacementError::Replacement`, with the complete-audit
   error and offset.
3. If both pass, the result is `ReplacementProof::Complete`.

This is behaviorally the same as the pre-candidate path: with `stopped` already
true, `verify_observed` performed a complete original scan while its observer
received no tokens, then the replacement was audited completely. The candidate
removes that inert observer path without changing `Policy::SOURCE`, limits,
state transitions, error construction, or the original-before-replacement
ordering. The identical-byte fast path remains ahead of the new dispatch.

The strict boundary is retained. A difference exactly half the replacement
still enters the old search/proof path and may fall back to a complete audit if
the enclosing window cannot fit; a difference just over half takes the direct
path. This distinction is covered by the added source-replacement tests.

## Tests and invariant review

The added tests cover:

- original-first failure precedence when both payloads are malformed;
- original and replacement byte-budget failures on the broad path;
- a valid broad replacement returning `Complete`;
- exact-half versus just-over-half dispatch boundaries;
- namespace, Unicode, mismatched-end-tag, and invalid-UTF-8 replacement
  failures compared with independent `verify_source` errors.

The candidate does not alter `State::within`, `WindowSearch::consider`,
`Window::prove`, `audit_window`, `scan`, or any source-policy check. Thus the
0747 window-proof invariant and the 0750 replay of document-start and namespace
state remain intact. No new source-policy state or check is introduced, so no
window-replay obligation is added. Existing complete-audit fallback behavior
for an exactly-half difference, a window that cannot be found, or a window that
fails proof is unchanged.

No source-level correctness blocker was found. Final acceptance still depends
on the coordinator's quality gates and matched performance evidence; this
review makes no performance claim.

## Quality-run correction

The first candidate XML test attempt was retained as a compile-only failure:
the newly introduced `verify_source(original, limits)?` result was not bound,
and the crate denies unused `must_use` values. The final candidate binds that
result as `_report`; this changes no control flow, error mapping, limits, or
verdict. The corrected candidate all-target XML test run exited 0, with input
hashes unchanged. The retained receipts are
`candidate-xml-tests-run.json` (exit 101) and
`candidate-xml-tests-v2-run.json` (exit 0).
