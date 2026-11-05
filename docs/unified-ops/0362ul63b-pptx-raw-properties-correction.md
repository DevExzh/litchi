# 0362ul63b — PPTX raw-properties MCE and opaque-fragment correction

- Date: 2026-10-25.
- Status: source-ready for independent review; native Cargo validation remains
  with the coordinator's isolated build slot.
- Parent: `bb10890b19` (`fix(pptx): close raw properties review findings`).

This correction closes the second independent review of the raw-properties
slice:

- MCE ingress is admitted against `max_input_bytes`, while the bounded
  processed stream and every internal DOM/opaque-fragment projection are
  admitted against `max_output_bytes`; structural depth, node, string and
  extension/color limits remain active in both passes;
- the read-only recommendation semantic projection now consumes the same
  output-admitted processed stream instead of reapplying the ingress ceiling;
- opaque raw-node namespace bindings and borrowed fragments use bounded,
  fallible reservations, while already-owned MCE fragments are moved without a
  second copy; writer validation is performed once before streaming rather
  than once per count/emission pass;
- exact byte and namespace-context assertions cover both view and presentation
  opaque fragments, and expansion-boundary regressions cover public view and
  presentation parsing plus the source-preserving read-only semantic path.

Source checks performed here are rustfmt, `git diff --check`, and focused
source/API sweeps. Cargo tests, rustdoc, clippy and ABBA measurements remain
outstanding for the coordinator's isolated run; no binary or large
performance evidence is committed.
