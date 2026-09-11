# XLSX profile scaffold readiness

This handoff contains requirements and runner infrastructure only. No raw
receipts, generated metrics, or performance claims are present because the
source-backed XLSX SVG owner and its ordinary public selector / attach /
detach API are not frozen.

The scaffold is ready for API wiring when the design review supplies a
callable source-backed worksheet picture view and mutation transaction. The
adapter must then implement the recipe manifest, preserve all three anchor
forms, and satisfy the semantic and refusal gates in `requirements.md` before
setting `XLSX_SVG_PROFILE_API_WIRED=1`.

Validation performed on the unwired scaffold:

- standalone harness `cargo check --locked --offline` passes;
- standalone harness `cargo fmt -- --check` passes;
- shell and Python syntax checks pass;
- the runner refuses an unfrozen run before creating a target;
- the runner refuses the unwired scaffold after the freeze flag alone; and
- an explicit repository `target` path is rejected by the safety guard.

The only retained files under this directory are the plan, deterministic
recipe manifest, source manifest and receipt tooling, the intentionally
refusing harness, and this readiness note. Temporary Cargo targets and Python
cache files were removed after validation.


Root independently repeated Python/JSON/shell syntax checks, the unfrozen
runner refusal, a locked offline harness check in a newly created temporary
Cargo target, and Rust formatting. Those checks passed; the temporary target
was removed. The scaffold contains no production adapter or performance
measurements and must remain gated until the callable lifecycle API is wired
and frozen.
