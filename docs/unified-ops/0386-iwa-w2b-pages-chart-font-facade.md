# W2-B Pages chart-font facade (draft)

## Scope

This slice adds selector-first Pages body-chart font identity/face-trait and
point-size routes over the shared `litchi_iwa_style::chart_font` graph engine.
The public surface is `Package::{body_chart_font, edit_body_chart_font,
apply_body_chart_font}` and the corresponding `_font_size` methods. Native
object identifiers remain private to the Pages package adapter.

The implementation owns exact immutable snapshots, semantic `Edit` verbs,
exact-source `Patch`/`inverse`, no-op preservation, candidate reopen/readback,
ZIP/component/object locality proofs, and typed graph/wire/archive limits. The
engine limits are derived from the opened package's physical ceilings; this
facade does not replace caller/package admission with a payload-only
`Limits::for_source` profile.

## Proof status

Host-authored source-built Pages inputs from `chart_style_cases::input` select
the newly added Column2d chart at body-chart position zero. The facade proof
covers direct host readback before the leaf read, font set/reset, point-size
set/reset, preservation of the other property, exact inverse/no-op bytes, and
candidate semantic readback. The existing native fixture's mixed-font and
mixed-size chart remains a typed refusal control.

The six facade routes remain host-only in the exit ledger until the final
integrated performance qualification and native mutation evidence. The Mac
queue remains pending; synthetic/self-round-trip evidence does not claim
native acceptance.

The final scoped proof also compares host and facade reads after host-authored
font/size setup, host and facade set/reset readback, and the preserved
unmodified property. A stale exact-source patch is refused, an out-of-range
selector is a typed `ChartNotFound`, and a tightened source ceiling returns a
typed finite resource limit before publication. The mutation proof compares
ordered ZIP member names and confirms both changed and untouched members; the
production locality verifier additionally checks untouched objects and raw ZIP
records.

## Gaps and boundaries

- No native E3/E4 evidence is claimed; the routes remain queued in
  `docs/unified-ops/0363-mac-queue.md`.
- Full release/ABBA facade measurements are still required on the final
  integrated candidate. Existing engine probes are not facade qualification.
- Unified `litchi-ops` dispatch/schema/capability registration is the separate
  D7 follow-up.
- Caption, legend, value-axis and lifecycle families are outside this slice.

## Verification

The final branch source was checked with:

- `cargo fmt --all`
- `cargo check -p litchi-pages --lib --tests -p litchi-iwa --test parity --jobs 2`
- `cargo clippy -p litchi-pages --lib --tests --all-features --jobs 2 -- -D warnings`
- `cargo doc -p litchi-pages --no-deps --all-features`
- `cargo test -p litchi-pages --lib body_chart_font -- --nocapture`
- `cargo test -p litchi-iwa --test parity chart_font_cases:: -- --nocapture`

The Pages facade controls and the existing host-vs-leaf chart-font cases pass;
the seven release/cost probes remain intentionally ignored for the
coordinator's final CPU slot. The shared substrate and shape/adapter checks
are recorded in their respective records; the coordinator must rerun the
complete gate on one integrated commit.

No raw captures or binaries were added.
