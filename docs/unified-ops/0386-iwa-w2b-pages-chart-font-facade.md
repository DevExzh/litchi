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

## Gaps and boundaries

- No native E3/E4 evidence is claimed; the routes remain queued in
  `docs/unified-ops/0363-mac-queue.md`.
- Full release/ABBA facade measurements are still required on the final
  integrated candidate. Existing engine probes are not facade qualification.
- Unified `litchi-ops` dispatch/schema/capability registration is the separate
  D7 follow-up.
- Caption, legend, value-axis and lifecycle families are outside this slice.

## Verification

Scoped Pages checks and the host-authored facade parity tests are required
before integration. The shared substrate/adapter gates are recorded in their
respective W1-E and W2-B records; the coordinator must rerun the complete
gate on one integrated commit.

No raw captures or binaries were added.
