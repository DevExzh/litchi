# 0783 — current-source phase review

This is a read-only review of the 0783 probe and the production paths it
measures. It proposes future diagnostic seams; it does not attribute the 0780
large-lifecycle result to any one implementation detail and does not recommend
changing production behavior from this source inspection.

## Probe contract

The probe's phase build is appropriately isolated in the evidence project.
`probe-src/src/main.rs:375-453` keeps the public lifecycle calls in the same
order in both binaries. The feature-off leg has one start clock and one end
clock; the `phase-timing` leg adds five boundary reads after capture, staging,
commit, apply, and serialization. The phase values are consecutive intervals
whose sum is checked against the lifecycle clock before the sample is emitted.
The six alternating blocks and three warmups per process preserve a matched
control for the extra clock reads and feature-build layout.

The current intervals mean the following:

| Recorded interval | Public work included | Boundary detail |
| --- | --- | --- |
| `phase_capture_ns` | `Package::opened_presentation()` | Begins at the lifecycle start and ends after the capture call. Package ingress is outside it. |
| `phase_stage_ns` | `snapshot.edit()` and `set_shape_text(0, 0, MARKER)` | The edit constructor and the two scene reads used by the text setter are one staging interval. |
| `phase_commit_ns` | `Transaction::commit()` | The setter is already complete, matching the standalone commit mode's setup. |
| `phase_apply_ns` | `apply_opened_presentation_commit(commit)` | The returned `Snapshot` is an unbound temporary and is dropped before the following timestamp. Its destruction is therefore part of this interval in both legs. |
| `phase_serialize_ns` | `Package::to_bytes()` | Includes `flush_presentation()` and the OPC writer. Readback, output hashing, and owner destruction remain outside the clock. |

The lifecycle owner placement is sound for this question: `package` and the
serialized output remain live through `allocation_metrics::finish()` and the
same temporary snapshot drop occurs in control and phase builds. The probe does
not move `verify_output` into a measured phase. `Package::from_bytes` and corpus
construction are outside the clock, so a future report must keep “capture”
separate from ingress rather than relabeling the current interval as open or
deserialize time.

The report metadata has one small ambiguity: `Mode::timing_scope()` at
`probe-src/src/main.rs:96-107` returns a lifecycle string mentioning
`phase-timing` for both feature variants. The receipt's `control`/`phases` leg
still disambiguates the files, but a durable report should carry an explicit
`phase_timing_enabled` field or use a feature-neutral scope string for the
control leg. This is metadata hygiene, not a timing or correctness defect.

## Safe future instrumentation seams

The external phase boundaries are the right first diagnostic. If source-level
phase counters are later needed, keep them opt-in and content-free. A safe
shape is a private `#[cfg(feature = "performance-diagnostics")]` recorder with
fixed-capacity numeric events, and a feature-off ordinary method whose call
sequence and ownership are unchanged. Do not put formatting, `String` labels,
per-part records, or a heap-allocating callback in the measured route.

Capture has one public seam in
`crates/litchi-pptx/src/package/model.rs:243-268`. It selects either
`capture_with_parent_digests` or `capture_with_provenance` and then offers the
new memo to the facade. If a later diagnostic needs internal capture phases,
instrument the private `capture_internal` blocks at
`crates/litchi-pptx/src/opened/model.rs:730-876` with numeric boundaries:

1. presentation root and slide-reference catalog (`:738-747`);
2. slide/MCE/notes-root capture (`:748-750`, delegated to
   `crates/litchi-pptx/src/presentation/package.rs:65-135`);
3. slide identity, relationship, and name-index validation (`:755-827`);
4. notes graph and index completion (`:818-827`); and
5. owned package clone, revision/digest handling, and memo finalization
   (`:828-876`).

These are nested capture operations when called from commit or apply. Every
event therefore needs a route/context tag, or the report must explicitly call
the values nested capture observations. A callback in `capture_internal` alone
must not be interpreted as only the outer `opened_presentation` capture.

Commit is the most useful narrow source seam at
`crates/litchi-pptx/src/opened/transaction.rs:1241-1298`. Preserve the current
move and error order, and if diagnostics are added, time these blocks without
recomputing their values:

```text
compact_changed_slides                 1242-1248
package_fingerprint_with_memo          1252-1253
signature check/unsign/re-fingerprint  1254-1270 (conditional)
Patch::capture                          1272-1277
candidate capture/validation            1284-1296
```

`Patch::capture` is itself a useful descriptive subphase at
`crates/litchi-pptx/src/opened/patch.rs:210-265`; it enumerates and sorts the
union of all source and target part names, captures both resource states, and
then captures root relationships. A future recorder should report the patch
resource count and signed-branch decision as metadata, rather than adding
work to derive them a second time.

Apply has a similarly exact private seam at
`crates/litchi-pptx/src/opened/patch.rs:892-970`:

```text
presentation-root check and validate_before  900-910
candidate clone and delta application       925-955
capture_candidate                            956-962
validate_after and revision check             963-968
single package assignment                     969-970
```

`capture_candidate` has an important branch at `:814-842`: a committed
snapshot can be rebound after `packages_equal` succeeds (`:821-827`), or the
candidate is recaptured with a parent digest memo (`:828-842`). Record the
branch as a numeric outcome if needed. Keep the equality check, memo
projection, candidate capture, both validations, and assignment exactly where
they are. The result's temporary drop belongs to the public apply interval and
should not be moved merely to make an internal event boundary convenient.

Serialization should initially remain one external interval at
`crates/litchi-pptx/src/package/codec.rs:333-342`. It includes the
`flush_presentation` no-op check (`:383-397`) and
`PackageWriter::to_bytes`. If an internal diagnostic is later justified, the
least intrusive writer boundaries are:

```text
PublicationPlan::from_package       crates/litchi-opc/src/pkgwriter.rs:236-354
try_write_preserved                  crates/litchi-opc/src/pkgwriter.rs:416-944
fallback materialize_pristine        crates/litchi-opc/src/pkgwriter.rs:364-384
plan.write and ZIP finish             crates/litchi-opc/src/pkgwriter.rs:386-407,
                                      1296-1323
```

The current generated fixture enters through borrowed `Package::from_bytes`
(`crates/litchi-pptx/src/package/codec.rs:243-255` and
`crates/litchi-opc/src/package.rs:1119-1131`). That route has no retained
owned source archive, so after the edit `PackageWriter::to_bytes` takes the
full publication-plan/materialization/deflate route: `try_write_preserved`
falls back when no preservation source exists, then `materialize_pristine`
and `PhysPkgWriter` emit every planned member. A source-preserving corpus can
take a different branch. Any future writer phase report must record that route
as metadata and must not combine its numbers with the current full-writer
fixture as though they were the same operation.

Do not begin with callbacks around `PhysPkgWriter::write` or each ZIP member.
Those callbacks would multiply clock and recorder overhead by package size and
could change compression/layout behavior. If per-member evidence is eventually
needed, use a separate writer diagnostic with a fixed numeric sink and pair it
with an unobserved writer build and a no-op observer control.

## Repeated work visible in the current source

The following are source-level work paths that the phase results can rank; they
are hypotheses about repeated work, not causal performance claims.

* Capture validates the presentation root and slide references, processes each
  slide's MCE/XML roots and notes-root classification, checks slide identities
  and relationships, builds the name index, clones the complete OPC graph, and
  computes or projects the complete-package revision. The digest memo skips a
  payload SHA-256 only when the payload allocation is shared; names, content
  types, and relationship lists are fed again on every pass
  (`opened/model.rs:911-990`).
* A single `set_shape_text` resolves the slide, parses a complete scene,
  scans the raw shape span, rewrites the XML, reparses the staged XML, and
  checks selected text and shape count
  (`opened/transaction.rs:216-250`). This is why staging must remain a
  distinct interval from commit even though the public lifecycle includes both.
* Commit scans every known slide for changed payload identity/equality and
  compacts changed XML (`transaction.rs:1354-1404`). It then fingerprints the
  complete package, captures the complete patch union, and captures the staged
  candidate again with the known revision and retained MCE/slide-root proofs
  (`transaction.rs:1242-1296`). Unchanged payloads can hit source memos, while
  metadata and relationship validation still run.
* Apply validates the patch write set before mutation, clones the complete OPC
  graph, applies deltas, captures or rebinds a candidate, validates the write
  set after mutation, and assigns once (`opened/patch.rs:909-970`). The
  `ResourceState` and relationship snapshots used by `validate_before` and
  `validate_after` are built separately (`:973-1012`).
* The writer plans all parts and relationships, audits required XML, attempts
  source preservation, and otherwise materializes and deflates the package
  (`litchi-opc/src/pkgwriter.rs:236-407,416-944,1296-1323`). On the current
  borrowed generated corpus, this is a whole-package serialization path even
  when one slide changed.

## Measurement pitfalls to preserve

* Boundary clocks are part of the observed intervals. Compare phase and
  feature-off legs in the frozen alternating plan; do not treat a phase value
  as an uninstrumented CPU duration.
* Nested phase sums are descriptive. Do not add internal capture or apply
  subphase values to the outer lifecycle total, and do not turn a residual into
  a causal explanation.
* Keep the output oracle, source/output identities, and semantic readback
  after the clock, while retaining the same handles until the same allocation
  boundary. Moving verification or a drop can change both timing and retained
  allocation measurements.
* Never call a second fingerprint, scene parser, serialization pass, or hash
  solely to populate diagnostics. Existing memo-hit behavior and the signed
  `unsign` branch must be observed without changing the work graph.
* Avoid per-event allocation and formatting. A fixed event array with integer
  phase IDs and one completion/error bit is sufficient for a later opt-in
  diagnostic. Errors must preserve the existing `Result` path and leave the
  same partial-move/drop behavior.
* Record route facts such as `capture` parent-memo presence, apply
  rebind/fallback, signed-branch execution, and exact-source/preservation
  availability as metadata. They are useful for stratification and do not by
  themselves establish why a phase took a given time.

The current 0783 source supports safe descriptive phase timing at the probe
boundary. It does not support a causal explanation of the 0780 large-lifecycle
regression without measuring matched historical legs under a separately frozen
plan.
