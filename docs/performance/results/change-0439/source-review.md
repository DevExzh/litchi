# 0439 ODP existing-document append source review

Scope: source-only review of the external `odp_existing_append.rs` draft. No
production files, builds, tests, scripts, or profiling were run.

## Blocking integration issues in the reviewed draft

The current external module is not ready to copy into `tools/perf-baseline`:

1. `run_odp_existing_append_lifecycle` still calls undefined
   `APPENDED_TITLE` and `APPENDED_BODY`. Use the shape-bound
   `appended_title(corpus.shape)` and `appended_body(corpus.shape)` values, or
   define an equivalent per-corpus binding.
2. The summary schema and runner are from different revisions. The summary
   requires `role`, `shape`, output archive identity, member counts, manifest
   flags, semantic/order/text digests, and runtime verification flags, while the
   runner still initializes removed names such as `expected_output_sha256`,
   `expected_output_bytes`, `opaque_member_decoded_bytes`,
   `opaque_member_decoded_sha256`, `untouched_member_decoded_identity_verified`,
   and `manifest_byte_identity_verified`. It also omits the newly required
   fields.
3. `OdpExistingAppendCorpus` no longer has `manifest_byte_identity_verified`,
   but the initializer and runner still reference it.
4. The untouched-member expression contains `.iter().iter()`, which is not a
   valid iterator operation.

## Fixture and oracle requirements

- Keep the exact six ZIP members and five manifest bindings gate: root `/`,
  `content.xml`, `styles.xml`, `meta.xml`, and the opaque member. The existing
  `Manifest` representation includes `/` in `entries`, so five is correct for
  this fixture.
- Assert the buffered base's exact five-member topology before extracting its
  three XML parts. The current source builder silently drops any future base
  members; this must not become an implicit fixture mutation.
- `expected_archive_shape` currently checks only that the first member is the
  stored `mimetype`. Also require Deflate for every other source and output
  member, and compare the relevant method/descriptor/CRC metadata where the
  report claims it.
- `expected_manifest_bindings` should retain the existing buffered checks for
  absent `manifest:size` and encryption metadata, in addition to media types.
- The output verifier compares selected decoded members and the opaque
  compressed span. Retain an explicit manifest-byte equality flag in the
  final summary; the ODP writer's `preserve_source_manifest` path can make this
  true when the inventory remains equivalent.
- The opaque compressed-byte equality is evidence of deterministic
  recompression by the same owned writer. It must not be described as
  source-backed raw compressed preservation.
- The stale fixture is useful because only the opaque payload differs. Record
  that source/content/styles/meta/manifest are equal while the opaque member is
  different, so refusal is demonstrably source-lineage based.
- The stale patch check currently uses only `.is_err()`. Match the exact
  `litchi_core::Error::InvalidFormat("stale ODP presentation patch source")`
  contract (there is no dedicated stale variant), and verify the failed apply
  leaves the stale snapshot unchanged.
- The exact no-op check currently proves byte equality and `Patch::is_noop()`.
  Add the shared-byte ownership check (`bytes().as_ptr()` plus length) if the
  lifecycle contract requires no-op snapshot sharing.
- `semantic_slides` proves visible title/body order, but it is still the public
  `Presentation` projection. Cross-check the source against the buffered
  corpus's independent expected semantic projection and retain a direct
  appended-tail/content XML check before presenting full semantic evidence.

## Schema and measurement boundaries

- `uncompressed_payload_bytes` currently adds `content.xml` bytes and opaque
  bytes. Existing buffered ODP manifests use semantic projection bytes. Give
  this field an explicit archive-payload meaning or keep semantic input bytes
  separate from opaque archive bytes; do not mix the two silently.
- The timing boundary is coherent: input cloning and sink construction are
  outside the clock; `Snapshot::from_bytes`, transaction staging/commit, and
  one sequential `write_all` are inside; digest finalization, validation, and
  cleanup are outside. The allocator region intentionally observes source and
  commit still alive at both endpoints, so nonzero live-after deltas are
  expected and must not be treated as leaks.
- `HashingDiscardSink` verifies successful digest and accepted length only. It
  provides no partial-write or sink-failure evidence, so the report must not
  claim those paths.

The real ODP APIs used by the draft are otherwise compatible: `Snapshot::from_bytes`
takes owned bytes, `transaction().add().commit()`, `Patch::apply/inverse`, and
the ODF `PackageWriter` methods used by the fixture all exist with the expected
signatures.

## Addendum: applied module review

The applied module now reconciles the earlier schema/runner mismatches, uses
shape-bound append strings inside the timed operation, checks the buffered
five-member base and six-member append topology, and uses the exact stale-source
error plus shared-byte pointer equality for the no-op gate. The measured owner
boundary is coherent: the input clone, append strings, and sink are outside the
clock; source opening, transaction staging/commit, and one sink write are
inside; source and commit remain alive through allocator/process endpoint
snapshots and are dropped afterward.

One preservation gap remains. `verify_append_output` compares decoded
`mimetype`, `styles.xml`, `meta.xml`, and the opaque member, while
`untouched_members_verified` repeats those four paths. Neither compares
`META-INF/manifest.xml` source bytes with output bytes. The separate binding
checks prove the two manifests have the expected five entries and metadata, but
not byte identity or ordering/format preservation. Either include the manifest
in the untouched-byte check or expose a separate `manifest_bytes_identity_verified`
flag before treating all untouched-member evidence as complete.

The sink's `input_bytes` is populated with the output semantic text projection
length and `authored_part_bytes` with output `content.xml` length after the
successful write. These are useful labeled projections, but they are not the
physical archive bytes consumed by `Snapshot::from_bytes` or the exact append
argument bytes. Keep the existing performance claim scoped to the owned
materialized lifecycle and document those fields as semantic accounting if they
remain in the generic sink envelope.

## Root resolution

The final root source includes the manifest in both untouched-byte comparisons,
uses the typed stale-source error, and checks the exact-noop byte pointer. The
sink projection fields are explicitly documented in source and `data-path.md`.
After the first CLI pilot exposed generic OPC routing, the missing exclusion
was added and the coder reviewed the complete enum/name/parser/help/default/
dedicated/generic/low-level routing paths without finding another omission.

## Final evidence review

The source-only reviewer inspected the final lifecycle, contract, profile
amendment, portable probes and sealed artifacts and found no remaining hard
schema, path, custody, chronology or oracle blocker. The 13 mutation probes
and unmodified copied control align with the final bundle. Replay writes its
receipt and log after inventory checks, so the root resealed after each
retained replay; compression metadata includes those replay artifacts.
