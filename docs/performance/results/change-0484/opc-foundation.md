# Authenticated replay and exact artifact restoration

This checkpoint adds the OPC capabilities needed by the replayable DOCX
paragraph-stream work described in [design.md](design.md). It is a correctness
and resource-ownership checkpoint, not a completed DOCX stream implementation
or a performance result. The full Office CRUD goal remains open.

The existing fixed-fragment splice API remains available. The new replay
provider supplies a fresh sequential reader and a sealed decoded length/SHA-256
proof. Preparation captures the expected proof; candidate validation and both
publication passes authenticate the emitted bytes against that captured proof.
A provider cannot expand the read bound by changing its advertised proof later.
Each live pass reserves its own parser and replay adapter windows from the
package execution context. Prepared replay plans retain neither window.
The adapter uses the smaller of the selected replay ceiling and 64 KiB; a
512-byte ceiling therefore admits a payload spanning many windows. The
allocation's actual capacity is checked against its reservation.

The consuming expected-artifact publication API authenticates a full preview
against the expected archive length and SHA-256 before touching the caller
sink. Preview reads, work, memory, cancellation, and source checks use the
same package context. Only the external publication consumes `OutputBytes`.
Preview failures preserve their typed cause without reporting internal sink
progress as caller output; failures during actual emission retain the exact
caller-accepted prefix. The private preview plan is used sequentially.

Exact artifact restoration accepts an explicitly supplied original artifact
and the complete current/original archive lengths and SHA-256 identities.
Runtime source versions are checked separately, so an independently reopened
provider can restore an artifact without persisting process-local identity.
Both complete archives are authenticated before output; the original is hashed
again while copying. The current package context owns operation I/O, work,
memory, and output charges, falling back to the original context only for an
unmanaged current package. Both contexts still enforce cancellation. Output
is charged once. A bounded 64 KiB copy/hash buffer is released on success and
failure, and post-output failures retain the exact accepted byte count.

The restored-output ceiling applies to the original artifact. A larger current
candidate remains a valid input. Publication now exposes its accepted archive
length together with the raw archive fingerprint, allowing format-owned durable
patches to bind both facts.
Immediate inverse publication also checks the candidate length before hashing.
Both ordinary exact snapshot-copy publication paths now reserve and release
their 64 KiB copy workspace through the selected context, including no-op
publication. Their output-focused tests retain their original output limits
and separately allow the required copy memory.

## Validation

The retained commands, stdout/stderr, source manifests, and receipts are under
[validation/](validation/) and [validation-sources/](validation-sources/).

| Receipt | Result |
| --- | --- |
| `opc-all-tests-dev37` | 571 passed, 1 ignored; all features, including doctests |
| `opc-clippy-dev39` | All features and all targets, warnings denied; passed |
| `opc-no-default-tests-dev38` | 549 passed, 1 ignored; no default features, including doctests |
| `opc-rustdoc-dev40` | All features, no dependency docs, warnings denied; passed |
| `opc-format-dev43` | Scoped rustfmt check after whitespace-only formatting; passed |

Both test runs, Clippy, and rustdoc report unchanged source manifests. These are scoped
development receipts, not a frozen final performance bundle. Earlier failed
attempts remain available, including the tests that exposed missing copy-memory
accounting and the resulting output-test policy updates.
The later format check also reports unchanged inputs. Between these gates,
OPC changed only in two indentation corrections and one test line wrap;
DOCX development continued independently.

The new focused coverage includes 19 replay-provider tests and 13 restoration
tests, alongside the 29 existing fixed-splice tests. It exercises Store and
Deflate preservation, immutable replay bounds, truncation, extra bytes, digest
changes, provider failures, per-live-publication reservations, cancellation,
short sinks, source changes, independently reopened artifacts, distinct/shared
contexts, memory refusal/release, and changes to original bytes during copying.
It also covers expected-artifact preview rejection, fresh providers on every
pass, preview versus external partial-output errors, and one output charge
for both changed and no-op publication.
The existing DOCX single-paragraph tail suite also passes all 40 tests in dev18.

## Remaining evidence

The sealed change-0483 evidence motivates replacing complete authored-fragment
retention with a replay capability; it does not measure this new multi-paragraph
path. DOCX stream integration, durable forward application, native-consumer and
fuzz checks, independent source/authored/chunk scaling, allocator/RSS evidence,
and final source/binary custody remain required before accepting change 0484.
There is no latency, throughput, or constant-RSS claim for this checkpoint.
