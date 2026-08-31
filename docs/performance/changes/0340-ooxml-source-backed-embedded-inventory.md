# Change 0340: OOXML source-backed embedded inventory

## Scope

This change adds a shared, additive inventory for OOXML OLE Object and Package
relationships. The inventory is lazy: it describes what the package catalog
and its bounded relationship/content-type metadata identify, without eagerly
materializing payload bytes or opening a stream for each entry. It does not
claim to inventory arbitrary opaque package members. The shared inventory is
owned by the common source-backed OPC path, while format facades expose only
the typed view appropriate to their format.

DOCX and XLSB are the first facade consumers. Their inventory accessors make
OLE Object and Package relationships discoverable without requiring callers
to understand OPC physical names, relationship parts, ZIP offsets, or the
internal source handle. An entry may identify its authored part name,
relationship/content type, inert category, size metadata, and source-backed
status. An unrecognised or unsupported OLE Object/Package detail remains an
inert catalog entry rather than being silently dropped; unrelated opaque
members are outside this inventory.

The catalog builds an indexed canonical-name lookup closure for the discovered
local payload parts. The index uses OPC-equivalent physical-name semantics and
is bounded to the inventory's authored relationship closure; it is not a
general arbitrary-member resolver. Duplicate or ambiguous canonical names
refuse rather than selecting the first physical member.

The initial scan is catalog-only. It examines bounded package-member and
metadata records for those relationship kinds, and does not inflate payloads,
parse an embedded file as another Office document, evaluate formulas, execute
macros or controls, follow external links, or infer meaning from payload
contents. Explicit `data()` and `stream_to()` operations are implemented for a
selected local payload. They read one payload at a time through the owning
`PartView`; inventory construction still never causes an aggregate payload
materialization.

## Lossless and mutation boundaries

- An exact source-backed read, including inventory construction, retains the
  original package bytes. Enumerating the inventory is not a publication or
  normalization operation and does not rewrite an embedded member.
- Inventory entries are observations of authored package state. Ambiguous,
  malformed, duplicate, or unsupported catalog metadata is surfaced as a
  typed refusal or conservative opaque entry according to the owning parser's
  existing policy; the scan never selects an arbitrary duplicate physical
  member or silently discards one.
- Canonical-name lookup is restricted to the indexed local OLE Object/Package
  closure. A name outside that closure, an unresolved relationship, or a
  canonical-name collision is not followed or guessed at; it produces the
  owning typed absence or refusal without changing source bytes.
- The shared inventory is additive and format-neutral. DOCX and XLSB facades
  may expose entries, but neither facade claims that an embedded object is a
  typed document, workbook, executable, control, macro, or safe-to-follow
  link. Embedded OLE, ActiveX, VBA, preview, and unknown payloads remain
  inert bytes and metadata.
- No data or stream is opened merely because an entry is listed. The
  implemented `data()` and `stream_to()` operations select one payload at a
  time and recheck source freshness and cancellation through `PartView` before
  and during the bounded read. `PartView` applies the configured `ReadLimits`
  for that operation; the API does not materialize an aggregate collection of
  embedded payloads.
- `data()` returns `SourcePayloadData`, which retains the private managed
  `PartData` reservation needed by the source-backed package. It does not
  expose `Arc`, `PartData`, package identifiers, runtime handles, or another
  internal ownership escape to callers.
- `stream_to()` bypasses the retained payload cache and streams directly from
  the source-backed part. It preserves the OPC typed partial-output failures
  reported by the underlying writer instead of flattening them into generic
  I/O errors or reporting success after an incomplete write.
- External relationships and external targets are inventory metadata only.
  The scan does not fetch network or filesystem content, resolve an external
  package, execute a target, or treat an external relationship as proof that
  local payload bytes exist.
- A stale or changed source cannot be bypassed through the inventory,
  `data()`, or `stream_to()`; each operation rechecks freshness through
  `PartView` before exposing or writing payload bytes. Signed packages remain
  readable: listing and selected payload reads do not rewrite members,
  invalidate signatures, or strip signature metadata. No mutation path is
  added by this change.
- This change does not add embedded-part editing, replacement, deletion,
  relationship retargeting, or re-packaging. Mutation of embedded content
  remains a separate capability requiring complete ownership, dependency,
  signature, and inverse-patch semantics.

This is a source-backed OLE Object/Package catalog and bounded payload-read
capability for existing OOXML packages. It does not promise payload decoding,
recursive format detection, or semantic editing. Unsupported payload content
remains inert and preservable through the existing source-backed package path.

## Deferred format and capability boundaries

PPTX is intentionally deferred from the initial facade surface. Its embedded
media, OLE, relationship, preview, and presentation-specific ownership rules
require a separate bounded inventory review before exposing the same view.
The shared catalog must not be widened by treating PPTX as an unexamined
DOCX/XLSB alias.

Payload decoding beyond `data()` and `stream_to()` remains deferred. A later
change must specify any typed decoder or long-lived reader contract without
weakening the one-payload-at-a-time limit, source freshness checks,
cancellation, `ReadLimits`, or signed-package read-only behavior.

## Validation status

- The focused common-crate suite passed 17/17 tests, including the embedded
  inventory, indexed canonical-name closure, bounded payload access, and
  refusal behavior covered by that module. The exact command was:

  ```sh
  CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/litchi-0340-target \
  CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_TEST_DEBUG=0 \
  cargo test -p litchi-ooxml-common --lib embedded::tests -- \
  --test-threads=1
  ```

- The DOCX facade library check passed with the same disk-backed target and
  serialized build settings:

  ```sh
  CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/litchi-0340-target \
  CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_TEST_DEBUG=0 \
  cargo check -p litchi-docx --lib
  ```

- The XLSB facade library check passed with the same settings, after the DOCX
  check completed:

  ```sh
  CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/litchi-0340-target \
  CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_TEST_DEBUG=0 \
  cargo check -p litchi-xlsb --lib
  ```

- These Cargo operations were strictly sequential on one disk-backed target.
  The exact target was deleted afterward with:

  ```sh
  find /home/zhuhe/CodeProjects/.cargo-targets/litchi-0340-target -xdev -depth -delete
  ```

- Post-cleanup, the root filesystem had 136 GiB available, `/dev/shm` used 53
  MiB, `/tmp` used 9.7 GiB from unrelated state and was not touched, and
  approximately 18 GiB of memory was available. Swap was nearly full.
- PPTX coverage and mutation coverage: **deferred**, not implied by this
  change.
- Broad workspace gates, strict lint, rustdoc, benchmark, and additional
  resource measurements were not run. No benchmark or process-memory result is
  implied by the checks above.

## Performance claims

`performance_claim: none`

No latency, throughput, allocation, RSS, decompression, I/O, or process-memory
claim is made. Catalog locality, one-payload-at-a-time access, cache bypass for
streaming, and bounded `PartView` reads are correctness and resource-safety
properties, not performance measurements. If later measurements are
collected, they must be attached to a bounded workload and replace this
declaration explicitly; until then, this record makes no performance claim.

## Follow-up

Extend payload decoding only with an explicit bounded contract around the
implemented `data()` and `stream_to()` operations. Add a separately reviewed
PPTX facade, then design embedded mutation only after ownership and
package-topology rules are explicit. Keep all external and executable-looking
payloads inert and preserve unknown OLE Object/Package content losslessly
throughout those changes.
