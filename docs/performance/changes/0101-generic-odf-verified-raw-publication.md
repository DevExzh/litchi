# Change 0101: verified raw publication for generic ODF content edits

Date: 2026-08-14

Status: Accepted as correctness and preservation coverage. No latency,
allocation, peak-memory, RSS, or source-I/O result is claimed.

## Scope

The generic packaged ODF owner used by the chart facade now publishes a changed
`content.xml` through the accepted family-neutral raw ZIP preservation
boundary. Eligible unchanged Store and Deflate members retain their exact local
spans and central-directory metadata, including extra fields and physical
member order, while the changed content member is regenerated and the complete
package is reopened.

The migrated generic path opts into verification of every unchanged source
payload before raw publication. This preserves the previous logical writer's
CRC/decompression safety behavior while removing recompression and retaining
unknown physical framing. The ordinary shared `replace_content_xml` helper used
by existing ODT, ODS, and ODP fast paths remains lazy and does not inflate
opaque media. An exact semantic no-op returns the original package bytes before
either verification mode is considered.

## Fail-closed boundary

Raw preservation requires a first, stored, extra-free canonical `mimetype`
entry whose bytes match the manifest root media type. Every copied member must
have consistent local/central framing and, where present, a valid data
descriptor. Noncanonical `content.xml` manifest aliases, stale manifest sizes,
signatures, encryption, unsupported compression, and unsupported ZIP layouts
retain the established logical rebuild or typed refusal behavior.

The generic opt-in additionally reads every unchanged file so malformed
payload CRC or decompression state cannot bypass the former writer's checks.
This is intentionally not applied to the already accepted shared raw fast
paths.

## Verification

Regressions cover deliberately different central and physical local order,
Store/Deflate members with local and central extra fields, exact raw framing,
complete semantic reopen, signed and encrypted exact no-ops, signature and
manifest-size fallbacks, canonical MIME/root consistency, `/content.xml` and
`./content.xml` aliases, corrupted descriptor CRC and sizes, and corrupted
opaque payloads. Common tests also prove that the default shared helper does
not inflate an unreadable opaque member while the opt-in verifier rejects it.

The complete ODF-common and ODT all-target suites, warning- and
deprecation-denied focused tests, strict Clippy, rustfmt, diff checks, and
independent adversarial review gate this change. A matched release corpus is
still required before making any performance claim for the generic facade.
