# 0362ul114 — OOXML metadata refusal property context

Status: prepared source repair from `e130751fc503099c4d5fff74a5746b5099d625e1`.

The isolated runtime receipt for the prior tip is:

`(deleted-coord-artifact)/10-tests-ops-ooxml-metadata-attempt-actual-f5-10.log`

It is 17,975 bytes with SHA-256
`5d648b68b997a0de1fc8c8a3adae5644484f2da2ac54afcd1c1cbf4d3654287c`.
Five OOXML metadata tests passed; the two failures were the create-time and
set-time `contentType` cases, both reporting `property() == None` instead of
the required shared field name `content_type`.

## Cause

The common OOXML metadata leaf correctly returns a typed
`RefusalKind::UnsupportedProperty` with a fixed, content-free reason.  The
refusal intentionally carries no authored field.  The facade's generic
`native_error` conversion preserved the refusal code and format but had no
request context, so it also left the facade property's optional field empty.
Moreover, the assignment context helper was DOCX-only and returned the
OfficeCLI input spelling `contentType`, rather than the shared metadata field
name.

## Repair

- `Assignments::unsupported_property()` now exposes the canonical
  `content_type` field name for all OOXML facade consumers.
- Complete metadata creation identifies the first populated field outside the
  shared OOXML core/application field sets and retains its canonical
  `Field::name()` for refusal attribution (`content_type`, `last_backup`, or
  `codepage`, among others).
- XLSX, PPTX and XLSB edit and create adapters reattach that contextual field
  only when the mapped native error is an otherwise unnamed
  `UnsupportedProperty`; existing error context and fixed refusal reasons are
  untouched.  The XLSB source-backed path uses the same adapter.
- DOCX uses the same canonical assignment spelling, keeping the unified
  facade consistent across OOXML formats.

The direct generic conversion remains property-free when no request context is
available.  The source-backed and create consumers therefore retain both the
redacted native reason and the public facade field required by record 0253.

Validation in this leaf is source-only as authorized; no Cargo/build/rustdoc
run is claimed.  The prior runtime receipt above is preserved as the failure
baseline.
