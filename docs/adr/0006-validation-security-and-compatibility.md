# ADR 0006: Validation, security, and compatibility

- Status: Accepted
- Date: 2026-07-31

## Preservation and validation

`Preserve` is the default serialization mode. Untouched entries, streams,
ordering, compression, timestamps, unknown markup/records, namespace choices,
and lexical details are retained when possible. `Normalize(profile)` is an
explicit planned transformation.

Validation never mutates. It returns a deterministic `Report` of structured
issues with severity, stable code, semantic/physical location, evidence,
specification citation, compatibility impact, and repair availability. `Err`
means validation could not complete because of I/O, cancellation, or limits.
Fatal safety failures stop opening immediately.

Repair uses previewable `RepairPlan<NonDestructive>` or explicitly requested
`RepairPlan<Destructive>`. Applying a plan yields a patch and before/after
diagnostics. Normal save never repairs. Normalizing conversion enumerates every
unrepresentable feature and requires an explicit keep, approximate, rasterize,
embed, flatten, or drop policy.

Normative Microsoft/ISO specifications under `3rdparty/` define canonical
semantics. Observed Office deviations live in versioned compatibility profiles
with cited rules, affected Office builds, and regression fixtures. Readers
preserve real-world quirks; writers do not bake quirks into semantic models.

Existing files retain their detected profile. New OOXML uses a broad
Transitional-compatible baseline unless explicitly pinned to Strict or another
target. Format and flavor are content-derived, never inferred from a filename.
Known incompatible extensions are typed errors; extensionless paths are valid.

Numbers rooted traversal never treats successful protobuf decoding as type
evidence. The document root must contain one type-1 payload, a referenced sheet
must contain one type-2 or type-3 payload, and a table drawable must contain one
recognized table-info payload. Canonical type 6000 and the fixture-backed
legacy type 6003 table-info encodings are mutually exclusive at one owner.
Referenced models prefer one type-6001 payload and accept one type-6000 legacy
payload only when the canonical type is absent. Missing references, duplicate
sheet/drawable/model ownership, duplicate typed payloads, ambiguous canonical
and legacy ownership, and malformed known payloads fail before publication;
valid table-info bytes under an unrelated message type remain opaque.

The separate structured compatibility scan preserves its historical primary-
message classification. A malformed preferred type-6001 model is an error,
while a type-6000 candidate is emitted only after complete table extraction
succeeds, allowing genuine table-info objects and malformed legacy false
positives to be skipped. Secondary messages cannot promote an object into the
global candidate set. These rules are parity behavior, not permission for the
ordinary document graph to expose orphans.

Structural identity changes traverse recognized incoming references as one
atomic graph edit. References to the same document are updated; external-book
references, VBA source text, and other inert external identities are not
rewritten by textual guesswork. If an unmodeled formula-like field can refer to
the changed identity, the safe facade returns a typed dependency block. Markup-
compatibility alternatives are edited only when their effective semantics are
modeled; otherwise the operation is blocked before any bytes are published.

Serialization is deterministic unless a `Clock`, actor identity, or
cryptographic RNG is explicitly supplied. Existing timestamps are preserved;
new files do not consult ambient time, process identity, filesystem metadata, or
global randomness. Cryptographic operations never fall back to deterministic or
ambient pseudo-randomness.

DOCX/DOTX/DOCM, XLSX/XLTX/XLSM, and PPTX/POTX/PPSX/PPTM-style variants share
their concrete document type with a validated runtime `Flavor`. Flavor-specific
operations use capability views. Promotion/demotion is explicit, checked for
active content, and atomic. Template instantiation and applying a template are
distinct planned transformations.

## Security boundaries

- External links, DDE, OLE objects, VBA, ActiveX, fields, formulas, and embedded
  documents are inert. Fetching or execution requires an explicit provider;
  core crates contain no ambient network client.
- Encryption uses consuming `Locked<T> -> Sensitive<T>` transitions. Sensitive
  documents save encrypted unless explicitly consumed into plaintext.
  Credentials are non-clone, redacted, zeroized, and non-serializable.
- Signed input is `Signed<T>`. Editing requires consuming it through explicit
  signature stripping or edit-and-resign. Integrity, coverage, cryptographic
  validity, trust, revocation, and time are distinct statuses.
- Office protection is enforced by default. Unlock and audited bypass are
  explicit capabilities and are not confused with encryption.
- Weak legacy encryption is decode-only by default. Authoring it requires an
  explicit policy and diagnostic; algorithms are never silently downgraded.
- Only a complete verified sanitization produces `Sanitized<T>`. Best-effort
  cleanup returns an ordinary document and diagnostics. Redaction seals its
  patch unless inverse material is explicitly retained under an authorized
  policy.
