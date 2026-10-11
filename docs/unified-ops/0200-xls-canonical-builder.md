# 0200 — Canonical XLS Builder

Fresh XLS authoring now uses the concrete `Builder` type, exposed at the crate
root and focused `builder` module. The old `Writer` name is removed. Existing
worksheet, style, formula, metadata, drawing and encryption authoring methods
remain on this single staged model.

`Builder::build()` returns a strictly admitted immutable `Package`. An empty
builder acquires one blank `Sheet1`; `Package::new()` uses that same path.
Repeated builds retain staged content, and changing or dropping the builder
does not mutate any previously returned package. The builder has no save,
stream or byte-output methods: ordinary publication uses the Package methods
and their atomic save policy.

Encrypted builds admit the generated image with the builder's configured
password. The owner retains the encrypted source and readable projection;
ordinary output remains refused. Explicit `to_encrypted_bytes()` exports the
retained image. Native tests, examples, facade fixtures and performance harness
consumers now use the owner-based output route, including encrypted fixtures.
Typed authoring failures remain available beneath `package::Error::Reader`.
Two drawing fixtures had page entries but zero pivot fields. Strict admission
exposed their invalid SXPI sequence; they now author valid pivots while keeping
their drawing/object-ID assertions unchanged.

## Verification

The combined XLS and XLS/encryption facade suite passes 2,234 tests and
doctests, with one ignored. Four focused Builder tests cover minimal/repeated
builds, immutable ownership after later authoring, encrypted image retention
and strict refusal of the malformed pivot page-entry case.

Warning-denied all-target Clippy and documentation pass. The glossary audit
checks 24 banned names and one transaction family. The boundary audit passes
for 66 packages and 268 internal dependencies, with 11 existing debt items.
The performance harness library/binaries compile with `xls-source-attribution`.
Independent review confirms the ownership shape and flags the encryption
entropy gap below. Both repository lockfiles are unchanged; builds use an
external lockfile and target. The rewritten base remains an ancestor.

## Remaining scope

The inherited opt-in encryption path still obtains salts/verifiers from
`SysRng` during build. This does not meet the ADR's final Builder determinism
rule: explicit entropy or Package security verbs must replace that path.
Default unencrypted builds are deterministic. Canonical encrypted atomic-save
verbs and the broader security API remain pending; this record claims the
Builder/Package ownership and publication shape, not full security compliance.

Fresh build still serializes and parses the admitted image. No zero-copy,
zero-reparse or performance acceptance is claimed. Unified XLS node reads,
CRUD, batching and history remain pending, along with the 10k operation target,
broader format coverage, Python/CLI integration and full-goal acceptance.
