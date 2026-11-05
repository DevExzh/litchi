# 0196 — XLS eager source limits and open options

## Change

The eager XLS reader accepts a checked `SourceLimits` profile through
`OpenOptions::source_limits`. The physical dimensions are `input_bytes`,
`entry_bytes`, and `aggregate_bytes`, using the shared 512 MiB / 512 MiB /
2 GiB defaults. Construction refuses zero and values above finite hard caps
with `Error::InvalidLimit`; ingress refuses excess with `Error::Limit`.
The facade preserves these typed errors as `invalid_argument` and
`limit_exceeded`.

Reader ingress checks physical input before parsing CFB and checks every
declared stream length before reading XML or BIFF payloads. Already parsed
CFB adoption checks its recorded physical length and stream catalog without
reparsing. Opaque, metadata and nested streams count; the root mini-stream
backing does not count twice. The raw adoption preflight lets the facade
check its default profile before its metadata read. Adoption cannot undo
allocations made during the caller's earlier CFB parse.

`OpenOptions` now uses consuming `password`, `leniency`, and
`compatibility_profile` setters, with `get_*` accessors. The replaced
`with_*` methods are removed and banned in the XLS glossary profile.
Passwords are redacted in Debug output. Native tests, examples and the facade
use the new spellings.

## Verification

The final native XLS + facade XLS/encryption suite passes 2,199 tests and
doctests, with one ignored. Coverage includes finite builder boundaries,
input refusal before any container read, all-stream accounting and precedence,
adopted-container and password-returning ingress refusal before payload reads,
exact-boundary valid workbook acceptance, source-profile preservation through
option setters, credential redaction, and fallback reader-position restoration
on success, limits and I/O failure. Existing source materialization, edit,
replay, encryption and both facade read routes remain covered.

Warning-denied all-target Clippy and rustdoc pass. The XLS glossary profile
passes all seven banned names. Independent review found
no remaining blocking defect. The crate-boundary audit passes for 66 packages
and 268 dependency declarations with 11 existing debt items. Cargo uses an
external copied lockfile with cached cc 1.6.0; repository lockfiles are
unchanged. The performance-harness library and binaries compile. Scoped
rustfmt and diff checks pass. The rewritten base at `32ad64e320` remains an ancestor. No iWork
crate is modified. No benchmark acceptance is claimed.

## Remaining scope

This record does not complete the canonical XLS owner. The positional reader
still has its existing `SourceBackedLimits` and options, and the focused edit
engine retains its object/CFB profiles from record 0195. Those policies must
be joined by the canonical Package ingress and retained through publication;
the eager envelope alone does not claim that migration. The positional
materialization adapter keeps its prior selected source/materialization
checks and CFB defaults through a private admitted-owner decode handoff; it
does not silently substitute the new eager default. Internal focused-edit
validation likewise retains its explicit CFB profile. No total-memory
bound or performance improvement is claimed.

The next owner wave must preserve both eager and positional read behavior:
BIFF5 and encrypted inputs cannot simply be redirected to the stricter
source-backed or cell-values readers. An Arc-backed Package should retain the
original parsed owner, expose an owned read-only Workbook projection, and
admit focused edits only on demand. Exact publication must retain source bytes
and source fences. Signed and encrypted state must remain explicit; opened
workbooks must never be serialized through the fresh Writer as a shortcut.
Native owner/Builder migration, facade XLS writes, broader Tier 2 and Tier 3,
Python/CLI, and full-goal acceptance remain open.
