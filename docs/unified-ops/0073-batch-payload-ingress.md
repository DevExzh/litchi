# 0073 — Complete OfficeCLI batch payload ingress

This is a W6 prerequisite within M2, following the DOCX single-operation and
publication loop in records 0070–0072. It does not complete M2 or the full
goal. The previous goal turn made verified implementation progress; the Git
metadata restriction still permits work on the implementation.

## Implementation

The former `batch.rs` placeholder now owns `BatchItem` and `BatchMode`, both
re-exported by the facade. `BatchItem` accepts all 21 keys specified by design
§3, including `op` as an alias for `command`. The implementation was checked
against the base checkout's `3rdparty/OfficeCLI/src/officecli/BatchTypes.cs`;
no OfficeCLI code was copied.

- Borrowed strings remain borrowed when the deserializer lends them. Escaped
  values own their decoded strings; `into_owned` detaches the whole payload.
- The existing `Props` decoder handles object and `key=value` array forms.
  Array entries split on the first equals sign. Typed booleans, exact integers,
  numbers, strings, null-clears and ordered repeated assignments are preserved.
- Matching envelope keys is ASCII case-insensitive. Serialization emits
  canonical keys, with `type` represented by Rust's `kind` and `dumpVersion`
  by `dump_version`. Optional null fields are absent; property nulls stay null.
- Unknown and duplicate envelope fields refuse. A duplicate cannot silently
  overwrite a prior value, even when one spelling is `op`, differs in case,
  or carries null. A command is required and cannot be whitespace-only.
- Debug output and decoder errors omit authored values. The explicit Rust
  envelope validator returns `corrupt_batch`; Serde decoding uses the caller's
  decoder error type with a content-free message.
- `BatchMode` defaults to `Atomic`; `BestEffort` must be selected explicitly.
- `batch::from_bytes` and `from_bytes_with_limits` decode a complete array,
  refuse trailing JSON and return the unified typed error. Checked
  `batch::SourceLimits` bound complete input bytes before decoding and item
  count before decoding any excess item. Defaults are 512 MiB and 100,000
  items. The byte limit also bounds property payloads and escaped strings.
  Errors retain a content-free `batch::Error` source, structural limit or
  line/column facts, and the item index when the failure belongs to an item.

The payload transports rather than executes: unknown command names, dump
versions, signed 32-bit numeric fields and text are preserved for later
command admission. Integers outside the OfficeCLI signed 32-bit wire range
are malformed transport; negative in-range values remain for command-specific
range validation. There is no implicit newline conversion. Serde's raw
`Deserialize` remains a compatibility path; untrusted input uses the bounded
decoder. `serde_json`, already a test dependency, is now a runtime dependency
for that decoder. No new format dependency was added.

The crate-level status text also now reflects the DOCX structural editing,
creation and saving implemented in the preceding checkpoints.

## OfficeCLI deltas

Strict ingress rejects unknown envelope fields instead of silently dropping
them, and rejects repeated envelope fields instead of taking the last value.
Malformed property array entries already refuse through `Props` (record
0069). These are deliberate strict-default changes under ADR 0034. Property
values retain their scalar types rather than being eagerly stringified; the
registry owns their later coercion. Capability reports and the generated
registry delta remain unchanged because this checkpoint exposes no additional
executable verb.

## Executor findings and next implementation

A read-only review of the current leaf and ops APIs established the next
concrete work, replacing stale assumptions in record 0015 §5.4:

1. `document::Edit<P>` supports eager/source owners and fallible cloning.
   Main-document source commits chain through effective overlays without a
   ZIP serialization/reopen. Physical graph, story and section families still
   have their own effective-owner restrictions.
2. `WorkEdit::commit` currently drops the leaf patch and diagnostics. The
   executor needs a result retaining both and the new `Docx` owner.
3. Current navigation resolves through the pre-run `Docx`. Structural edits
   and content-dependent selectors require staged navigation or a commit
   boundary before the next dependent item. Resolving every item against the
   initial snapshot would violate sequential batch semantics.
4. Homogeneous main-document sets are the first coalescing target. Runs must
   respect operation limits and engine/family boundaries. Error-path replay
   must attribute commit failures and isolate failed best-effort items.
5. Composite patches must retain ordered leaf patch steps with their own
   bases. Inversion reverses the steps; concatenating leaf operations from
   distinct bases is invalid.

Execution acceptance must cover repeated targets, sequential structural
addressing, exact no-ops, atomic rollback, per-item errors, source fences,
managed-source policies, operation-cap boundaries, eager/source parity and
inverse replay. Paired 1/10/100/1k/10k measurements remain required. This
record makes no batch-execution or performance claim.

## Verification and cleanup

- Minimal ops: 413 tests/doctests passed, including 16 new batch tests and
  eight new runnable examples.
- All-feature ops: 994 tests/doctests passed.
- All-feature all-target ops Clippy passed with warnings denied.
- All-feature ops rustdoc passed with warnings denied; DOCX-only facade
  compilation verifies the public re-exports.
- Crate-boundary checker, OfficeCLI delta regeneration check, scoped rustfmt
  and `git diff --check` passed.
- Independent review checked borrowing, alias collisions, bound ordering,
  typed error retention, item attribution and property-array replay. The
  decoder module is public and the facade re-exports it.

Representative reproduction commands:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --all-features
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --all-features --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-ops --all-features --no-deps
CARGO_NET_OFFLINE=true cargo check -p litchi --no-default-features --features docx
python3 tools/check_crate_boundaries.py
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
```

The local rewritten base remains `32ad64e3207655b4ba558fc2793d9a7396655cdf`
and is an ancestor of HEAD, so that revision needs no rebase. No temporary
worktrees or fixture links were created. Required staged evidence/binary
checks printed nothing. The final audit covered 324 changed/untracked paths
and found no prohibited evidence or compiled executables. Cleanup removed
21,290 build files (30.5 GiB) and 11 scratch/cache paths, including the
generated untracked lockfile.

`git add -A` remains blocked because the linked worktree's
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` cannot be
created on the read-only filesystem. No commit was created. All implementation
changes remain in the checkout, and the full goal remains active.
