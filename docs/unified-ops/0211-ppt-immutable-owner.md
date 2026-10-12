# 0211 — Immutable PPT owner and owned presentation

The primary `litchi_ppt::Package` is now an immutable, one-Arc owner.
Its cheap, `Send + Sync` clone retains the validated CFB and parsed presentation.
The former generic mutable package is explicitly `raw::Reader<R>`; native
low-level consumers and frozen benchmark routes use that spelling.

`Presentation` is also a one-Arc owned projection. It outlives its package,
shares the parsed document and successful lazy Pictures cache across clones,
and retains source freshness checks for positional reads. Captured infallible
metadata such as slide count remains available without a source read.
Fallible source-dependent reads reject a stale source.

Byte, bounded sequential-reader, path, positional and already-parsed CFB
constructors preserve checked policies. Preparsed adoption moves the original
CFB index and bytes into shared ownership without reopening or copying them.
Positional adoption reuses the existing shared CFB. The facade now retains this
native owner rather than maintaining its own sealing implementation. Positional
detection applies the selected PPT limits before admission, including profiles
above default ceilings. `retain_source(false)` disables edit ownership; the
read projection still retains sources needed for lazy reads and freshness.

`to_bytes`, `write_to`, `save` and `save_with_durability` publish exact original
CFB bytes. Atomic save fences the source before streaming, after streaming and
before replacement, preserving the destination on failure. Sequential output
can have written a prefix on I/O failure. Encrypted owners retain their actual
40/128-bit profile; ordinary publication returns typed `EncryptionPolicy`
before output, and `to_encrypted_bytes` explicitly exports unchanged ciphertext.
Passwords are not retained by the owner. Password-required handoff preserves
the parsed CFB for the facade, including builds without decryption support.

## Verification

Full all-feature suites passed 1,372 native PPT tests/doctests (11 existing
ignored cases) and 1,549 facade tests/doctests. Final focused runs cover ten
canonical owner regressions and seven facade PPT cases, including above-default
positional policy retention. Tests check exact bytes/reopen, Arc identity,
projection lifetime/thread transfer, lazy cache sharing, stale reads/output,
atomic destination preservation and temporary-file cleanup, typed sink errors,
bounded input and encrypted output policy.

Warning-denied all-target Clippy and rustdoc pass for PPT and the facade.
The minimal PPT glossary audit passes 23 banned-name checks; it registers no
primary transaction family yet. The PPT-only facade build and all nine
minimal-feature canonical owner tests pass. The crate
boundary audit passes for 66 packages, 268 internal dependency declarations and
11 existing debt items. All 15 OfficeCLI parity generator tests pass and the
generated table is current against the base checkout's OfficeCLI source.

The standalone performance harness and umbrella consumer compile, and
`ppt_picture_selectors_preserve_semantics_and_phase_read_evidence` passes.
The frozen routes retain their original reader implementation through explicit
`raw::Reader` spelling. This check establishes semantics and phase-read evidence,
not a new timing result. Reproduce with `cargo test --manifest-path
tools/perf-baseline/Cargo.toml --lib
ppt_picture_selectors_preserve_semantics_and_phase_read_evidence`.

Validation uses external offline lockfiles and build targets. The workspace
copy selects cached cc 1.6.0; repository lockfiles remain unchanged. Core commands:
`cargo test -p litchi-ppt --all-features`,
`cargo test -p litchi-ops --all-features`, and
`cargo clippy -p litchi-ppt -p litchi-ops --all-features --all-targets -- -D warnings`.

## Remaining scope

Primary PPT Edit/Patch/Commit, canonical Builder/new, unified PPT navigation and
write dispatch remain pending. This milestone does not claim benchmark speedup,
performance acceptance, Python/CLI completion or all-format goal completion.
Existing XLS scaling/performance follow-ups remain open.

The base remains an ancestor at `32ad64e3207655b4ba558fc2793d9a7396655cdf`, so
no additional rebase was needed. The full goal remains active.
