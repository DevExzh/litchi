# XLSX SVG source-projection contextual profile

This directory owns a bounded, source-only profile for ordinary worksheet
`SpreadsheetDrawing` SVG discovery and contextual SVG projection. It is a
diagnostic companion to the lifecycle design; it does not measure package
editing, worksheet commits, relationship planning, cloning, attaching,
detaching, native Office acceptance, or final lifecycle performance.

The harness opens deterministic drawing XML through the frozen public
`litchi_xlsx::drawing::SourceDrawing` scanner and stops at the shared
`litchi_drawingml::svg_blip` codec boundary. It exercises four separately
reported operations:

* contextual scan/read of retained raw fragments and namespace context;
* standalone lazy export/readback of each SVG value;
* scalar embedded-reference replacement followed by export/readback; and
* a deliberately small output-cap refusal lane.

Every fixture contains direct `xdr:pic` pictures inside `twoCellAnchor`
elements. The regular corpus has 16, 32, or 128 pictures and 0, 32, or 128
root namespace declarations, each declaration carrying a deterministic
approximately 1 KiB URI. Each picture contains an opaque QName attribute and
element so the export/edit lanes check that contextual projection keeps those
values. The retained native-inspired corpus is the exact 149,433-byte,
32-picture, 128-declaration input described in `corpus-manifest.json`.

The planned collection is three fresh processes per lane, two warmups per
process, and twenty measured samples per process. Each receipt records the
fixture hash, elapsed time, process-local allocation calls and bytes, peak
live bytes, retained raw-source sum, `source()==None` count, namespace
context counts, public `NamespaceContext::shares_storage` identity, semantic
readback checks, and refusal count. `/usr/bin/time -v` supplies process RSS.
The allocator and RSS observations are diagnostics for this host/scanner
boundary; they are not claims about whole-package memory or XLSX lifecycle
cost.

Receipts carry both FNV-1a and SHA-256 corpus identities; the compact
`results/*/corpus-sha256.tsv` files make the full fixture identity reviewable
without parsing every sample.

## Frozen comparison

The before snapshot is commit `8b5838c59` and the after snapshot is commit
`fc44c4e6c`, each built in its own detached worktree. The after snapshot is the
corrected host freeze with raw-fragment-cap preflight and its boundary
regression test. The after source files are recorded by these SHA-256 values:

```text
crates/litchi-xlsx/src/drawing/source.rs
  952cbc65b89314c61ecee623be91af64665d3c92b6827a58fbceaf8a722efb00
crates/litchi-xlsx/src/drawing/source_tests.rs
  7278bc01cf6bb84234e2f359064a3cf85cff6a8f270ec8ea9b7a83d868c9e34e
crates/litchi-drawingml/src/svg_blip.rs (unchanged shared boundary)
  e8a1e480d249b834f286e0deec083f8044257faf3b95905fb5b771e83f71605a
```

The before source hashes and every built binary hash belong in each result
bundle. A source manifest is generated before the build and after all lanes;
the runner refuses to continue if that manifest changes. The runner requires
`PROFILE_FROZEN=1`, a fresh caller-owned Cargo target directory, and an
offline locked build. It refuses to run with injected Rust flags.

Run a frozen snapshot from its detached worktree with:

```text
PROFILE_FROZEN=1 \
  PROFILE_LABEL=after-fc44c4e6c \
  CARGO_TARGET_DIR=/var/tmp/litchi-xlsx-contextual-profile-target \
  ./docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/run_profile.sh
```

Use `PROFILE_LABEL=before-8b5838c59` for the before worktree. The label pins
the expected full commit and frozen source hashes; the runner also refuses a
dirty production `crates/` tree. The target directory must be fresh and is
removed on normal exit. Do not point it at the repository target directory or
at another agent's scratch. `validate_provenance.py` checks the executable
hash, commit, and three source hashes while the binary exists; `verify.py`
checks the retained receipt after cleanup. `summarize.py` recomputes the
report from raw JSON and `/usr/bin/time -v` files. Lane stderr files are
retained and a warmup error aborts the process instead of becoming a sample.

No acceptance gate is set by this directory. In particular, a lower
`retained_raw_source_bytes` sum or a `source()==None` count of zero is not by
itself a proof of lower memory use. The context-sharing field is meaningful
only where all pictures expose a context and the public storage-sharing
predicate agrees.
