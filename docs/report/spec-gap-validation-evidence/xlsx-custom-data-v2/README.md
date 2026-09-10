# XLSX Custom Data and Connections v7 evidence

This isolated candidate records the v7 Custom Data and Connections source-preservation batch. The validation source is based on commit `b2e2ff8e1fdabd4ab16a67ddf077e6968bf5e668`; the workspace audit was performed against current HEAD `463ab2a5989bba49ffd95277dc29422a363de7f8`. The frozen source receipt is `source-manifest-v7.json`, with 17 SHA-256-bound source/documentation paths. `source-audit.json` records the byte checks and the current-HEAD preservation audit.

The batch covers bounded typed Custom Data catalogs and source-bound Connections transactions. Custom Data admission validates the recognized workbook/relationship graph, keeps typed properties and inert payloads source-bound, preserves relationship/content-type tokens and explicit empty relationship parts through removal/save/reopen/inverse, and keeps edits atomic. Connections parsing and publication validate namespaces, schema structure, MCE selection, processing instructions, identifiers, graph edges, and configured XML/resource limits. Supported source edits splice scalar attributes, update known nested tables/fields/parameters at their original spans, retain comments/PIs/foreign attributes and MCE context, and move raw items for same-parent reorder. Persistent namespace contexts and shared URI storage avoid multiplying inherited declarations; the source publication path preflights the final size and uses bounded forward copying. Top-level connection matching and reorder use fallibly reserved ID indexes/membership sets. The v7 tests exercise 8,192-connection scalar commit/reopen and reorder paths without timing claims.

An explicit `extLst` replacement owns the complete selected owner subtree; old attributes, children, comments, PIs, and inactive branches are not merged into a requested replacement. Selected subtree removal may remove its descendants, while unrelated source remains retained. Ambiguous cross-branch moves or edits that cannot prove source ownership are refused. These are bounded source-preservation guarantees for the supported OOXML structures, not a claim that arbitrary unknown XML can be safely edited.

The implementation leaves payload and connection content inert. It does not evaluate queries, refresh data, render workbooks, establish external connections, or provide native producer interoperability evidence. It makes no latency, peak-memory, or performance-improvement claim. Native producer coverage, unsupported cross-branch opaque imports, and semantics outside the recognized graph remain open. The committed Survey implementation is preserved separately: all four Survey source files at current HEAD are present and byte-identical in the candidate, but Survey is outside this v7 source manifest. No deletion hazard was found: every tracked `crates/litchi-xlsx` file at current HEAD exists in the candidate, and there are no `crates/litchi-xlsx` changes between the validation base and current HEAD.

The root v7 gate passed 1,378 unit/integration tests and two doctests. Strict all-target/all-feature Clippy with `-D warnings`, warning-denied rustdoc, formatting, and whitespace checks passed. The compressed raw gate logs are `tests.log.gz`, `clippy.log.gz`, and `rustdoc.log.gz`; `Cargo.lock.gz` is the workspace lock snapshot used by this candidate. The receipt records hashes for every evidence artifact and distinguishes the validation base from the later workspace HEAD. The containing Git commit identifies feature publication; the recorded base and audit commits identify the sources used during validation.

## Reproduce the recorded gates

Run from the candidate root with Rust `1.95.0`:

```text
cargo +1.95.0 test -p litchi-xlsx --all-features
cargo +1.95.0 clippy -p litchi-xlsx --all-targets --all-features -- -D warnings
RUSTDOCFLAGS="-D warnings" cargo +1.95.0 doc -p litchi-xlsx --all-features --no-deps
```

The recorded root commands did not pass `--offline` or `--locked`; the stored `Cargo.lock.gz` permits an optional hermetic replay after restoring it as `Cargo.lock`. For source fidelity, verify `source-manifest-v7.json` against the candidate paths and inspect `source-audit.json`. Format the 16 Rust paths in the manifest with the repository's Rustfmt configuration, then run `git diff --check`; the one remaining manifest path is the documentation matrix.

Evidence files:

- `source-manifest-v7.json`: 17 frozen source/documentation hashes.
- `source-audit.json`: candidate/current-HEAD/source-preservation and deletion-hazard audit.
- `tests.log.gz`, `clippy.log.gz`, `rustdoc.log.gz`: compressed root gate output.
- `Cargo.lock.gz`: candidate workspace dependency lock snapshot.
- `receipt.json`: artifact hashes, gate counts, commands, toolchain, and commit provenance.
