# Shared Custom Data codec validation

The common X14 model and XML codec replace the XLSX-owned duplicate codec and serve the XLSB owner under development. Explicit limits cover input/output bytes, strings, nodes, depth, namespaces, attributes, XML events and UTF-16 UID units. Parsed extension fragments retain standalone namespace closure on the value; there is no process-global namespace cache. Source-bound semantic no-ops preserve the original raw XML.

The current nine-file candidate is `source-v3.json`. Direct extension children use the core Transitional SpreadsheetML namespace; each extension has an optional URI and exactly one wildcard element child. Regressions cover XML-only whitespace, unbound prefixes, raw attribute delimiters, explicit event limits and UTF-16 UID boundaries. Earlier `source.json`, `source-v2.json` and their reviews document rejected captures, not this revision. The final regression rejects direct Strict SpreadsheetML extensions.

The root overlaid only these nine files and the retained lockfile on an isolated checkout of the recorded base. Full common crate tests, including doctests, passed (300 tests); the targeted XLSX Custom Data library selection passed 26 tests. Strict library Clippy and warnings-denied rustdoc passed for both crates. All nine source hashes matched the built checkout and shared source after validation. The XLSX selection does not claim all integration targets ran.

```sh
cargo test --locked -p litchi-ooxml-common --offline
cargo test --locked -p litchi-xlsx --lib --offline custom_data
cargo clippy --locked -p litchi-ooxml-common -p litchi-xlsx --lib --offline -- -D warnings
RUSTDOCFLAGS="-D warnings" cargo doc --locked -p litchi-ooxml-common -p litchi-xlsx --no-deps --offline
```

Current logs and exit codes are retained in `v3-*`. Independent final review approves this shared codec/model scope; all seven prior blockers are cleared (`v3-review.json`). XLSX package-level XML/UID limit exposure remains a separate host follow-up. XLSB graph ownership and binary connection bindings are outside this codec validation scope. No runtime speedup is claimed.
