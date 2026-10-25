# 0362ul67b — PPT VBA policy inheritance and decompression admission

This return packet corrects the PPT owner-policy implementation from
`0362ul67a`. It is source-ready on `uo2/vba-raw-policy-0362ul67`, with no
Cargo, native, rustdoc, benchmark, or macro-execution command run.

## Boundary corrections

`OpenOptions::vba_project_limits` now records whether the caller selected the
profile explicitly. `Reader::presentation_with_options` and
`SourceBackedPackage::presentation_with_options` inherit the retained profile
when the supplied options are otherwise new, while an explicit setter always
replaces it. This keeps the owner policy through `Package::from_source` and
does not constrain unrelated outer PPT streams.

The retained VBA profile is copied through metadata, primary edit, scoped,
text-edit, and slide-order reopen helpers. `slide_order::Snapshot` stores the
profile for adoption and every publication reopen. The encrypted builder
option handoff also preserves the selected profile. `Presentation::vba()`
continues to be the only canonical read operation; no mutable owner or
`vba_with` twin is introduced.

For the nested VBA payload, PPT now passes
`min(VbaProjectLimits::max_cfb_bytes, Limits::max_cfb_bytes)` to the outer
decompressor before it allocates output. The inner CFB policy uses the same
physical ceiling, so a large outer cap cannot bypass a stricter project cap.

## Regression coverage

The VBA writer tests cover the inverse outer/inner cap, eager reader policy
inheritance, explicit profile replacement, source-backed projection,
`Package::from_source(source, OpenOptions::new())` adoption, and metadata
republish retention. The retained low project cap is asserted through the
canonical inert `vba()` path, which never executes macros.

## Source checks

```text
api_glossary.json JSON parse and duplicate-key check  PASS
rustfmt --edition 2024 --check (changed Rust files) PASS
git diff --check                                  PASS
staged-size and binary guards                     pending until commit staging
```

The coordinator should run the PPT feature matrix, the DOCX encrypted gate,
glossary/rustdoc profiles, and the full consumer gate after merging.
