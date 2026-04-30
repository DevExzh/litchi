#![no_main]

use libfuzzer_sys::fuzz_target;

// NOTE: `xml-minifier` is declared as `proc-macro = true` in its Cargo.toml,
// meaning it only exports compile-time procedural macros (`minified_xml!`,
// `minified_xml_str!`, `minified_xml_format!`) and has no runtime callable
// API. Proc-macro crates cannot be invoked at runtime by their dependents.
//
// This harness is preserved as scaffolding so the fuzz directory layout
// matches the conventions of sibling crates (e.g. `litchi-rtf/fuzz`,
// `litchi-formula/fuzz`). To make it functional, the upstream crate would
// need to expose `minify_xml` (currently private in `src/lib.rs`) through a
// non-proc-macro library -- for example by splitting the implementation
// into a sibling crate or by removing `proc-macro = true` and re-exporting
// the helper. Per task constraints, neither `crates/xml-minifier/Cargo.toml`
// nor `crates/xml-minifier/src/` may be modified here.
//
// The body below coerces the fuzzer input to UTF-8 the way `minify_xml(&str)`
// would expect, mirroring the eventual public-API contract.

fuzz_target!(|data: &[u8]| {
    let _ = String::from_utf8_lossy(data);
});
