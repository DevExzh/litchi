# flate2 1.1.10 provenance

This directory contains the upstream `flate2` 1.1.10 package source from the
crates.io registry cache (`index.crates.io-1949cf8c6b5b557f/flate2-1.1.10`).
The Rust sources, tests, examples, documentation and license files are copied
byte-for-byte. The generated `Cargo.toml` is the only package file changed: it
pins the reviewed relative `zlib-rs` 0.6.8 dependency so this package remains
the exact decoder instance used by the vendored Allsorts bridge. The upstream
`Cargo.toml.orig` is retained unchanged as provenance; Cargo builds the
generated manifest. No flate2 Rust implementation or feature behavior is
changed.

The upstream package is dual-licensed under Apache-2.0 and MIT; the original
`LICENSE-APACHE` and `LICENSE-MIT` files are retained.
