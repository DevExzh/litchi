# Core SpreadsheetDrawing content-part read gate

The final isolated Rust 1.95.0 run passed 12 content-part tests, 15 SVG read
tests, 51 SVG lifecycle tests, strict library Clippy, and strict rustdoc.
`root-verification.json` identifies the successful raw run files and hashes
the retained evidence and owned source files.

The snapshot was based on `8702fd4db8723acceb7deb51bcb40ff66604bf10` with the
core read changes overlaid. It excluded the unfinished `form_control` export.
This is a focused feature/regression gate, not a full shared-workspace gate.
The later pivot list commit was validated separately. The content-part test
was a symlink to the frozen owned test; its hash is recorded separately from
the regular-file manifests.

The first content-part command lacks `--locked --offline`; its before/after
lock hashes match the retained lockfile. The remaining final commands include
those flags. Initial Clippy runs used incomplete snapshot configuration and
failed; those diagnostic logs remain unchanged. The authoritative Clippy run
is `runs/clippy_litchi_xlsx_with_config.txt`, after copying the repository's
toolchain, Cargo configuration, and `clippy.toml`. No warning suppression was
used for that run. Earlier Rust 1.98.1 results are not this final gate.

Coverage establishes direct Strict/Transitional core ownership, source/typed
anchor pairing, required relationship grammar, bounded borrowed XML payloads,
opaque extension preservation, and refusal of malformed or ambiguous owners.
It does not establish nested `xdr14` ownership, mutation, recursive graph
closure, typed Ink payloads, native application compatibility, or performance.

For a fresh replay, use the feature commit containing this directory, copy
the retained `Cargo.lock` to the checkout root, and run the final recorded
commands with a private Cargo target. The absolute paths in the raw logs
identify the original snapshot and do not need to be reused.
