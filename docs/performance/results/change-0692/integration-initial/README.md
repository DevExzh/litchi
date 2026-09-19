# Narrow facade feature gate attempts

The warning-denied `pptx` facade test build fails on two existing unused
helpers in `unexpected_format.rs`. A `pptx,xlsx` retry hits an existing
unused-mut in a non-OOXML detection test. Adding `ods` hits an existing unused
`PreparedOdfProbeError.bytes` field. All three commands/logs are retained here.

The final facade run returns to the requested `pptx` scope with normal rustc
warning policy (no RUSTFLAGS override). Owner checks, Clippy, tests and rustdoc
still deny warnings. No facade production/test source or dependency changed.
These pre-existing warning gaps are disclosed rather than suppressed in source.
