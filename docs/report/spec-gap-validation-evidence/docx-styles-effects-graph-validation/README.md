# DOCX styles-with-effects graph validation

Graph inspection validates unselected effects parts through the existing package checks and eager Styles parser without constructing an additional retained Resource projection. Selected exact-source proof and refusal behavior remain unchanged.

Validated source: `9a1f99b15a572de5366b4470d7bbbc63f4567d057fed38468379678604224346`, on baseline `4d01e69d4f9e8a6a87d3445e2b0e1f136236239e` in an isolated checkout. Independent review approved this exact source hash.

Root validation passed 1,100 library tests and 24 styles_with_effects integration tests, strict Clippy for those targets, and warning-denied rustdoc. Commands and exit codes are retained alongside compressed logs and the resolved Cargo.lock. Runs used TMPDIR=/var/tmp, no inherited Rust warning suppressions, and RUSTDOCFLAGS=-Dwarnings for documentation.

Regression counters verify one retained projection and four graph validations, with existing source, semantic readback, and inverse assertions. This batch makes no latency or allocator-measurement claim; separately collected profiler results require their own provenance review.
