# Existing-resource projection reuse

The measured effects publication path repeatedly reconstructed typed style
projections. This change reuses an already validated resource during existing
resource publication and candidate readback, after checking the actual part
bytes, content type, conformance, and caller limits.

The ordinary graph inspector still constructs and validates resources for all
recognized targets. Only the selected target can use the private proof-reuse
path. Every other main/glossary target retains the previous MCE and typed-style
refusal behavior. The proof is an exact comparison against the resource bytes,
not an assumed relationship identity or a matching byte count.

Publication still uses the source-checked OPC replacement seam. Readback
captures the actual graph token and checks aggregate metrics as well as the
existing owner, relationship, content-type, and limit state. Add/remove and
ordinary `put` readback retain their ordinary loading paths.

## Correctness evidence

Root checked these source versions:

- `crates/litchi-docx/src/styles/effects.rs`:
  `aebd25d0fba937b4760c39cbe67839fee496f9036a9cacbc36265175873d6bb8`
- `crates/litchi-docx/tests/styles_with_effects.rs`:
  `0c5a011cfa2c5ed34f1421b31ab1b3ff75bfbeb92171e1f7c5d3bcbb37c8ebbd`

The following commands passed without lint suppression, with `TMPDIR=/var/tmp`:

```text
cargo test --locked -p litchi-docx --test styles_with_effects --offline
24 passed

cargo test --locked -p litchi-docx --lib styles::effects::tests --offline
3 passed

cargo clippy --locked -p litchi-docx --lib --tests --offline -- -D warnings
```

The new regression puts unsupported `mc:MustUnderstand` and duplicate typed
`w:numId` content in the secondary glossary effects part. Main-owner reads and
removal refuse, and package bytes remain unchanged. Existing tests retain
source/inverse, opaque XML, signature, graph-closure, and exact/one-under limit
coverage. `git diff --check` passed for both changed Rust files.

The native two-owner unit fixture records one actual load, one separate
readback, and five typed projection builds during patch application. The
facade performs no nested candidate clone; direct patch application retains
one candidate clone. These are fixture-specific operation counts, not latency
or allocation measurements.

## Performance evidence boundary

The retained profile for production commit `506e6f8e5` remains evidence for
that earlier revision. It does not measure this change. A new matched capture
is required before attributing a runtime or allocation improvement to this
revision. No speedup, native application acceptance, or scaling claim is made
here.
