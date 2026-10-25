# 0362ul89am — PPTX minimal-target compile repair

This repair is a fresh leaf from `3377583e2b` (the reviewed 89al
composition). It addresses the three failures reported by the minimal
all-target qualification without changing runtime behavior:

- use the imported `Error` type in the presentation capture fallback;
- compile the test-only `SlideRootProof::new` constructor only for tests;
- extract the error with `Result::err()` in the borrowed XML scanner test so
  `XmlScan` does not need a production `Debug` implementation.

The existing source consumers and tests remain unchanged apart from these
compile/lint corrections. Source-only rustfmt and diff checks are required
before handoff; Cargo, native builds, runtime tests, rustdoc, and the complete
qualification remain coordinator work.

No generated evidence, fixture, executable, or `3rdparty/` path is included.
