# 0362ul89at — PPTX downstream raw-limit consumers

Status: source-ready downstream repair based on
`8304c5a03085d9fb49c5d3db46f4eae3b1adcaae`.

The remaining workspace consumers now pass explicit bounded policies to the
changed PresentationML raw codecs:

- eager shape-scene reads in the PPTX scene cache and the two PPTX benchmark
  examples use `shape::Limits::DEFAULT`;
- the comprehensive transition example passes `transition::Limits::DEFAULT`
  to both write and read;
- the SmartArt scene check passes the shape policy explicitly;
- the advanced-feature example aliases animation and metadata `Limits` so
  each codec receives its own policy without an import collision.

The four named PPTX example command snippets now select the `litchi` package
and its `pptx` feature.  The advanced-feature command was corrected to the
same valid feature spelling while its limits aliases were repaired.

`git grep` auditing of non-PPTX crate consumers found no additional raw calls
without limits; PPTX-internal callers already carry their selected shape or
transition policy.  `rustfmt --edition 2024 --check` and `git diff --check`
pass.  No Cargo/build/rustdoc run, generated evidence, fixture, executable,
or `3rdparty/` path was used or added; combined consumer validation remains
coordinator work.
