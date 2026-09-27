# 0796 direct attribute-boundary probe

`baseline.rs` is a byte-for-byte copy of the OPC helper at
`crates/litchi-opc/src/xml_attributes.rs`.  `candidate.rs` is a byte-for-byte
copy of the rejected 0794 OPC candidate.  The two files are deliberately local
modules so one evidence binary measures the same quick-xml workload against
both implementations.  They are not workspace dependencies and this probe
does not edit production sources.

The probe uses quick-xml 0.41.0 directly.  It constructs every `BytesStart`
and its backing input before the timed region.  The named non-inlined owners
are exactly:

* `before_construct` and `after_construct`: construct the selected checked
  iterator, pass a reference to `std::hint::black_box`, and drop it, once per
  loop iteration.  Their checksum is a leg-independent accumulation of loop
  indices; it does not use the iterator layout.
* `before_consume` and `after_consume`: construct a fresh checked iterator,
  drain it through its first error or natural end, and black-box each yielded
  attribute.  It stops immediately after an error; an error-free drain makes
  only the natural terminal `next()` call.  The returned accepted count,
  value-length checksum, and first-error marker are independent of the
  implementation when the semantic oracle passes.

The clock surrounds the selected named owner through a function pointer; the
same one-call dispatch and timer read are part of every native sample.  Case
generation, input construction, the quick-xml/baseline/candidate semantic
oracle, clone checks, checksum validation, JSON serialization, and file I/O
are outside the clock.
Warmups invoke the same owner with the same iteration count and are not
reported.  A report contains the exact source input as UTF-8 for short cases
or hex for the 4096-byte cases, iterator sizes for both legs, the full
structured error/position oracle summary, and per-sample elapsed time plus
returned counters.

The oracle compares both helpers with quick-xml's checked iterator until its
first error or end.  It compares borrowed key/value bytes, all error variants
and positions, terminal `None` behavior, and clones made after transitions
0, 1, 4, 5, 32, and 33.  `--self-check` executes only this independent oracle
for all frozen cases; it never calls a named timing owner.

The copied helpers retain external cfg(test) module declarations. This standalone probe excludes cfg(test); cargo test and --all-targets are not used. Helper unit-test evidence is inherited from 0794. All timed error cases end at the error: a valid tail after an error is not independently exercised, limiting recovery-adapter coverage.
