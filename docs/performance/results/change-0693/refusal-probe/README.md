# `probe0693-refusal`

This is a retained supplemental native probe for the opened-PresentationML
capture path from change 0693. It is a guardrail for refusal precedence and
the cost of rejecting malformed packages. It does not modify production
sources and does not support a production performance claim.

The fixtures are authored with the public `Package::new`, presentation
authoring, and slide APIs. Each authored package is serialized and reopened
before mutation. Malformed variants then clone the public OPC graph, mutate a
root payload or relationship through `get_part_mut`, and are adopted with
`Package::from_opc_package`. This keeps fixture construction outside all
timers and exercises the same immutable package shape used by the native
capture path.

The default matrix contains ten bounded cases:

* a small valid package with a complex first slide and speaker notes;
* a valid generated 12-slide by 8-text-box package;
* an early duplicate slide-name attribute;
* a late invalid slide root after the complex first slide;
* the same late invalid root with valid MCE markup on the first two slides;
* a late missing presentation-to-slide relationship;
* the same late missing relationship with valid MCE markup on the first two
  slides;
* a valid notes graph with a malformed notes XML tail (direct typed XML refusal);
* mixed transitional/strict slide conformance;
* a valid first slide payload padded above 16 MiB and below 64 MiB, which
  reaches notes validation's generic slide-root refusal.

The matrix command defaults to 100 samples and 5 warmups:

```sh
probe0693 matrix
probe0693 matrix 100 5
```

Each case reports the serialized archive hash of its valid authoring base and
the exact prepared in-memory input graph hash. The graph hash includes every
part name, content type, payload digest, per-part relationship tuple, and
package-root relationship tuple. Malformed variants are never serialized:
their graph hash is the input identity used by the timed capture. Refusal
cases retain an expected typed `Debug` form, including the exact
offset-bearing duplicate-name message captured during setup. Every warmup and
timed sample asserts the same expected result. Valid controls also assert the
expected slide count and first slide name. Raw nanoseconds are printed after
each capture call returns.

Only `Package::opened_presentation` is inside a sample timer. Snapshot
inspection, error formatting, correctness assertions, and snapshot drop occur
after the timer. The probe has no allocator feature or allocator counters.
