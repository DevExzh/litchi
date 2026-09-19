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

The default matrix contains eight bounded cases:

* a small valid package with a complex first slide and speaker notes;
* a valid generated 12-slide by 8-text-box package;
* an early duplicate slide-name attribute;
* a late invalid slide root after the complex first slide;
* a late missing presentation-to-slide relationship;
* a valid notes graph with a generic-invalid notes XML tail;
* mixed transitional/strict slide conformance;
* a notes payload larger than 16 MiB and smaller than 64 MiB.

The matrix command defaults to 100 samples and 5 warmups:

```sh
probe0693 matrix
probe0693 matrix 100 5
```

Each case reports its serialized input archive hash and a sorted OPC graph
hash. The graph hash includes every part name, content type, payload digest,
per-part relationship tuple, and package-root relationship tuple. Refusal
cases retain an expected typed `Debug` form; the duplicate-name parser's
offset-bearing message is checked by typed XML error family plus its stable
duplicate-attribute text. Every warmup and timed sample asserts the same
expected result. Raw nanoseconds are printed after each capture call returns.

Only `Package::opened_presentation` is inside a sample timer. Snapshot
inspection, error formatting, correctness assertions, and snapshot drop occur
after the timer. The probe has no allocator feature or allocator counters.
