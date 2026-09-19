# `probe0699-refusal`

This is a diagnostic native probe for the opened-PresentationML capture path.
It checks refusal precedence and the cost of rejecting malformed packages. It
does not modify production sources and does not support a production
performance claim. The fixture construction remains byte-for-byte equivalent
to the retained refusal matrix; the emitted probe identity is `0699-refusal`.

The fixtures are authored with the public `Package::new`, presentation
authoring, and slide APIs. Each authored package is serialized and reopened
before mutation. Malformed variants then clone the public OPC graph, mutate a
root payload or relationship through `get_part_mut`, and are adopted with
`Package::from_opc_package`. This keeps fixture construction outside all
timers and exercises the same immutable package shape used by the native
capture path.

The matrix contains exactly these ten bounded cases:

* `small-valid`: a small valid package with a complex first slide and speaker notes;
* `generated-12x8-valid`: a valid generated 12-slide by 8-text-box package;
* `early-name-error`: an early duplicate slide-name attribute;
* `late-root-error`: a late invalid slide root after the complex first slide;
* `late-root-error-mce`: the same late invalid root with valid MCE markup on the first two slides;
* `late-missing-relationship`: a late missing presentation-to-slide relationship;
* `late-missing-relationship-mce`: the same late missing relationship with valid MCE markup on the first two
  slides;
* `notes-invalid-tail`: a valid notes graph with a malformed notes XML tail (direct typed XML refusal);
* `mixed-conformance`: mixed transitional/strict slide conformance;
* `slide-raw-overlimit-16m-to-64m`: a valid first slide payload padded above 16 MiB and below 64 MiB, which
  reaches notes validation's generic slide-root refusal.

All ten fixtures are authored, serialized, reopened, and mutated before any
timed capture loop. The matrix command defaults to 100 samples and 5
warmups:

```text
probe0699-refusal matrix
probe0699-refusal matrix 100 5
```

To time one existing case, use its exact case name:

```text
probe0699-refusal case late-root-error 100 5
```

For profiler runs, `profile` repeats one case with bounded iterations and
prints only final success/refusal counts and an assertion checksum:

```text
probe0699-refusal profile late-root-error 1000
```

The profile loop includes capture, the exact result assertion, and snapshot or
error drop in its denominator. Fixture construction, including construction
of the unselected cases, is outside that loop. Zero counts and counts above
the finite probe bounds are rejected.

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

For matrix and case samples, only `Package::opened_presentation` is inside a
sample timer. Snapshot inspection, error formatting, correctness assertions,
and snapshot drop occur after the timer. The probe has no allocator feature or
allocator counters.
