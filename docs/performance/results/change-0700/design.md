# Change 0700 design review

Status: source/design review passed; terminal evidence recommends rejecting the
direct `memmem` candidate for the shared production path. The hypothesis was
deliberately one local search change in the shared buffered MCE precheck. This
packet does not propose a frame, namespace ownership, stream, parser, or
output-model change.

## Measured target

Change 0699's source and profile review found that the marker-free path enters
`process_markup_compatibility` and spends most of its sampled work in the
fixed-window search for the MCE namespace URI. The current precheck is:

```rust
if !xml
    .windows(NAMESPACE.len())
    .any(|window| window == NAMESPACE.as_bytes())
{
```

The existing direct `memchr` dependency already exposes the safe, byte-slice
substring search used by other OOXML code. The candidate should change only
the predicate to this exact expression:

```rust
if memchr::memmem::find(xml, NAMESPACE.as_bytes()).is_none() {
```

The fully qualified spelling keeps the production diff to the predicate and
does not add an import or a dependency. `memmem::find` searches arbitrary byte
slices, returns the first matching byte offset, and returns `None` exactly when
no complete occurrence exists. For this non-empty constant needle, its boolean
result is equivalent to the existing `windows(...).any(...)` predicate for
empty, shorter-than-needle, boundary, repeated-prefix, and arbitrary-byte
inputs. The library's existing `memchr` workspace dependency is already
locked and owned by this crate; no dependency or feature change is needed.

## Semantic boundary that must remain fixed

The input-limit check must remain before either search. The no-marker branch
must retain its output-limit check, `Cow::Borrowed(xml)` result, and default
`Report`. A marker anywhere in the byte stream, including a comment, text,
malformed XML, or a repeated occurrence, must still enter the existing reader
path. The search is only a dispatch precheck; it is not XML validation and
must not be narrowed to element names or namespace declarations.

Once the predicate reports a marker, the candidate must leave the existing
`Reader`, `BoundedOutput`, `Frame`/`Ctx`, namespace, selection, output, error,
and report code byte-for-byte unchanged. The streaming MCE APIs are separate
and must not be edited or claimed by this change. The active-offset helper's
marker and delimiter searches are also outside this one-line hypothesis and
should remain unchanged unless a separately reviewed experiment justifies
them.

## Required correctness controls

The focused test set should compare baseline and candidate behavior for the
search parity matrix before relying on end-to-end timings:

- empty and shorter-than-URI buffers;
- the URI at offset zero, at the final valid window, and across repeated
  prefixes/overlaps;
- near matches differing at the first, middle, and final byte;
- arbitrary non-UTF-8 bytes and repeated URI occurrences;
- the URI in element text, a comment, and malformed XML, proving that the
  precheck remains lexical and byte-oriented;
- late URI placement after a long marker-free prefix, plus an adversarial
  repeated-prefix buffer, to exercise the measured scan without changing the
  branch taken;
- input and output limits on both marker-free and marker-present inputs,
  including the ordering where an input-limit refusal precedes the search and
  an output-limit refusal remains specific to the borrowed path;
- exact `Cow` ownership, output bytes, `Report`, and error identity for each
  applicable case.

Representative A/B controls must include marker-free refusal, real edit and
no-op workflows, and an MCE-positive workflow. The MCE-positive control is
required because the new precheck also runs when a marker is present; any
search win must not hide setup cost or alter the reader path. A late-URI and
repeated-prefix synthetic control should be reported separately from normal
documents. Timing claims require balanced process legs, warm-up, retained
distributions, and the packet's allocation/output/oracle checks; a single
microbenchmark is insufficient.

## Review decision

The one-line `memchr::memmem::find` predicate is source-compatible, and the
focused, oracle, marker-control, integration, and evidence gates pass. The
terminal follow-up nevertheless shows +1.466% to +3.442% median total latency
on the three real workflows across balanced pairs, alongside the measured
code and stack growth. The direct candidate is therefore rejected for shared
path retention. A future lower-setup first-byte search plus exact
`starts_with` check is a separate hypothesis requiring a fresh source review
and the same semantic and timing controls.
