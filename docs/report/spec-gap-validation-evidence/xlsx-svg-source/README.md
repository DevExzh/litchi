# XLSX drawing source helper

`drawing::SourceDrawing` borrows a complete worksheet drawing member and
indexes direct pictures in source order. Each picture retains numeric EMU
geometry for two-cell, one-cell, or absolute placement, exact element ranges,
namespace context, raster relationship identity, and its admitted SVG owner.
Global relationship references also include opaque descendants and non-picture
objects, so later graph cleanup can account for those references.

Escaped picture and SVG-owner records retain the original source lifetime.
The compile-fail doctest prevents mutation of the source while a cloned record
is still used. Source leases compare identity and length and keep complete member
bytes out of derived debug output. Namespace completion is exposed through validated
record methods; the low-level function accepting namespace pairs is private.

The scanner bounds input, events, depth, pictures, namespace state, relationship
references, and completed fragments. It refuses ambiguous or MCE-owned SVG
candidates and avoids typed inference inside foreign or unknown extensions.
Namespace and relationship admission precede retained allocations.

`review.json` binds the independently reviewed files and isolated root checks.
The root checkout contained the listed four files over the recorded base
commit, excluding in-progress worksheet lifecycle edits. The retained
`workspace.Cargo.lock` is the lockfile used for those checks. To repeat the
checks, use the source-helper commit, copy that lockfile to the workspace root,
and run the recorded commands with the repository's test fixtures available.
The temporary checkout and its build output are removed after verification.

This prerequisite does not establish worksheet attachment/detachment support,
generated-package schema validity, performance results, or native application
acceptance. Those require separate lifecycle evidence. Unit-bearing
`ST_Coordinate` values remain outside the numeric EMU profile.

## Retained namespace follow-up

The standalone diagnostic `retained-scope-probe.rs` found a remaining resource
issue after the prerequisite review. It builds 32 pictures with distinct SVG
relationship IDs and 128 inherited namespace declarations of about 1 KB each.
On the helper source recorded in `review.json`, it prints:

```text
input_bytes=149433 pictures=32 retained_svg_source_bytes=4205334
```

The final count sums the lengths returned by each owned SVG projection's
`source()` accessor. It excludes the namespace models and other allocations;
it is neither a peak-memory measurement nor a timing result. The helper
currently namespace-completes and retains the inherited declarations for each
SVG projection, repeating shared context. Reducing that amplification remains
open. Any correction must preserve namespace meaning in opaque payloads and
the validity of independently serialized fragments.

The diagnostic uses only the public drawing API. Build `litchi-xlsx` at the
prerequisite revision, then compile this file with Rust 2024, the built
`litchi_xlsx` library supplied through `--extern`, and its dependency directory
supplied through `-L dependency=...`. Later corrected revisions may expose
contextual source through a different accessor; this file records the original
reproduction rather than a forward-compatibility promise.
