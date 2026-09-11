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
