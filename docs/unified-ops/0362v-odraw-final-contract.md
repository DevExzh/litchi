# 0362v — ODraw final raw admission and builder namespaces

- Date: 2026-10-23.
- Base: `2b42e80421`.
- Status: source migration complete; coordinator validation is pending.

This slice closes the remaining ODraw raw and authoring-surface gaps. The
zero-copy `Children` iterator now admits at most `Limits::max_records` parsed
children and reports the attempted count in the typed `LimitExceeded` error.
Its size hint is capped by the remaining admission budget as well as by the
wire bytes available.

Image-file discovery keeps bounded allocation at both scanner boundaries.
`image::all` caps its initial vector reservation by the retained file policy,
and recursive collection appends nested `BStore` files directly while checking
the aggregate file ceiling. The change removes the temporary per-store vector
and does not alter depth-first order, delay resolution, source bytes, or the
retained raw/BLIP policy used by `File::blip` and `File::into_owned`.

Authoring drafts now live in canonical builder modules:
`litchi_odraw::builder` owns property and shape builders, while
`litchi_odraw::image::builder` owns BLIP, FBSE and store builders. The
`write` modules retain only framing, copy and digest primitives. DOC, PPT,
XLS, imgconv, ODraw tests and the checked qualification probe migrate in the
same source change; no compatibility aliases or deprecated paths remain. The
ODraw glossary profile bans the removed draft paths so a later public-surface
regression is caught by the existing gate.

The leaf was kept source-only as required while the coordinator's concurrent
builds were active. Focused validation for integration is:

* `cargo test --offline -p litchi-odraw --all-targets --all-features`;
* all-target/all-feature checks for `litchi-doc`, `litchi-ppt`,
  `litchi-xls` and `litchi-imgconv`;
* the ODraw glossary/rustdoc check and workspace formatting.

No performance or native qualification result is claimed here. Generated
targets, binaries and raw evidence remain outside the repository.
