# ODP legacy migration

The files beside this document are source-preserving relocations of the
former `litchi-odf::odp` implementation, together with the former top-level
handout-master helper under `package/`. They are not wired because their
legacy attachment methods depend on the umbrella's generic package and old
cross-family root types.

The eventual owner split is presentation `model`, `codec`, and `package`
layers, with any neutral page-layout vocabulary moved below the family
boundary. The existing deferred ODP tests record the remaining authoring
surface in `tests/deferred/README.md`.
