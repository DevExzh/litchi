# OTH legacy migration

The files beside this document are the source-preserving relocation of the
former `litchi-odf::oth` implementation. They are intentionally unwired: the
legacy authoring facade imports ODT directly, which would violate the family
boundary, and still relies on umbrella package types.

The eventual owner split is `codec`/`model`/`package`/`authoring`, using a
neutral text capability where needed. The deferred OTH regression remains
under `tests/deferred/legacy` and is outside Cargo's direct integration-test
discovery until that dependency boundary is resolved.
