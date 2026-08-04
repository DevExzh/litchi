# ODS legacy migration

The files beside this document are source-preserving relocations of the
former `litchi-odf::ods` implementation. The package adapters in `package/`
are the former top-level data-pilot, named-expression, and definition-package
implementations. They remain unwired because they refer to the old umbrella
package/rebuild helpers and pre-split ODS modules.

The active ODS layers already own the canonical spreadsheet model and codecs.
These adapters will be reduced to package-scoped transactions there; no
umbrella aliases are retained. Existing deferred ODS coverage documents the
same boundary in `tests/deferred/README.md`.
