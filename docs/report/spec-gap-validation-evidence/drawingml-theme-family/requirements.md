# DrawingML theme-family verification requirements

Scope: the exact `http://schemas.microsoft.com/office/thememl/2012/main`
`themeFamily` fragment, specified by vendored MS-ODRAWXML sections 2.4 and
5.17. This batch does not establish host package placement or native acceptance
of newly authored output.

| Requirement | Required evidence |
| --- | --- |
| Required unqualified `name`, `id`, `vid` attributes | Positive and missing/duplicate/namespace-confusion tests |
| GUID domain matches imported `a:ST_Guid` | Local normative grammar and valid/invalid lexical tests |
| Optional family-namespace `extLst` uses DrawingML extension-list type | Namespace/placement tests and opaque extension replay |
| Name follows XML string semantics | Empty, escaping, literal whitespace and character-reference tests |
| Immutable shareable snapshot | Compile-time Send + Sync and allocation profile |
| Source-preserving edits | Exact no-op, unrelated-byte preservation and semantic readback tests |
| Atomic reversible publication | Inverse, independently reopened source, stale-source and failure tests |
| Bounded untrusted XML and output | Exact and over-limit tests, including opaque ancestry |
| Useful detached authoring | Compiled public API example and canonical readback |
| Production checks | Full crate tests, strict all-target Clippy, warning-denied rustdoc, formatting |
| Scoped performance evidence | Reproducible harness, raw samples, toolchain and source hashes |

The wider audit and performance program remain open. Fragment creation and
metadata edits must not be described as complete theme-part or package CRUD.
