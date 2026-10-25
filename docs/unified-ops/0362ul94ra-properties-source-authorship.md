# 0362ul94ra — properties source/authorship classification

The `0362ul94r` repair correctly changed the shared-relationship graph
assertion to a typed refusal, but its application-properties planner still
classified an XML lexical error by searching the error message. That could
promote a retained source failure to an authored `invalid_value`; the same
ambiguity existed when prepared core/application output reached a writer.

This follow-up establishes the provenance at the codec boundaries. The
application and core readers reject decoded XML 1.0 control characters while
they are still retained-source reads, so the transaction wraps them as
`properties::Error::Package` and the facade reports `corrupt_file`. Authored
application values are validated before source splicing or new-part encoding,
and authored core values are prevalidated before merge; those failures remain
`properties::Error::Invalid` and the facade reports `invalid_value`.

Generated application invariants now use the shared typed `Internal` error,
and prepared/generated versus authored codec adapters no longer inspect
user-controlled error text. Tests cover source entity controls, authored NUL
controls, the transaction classification paths, and the `native_error`
facade mapping.

This is a source-only repair. No Cargo, build, rustdoc, runtime, native, or
performance command was run; the coordinator must rerun the affected common,
format-adapter, and facade gates against this commit.
