# 0362uu — Proposal-only replay result-rule registration

This record adds a review helper for the final API inventory. It does not
accept or rewrite glossary policy rows.

`tools/propose_api_result_rules.py` consumes one exact rustdoc JSON document
and one checked `api_glossary.json` profile:

```text
python3 tools/propose_api_result_rules.py \
  --json <frozen-rustdoc.json> --profile <profile-id>
```

For each nominated `immutable_owners` type it considers only direct public
`apply`/`apply_*` methods not already present in `result_methods`. A candidate
is emitted only when `ReturnTypes.matches_result` and `matches_owner` both
prove the exported `Commit`/`BatchCommit` identity. Local aliases are matched
by nominal identity. Each path is first checked with its own Rust defaults (so
a zero-parameter alias remains a zero-argument policy row); a non-default
generic row is rendered from the alias target's declared bindings, never by
guessing the underlying argument order. Concrete family `Commit` paths are
preferred, and any equivalent checked paths are reported as alternatives.
Unsupported, ambiguous,
missing-identity, wrong-owner, and non-Commit results remain explicit
`unresolved` findings. Raw, record, Patch-owned, and History helpers are not
nominated and therefore are not inferred. Existing rows are listed separately.

The JSON output is review input only: the helper never edits policy or marks a
row accepted. Real rustdoc fixture tests cover an owner result, an alias,
wrong-owner and non-Commit results, a missing identity, and preservation of an
existing rule plus the read-only CLI path. Final source-freeze, per-profile
rustdoc capture, full inventory disposition, and qualification remain external
steps; this record makes no coverage or performance claim.
