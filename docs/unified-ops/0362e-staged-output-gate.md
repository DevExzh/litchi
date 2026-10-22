# 0362e — Detect output that bypasses edit commit

- Date: 2026-10-21.
- Base: `f5d1a64643`.
- Status: tooling correction validated; the OLE implementation correction and
  final integrated gate remain required.
- Contract: 0351 §2.3; source-backed publication additionally follows 0355a §1.

Coordinator review of the pending OLE owner migration found two public
staging methods that return rendered artifacts:
`object::Edit::put_stream_shared_with_rendered` and
`put_streams_shared_with_rendered`. Their consumers can obtain output before
`Edit::commit`. The former inventory reported neither method, and checking
only the presence of a `commit` method did not catch this second publication
path. The OLE wave is sent back for correction before acceptance, including
all DOC/XLS consumers and preservation of the existing single-render path,
failure atomicity, batches, and exact no-op behavior.

The glossary gate now rejects known artifact-output verbs on the `Edit` of
a registered family with `family_methods` enabled. These include byte output,
`write_to`, save variants, and the `_with_rendered` and `_to_stream` suffixes.
The inventory uses the same predicate and emits `edit_output` rows for
staging types, including legacy `Editor` and `*Transaction` names, even if the
family has not been registered. Builder drafts are excluded from that
inventory category. Output on `Package`, serialization of a detached `Patch`,
and ordinary preparation/read methods are unaffected.

This is a name-based guard, not proof that every possible publication bypass
has been removed. The final workspace inventory and semantic review remain
required. An exploratory pass over the already available rustdoc files found
the two OLE methods above; it is not a fresh whole-workspace acceptance run.

Real-rustdoc fixtures cover the single-stream and batch handoffs, ordinary
output methods, a consuming stream publisher, a public re-export, private
methods, builder drafts, package output and patch serialization. All 30 tests
in `python3 -m unittest tools.test_check_api_glossary` pass. All 50 repository
profiles pass `checked_policy(read_json(POLICY))`, and `git diff --check` is
clean. An attempted `--check-policy` CLI invocation was rejected because the
tool has no such option; policy validation was rerun through that supported
Python API. No format-runtime code changes in this slice, so no native
performance claim is made.

The same review sent the pending OLE wave back for content-free `Debug`
output on its newly exposed owner/edit/patch/prepared types and the required
non-exhaustive error enum. Those corrections are separate implementation work,
not exceptions to the contract. The redaction audit remains the only deferral.

W4, W4b, W5 and W8 integration, performance repairs, the remaining shared/raw
API work, and the final full gate remain open. Nothing is pushed.
