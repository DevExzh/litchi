# 0362e — Detect output that bypasses edit commit

- Date: 2026-10-21.
- Base: `f5d1a64643`.
- Status: tooling correction and the OLE implementation correction are
  integrated; the final integrated gate and semantic review remain required.
- Contract: 0351 §2.3; source-backed publication additionally follows 0355a §1.

Coordinator review of the pending OLE owner migration found two public
staging methods that returned rendered artifacts:
`object::Edit::put_stream_shared_with_rendered` and
`put_streams_shared_with_rendered`. Their consumers can obtain output before
`Edit::commit`. The former inventory reported neither method, and checking
only the presence of a `commit` method did not catch this second publication
path. Commit `39a7a2729b` removed those publication methods and migrated the
DOC/XLS consumers to `Edit::commit().package()`/`into_package()`, preserving
the single-render path, failure atomicity, batches, and exact no-op behavior.

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
the then-present two OLE methods above; it is not a fresh whole-workspace
acceptance run.

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
non-exhaustive error enum. Those corrections are now integrated: the
structural/content-free implementations are recorded by `1a5f56f45d`, and
the error enum is non-exhaustive in `02c9f44c78`. The redaction audit remains
the only deferral.

W4 and W4b source integration is assembled (the W4/W4b assembly commits are
in the current ancestry); their native/performance qualification and the final
full gate remain open. DOCX raw/model/helper work,
signed-native replay, W8 native qualification, remaining shared/raw API work,
and the final full gate remain open. W8 content-matching replay is integrated
across the format consumers; it is not yet a performance acceptance claim.
W5's OLE correction is not an open exception. Nothing is pushed.
