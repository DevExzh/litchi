# 0362ul10 — W5 OLE owner/history qualification protocol

- Date: 2026-10-24.
- Source lane: W5 OLE owner and retained-history follow-up from 0359j.
- Status: executable source/build protocol committed; native qualification and
  performance acceptance remain pending coordinator execution.  The handoff,
  fidelity and independent-validator repair is recorded in 0362ul10a.

The existing W5 source and allocation lanes did not provide a paired native
qualification for the complete owner/property selector set.  This packet adds
`tools/native_qualification/prepare_ole_owner_inputs.py`,
`tools/native_qualification/manifests/ole-owner.json`, and
`tools/native_qualification/ole_owner_paired_runner.py`.

Preparation is deliberately external to production source trees.  It requires
the clean pre-owner migration `0ce4bb8b61f03a1e5497633f51b954e76f98aa8c`, a
clean runtime-selected candidate commit, and the reviewed consumer inputs:

| input | SHA-256 or revision |
| --- | --- |
| immutable full OLE probe workload | `ad4c52c01333fbc13c7b3cca18e2c0a0d2ccf89d5c07972e2e453fa9bd8d713a` |
| current explicit-CFB-limits consumer adapter | `dc11e6533c43434c314bd359ad68e458555291f4c134b44b13c1f3ce29fb4508` |
| small/large owner clone probe | `0a23c99ed8958e0df1aa51e5f37e7a949328149074b2096613c972d2b8cf6e98` |
| paired standalone Cargo lock | `bc5a133f0668142d18d637b2b196097e855c8890697cb3bbf43aeb937145f04a` |

The preparer writes two standalone drivers whose path dependencies point to
the selected source roots.  The baseline driver uses the immutable probe; the
candidate driver uses only the reviewed CFB consumer-arity adapter.  The
baseline production tree has an empty allowed-overlay set and its Git tree and
status are checked both before and after preparation.  Receipt validation
rehashes every driver file and rejects source, adapter, lock, clone, workload,
or policy mutations.

The runner builds fresh normal and allocator binaries for both endpoints with
one offline locked Cargo job.  It records toolchain, source, driver, lock,
target, command, event-journal, and executable identities.  ABBA refuses to
start if any identity changes.  The timed phase calibrates all eleven routes
(`open`, `begin_edit`, `noop_commit`, `replace_stream`, `read`, `edit`,
`doc_tracked_revision`, `xls_cells`, `xls_cells_stage`, `doc_history`, and
`xls_history`) and both clone sizes to a 500 ms target.  Five pinned-CPU
`before,after,after,before` blocks are retained, with every accepted sample
leg required to span at least 250 ms.  The probe remains the owner of exact
output and forward/inverse replay checks; the runner adds no fidelity work to
the timed operation boundary.

The allocator phase is paired and explicitly non-claimable.  It retains the
129 DOC/XLS packages and 128 patches while capturing ordinary property
read/edit, XLS cells/staging, and both history routes.  The existing counting
allocator's live/peak counters and its `record_dealloc`/`record_realloc` paths
are bound to the receipt; allocator elapsed values are never treated as native
timing.  Host quietness is checked before, after, and every 100 ms during each
child, with fixed 19.2 GB target preflight, 20 GB target, and 20 GB address
space limits.

Validation performed in this source-only packet:

```text
python3 -m unittest tools.test_prepare_ole_owner_inputs tools.test_ole_owner_paired_runner
python3 -m py_compile tools/native_qualification/prepare_ole_owner_inputs.py tools/native_qualification/ole_owner_paired_runner.py
```

No Cargo, native binary, timing, or acceptance result is claimed by this
record.  The final candidate source ref, preparation receipt, build receipt,
raw paired output, and reduced coordinator decision must be recorded outside
the repository and reviewed before W5 closure.
