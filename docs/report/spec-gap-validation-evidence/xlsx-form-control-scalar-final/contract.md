# Existing XLSX form-control scalar lifecycle

This batch extends the existing inert worksheet owner and properties codec with
paired properties/VML scalar edits. It does not claim control execution,
rendering, creation, deletion, or native-application acceptance of edited files.
The retained fixtures establish input provenance; library save/reopen tests are
separate evidence.

## Admitted operation

An existing control is selected by zero-based position or exact authored name.
The owner proves the worksheet, control-properties, DrawingML and VML identity
closure under the pinned `OwnerProfile`. A scalar edit validates the selected
properties value and its corresponding VML `ClientData` representation before
publishing both parts atomically. Unsupported or ambiguous mirrors fail closed.
The mappings and omitted-value rules are grounded in
[the mirror evidence](../xlsx-form-control-mirror-evidence.md).

One control per worksheet batch is admitted. Multiple scalar fields on that
control compose in order. Two independently staged ordinary form-control edits
on the same worksheet produce a structured `Conflict::FormControls` at join;
the rejected edit is returned intact. Three-way planning uses the same worksheet
owner key and requires explicit resolution for overlapping branches. This
conservative conflict scope reflects
the shared VML owner and the bounded single-control batch.

The ordinary worksheet API stages through `set_form_control_scalar` or a
selector-pinned `edit_form_control` handle. `PackageChange::FormControl` records
the logical control position and before/after properties, including its inverse.
Ordinary in-memory patches require the exact immutable source workbook; their
inverses require the corresponding result. This pins the complete owner graph
without treating two equal changed payloads as proof of unchanged ownership.
Durable patches retain the existing complete serialized-source precondition.
The source-backed API retains the owner read set and checks it during replay
and publication. Neither path follows external resources or executes macros.

## Preservation and limits

Exact no-ops retain source bytes. An effective scalar edit replaces only the
selected properties and VML payloads; unrelated package members and unrelated
XML are preserved. Inverse patches restore the retained source payloads.
Source versions, signatures, owner relationships, and incoming edges participate
in validation. Paired output is validated before stream publication begins;
an I/O error from a caller's sink is still an I/O failure, not a transactional
rollback of bytes already accepted by that sink.

Owner limits bound scalar operation count, changed parts, source and output
bytes, staging, relationships, shapes, XML structure, and retained read sets.
Cancellation and resource refusal must occur without publishing a partial pair.
Generated source-backed properties/VML allocations retain clone-shared memory
and object reservations. The generated collection retains its projection lease;
properties and opaque XML aliases retain their source payload lease. Charges
are conservatively grouped and can remain until the last alias is dropped.
Managed before-payloads remain shared through their existing source handles.
Export to an eager package explicitly materializes a bounded copy rather than
letting an unleased shared allocation escape. The unmanaged compatibility path
does not imply execution-budget accounting.

Retained source-only formula tokens such as `#REF!` are preservation-only:
exact no-ops and unrelated scalar changes are admitted, but replacement through
this scalar API is refused.

Graph-changing formulas, list edits, multi-control batches, broader dialects,
and control creation/removal remain outside this scalar profile. The leaf codec's
detached editing capabilities do not imply support for these owner operations.
