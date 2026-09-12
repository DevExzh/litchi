# XLSB Data Model table rename

`data_model::Transaction::rename_table(table_id, new_name)` stages a table
rename across workbook Data Model records and a complete known XLDM 140
identity closure. The selector is the existing table's stable `TableID`, not
its display name or a storage path. It returns `false` for an equal name;
errors leave the transaction unchanged.

The ordinary detached edit flow is:

```rust,ignore
let snapshot = package.data_model()?;
let mut edit = snapshot.edit();
edit.rename_table(table_id, "Sales")?;
let commit = edit.commit()?;
let changed = package.apply_data_model(&commit)?;
```

The operation proves the full outer table/relationship graph against the
inner closure before rewriting. It changes the table name, matching outer
relationship endpoints, and source-qualified time-grouping references
together. Inner metadata rewrites preserve unrelated lexical content;
TableIDs, column identities, native index/value bytes, and storage member
paths remain unchanged. Applying a commit retains the existing source-bound
conflict checks and inverse behavior.

Unknown or incomplete inner closures, non-XLDM-140 storage, and identity
changes following an opaque payload replacement are refused. This operation
does not refresh a data source, evaluate DAX/MDX, regenerate values, or
perform external I/O. It is not an arbitrary Data Model schema editor.

The neutral XLDM writer computes changed metadata and final storage sizes
before constructing output, using the host's model-part byte cap. Exact-cap
and one-under checks cover UTF-8/UTF-16 directory sizing and storage alignment.
This establishes output admission and preservation behavior; it is not a
claim that all planning allocations are absent or that workflow performance
has been accepted.

The focused `data_model_identity` integration target contains 18 tests. Its
positive relationship-bearing synthetic fixture includes native relationship
index ownership and checks the exact expected metadata rewrites while every
other inner member remains byte-identical. Other cases exercise equal-name
no-op, inverse/save/reopen, source conflicts, limits, escaping, endpoint case
variants, time-grouping composition, signed-package policy, and opaque
replacement refusal. Neutral tests separately cover comment/opaque lookalike
preservation and qualified index ownership.

Positive evidence is synthetic XLDM 140. The native `date.xlsb` fixture is
model-free, and `tdf167689_x15_namespace.xlsx` is a different storage profile;
neither demonstrates native-producer acceptance of this rename operation.
The separate `xlsb-model-identity-performance/requirements.md` defines the
pending matched host/neutral performance measurements. No native Office
round-trip or speedup claim is made here.
