# 0541 combined worksheet planning error-order guard

0540 measured about 19–20% of selected planning instructions in the direct XML
reader/namespace children of each separate validation and parser loop. Before
sharing those events, a candidate must preserve whole-validation precedence.
This batch adds public integration tests to establish that behavior. It does
not change production, remove a pass, or make a performance claim.

Each named pair first establishes the independent raw or semantic failure,
then the independent validation failure, then asserts the combined first
error through SourceBackedEditor::edit_sheets. Fixtures also exercise retry
and source retention where supported by the public API. Coverage and limits
are reviewed independently rather than claiming all possible invalid inputs.

The root runs every Rust check serially in the single owned target directory.
Each source attempt is frozen before execution, with full manifest, exact test
sources and patch. run.py binds every command to before/after source equality,
plan and script hashes, timestamps, stdout/stderr, exit status and artifacts.
A failed attempt remains retained under its original snapshot; a correction
requires a new snapshot. Focused tests precede the full all-features XLSX suite,
workspace check, targeted integration Clippy, rustdoc, formatting and boundary
checks. Test-only scope needs no native performance comparison.

After independent review and final checks, verify all receipts and source
bindings, remove the owned target and test temporary directory, then seal and
replay the bundle without mutation. Commit the test and evidence batch.
OLE2/OOXML remain the priority, ODF stays deferred, and iWork is excluded.
