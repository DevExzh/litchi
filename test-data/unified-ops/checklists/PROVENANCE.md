# Unified API checklist reports

These three tab-separated reports are written by this repository's own tests.
They copy nothing upstream and carry this repository's license (see the
workspace `Cargo.toml`).

| File | Written by |
|---|---|
| `crud-checklist.tsv` | `crates/litchi-ops/tests/crud_matrix.rs`: one row per claimed (format, element, verb) cell and what the Rust unified API did with it |
| `crud-adr0008.tsv` | the same test: the per-format atomic-failure, patch-inverse and composition rows |
| `fidelity.tsv` | `crates/litchi-ops/tests/corpus_fidelity.rs`: the empty-edit outcome of every claimed-format file under `test-data/` |

Readers:

- `crates/pyo3-litchi/tests/test_crud_mirror.py` and
  `crates/pyo3-litchi/tests/test_fidelity_mirror.py` replay the rows through
  the Python bindings.
- The `path_parser_differential` test of `crates/litchi-core` takes its path
  strings from these files as well as from the workspace sources.

To refresh them, run both tests with `LITCHI_REPORT_DIR` set to a scratch
directory outside the checkout, then copy the three `.tsv` files here:

```sh
LITCHI_REPORT_DIR=<scratch> cargo test -p litchi-ops --all-features --test crud_matrix --test corpus_fidelity
```

The tests also write Markdown renderings of the same rows; those are not kept.
