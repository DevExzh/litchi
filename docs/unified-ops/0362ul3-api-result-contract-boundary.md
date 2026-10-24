# 0362ul3 — result-contract identity and facade boundary

The API inventory now distinguishes the two owner-accessor contracts without
trusting a type's spelling.

* `ReturnTypes.commit_package` accepts the ODF metadata alias only after it
  proves the external nominal identity
  `litchi_odf_common::metadata::Commit` and exactly one generic owner
  argument. An unrelated external `Commit`, or an ambiguous duplicate
  canonical identity, fails closed.
* `file()`/`into_file()` are recognized as owner accessors only for the
  nominated `litchi_ops` facade Commit paths. A format-owned Commit with those
  names still has to provide `package()`/`into_package()`.
* Inventory checks foreign aliases in their defining crate. The shared core
  Commit remains the explicit exception because its triad is a shared contract.
* The glossary registers every identity-proven replay row from the frozen
  all-features capture: 262 ordinary rows, the nine ODF metadata rows enabled
  by the resolver, and the two unified-File rows. Minimal profiles omit only
  feature-only VBA rows; full profiles register them. `ops-all` nominates
  `litchi_ops::File` and its two `Applied` Commit results.

The direct 68-capture inventory command completed with zero `replay_return`
rows. Its remaining five `commit_type` rows are the known DrawingML (four)
and Markdown (one) contract debt; no row was waived. The exact diagnostic
category counts were:

| category | rows |
| --- | ---: |
| `commit_return` | 3 |
| `commit_snapshot` | 7 |
| `commit_type` | 5 |
| `edit_type` | 36 |
| `exhaustive_error` | 19 |
| `is_changed` | 3 |
| `mut_draft` | 44 |
| `mut_owner` | 3 |
| `replay_return` | 0 |
| `transaction_type` | 15 |
| `with_limits` | 21 |

This diagnostic used capture/source `49c25ef67f62a8bb943e7c6a0bf464e49ae87d07`
under `/home/zhuhe/.cache/litchi-api-contract/`. The implementation receipt
is commit `213fefbfd2df93946b46e6b3ebf2fa91662f0a0c`. The retained external
TSV is `/tmp/uo2-result-contract-inventory-final.tsv` (19,440 bytes,
SHA-256 `a0166f99f9c1c30d30967bd5ad14d7d6ef56dbad5363788821d62d0ab90db840`).
This is frozen diagnostic evidence, not final-source closure; the inventory
must be rerun against final-source rustdoc captures after the remaining
production work is integrated.

The command was:

```text
python3 tools/api_debt_inventory.py \
  /home/zhuhe/.cache/litchi-api-contract/49c25ef67f62a8bb943e7c6a0bf464e49ae87d07/inventory/*.json
```

Receipts:

```text
python3 -m unittest tools.test_check_api_glossary tools.test_api_debt_inventory
# 64 tests, OK
```

The change is tooling and glossary policy only; it does not alter format
runtime paths or performance behavior.
