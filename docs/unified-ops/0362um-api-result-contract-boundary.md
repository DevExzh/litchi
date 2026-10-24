# 0362um — result-contract identity and facade boundary

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
and Markdown (one) contract debt; no row was waived. The command was:

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
