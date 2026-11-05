# 0362ul1 — API inventory delegation and exact-row rerun

This is a tooling/provenance follow-up to 0362ul. It does not close any API
debt or turn the final inventory into an acceptance result. Item 22 remains
the only permitted deferral; the semantic review is still diagnostic.

## Inputs and source boundary

The 68 rustdoc inputs were captured at frozen source
`49c25ef67f62a8bb943e7c6a0bf464e49ae87d07`. All 68 receipt/artifact bindings
were independently validated. The external manifest is
`/home/zhuhe/.cache/litchi-api-contract/partial-diagnostic-49c25ef67f62a8bb943e7c6a0bf464e49ae87d07/receipt-manifest-full.tsv`,
32,741 bytes, SHA-256
`8cbba5abb762dfb86748cf28603b087c0aefa89fd44dc3cc3f62d225cd710147`.
The raw orchestrator run remains `incomplete`: its original tool stopped at
the foreign `litchi_core::detection` glob before producing inventory rows.
No raw JSON or receipt was changed.

The repaired delegated-glob inventory tool was `bb36ac17854b99076315b8a1a344f3d52e17e90c`;
the exact-row follow-up and tests are `06ea36e8c1fa19e86d06245d905cc1268bb98f02`.
Thus the input source (`49c25ef`) and inventory tool source (`bb36ac`, then
`06ea36`) are deliberately different and this is a tooling rerun, not a
claim that the stale capture is a final-source capture.

## Rerun and result

The external command was run sequentially over the 68 JSON files:

```text
/usr/bin/time -f 'elapsed=%E maxrss=%MKB' python3 -B tools/api_debt_inventory.py $(find /home/zhuhe/.cache/litchi-api-contract/49c25ef67f62a8bb943e7c6a0bf464e49ae87d07/inventory -maxdepth 1 -type f -name '*.json' -print | sort) > /home/zhuhe/.cache/litchi-api-contract/partial-diagnostic-49c25ef67f62a8bb943e7c6a0bf464e49ae87d07/inventory-full-dedup.tsv 2> /home/zhuhe/.cache/litchi-api-contract/partial-diagnostic-49c25ef67f62a8bb943e7c6a0bf464e49ae87d07/inventory-full-dedup.stderr.log
```

Before exact-row filtering, the repaired tool emitted 442 rows in
`inventory-full-repaired.tsv` (64,517 bytes, SHA-256
`7f3417da99158b2eb96ba43f76fbea8d8dc7670d76a330009059ab9603b46118`). Two
nominal/alias traversals each repeated an identical `litchi-iwa-protos`
`HeaderRowEdit` or `CommentStorageTableReferenceEdit` row. The follow-up
emits 440 data rows plus the header, all 440 tuples unique, in
`inventory-full-dedup.tsv` (64,233 bytes, SHA-256
`bd68ed36b3b57b89d5a0e3b4707f495212e23eae6c15cac98d2ea793285a2d9e`). The
stderr timing log is 33 bytes with SHA-256
`be5a21e903a5d00a0c4c2646cae358531daec43a740858a20b02309a06c12b87`.

The wrapper removes only exact duplicate tuples; distinct public alias paths
remain distinct. Focused fixtures also prove that a facade's local rows are
retained, a supplied delegated-crate document is itself scanned for debt, and
an undeclared foreign glob still fails closed. The 16-test command

```text
python3 -m unittest -q tools.test_api_debt_inventory tools.test_check_api_glossary.PolicyTests
```

passed. The existing semantic report remains the pre-dedup 442-row diagnostic
(`semantic-review-full.md`, 97,720 bytes, SHA-256
`6661513aa140eb4db50be2d456f4dbf03c66264fd0a083bab6c8754494ffb8e8`); it is
not a disposition or closure record.

After integration at `fae96321e3`, the coordinator ran the inventory,
disposition, capture, glossary and three registered W8 test modules together:
119 tests passed in 5.065 seconds. The command was
`python3 -B -m unittest tools.test_api_debt_inventory tools.test_validate_api_disposition tools.test_final_api_inventory tools.test_check_api_glossary tools.test_w8_projection_policy tools.test_w8_file_provenance tools.test_w8_original_baseline`.
The external 297-byte log is
`(deleted-coord-artifact)/composed-inventory-w8-fae96321e3.log`,
SHA-256 `24485e6f6c74b3528fe1e66c39af7348174d1c86eee620542ffd21ceb51a3e0c`.
This validates the composed tooling; the 440 diagnostic findings still need
repairs or exact semantic dispositions and final-source regeneration.
