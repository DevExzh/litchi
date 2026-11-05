# 0362ul3 — result-contract identity and facade boundary

The final reviewed tool tip `435e1547bd` is merged at `f620a8dd26`.
All fourteen registered tooling modules pass together on composed source
`83a0ace78bfaf2cf6e14aa066b27e6657366a811`: 1,428 tests in 46.721 test-runner
seconds (48.210 wall seconds including process setup). The source/command/log
binding is in [the composed receipt](results/0362ul3/composed-tooling.json).
The 633,734-byte raw log remains external with SHA-256
`f12da526f0dea17c39c98e30ae086495a0726e546a1ba6f8a64f881e9d048a22`.
This is a composed tooling pass; the full runtime gate remains pending.

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
* Inventory checks foreign aliases in their defining crate when that crate's
  rustdoc is among the supplied inputs. A missing non-core defining graph
  emits a `commit_type` debt row instead of being silently skipped. The shared
  core Commit remains the explicit exemption because its triad is a shared
  contract.
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
# 67 tests, OK
```

The change is tooling and glossary policy only; it does not alter format
runtime paths or performance behavior.

## Returned-review repair

Review of `213fefbfd2df93946b46e6b3ebf2fa91662f0a0c` found two fail-open
edges, now covered in the follow-up source and fixture tests:

1. The inventory CLI passes the complete supplied crate set and an exact
   canonical-coverage manifest into the row walker. A foreign `Commit` whose
   defining crate, nominal target, or required triad is absent produces a debt
   row; a crate name alone is not enough to delegate the audit.
2. Namespace identity now follows local type-alias chains transitively, with
   cycle/malformed-target guards. Aliases through separate chains to one
   nominal target survive either traversal order, while unrelated targets
   invalidate the path rather than being silently merged. The derive-macro
   namespace remains independent of the nominal type namespace.

The exact fixtures are `fixture::external_metadata::Commit` for the missing
`litchi_odf_common` defining graph, `fixture::alias_chain::Error` for
transitive aliases and unrelated targets, and
`fixture::namespace_collision::Error` for both derive-macro visit orders.

The subsequent review of `837815f1c0` found one coverage ambiguity, also now
covered by fixtures: duplicate supplied crate roots are rejected before the
coverage union is formed, and a covered path must have a local definition
crate ID and local nominal identity. An indexed foreign item or a re-exported
foreign target cannot masquerade as the defining crate.

The returned-review receipt is 67 focused Python tests passing. No Cargo build
was run. The Python inventory was rerun against the frozen `49c25ef...`
captures after the repair and produced the same 19,440-byte TSV and SHA-256
listed above; those category counts remain diagnostic evidence and are still
not final-source closure.
