# 0710 DOCX custom-properties preservation evidence

The patch gates DOCX custom-properties serialization on existing edit intent.
This correctness batch carries `performance_claim: none` and excludes iWork.

Five added tests cover clean and refused-edit saves through both public routes,
dirty-property retry after a partial sink failure, raw OPC replacement, and
empty-part preservation until explicit clear. The baseline has four expected
preservation failures; the candidate passes all 1,453 DOCX tests with all
features and targets, formatting, and warning-denied Clippy.

The strict public-API oracle is reused byte-for-byte from change 0709. Its
original failed report remains in that packet; this packet passes the same
refusal preservation assertions and both no-edit controls. All four outputs
are identical to the source archive. An independent Python ZIP reader confirms
every decoded member is unchanged. The admitted paragraph-edit outputs match
0709 byte-for-byte; their relationship rewrite remains outside the scoped
paragraph-preservation proof. The probe name/schema retain 0709 to make reuse
explicit.

| Evidence | Artifacts |
| --- | --- |
| Change and constraints | `change.patch`, `constraints.json`, `revision.json`, `fixture.json` |
| Failing baseline control | `baseline/checks.json`, `baseline/source.json`, `baseline/custom-properties.log`, `baseline-result.json` |
| Candidate quality | `candidate/checks.json`, `candidate/source.json`, `candidate/*.log`, `test-summary.json` |
| Strict oracle | `oracle-reuse.json`, `oracle.py`, `oracle/report.json`, `oracle/result.json`, `oracle/artifacts/` |
| Independent ZIP comparison | `reproduce.py`, `zip-verification.json` |
| Review and evidence gates | `review.json`, `evidence/results.json`, `evidence/*.log`, `final-report-gate.json` |
| Cleanup and final integrity | `cleanup.json`, `audit.py`, `artifact-manifest.json`, `artifact-seal.py` |

`check.py` serializes baseline and candidate Cargo checks with source custody.
`oracle.py` binds the unchanged probe, production source, output archives, and
binary. `evidence.py` runs all six repository evidence gates. The preparatory
`dependency-build` only compiled baseline library dependencies while tests were
authored; it is not a test acceptance result. Owned build, binary, and filesystem
scratch roots are removed, with the oracle binary identity retained.

From the repository root:

```sh
python3 -B docs/performance/results/change-0710/audit.py
python3 -B docs/performance/results/change-0710/artifact-seal.py --check
```

No performance or native Office compatibility claim is made. Calling
`custom_props_mut()` still conservatively records edit intent, even if the
caller makes no value change. The broader non-iWork goal remains active.
