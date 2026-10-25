# 0362ul145a — final composed source census matrix

Source S is `035615e0de91575af2fe8872f502a7cea32e0565`, containing reviewed
runtime/glossary repairs, native v2 consumers through 135f/136e/137e, and
coordinator records through 145. The progress merge resolves its single
conflict using the coordinator's newer combined 131/132 account and all later
entries. No source change is introduced by that documentation merge.

This matrix-only source-binding commit C regenerates the census matrix with:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -B tools/prepare_raw_api_census.py --root . \
  --output tools/raw_api_census_matrix.json \
  --source-commit 035615e0de91575af2fe8872f502a7cea32e0565
```

The 84 profiles remain identical: 69 glossary profiles and 15 package-only
profiles, covering 68 direct packages. Only the S source commit changes;
entries, lanes and captures remain empty. The deterministic matrix check
passes. All three role locks are unchanged. The historical coverage manifest
and C780 diagnostics retain their original identities.

Final compiler/runtime/glossary/corpus and raw-census packets bind C and its
exact tree, parents and role locks. Future native protocol P commits bind
this immutable C without changing it. Full53 plus seven compiler and five
runtime supplements, raw capture/callable review, guarded native builds,
fidelity, quiet measurements and trunk integration remain open. Nothing is
pushed, and matrix preparation is not a qualification result.
