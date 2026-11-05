# 0362ul131a — composed source census matrix

Source S is `faab79bc388e09101f8f4cf53ff722e0c4de5046`, containing the eight
reviewed repairs and composition record131. This commit C refreshes
`tools/raw_api_census_matrix.json` against S with:

```sh
python3 -B tools/prepare_raw_api_census.py --root . \
  --output tools/raw_api_census_matrix.json \
  --source-commit faab79bc388e09101f8f4cf53ff722e0c4de5046
```

The profile set is unchanged: 84 profiles, comprising 69 glossary profiles
and fifteen package-only profiles, covering 68 direct packages. Only the
source pin, current glossary identity and common crate manifest identities
change. The historical raw coverage manifest is untouched. The matrix stays
explicitly incomplete, with empty callable entries, lanes and captures.

After committing C, `python3 -B tools/prepare_raw_api_census.py --check`
checks the scoped glossary/workspace/manifest identities against S rather
than requiring HEAD to equal S. Actual compiler, runtime, glossary and raw
capture packets must bind C and its unchanged three locks. This avoids a
self-referential source hash and makes no capture or final coverage claim.

Raw capture and reviewed finalization are separate: collect immutable
84-profile/68-package evidence first, review callable dispositions, then
strictly validate a distinct complete manifest against those captures.
No existing failure or historical capture is relabeled. All final gates and
performance/fidelity obligations remain open; nothing is pushed.
