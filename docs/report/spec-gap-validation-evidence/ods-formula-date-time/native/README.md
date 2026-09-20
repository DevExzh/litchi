# Date/time LibreOffice compatibility capture

This directory retains a fresh headless LibreOffice conversion of all 132
contract-bound date/time oracle formulas. It is compatibility evidence for a
fixed fixture and profile; it does not establish production evaluator support
or normative agreement.

The final capture is bound to contract
`cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f` and
oracle corpus
`273421f087314733c9b383e6a1b5c1bb1e43de3b65b0be24aeb1945864e6438d`.
The input FODS, recalculated ODS, stable `content.xml`, typed observations,
and their hashes are recorded in `provenance.json`; `capture.log` retains the
converter, locale, command shape, and comparison counts.

The capture used `/usr/bin/libreoffice` 26.2.5.2 (Build 2), `--headless`, a
fresh temporary `UserInstallation`, and `C.UTF-8` for `LANG`, `LC_ALL`,
`LC_CTYPE`, and `LC_NUMERIC`. Temporary profile and output directories were
removed after capture. Run the retained check with:

```text
python3 reproduce.py
```

The reproducer validates the retained input/output/content hashes and then
converts the FODS again with a fresh profile. It compares deterministic typed
rows and checks the six volatile rows as host observations. `NOW`, `TODAY`,
and no-Year `EASTERSUNDAY` are not compared to the contract's injected
timestamp because LibreOffice reads its own calculation clock.

The 132 observations classify as 94 profile parities, 32 explicit native
divergences, and 6 host-clock observations. The largest documented
divergence groups are:

* LibreOffice uses its own date domain for year 1 and normalizes DATE month
  zero/overflow instead of returning the profile's #NUM! errors.
* Its DATEDIF month case, MINUTE boundary formula, negative/multi-day TIME,
  zero-fraction WORKDAY, non-integer WEEKNUM, and custom/all-off WORKDAY
  cases differ from the selected profile.
  The retained exact binary64 tie `MINUTE(239.5/86400)` evaluates to `3` in
  LibreOffice, while the selected rounded-total-seconds profile expects `4`.
  The exact negative tie `SECOND(-0.5/86400)` evaluates to `0` in
  LibreOffice, while half-away rounding followed by day normalization expects `59`.
* DATEVALUE/TIMEVALUE numeric and simple-fraction fallback rows expose the
  host parser profile; `DATEVALUE("1/4")` is parsed as a calendar date and
  the numeric fallback rows are rejected.
* ReferenceList and formula-error sequence cases expose LibreOffice's own
  argument and error behavior. `Err:*` values are retained as native raw
  observations and are not relabeled as the contract's formula error codes.

The prior capture bound to contract
`3d70836316ccbd671a37714b6f3f773899e3dec9d1cef81d7ba50e8f3ed75413` is kept
under `history-contract-3d708363/` as a historical diagnostic. Its hashes and
provenance are unchanged.
