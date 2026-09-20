# Financial root solver LibreOffice compatibility capture

This directory retains a fresh headless LibreOffice conversion of all 31
vectors in `oracle-root-vectors.json`: 10 `IRR`, 10 `RATE`, and 11 `XIRR`
cases. It is compatibility evidence for this fixed fixture and LibreOffice
profile; it does not establish production evaluator support or replace the
independent high-precision root corpus.

The capture is bound to financial contract
`fbe2283886326582b0e34fd277b38e6fb3e1bdeb2ecfbec5522da0f6f1467e4b`, date/time
contract `cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f`,
and root corpus
`7ccf4912d470a8d895d92343952d34c199f746d25ab1d88e5df2d9a019501c2e`.
`provenance.json` records the formula input, recalculated ODS, stable
`content.xml`, converter profile, and hashes.

The strict rows classify as 15 numeric profile matches, 7 same-kind formula
errors whose LibreOffice raw error text is retained, and 8 explicit native
divergences. The row `irr.multiple_root_mid_guess_bounded` is classified only
as `host-observation`: the corpus explicitly permits bounded-search
selection/success uncertainty for that multiple-root case, so its native
`0.20000000000005` result is retained without calling it either a strict pass
or a divergence.

The explicit divergences include LibreOffice's treatment of tangent/no-bracket
IRR, `IRR` guess `-1`, RATE negative-integral and `Rate=-1` boundary cases,
non-integral RATE boundary handling, and XIRR sign/date-domain cases. These
are compatibility observations only; no normative expected result was changed.

Reproduce the capture with:

```text
python3 reproduce.py
```

The reproducer verifies all retained hashes, converts the FODS with
`/usr/bin/libreoffice` 26.2.5.2 using `--headless`, `C.UTF-8`, and a fresh
temporary `UserInstallation`, compares all 31 typed rows, and removes the
temporary profile and output directory on exit.
