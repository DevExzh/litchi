# Financial reducer LibreOffice compatibility capture

This directory retains a fresh headless LibreOffice conversion of all 41
vectors in `oracle-reducer-vectors.json`. The six functions are `CUMIPMT`,
`CUMPRINC`, `FVSCHEDULE`, `MIRR`, `NPV`, and `XNPV`. The receipt is
compatibility evidence for this fixed fixture and LibreOffice profile; it does
not establish production evaluator support or change the independent Decimal
oracle.

The capture is bound to contract
`fbe2283886326582b0e34fd277b38e6fb3e1bdeb2ecfbec5522da0f6f1467e4b` and reducer
corpus
`9701fb0e3f6917be16f538189245c1ae3fe53af956aa98241ce1aa4605ee7a26`.
`provenance.json` records the input FODS, recalculated ODS, stable
`content.xml`, converter profile, source mappings, and hashes. The scalar
11-function receipt in the sibling `native/` directory is unchanged.

The 41 rows classify as 15 numeric profile matches within the corpus's
relative tolerance, 1 exact native error-code match, 13 same-kind formula
errors whose LibreOffice raw error text is retained, and 12 explicit native
divergences. Divergences include host acceptance of invalid cumulative and
XNPV domains, different integer/control handling, binary64 cancellation in
NPV/XNPV, and LibreOffice's conversion of an XNPV text element.

Two compact corpus descriptors require valid native formula representations:

* `npv.scaled_underflow_rescue` expands the 2,000 repeated zeros and `1e300`
  tail into an explicit 2,001-row array literal. The expansion preserves
  source order and geometry and is recorded in `native-results.json`.
* `mirr.text_empty_logical_compact` uses `Data.A1:Data.E1`, a one-row,
  five-cell reference containing `-100`, text, a true blank cell, `TRUE()`,
  and `110`. LibreOffice has no valid inline `EMPTY` token, so the blank cell
  preserves the corpus element mapping without silently replacing it with an
  empty string.

All other formulas retain their corpus array/reference geometry. Native raw
error text such as `Err:502` and `Err:539` is preserved and never relabeled as
the normative error code.

Reproduce the capture with:

```text
python3 reproduce.py
```

The reproducer validates every retained hash, converts the FODS with
`/usr/bin/libreoffice` 26.2.5.2 using `--headless`, `C.UTF-8`, and a fresh
temporary `UserInstallation`, compares all 41 typed rows, and removes the
temporary profile and output directory on exit.
