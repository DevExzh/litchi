# Financial scalar LibreOffice compatibility capture

This directory retains a fresh headless LibreOffice conversion of the 65
vectors in the committed scalar corpus. The slice covers the eleven closed
form functions `EFFECT`, `FV`, `IPMT`, `ISPMT`, `NOMINAL`, `NPER`, `PDURATION`,
`PMT`, `PPMT`, `PV`, and `RRI`. It is compatibility evidence for this fixed
fixture and LibreOffice profile; it does not establish production evaluator
support or replace the normative contract.

The capture is bound to contract
`fbe2283886326582b0e34fd277b38e6fb3e1bdeb2ecfbec5522da0f6f1467e4b` and scalar
corpus
`15bb7e32bd47dc64de2656cb24a8995f59916a9d87d13fffaa836aa82c302531`.
`provenance.json` records the input FODS, recalculated ODS, stable
`content.xml`, typed rows, converter profile, and all SHA-256 identities.

The 65 rows classify as 40 numeric profile matches within the corpus's
relative tolerance, 3 exact native error-code matches, 11 typed error-kind
matches whose LibreOffice raw error text is retained, and 11 explicit native
divergences. The divergences are compatibility observations only. They cover
LibreOffice accepting fractional `Nper` instead of the contract's truncating
Integer profile for `PMT`/`PPMT`, accepting `PPMT` period/rate cases that the
contract refuses, ignoring non-0/1 `Type` or `PayType` controls for the
annuity functions, and losing the expected result on the tiny-rate `NPER`
case through host numerical cancellation.

Every formula row retains the independent expected value in
`native-results.json` and the native typed value or raw `Err:*` text. A
same-kind error is labeled `error-kind-match`; the native error code is never
rewritten as the contract's code. Divergent numeric or type results are
labeled `divergence` and do not update the expected corpus.

Reproduce the capture with:

```text
python3 reproduce.py
```

The reproducer verifies the retained hashes, converts the FODS with
`/usr/bin/libreoffice` 26.2.5.2 using `--headless`, `C.UTF-8`, and a fresh
temporary `UserInstallation`, compares all 65 typed rows, and removes the
temporary profile and output directory on exit. ZIP metadata from the
retained ODS is checked separately from the semantic `content.xml` rows.
