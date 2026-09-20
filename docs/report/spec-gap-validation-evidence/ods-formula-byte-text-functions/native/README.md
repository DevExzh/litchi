# Native LibreOffice byte-function receipt

This directory retains one small, local LibreOffice conversion fixture for
the seven ODF 1.4 §6.7 functions: `FINDB`, `LEFTB`, `LENB`, `MIDB`,
`REPLACEB`, `RIGHTB`, and `SEARCHB`. The source is
[`byte-functions-native.fods`](byte-functions-native.fods); its formulas are
recalculated by `/usr/bin/libreoffice` 26.2.5.2 with a fresh headless profile,
and the resulting document is retained as
[`recalculated.ods`](recalculated.ods).

The seven ASCII rows agree with the selected evaluator profile in
[`../contract.md`](../contract.md). The remaining rows deliberately use
`é`, `界`, `🙂`, and `ß` to expose host behavior. Under `C.UTF-8`, LibreOffice
reports native width units for `LENB` and maps byte positions using those
units. Consequently ten of the twelve non-ASCII rows differ from the
profile's UTF-8-octet results. The two matching rows are retained as
observations; they do not establish general compatibility. The receipt does
not identify the native units as a particular code page and does not claim
that LibreOffice byte-function results are portable.

[`native-results.json`](native-results.json) records the typed native value,
the expected UTF-8-profile value, and the comparison for every formula.
[`provenance.json`](provenance.json) records the input/output hashes,
converter version, locale, stable `content.xml` hash, and the explicit
coverage limits. The raw ODS ZIP hash is a capture hash; ZIP metadata can
vary between conversions, so reproduction checks the stable content hash and
typed results.

Run the bounded local check with:

```text
python3 reproduce.py
```

The script performs no network access. It checks the source hash, converts
the fixture in a temporary LibreOffice profile, checks the recalculated
formula sequence and typed results, verifies the stable content hash, and
removes its temporary profile and output tree on exit. Set `LIBREOFFICE` to a
different executable only when deliberately reproducing with another local
LibreOffice build; that is a new native observation and may require updated
receipt hashes.
