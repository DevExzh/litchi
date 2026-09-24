# Upstream cached observations

The 56 selected numeric caches cover all eleven elementary functions in this
batch. They come from LibreOffice's mathematical FODS fixtures at commit
`d804d6aff49054bad1719ec3c2d136b545bbc7e7`. All eleven local input files were
compared byte-for-byte with their files at that upstream commit. The
[provenance receipt](provenance.json) records SHA-256 identities and source
coordinates are retained in [cached-results.json](cached-results.json).

One cache is explicitly excluded: `MOD(26^15;77)` stores zero. The represented
binary64 numerator has remainder 9, while the exact integer expression would
have remainder 34. The finite evaluator uses the remainder of the represented
operands without quotient/product cancellation. The independent exact-rational
[oracle](../numeric_oracle.py) covers that numerical distinction.

Run `python3 extract.py /path/to/libreoffice-core /temporary/output` to reproduce
the observations without changing committed evidence. The extractor verifies
the pinned input hashes. The public native-cache test allows relative error
`1e-13` for serialized decimal caches and requires exact cached zeros.

These are cached observations, not native application execution, resave,
acceptance or complete-fixture evaluation. Only closed expressions using the
selected functions and supported literals/operators are extracted. Upstream
fixture data is distributed under the retained MPL 2.0 license.
