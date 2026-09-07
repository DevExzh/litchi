# 0464 native-pair static audit

This directory preserves the bounded static audit that was run before the
0464 probe receipts were available. `graph-audit.py` and `graph-audit.json`
are byte copies of `/tmp/litchi-0464-graph-audit.py` and
`/tmp/litchi-0464-graph-audit.json`; the archive copies have the same hashes:

```
graph-audit.py   4b516c5cc35e8c042785775f3c28c4bd8ca3df125055052789bd9f9abe756c64
graph-audit.json 4c02493c3d96b54af1d2fc7a75842b6e6bffd55d7254d7d82b6fbc3c4139c6df
```

The original command, from the repository root, was:

```
python3 -B /tmp/litchi-0464-graph-audit.py > /tmp/litchi-0464-graph-audit.json
```

Its relative scan roots were `3rdparty/libreoffice-core`, `3rdparty/poi`,
`3rdparty/Open-XML-SDK`, `test-data`, and `docs/performance/results`. The
recorded inventory contains 928 files, 912 parsed archives, 16 malformed
archives, 1,727 graph slides, 40 distinct-hash graph groups, and one group
with a member having zero heuristic blockers.

Those counts are static leads, not an admissibility result. The earlier
wording that the audit found “no admissible independent pair” is withdrawn.
The supported conclusion is narrower: under this heuristic, the only group
with a zero-blocker member was the POI unsigned/signed XMLDSig group. The
signed member carries `_xmlsignatures/*`, so the current planner's signature
policy still makes that particular pair unsuitable; the audit did not prove
that no other distinct pair can be admitted.

Known limitations in the preserved script:

* Its chart predicate labels some chart relationships/workbook closure as
  `chart leaf/markup closure`. Current Rust planner semantics permit some
  chart/workbook closure that this static predicate does not model. Such a
  blocker requires an actual `plan_cross_slide_copy` probe and is not a Rust
  refusal.
* Its macro test is `('vbaProject' in n.lower())`. The case mismatch can miss
  ordinary `vbaproject` member names, so a missing macro blocker in the JSON
  is not evidence that a package is macro-free. The planner's content-type
  and relationship checks remain the authoritative test.
* XML element, namespace, relationship-closure, and package-feature checks
  are approximations. Exact layout/master/theme graph equality is necessary
  for the current planner but does not establish source-slide admission,
  output correctness, or independent producer provenance. The script never
  invokes Rust or a native office application.

The 0464 receipts supersede the earlier “six 0456 probes only” status. They
remain outside this archive and were not changed:

* `checks/shapes-derived-probe.json` (SHA-256
  `dfdab255103b6bd45b3488ce3353b4d4950f1f10d491e3d91972b406cfc80662`) records
  the actual `UnknownSemanticSurface` refusal for
  `test-data/ooxml/pptx/shapes.pptx` to
  `test-data/office-interop/litchi-changed/shapes-litchi.pptx`.
* `derived-pair-probe.json` (SHA-256
  `cbe2f13a53c3eeb311204cb0c18151d3bd1991191aa6e84044747bcd48d1c35a`) records
  a positive original-to-Litchi-derived publication (`PUBLISHED bytes=55891
  slides=3`). This is a distinct derived control, not an independently
  authored native pair.
* `native-pilot/inventory.json` (SHA-256
  `51dd93572de91debf7f962c0143e88ea4f93c585e315e2454bfeef7193dd2ae7`)
  records the subsequent LibreOffice 26.2.5.2 resave of that derived output;
  it does not establish independent-document provenance.

No production source was edited for this archive, and the static scan was not
rerun during archival.
