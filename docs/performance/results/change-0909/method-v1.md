# 0909 raw ZIP census method

`zip-census.py` is an independent, read-only comparison for the owning-input
PPTX qualification. It has no Litchi import and does not build, run, or profile
the workload. Python's standard-library `zipfile` is used only to decode every
member and compute an independent SHA-256. The archive framing is parsed from
raw bytes in the helper itself:

- the EOCD (and ZIP64 EOCD when the ordinary fields are saturated), including
  the raw archive comment;
- every central-directory record, including its raw name, extra, comment,
  fixed metadata, and physical local-header offset;
- every local header, including fixed metadata, name and extra bytes;
- the compressed data range selected by the central record; and
- a signed or unsigned 32-bit or ZIP64 data descriptor when bit 3 is set.

Central metadata equality excludes exactly the central record's four-byte local
header offset. That offset is physical placement and moves when an earlier
member changes. All other central fixed fields, central name, extra, and
comment bytes remain part of the comparison. Local header offsets are retained
in the report for custody but are not an equality component. Compressed payload
equality is a direct byte comparison, with size and SHA-256 retained on both
sides. Decoded payload equality is computed separately through `zipfile`.

The report is create-only: an existing `--report` path is rejected. A changed
member is permitted only when its decoded archive name is listed with repeated
`--allow-changed MEMBER`. The generated-fixture convenience mode supplies the
eight expected slide XML names and checks the 445-member shape:

```text
python3 docs/performance/results/change-0909/zip-census.py \
  --source source.zip \
  --output output.zip \
  --report zip-census.json \
  --expected-generated-slides
```

The equivalent explicit allowlist is:

```text
--allow-changed ppt/slides/slide1.xml \
--allow-changed ppt/slides/slide29.xml \
--allow-changed ppt/slides/slide58.xml \
--allow-changed ppt/slides/slide86.xml \
--allow-changed ppt/slides/slide115.xml \
--allow-changed ppt/slides/slide143.xml \
--allow-changed ppt/slides/slide172.xml \
--allow-changed ppt/slides/slide200.xml
```

The selected positions in the generated 200-slide fixture are zero-based
`[0, 28, 57, 85, 114, 142, 171, 199]`, which map to the names above. The
allowlist is a qualification boundary, not an assertion that every selected
member must differ. The helper reports `changed_members.by_component` for all
differences and fails if any component of an unlisted member changes. It also
fails on changed member order, member count, or archive comment.

The JSON retains source/output archive size and digest, ordered names, comment
bytes and digest, per-member raw component fingerprints, decoded hashes,
relocation offsets, changed-member lists, and violations. A pass therefore
establishes exact resulting-byte equality for untouched members at the checked
components. It does not establish which writer branch copied those bytes, how
many copy calls occurred, or that equality came from a particular preservation
implementation.

The report distinguishes generated-fixture evidence from parser feature scope.
The generated corpus is expected to have 445 members and the selected slide
names above. ZIP64 and data-descriptor handling is implemented and the report
records whether either form was observed, but a normal generated fixture does
not constitute adversarial ZIP64/data-descriptor coverage. Such coverage needs
separate malformed and boundary fixtures; this helper does not silently claim
it from the owning-input run.

