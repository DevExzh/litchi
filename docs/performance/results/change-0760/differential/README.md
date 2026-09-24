# Base-versus-branch differential, change 0760

Probe: probe/src/main.rs, mode 'diff'. The recorded run uses the probe
binaries built with one command ('cargo build --release --offline', with
tools/perf-baseline/Cargo.lock copied beside the manifest) against the two
equal-length detached worktrees, 0760-before-src at the base 1d1044e3ac (leg pa)
and 0760-branch-src at the production commit dfde1e43bb (leg pb); their SHA-256
digests are in ../binaries.sha256. Both read the same fixture tree
(test-data/, 78 .pptx files).

Command (each leg): pptx-memo-probe diff <branch-src>/test-data 850 760 > final-<leg>.txt

Per fixture: the unmodified archive through Package::from_bytes; 18 structured
mutations of the first, middle and last slide (applied at the OPC level and
opened with Package::from_opc_package); and 850 seeded random byte-level
mutations of a random slide. Each package runs: open, capture, a no-op
commit and publication, one-shape edits of the first/middle/last slide (each
committed, chained into a second commit from the committed snapshot,
published on a fresh facade, then edited, committed and published again from
the published snapshot), an all-slide edit, a notes edit, a shape removal, a
slide move, and a capture with MCE retention disabled plus an edit. Every
step prints one outcome line: a refusal's Debug text, or the snapshot
revision, slide identities digest, patch digest and published-bytes digest.

Result: 67,790 packages, 340,198 outcome lines; the two outputs are
byte-identical:

```
8a4b8454578ce51a19612e7e92caf80699463e2dc886a4d059cb62d400c6ef4e  final-a.txt (base, pa)
8a4b8454578ce51a19612e7e92caf80699463e2dc886a4d059cb62d400c6ef4e  final-b.txt (branch, pb)
```

Two earlier runs, whose branch leg was built from the worktree just before the
production commit (its later edits reformatted test code only), gave the same
result: one with the structured mutations only (1,490 packages, 14,475 lines)
and one with the full set, whose output digest equals the one above. The 52 MB
outputs were not retained.

```
outcome lines per step kind (identical in both builds):
  *.capture            total=  13928 ok=  13928 refused=      0
  *.chained            total=  12947 ok=  12597 refused=    350
  *.commit             total=  38938 ok=  37732 refused=   1206
  *.publish            total=  37732 ok=  37732 refused=      0
  *.stage              total=  73217 ok=      0 refused=  73217
  capture              total=  67790 ok=  13928 refused=  53862
  noop.commit          total=  13928 ok=  13928 refused=      0
  noop.publish         total=  13928 ok=  13928 refused=      0
  open                 total=  67790 ok=  67790 refused=      0
most frequent capture refusals:
   41365  ERR Invalid("invalid sld root or namespace")
    1991  ERR Xml("syntax error: attribute value not closed: `\"` not found befo
    1491  ERR MarkupCompatibility(Xml("ill-formed document: entity or character 
    1399  ERR Invalid("PresentationML part has an invalid root namespace")
    1033  ERR MarkupCompatibility(Xml("syntax error: attribute value not closed:
     737  ERR MarkupCompatibility(NonConformant("invalid QName: invalid XML QNam
     734  ERR MarkupCompatibility(Xml("syntax error: tag not closed: `>` not fou
     458  ERR Xml("ill-formed document: entity or character reference not closed
     379  ERR Xml("syntax error: XML declaration not closed: `?>` not found befo
     378  ERR Invalid("PresentationML part lacks an element root")
     305  ERR Xml("syntax error: tag not closed: `>` not found before end of inp
     300  ERR MarkupCompatibility(Xml("invalid utf-8 sequence of 1 bytes from in
```
