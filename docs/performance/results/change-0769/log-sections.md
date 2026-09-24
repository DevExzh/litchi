# 0769 log sections (ready to paste)

## For HOTSPOTS.md

## 0769 — CFB open and reads agree on the partial last mini sector

[0769](0769-cfb-mini-sector-open-read-agreement.md) closes the gap record
0767's fault lane found. A5 admitted a mini stream in the partial last mini
sector of a root whose size is not a multiple of 64, and the whole-stream
readers then refused it, which caused all 52 of that lane's read errors after
open. Every layer now bounds the bytes a stream takes. The cost is none on
this path: A5's ownership loop is the base's, with a once-per-stream check
only for an unaligned root (+3 instructions per stream on opens, +0.04%),
and mini-stream reads run 1.5% fewer instructions. The record's first
version put the check inside the loop and ran +0.48% instructions per open
of 10,000 mini streams. It also measured `doc_semantic_open/large`
+14.1% (+2.4% in the retained build) with fewer instructions, which is code
placement. The non-iWork goal remains active.

## For REPORT.md

## 0769 — CFB open and every reader agree on mini streams at the root's end

[0769](0769-cfb-mini-sector-open-read-agreement.md) makes `litchi-cfb`
admit a mini stream exactly when every byte it takes lies inside the root
entry's stream size (MS-CFB 2.6.1, 2.7). A stream the root size covers now
reads through every reader (a real producer, Zeiss AxioVision in POI's
`BlockSize512.zvi`, writes that shape). A stream that needs a byte past the
root size is refused at open with the existing `CorruptedFile` "outside the
root mini stream" error, where it was admitted and refused only on read. On
two builds:

- record 0767's 30,400-fault lane goes from 52 cases admitted and then
  refused on read to 0, with no other change;
- a sweep of every root size within 128 bytes of 105 files' own goes from
  3,906 such copies (1,353 on which readers disagreed) to 0, matching an
  independent oracle on all 21,741;
- a census of 1,830 compound files changes no real file's verdict.

It also folds in record 0767's review follow-ups: the MiniFAT read's
move-out, all-clear assertions, and freeing chain buffers above 1 MiB. Costs
over 16 ABBA rounds on core 30: every median within −4.4% to +3.5%, opens
+0.03% to +0.09% instructions, identical allocations; the moves are code
placement. Gates pass for `litchi-cfb`, its OLE2 dependents and the facade.
`performance_claim: none`. [Evidence](results/change-0769/README.md).

## For GOAL_AUDIT.md

## 0769 — CFB validation and reads made to agree; one refusal moved to open

[0769](0769-cfb-mini-sector-open-read-agreement.md) keeps every CFB
validation, ownership, cycle, overlap, FAT, MiniFAT, directory and truncation
check, and strengthens one. A mini stream whose bytes run past the root mini
stream's declared size is now refused at open (ADR 0006: fatal safety
failures stop opening), not first admitted and then refused per read. A
stream whose bytes all lie inside is read by every reader. Before, the cursor
reader and the shared cache refused it while the range and direct readers
returned it (MS-CFB 2.6.1, 2.7; readers preserve real-world quirks, ADR
0006). Tests hold open and eight readers to an independent byte-bound oracle
on crafted, shuffled, fault-injected and fixture-derived files and on a
vendored real-producer file. A cross-build census, fault lane and root-size
sweep show no file admitted and then refused. Remaining debt: the Zeiss
producer's `FREESECT` storage start fields, and a version-3 header with
4096-byte sectors, still refuse those real files at open (separate
directory- and header-validation questions). No coverage, claim or
timing-contract promotion; the non-iWork goal remains open.
