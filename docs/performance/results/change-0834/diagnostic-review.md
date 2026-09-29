# 0834 — PPTX diagnostic replay review

This is an independent review of the retained diagnostic, not a qualification
result. The inputs are [`diagnostic-analysis.json`](diagnostic-analysis.json),
[`diagnostic-decoded.json`](diagnostic-decoded.json), the warm report, and the
cold-verified receipt and log.

## What the diagnostic establishes

The analyzer's `status: pass` means that its reconstruction checks passed. It
does not mean that the source classifier passed: the decoded diagnostic still
has `classification: classification-failed`, and the cold-verified command
exited with status 1.

The retained replay is internally consistent:

| Phase | Reads | Returned bytes | Payload overlap |
| --- | ---: | ---: | --- |
| `open` (records 0–880) | 881 | 169,092 | 16,230 bytes of unselected slides |
| `query` (records 881–882) | 2 | 615 | 522 bytes of the selected slide |
| all | 883 | 169,707 | 16,752 slide bytes, 0 media bytes |

The one exact tail record is at raw index 1:
`[16,953,344, 17,018,880)`, requested and returned length 65,536. It is
before the validated `open_read_count: 881`. The fixed 22-byte EOCD attempt is
index 0; the following 33,194-byte directory read is index 2. The analyzer
checks that the tail occurs exactly once, that its per-payload overlaps equal
the complete open-phase totals, and that the selected query range is fully
covered with no query unselected-slide or media overlap.

The raw vector also contains zero-length records. They contribute calls but no
returned bytes or payload overlap, and the analyzer retains them without
treating them as failures. This matches the bounded classifier design.

## Interpretation of the failure

The child mode is `verified-prime`, so this is the preliminary aligned-source
primer that a `cold-verified` request runs before its measured child. The
failure is therefore the old aggregate classifier rejecting the open tail's
physical overlap with unselected slide payloads. The `cold-verified` child was
not reached: there is no cold proof, fincore result, positive `read_bytes`
gate, or timed cold sample in these artifacts. The warm invocation passed and
is useful as a control, but it does not change that boundary.

The warm report identifies the unaligned source as 17,017,139 bytes with its
own hash and selected-slide-only classification. The failed primer identifies
the aligned source as 17,018,880 bytes with a different hash. That size change
is compatible with the prepared page-aligned copy, but these diagnostic files
alone do not prove the EOCD offset, comment-length transform, byte equality,
or zero suffix. Those geometry checks remain admission requirements.

## Design alignment

The evidence supports the bounded repair in [`review-design.md`](review-design.md):

* The immutable source envelope (`source_sha256` and `source_bytes` here),
  chronological raw records, and validated `open_read_count` provide the
  phase boundary without repeating identity or phase fields on every record.
* The exact tail allowance belongs only to `open`; the selected query retains
  its ordinary complete-selected-range and zero-unselected/media rules.
* The 16,230 bytes of open unselected-slide overlap must remain visible and be
  accepted only as the bytewise overlap predicted for this one exact tail. It
  must not be subtracted or relabelled as semantic I/O.
* `verified-prime` may use the aligned allowance only after the independent
  prepared-source geometry and identity checks bind this source. No allowance
  follows from the mode string or aligned length alone.

The next admissible result is a fresh run in which the primer classification
passes, followed by the actual `cold-verified` child and its independent proof
gates. Until then, this packet records an untimed replay explanation only; it
does not authorize a cold performance claim.
