# Assembly review for change 0463

This review covers the archived disassembly artifacts produced for the frozen
candidate source epoch
`d18406665f4fd9ad76cb0a2530bfe8e7459bf4ada874ea353b435a84ae4f921c`
(7030 files) and the retained baseline epoch
`fd86ab1c33f322d6550688833a6549968a1ccf96c537424d34b420c0bedefd3c`.
Both assembly receipts report `pass` and `source_unchanged`. The review is a
manual control-flow inspection of the `.asm.txt` files; it does not rerun
objdump, builds, or captures.

The retained artifact hashes are:

| target | baseline artifact | candidate artifact |
| --- | --- | --- |
| `transaction-commit` | `fada7f2e3083b7068c1568d4af3592a1f0819eac2d11989a774305cbdaf508f0` | `61e8eb1366e5d1fd0af7712fee883c392a585d719299537d32dd03f70aab3ee6` |
| `to-bytes-bounded` | `07c41759dd4d37df1280efcabc41eb65d44b9f4984c7e92461b0a3a6eb83c38c` | `0713541b78db0a33d4f1d6abfd620403c4e3cc5ef86cf28fbe694d75062ecbc2` |
| `validate-compact-xml-parts-against` | `a671e296ec67d60f2e0c40035bf2b06574317ef7f7e27683f540e8feddeea396` | `bf81f955bd8790cb47e75de5b8b950eb884e5c883f4082b539e3afbb3e484c9b` |

## Commit path and proof guard

The baseline commit body calls `MutablePresentation::to_bytes_bounded` with
the 128 MiB package bound at `baseline-transaction-commit.asm.txt` lines
154–164. The candidate instead calls
`MutablePresentation::to_owned_package_bounded` with the same bound at
`candidate-transaction-commit.asm.txt` lines 153–164. The candidate result is
handled as the reopened `OwnedPackage` result before the existing commit error
and readback flow continues.

The candidate retains the source-reference precheck. At lines 1696–1708, the
existing content-changed branch skips that check when its flag is set; on the
other branch it calls `validate_compact_xml_parts_against` with `%ecx = 0` and
propagates its error result. This is the same precheck shape as the baseline at
lines 1557–1569.

After that precheck, the candidate constructs the final limits profile at
lines 1709–1723. The immediate constants are 128 MiB aggregate bytes, depth
512, one million events, 250,000 attributes, 16 MiB token bytes, and 128 MiB
text bytes. An invalid limits result branches to the fallback block at
`18c259e`. A valid result then checks the proof option at lines 1724–1731;
an absent proof takes the same fallback.

For a present proof, lines 1732–1742 pass the proof, reopened candidate, source,
final limits, 128 MiB package bound, and 65,536 part bound to
`PublicationAuditProof::covers`. A true return branches to `18c25d0`, which is
the media-check path. Thus the proof hit skips only the final compact-XML
validator.

The proof miss path at lines 1743–1751 calls
`validate_compact_xml_parts_against` with `%ecx = 1`, checks its error result,
and then reaches the same media checks. The candidate therefore retains both
validator calls in its body: the raw-reference precheck with `%ecx = 0` and the
conditional full fallback with `%ecx = 1`. The baseline has the corresponding
precheck followed by an unconditional `%ecx = 1` call at lines 1557–1579.
Both candidate proof-hit and proof-miss paths continue through
`verify_embedded_media` and `verify_removed_media` at lines 1752–1767, with
the same error propagation shape as the baseline. The assembly supports the
intended guard placement and does not show a path that bypasses media checks
or the candidate package handling.

## Retained validator body

The candidate and baseline `validate_compact_xml_parts_against` artifacts both
report 570 instructions and a 536-byte generated stack frame. Manual inspection
of the prologue, argument moves, package-part iterator, compact-name dispatch,
error branches, and terminal paths shows the same control-flow structure;
addresses and RIP-relative displacements are relocated between binaries. The
candidate body remains present and callable at `18c7f60`, and the baseline body
is present at `18eded0`. This review does not treat a symbol-count or text
search result as proof of elimination.

## Generated-body measurements

The assembly receipt reports 3045 instructions and a 4096-byte generated frame
for baseline `transaction-commit`, versus 3395 instructions and the same
4096-byte frame for the candidate. The added body includes the proof/limits
branch and ownership handling. These are generated-body counts and compiler
frame metadata; they are not peak stack or memory measurements.

The baseline `to-bytes-bounded` target is present with 2582 instructions and a
3256-byte generated frame. The candidate query for that old symbol returns a
208-byte section-header-only artifact and is marked `missing_or_inlined`; the
candidate commit call visibly targets `to_owned_package_bounded` instead. This
does not establish that serialization code disappeared, and no such inference
is made here.

## Review result

The archived candidate assembly preserves the source precheck, adds an
explicit limits/proof guard, skips only the final validator on a successful
matching proof, and falls back to the original validator on invalid limits,
missing proof, or proof mismatch. The validator and post-validation media
checks remain present. I found no assembly control-flow blocker. Performance or
retention conclusions are outside this review and remain dependent on the
separate capture and harness receipts.
