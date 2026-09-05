# 0427 resource observation review

All 240 retained samples returned to their own absolute entry live-byte value
after the final sink drop. Every phase live-byte change was identical within
each 30-sample process and between the two repeats of its API/corpus pair.
No phase-change repeat row crossed the declared 5% review trigger. These are
callback-order observations of one frozen build, not a before/after resource
improvement, RSS-release, cache-eviction or leak finding.

| Corpus | API | After publication B | After document drop B | After caller Arc drop B | After sink drop B | Whole-probe peak above entry B |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| plain | owned | 618915 | 128626 | n/a | 0 | 1181495 |
| plain | source-backed | 277062 | 192119 | 128626 | 0 | 909874 |
| media-rich | owned | 168526835 | 67265282 | n/a | 0 | 203809932 |
| media-rich | source-backed | 124056121 | 100897023 | 67265282 | 0 | 143640102 |

Deltas in the table are from each sample’s entry, except the explicitly named
whole-probe peak above entry. Both repeats have the displayed values. The
whole-probe region includes cloned inputs and the reserved sink; its maximum
must not be substituted for the older operation-only region peak.

At the source-backed document-drop checkpoint, the caller still holds both
instrumented source Arcs and the output sink. Actual strong counts are one
for each source. Dropping those final caller source owners decreases callback
live bytes by 63,493 B for plain and 33,631,741 B for media-rich. The remaining
callback delta then equals the common reserved sink capacity for that corpus;
dropping the sink returns the delta to zero. This journal identifies the
release boundary without converting global counts into object-owned memory.

Owned publication keeps both package snapshots and the published result
alive at its checkpoint. Source-backed publication has already consumed and
dropped the editor. The different owner sets and API behavior prevent treating
the publication or document-drop values as a matched optimization comparison.

The exact output comparison passed for all 264 iterations (24 warmups plus
240 retained samples). Inputs match across APIs; expected output identity is
checked within each API/corpus across repeats. Whole-process resource logs
include construction, gates, binary hashing and serialization; their RSS and
elapsed time remain context only.

The full cache/managed-budget/near-limit gate remains open. The next-work note
identifies existing managed constructors and the missing fallible diagnostic
forwarders; no cache or budget counters were silently replaced with zeros.
