# The rejected first implementation: a retained `SharedOleStreamCursor`

Opportunity 3 of change [0574](../../../0574-ole2-next-opportunity-survey.md) says
"give the globals scan a retained stream cursor", and names
`SharedOleStreamCursor` — the primitive change 0566 built for the worksheet
scan — as the mechanism. That was implemented first, measured, and **rejected**.
This directory is its evidence, retained because the record's central finding is
why it lost.

The cursor form removed **more** chain steps than the change that landed
(1,076 links per flagship open against 2,099, because a cursor's read loop *is*
the chain walk and does not walk the prefix separately). It removed fewer
instructions, and it **added cycles**, because a cursor advances through four
`Result`-returning methods per sector — `read_exact`, `normalize_state`,
`physical_span`, `advance_within` — where `read_stream_range` has one flat loop.
`#[inline]` on the three helpers was tried and recovered almost none of it, which
is why the record attributes the cost to the per-sector dependency chain rather
than to call overhead.

| Path | What it is |
| --- | --- |
| `abba.json` | The paired A/B/B/A matrix against the same before binary, folded, 100 warmups and 3,000 samples. |
| `perf/cursor-*.csv` | `perf stat` isolation pairs, same method as the landed leg's. |
| `callgrind/{ann,incl}-cursor-*.txt` | Self and inclusive annotations for the three isolation pairs. |
| `callgrind/chain-edges.txt` | The `next_chain_sector`, cursor-method and `GlobalsBuffer::ensure` edge blocks at threshold 100. |
| `globals-buffer-with-a-retained-cursor.rs.txt` | The exact `crates/litchi-xls/src/workbook/source.rs` the measured binary was built from, retained as text so the rejected design is readable without reconstructing it. |

The binary that produced these is not retained; its sha256 is
`b67c3be3068da5c30fa08f9d477bc77d86ef4dd1b6cb636a6be77d40777e04d9`, built from
the same detached worktree of `32d25e088` with only
`crates/litchi-xls/src/workbook/source.rs` replaced by the file above.
