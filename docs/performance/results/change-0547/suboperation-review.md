# 0547 sub-operation address review

`Disposition: pass for the frozen diagnostic attribution, with interpretation
limits recorded below. No production candidate is admitted by this review.`

This is an independent read-only review of the category function in
[`suboperations.py`](suboperations.py) against the retained baseline
disassembly in [`baseline/assembly-2.stdout`](baseline/assembly-2.stdout) and
the current private `SectorChainScratch` implementation
([`file.rs`](../../../../crates/litchi-cfb/src/file.rs#L2756)). No Rust source,
build, test, benchmark, profile, or capture was run or changed.

The evidence is bound to revision
`2118c6fb1d023005e224aaa938e84d6ed0588b70`, `file.rs` SHA-256
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`, the
`collect_exact` symbol at `0x2f223c0` with size `0x59c` (1,436 bytes), and
the following final artifacts:

| Artifact | SHA-256 |
| --- | --- |
| `baseline/assembly-2.stdout` | `24dbff93d02f57f8fb9d3149a063348ae1487dff5dc34a216f06720d9b45b24f` |
| `baseline/assembly-index.json` | `91f0bf36bd5ad9e680af13a60d7036f265b6dbd8eff3f1b0065e92e0f3fa6531` |
| `instruction-analysis.json` | `f5bd56a45b825399bb51a880e7549d98caf8cb14eaeb0be99316a311f15457d5` |
| `suboperations.py` | `f350fea4c5c24dd3621843aea97985ae7155e40d03db74e82ec6532d71b94dca` |
| `suboperations.json` | `ebe7466d7decbdf94f5e15e64a6e7c5659932b0abf860dc457871c5a7178ecff` |

## Partition and source correspondence

The category ranges are half-open instruction-offset ranges, with the one
explicit singleton at `0x1f8`. They do not overlap. All 160 instruction
positions that receive positive self-Ir in each of the four five-dump jobs
fall inside the measured symbol, receive exactly one category, and reproduce
the reported `collector_self_ir` when their category totals are added. The
static disassembly contains additional cold instructions; the current valid
profiles execute none of those error bodies, so their dynamic category totals
are zero rather than evidence that the paths are absent.

The hot ranges correspond to these source operations:

* `0x1c0..0x1f8` and `0x1fd..0x201` contain the visited logical-length check,
  word-index/capacity check, mask/address calculation, word load, membership
  test and branch, followed by the bit set. The branches to the defensive
  error blocks remain charged to the visited predicate; only their cold
  formatting/reset bodies are outside it.
* `0x201..0x23c` contains the sector-vector length/capacity check, the
  possible `Vec` growth path, the sector store, and the length increment. This
  is the loop append boundary. Its possible growth call is also represented
  separately as a direct child, so it is not silently added to self-Ir.
* `0x23c..0x26f` contains the allocation-table load, remaining-count update,
  intermediate and final marker tests, successor state update, and the
  repeated table-length check. `0x40b..0x414` contains the final
  `ENDOFCHAIN` test and branch. These are the FAT/MiniFAT lookup and marker
  checks for the exact chain walk.
* `0x6a..0x76` writes the successful `Result` status for the empty and
  non-empty terminal cases, and `0x2c5..0x2da` is the common function
  epilogue. The latter is shared by successful and error returns; the name
  `success_status_and_shared_return` records that distinction and must not be
  read as success-only work.
* The listed cold ranges contain the formatted corruption/error construction
  and the post-error scratch reset. They are not mixed into the positive
  valid-loop totals except for the ordinary branch instructions that select
  them.

The mapping is therefore suitable for the stated diagnostic question: how
much measured collector self-Ir is associated with the repeated visited
operation versus append and successor/marker work. The emitted loop has the
bitset operation inlined; the separate `CheckedBitSet::insert` symbol being
present in the assembly inventory does not imply a call edge from this loop.
The analysis correctly keeps direct `memset` and `finish_grow` descendants
outside the self-Ir category sum.

## Specific boundary findings

The `0x1f8` instruction is:

```text
mov %r12,0x38(%rsp)
```

It is a per-iteration stack spill/temporary between the visited test and the
bit set. There is no matching load from that slot in the linear collector
body before return; it is compiler bookkeeping around the following append
and possible reserve/error machinery. The script deliberately keeps it out
of the visited range and charges it to the broader
`fat_lookup_marker_checks_loop_state` bucket. That preserves a disjoint
partition, but the instruction is not a FAT load or marker test. The FAT
bucket therefore includes one shared bookkeeping instruction per executed
loop iteration. Its contribution over the five timed dumps is 164,640 Ir for
XLS-owned, 120 for tiny CFB, 20,480 for many-small CFB, and 163,840 for
few-large CFB. Future work that treats the FAT bucket as a pure removable
mechanism should either report this spill separately or retain this explicit
qualification.

The first current-sector allocation-table bound check is at `0x1a3` and is
currently caught by the script's final fallback category,
`entry_reservation_visited_preparation`. The repeated successor bound check
at `0x266` is correctly in the FAT/marker bucket. The `0x1a3` check executes
once per collector call (10 Ir per timed dump in this profile), so its current
placement does not materially change the dominant percentages, but the
fallback category is not a pure reservation/preparation category. The
`0x1c0` logical-length comparison and `0x1d0` words-capacity comparison are
visited-map bounds, not additional FAT bounds.

The source order for the entry work is preserved in the address layout:
`sectors` capacity is checked first (`0xc8..0xd3`), its growth helper is in
the `0x499` block, and `prepare_visited` computes the word count, grows the
retained map, zeroes newly added words and then fills the retained map
(`0xd9..0x191`, with direct `memset` children). The map's logical length is
set before the loop. This confirms that the attribution does not imply that
zero-fill, reservations, or allocation labels were removed.

The same fallback category also contains the cold allocation-failure arms
for the two fallible reservations: the sector-reservation failure setup near
`0x4c0..0x4d1` and the map-reservation failure setup near `0x510..0x534`.
Those arms construct `OleError::Allocation` and reset the scratch state; they
are not ordinary entry preparation. They are unexecuted in these valid
profiles. If a later malformed/allocation-failure run is attributed by this
script, that category must be described as a broad setup/reservation bucket,
or the failure arms should receive a separate cold-allocation category.

The ordinary corruption blocks are correctly separated from success:

* `0x3f..0x64` is the non-ENDOFCHAIN empty-chain error;
* `0x7e..0xc8` is the declared-count/table-length error;
* `0x26f..0x2c5` and `0x2da..0x40b` contain invalid-index, bitset-bound,
  cycle, and reset/error-result paths;
* `0x414..0x499` contains late-length, early-ENDOFCHAIN and invalid-marker
  formatting paths.

The `0x2c5..0x2d9` epilogue remains intentionally shared. The allocation
failure arms noted above are the only cold error construction that the current
named cold bucket does not include; this is a naming/interpretation issue,
not a positive-profile sum or overlap defect.

## Measured rows and limits

The final report reproduces these disjoint self-Ir totals:

| Job | Collector self Ir | Visited self Ir | Visited / constructor |
| --- | ---: | ---: | ---: |
| XLS owned-source | 5,601,140 | 2,304,960 | 20.3684% |
| CFB tiny | 5,200 | 1,680 | 0.6866% |
| CFB many-small | 764,485 | 286,720 | 2.0517% |
| CFB few-large | 5,571,945 | 2,293,760 | 22.3481% |

The visited range is 41.15–41.17% of collector self-Ir in the two large
profiles. These are Callgrind instruction references from one baseline
repeat, not native time, instruction-retirement hardware counts, allocation
bytes, memory use, scalability, or a fraction of end-to-end latency. The
constructor denominator includes all work under the selected owner, while
the category numerator includes only exclusive self-Ir in this collector.
Direct callees remain separate and must not be added to the category totals.

This review does not validate a future terminal-proof or checkpointed
implementation. Such a candidate still needs fresh source-bound builds,
semantic/error-order guards, allocation and reset tests, malformed-input
latency evidence, and the existing two-repeat OLE2/OOXML gates. ODF remains
deferred until the OLE2/OOXML optimization goal completes.

