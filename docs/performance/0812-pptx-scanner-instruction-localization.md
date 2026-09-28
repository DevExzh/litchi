# 0812 — localizing the current PPTX scanner leaf

Fresh native samples localize the scanner's leading sampled instruction to
`Result::map_err` payload movement at `notes/codec.rs:389`. The release
assembly contains three successive 40-byte payload-copy sequences after
`Reader::read_event_impl` and before event dispatch. This supports a bounded
experiment replacing `read_event().map_err(xml_error)?` with an explicit
success/error match. It does not establish causal instruction cost or a
performance benefit. Production remains unchanged.

## Exact binary and fresh evidence

The attempt to reconstruct the [0811](0811-pptx-current-native-attribution.md)
frame-pointer executable produced the same 14,387,272-byte length but a
different SHA-256: `63b55206…c231` instead of `b8b89fc5…2062`. Source, probe,
compiler versions, Cargo arguments, and flags match; the reason for the binary
mismatch is not established. The guard refused historical instruction mapping.
The original plan, driver, failure receipt, and historical offset census remain
retained. Historical samples are neither relabeled nor pooled with new samples.

A separately frozen [amendment](results/change-0812/fresh-plan.md) captured two
new profiles of the actual rebuilt executable: CPU 12, `cycles:u` at 499 Hz,
frame-pointer stacks, 100 large captures per run, and zero warmup. All 200
measured outputs match the sealed current-source source/output/semantic
oracles. Both profiles were decoded while that exact binary existed; raw data
and frames are retained as four deterministic gzip members. Disassembly and
DWARF line attribution belong to this same binary.

| Fresh repeat | Whole samples | Exact-owner samples | Scanner leaf samples | At offset `0x24d` | Unknown interior samples | Lost-event lines |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 0 | 3,066 | 960 | 231 | 193 | 2 | 0 |
| 1 | 3,095 | 970 | 227 | 178 | 0 | 0 |

The exact owner is `namespace_uri_probe::capture_region_0793`. The scanner
symbol is 2,588 bytes. At offset `0x24d`, the instruction is
`movups 0x10(%rcx),%xmm1`. `addr2line` identifies the inline
`core::result::Result<T,E>::map_err` frame at Rust `result.rs:967`, called from
`notes/codec.rs:389`. The following copy sequences lead into the event
discriminant dispatch. Independent raw offset counts agree with the primary
instruction join; all reported offsets land on instruction boundaries.

Sampling skid, inlining, and frame-pointer perturbation limit interpretation.
The earlier 0811 large-capture frame-pointer/profile p50 ratio is 1.069250;
this is a prior instrumentation warning, not a fresh ratio for the rebuilt
binary. No fresh ordinary-build control was collected. Instruction samples
cannot be converted into causal savings or ordinary-build phase fractions.
No latency, allocation, RSS, or universal speedup is claimed.

## Next experiment and semantic boundaries

The [scanner review](results/change-0812/scanner-review.md) proposes one local
change: match `Ok(Event::…)` and `Err(error)` directly on
`Reader::read_event`, applying `xml_error` only on failure and keeping the
existing event-arm bodies. First inspect
the candidate's release assembly: abandon this hypothesis if it leaves the
copy chain unchanged. A code-generation difference alone does not qualify
adoption; require fresh public-workflow before/after benefit, tail and resource
guards, all applicable quality gates, and exact differential outcomes.

Keep parser configuration, namespace push/pop timing, Start/Empty/End behavior,
limit ordering, element inspection, checked attributes, UTF-8/unescape,
relationship collection, semantic proofs, and the buffered oracle unchanged.
The generic invalid-root masking and parser/refusal identities remain binding.
This is distinct from the rejected 0806 attribute iterator experiment.

The [namespace review](results/change-0812/namespace-review.md) rejects a simple
merge of namespace discovery and checked attributes. The resolver's unchecked
pass installs declarations before root resolution; the checked pass has
different duplicate and error behavior. A shared traversal would require
additional state and differential validation to preserve ordering. It is not
the next local candidate.

## Execution and reproducibility

The rebuild passes with eleven inherited unused-helper warnings. Production,
all 9,196 source-file hashes, six probe-file hashes, and 35 architecture inputs
remain unchanged; the prior source-qualified quality witnesses still apply.
There is no new production correctness or six-gate execution claim.

The fresh driver completed both profiles, both decodes, and compression, then
hit a post-capture `NameError` (`m` versus `nm`) while recording assembly.
The frozen driver and all receipts remain intact. A separate completion script
reran only `nm` and ran `objdump`; no workload was retried. Both `nm` outputs
are byte-identical. The failure and recovery are checked by the validator.

[Packet and replay instructions](results/change-0812/README.md),
[instruction analysis](results/change-0812/instruction-analysis.json), and
[independent fresh audit](results/change-0812/fresh-offset-audit.json) retain
the evidence. The broader OLE2/OOXML goal remains active; iWork is excluded.

Independent readers and aggregate custody checks pass before and after cleanup.
The recreated target was removed after binary/source verification: 1,010 files
and 671,342,204 logical bytes. Compressed raw data, decoded frames, assembly,
line attribution, and both failure records remain retained.
