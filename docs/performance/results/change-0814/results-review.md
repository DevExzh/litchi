# 0814 terminal results review

This is an independent, read-only review of the retained 0814 packet after
capture, decode, and assembly. It makes no production recommendation and ran
no workload, Cargo build, profiler, or binary tool. The review covers the
current source at `5eb8629254`, [analysis.json](analysis.json),
[root-audit.json](root-audit.json), [offset-audit.json](offset-audit.json),
the retained assembly receipt, and [assembly-recovery.json](assembly-recovery.json).

The packet contains 54 native reports and two large-only perf reports: 56
reports and 1,820 verified outputs. Source, output, full semantic, raw perf,
gzip, and exact-binary identities pass. The fresh probe has 36 passing tests;
the 1,241 passed, zero failed, three ignored production result is reused by
exact 0813-after identity. The three offline readers were independently
checked:

```
analysis.py --check       PASS 56 reports/1820 samples
root_audit.py --check     PASS 56 reports/1820 samples
offset_audit.py --check  PASS exact-binary leaf-offset replay
```

The timing lane measures build perturbation. Wrapper/control p50 ratios are
1.002120, 1.012834, and 1.040429 for tiny, medium, and large; fp/wrapper
ratios are 1.005104, 0.999433, and 1.006262. The large wrapper interval is
[1.034036, 1.044374], while the fp interval is [0.999645, 1.013582]. These
values do not establish a source speedup or an ordinary-build phase share.

Perf stacks are large-only and remain descriptive:

| repeat | whole | exact owner | scanner leaf | inspector leaf | SHA leaf | reader leaf | CheckedAttributes::next leaf | unknown-owner stack | lost lines |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 0 | 2867 | 851 | 150 | 39 | 91 | 69 | 25 | 2 | 0 |
| 1 | 2851 | 853 | 142 | 59 | 84 | 67 | 14 | 1 | 0 |

The scanner and inspector inclusive counts are 729/716 and 239/284, so they
overlap. Unresolved samples and frame occurrences remain retained. Tiny and
medium have timing controls and exact assembly only; they have no sampled
stack counts.

The exact fp offset replay maps scanner self leaves to
`movups 0x18(%rax),%xmm1` at offsets `0x25b` and `0x33b`, with 66/62 and 59/52
samples across the repeats. The ordinary and profiling-wrapper scanner ranges
contain the analogous two arm-local moves. The old three pre-dispatch copy
chain is absent. Offset replay conserves each owner sample: 150 + 39 + 662 = 851
and 142 + 59 + 652 = 853, with every mapped offset on an instruction boundary.
This mapping still cannot assign causal cycles because the stack build uses
frame pointers and samples can skid.

The first assembly attempt stopped correctly when the short inspector suffix
matched both `litchi_opc::content_type::inspect_element` and
`litchi_pptx::notes::codec::inspect_element`. Its control scanner artifacts
were preserved. `finish_assembly.py` then selected the fully qualified PPTX
mangled prefixes and recorded all six scanner/inspector ranges against the
unchanged control, profile, and fp binary hashes. The recovery record states
`workload_rerun: false`; recovery and offset-conservation validation pass.

A bounded source hypothesis supported by this evidence is to borrow
the `Start` and `Empty` event payloads in their existing match arms:
`Ok(Event::Start(ref element))` and `Ok(Event::Empty(ref element))`, changing
the calls to pass the borrowed `element`. `resolver.push` and
`inspect_element` already consume only a `&BytesStart`, and no reference
escapes either arm. The event therefore remains alive through the calls and
drops at arm exit; resolver pop scheduling, node/depth/attribute limits, typed
error mapping, and the explicit error arm retain their source order. The
inspector still unescapes every non-namespace attribute value, while only
relationship-namespace values are retained. This explains why the candidate
is bounded; it does not prove that LLVM will remove the two 32-byte moves or
that the change improves a workflow.

Any candidate trial needs fresh ordinary/profile assembly (and fresh fp
assembly if comparing sampled offsets), the existing differential and refusal corpus, resource-limit checks, and all eighteen
public workflow rows across the six generated shapes. Namespace discovery and
checked-attribute handling should remain separate. The broader native-producer,
cold-cache, range-source, non-seek output, scaling, and cross-format goals
remain open even if this bounded candidate succeeds.
