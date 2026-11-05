# 0362ul94d — OPC generated-part accounting compiler return

The actual combined OOXML compiler-first run on
`0db7ff3a6aee574102ae0319da2c1d9e14e5e4f2` passed formatting across 90
changed Rust paths, then stopped on one unused closure binding in OPC archive
metadata accounting. This repair uses `contains_key` for that existence check
and retains the same short-circuit provenance/member predicate. There is no
additional traversal, allocation, publication, or change to byte admission.

Retained actual state:
`/home/zhuhe/code/litchi-worktrees/.coord/ooxml-combined-0362ul94-f5-actual-0db7ff3a6a-v2`.
The earlier v1 packet was rejected before GO for driver receipt handling; it
remains inert and does not represent a failed build.

The exact compiler command was
`cargo check --offline --locked -q -p litchi-ooxml-common --all-targets --all-features`.
It exited 101 after 30.102 seconds; target size was 154,338,851 bytes.
Root and an independent reviewer verified source/three-lock fences,
log/trace identities, empty preflight builders, no live guard errors and
terminal owned-process samples. The compiler PID is absent.

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 12,097 | `7f055deeea31e5aba2ad80628f32c1df281de487798ba110a824f02f68101fb9` |
| `result.json` | 13,179 | `4dc43a83a5167b09de70f4e6520877b460e037576a901d9158e24dbf2b96c4f9` |
| `01-check-common-attempt-actual-f5-01.log` | 17,325 | `cd8d53ff1781709799a9cf894d9421cea5a191455424509e4d144cde81c718d1` |

Fresh compiler validation is required. The 65-row combined proposal and
Custom XML fidelity/limit returns remain open; this record claims no runtime
or performance acceptance.
