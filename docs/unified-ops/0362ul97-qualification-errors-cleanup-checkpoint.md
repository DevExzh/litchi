# 0362ul97 — Qualification, typed consumers and cleanup checkpoint

Date: 2026-10-25. This checkpoint records private source compositions and
actual failed qualification captures. Neither candidate is accepted into
integration or trunk yet; final runtime and performance obligations remain.

## PPTX

The first qualification on `02b2489eada2bfcd3d1b9bc285b0ea0d3c7ae0d6`
passes formatting and minimal all-target compilation, then fails strict
all-feature/all-target Clippy with 76 diagnostics. Root verifies the complete
receipt, source/three-lock fences, command logs/traces, target guards and
terminal ownership. The capture remains at
`.coord/pptx-combined-0362ul89-f5-qualification-v2-02b2489ead`.
Its `result.json` is 37,850 bytes, SHA-256
`5878d7e3ed980656fefa0ddb83823d3fa077e3041ee8cde4cf058f0f092f1274`.

Independent reviews pass 89x/89y/89z. Root composes their 30 Rust paths at
`3181342f0cb51ee787934500ad467513de3e59d7`, with three leaf records added
by documentation-only `2da172ece2298107daa4feebe9b514b30eabe509`.
Root reproduces three formatting-only adjustments from immutable leaf bytes.
Guides and sections pass selected limits explicitly instead of duplicating
them on every parsed node; no per-node heap allocation is added.

The actual qualification on `2da172ece2` passes formatting (393 paths,
2.835 seconds) and minimal all-target compilation (58.533 seconds).
Strict Clippy clears the original 76 findings, then reports three test-only
findings (81.023 seconds, exit 101). No later row runs. Root verifies every
receipt and terminal session exit. State:
`.coord/pptx-combined-0362ul89-f5-qualification-2da172ece2`.

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 35,905 | `fc3a80c91ef8c91b8a5c52fed521c03578e2779ea090b74a75a935b533545381` |
| `result.json` | 37,806 | `c9eb8ac7323e4ed0fc8a6516328ca9115e9de8e8ef3deb69229385de0803785e` |

Root fixes the test-module placement and two identical constant expressions
in `ae6f32d202a04e3d6675fa431b2579c0d5f27192` (record 89ab), tree
`361e2d24d8c9d0cb2c838f109646b3ede9150545`. Independent review confirms
no changed production logic or assertions. Its fresh seven-row qualification
is running in
`.coord/pptx-combined-0362ul89-f5-qualification-ae6f32d202`.
Root corrects two prior-capture references before GO; final plan SHA-256 is
`b0c48a6274ceb6e4a3a637cc6b3ac5e9e3d060cdc1701cc0b6492a15e3e5b2e5`.
No result is inferred for this running capture.

## Shared OOXML and consumers

The actual consumer capture on `5ec78a4858` passes formatting and OPC
compilation, then fails the four-format check with DOCX/XLSX errors. Root
verifies all completed receipts; the remaining rows did not execute. State:
`.coord/ooxml-combined-0362ul94-f5-consumers-5ec78a4858`.

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 16,030 | `b7cf8149942c3602a6699d49d48d53779f0cbff2394f7afbe3fd76cb63fe21b4` |
| `result.json` | 17,338 | `c01eedfe86b84f78a6d302914a269fa99d2d7664fc4da5304a1275b77470331a` |

Private composition `95973e8c6a2bcbb8724d1be50aea4c510fd8cdf3`, tree
`6854b92295c3aa15c97e783e9bee125be8edaeba`, joins Custom XML through
`44e0be0633`, DOCX through `5d8aec545a`, and XLSX through `cef014a61f`.
Record 94i identifies the complete immutable chains. All 38 paths apply
without conflict. Root reproduces 36 leaf blobs and the combined ops tests
from immutable preimages, including three formatting-only normalizations.

Custom XML preserves lexical/opaque content, tracks selected MCE source
offsets without marker inflation, checks selected limits before materializing
XML, and fixes paired-schema indexing and typed source-limit errors. DOCX
moves the owned VBA bytes and preserves exact nested properties limit kinds.
XLSX joins delegate every retained error to the leaf classifier, preserving
refusal kinds, patch-conflict, limit and source-change classifications.

Reviews return incomplete intermediate proposals instead of accepting them:
94g loses typed limits; 94ga inaccurately projects resource dimensions;
94h loses wrapped classification; 94ha tests only the ordinary fallback;
94cd still has two offset type errors. Final chains repair all those gaps.
The source reviews pass; actual combined compiler/Clippy preparation is in
progress. The retained 65-row proposal, runtime/glossary and performance
obligations remain open.

## Cleanup wave 67

Root authorizes removal only after source, branch, payload and holder checks.
The retired worktrees are removed without force:

- `uo2-pptx-2749-integration-tests` at
  `9672f444068393af4c57b4408b7321c26cbed62d`.
- `uo2-pptx-tracks-borrow-00a42` at
  `5bf328754a02e6959e6f903ff6cb3cdd833d15af`.

Their branches remain. All three changed blobs exactly match the composed
PPTX source. Same-device sparse restoration succeeds; both paths and registry
entries are absent. No ignored/untracked payload needs quarantine. This is
source-preservation cleanup, not candidate runtime acceptance.

Evidence root:
`.coord/worktree-cleanup-2026-10-25/quarantine/wave67-pptx-compiler-test`.

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `manifest.json` | 2,827 | `2b7e535cd8ca7c3fea23e17dcd6f47b07138c29e208ac0f14fa22c82e26abd62` |
| `restore-proof.json` | 2,278 | `5a32b3f0bd638448ef7f9aae359818c2547c39581fe00a0ab4df9a154f15c2ad` |
| `receipt.json` | 3,367 | `88cacecbf405c2010f295401dd6481b7dedf3024bbb9be0d4374c518f67c81b8` |

Checkout allocation recovered is 61,112,320 bytes. Subtracting all 20,480
bytes of wave evidence gives full-wave net recovery of 61,091,840 bytes.
Root verifies refs, blob identities, restoration receipts, four held heads
and the protected constructor probe. Cleanup holds are released.

Final formatting, full 53-row gate including isolated documentation, fresh
84-profile raw census and separate 68-package inventory/dispositions,
corpus/native qualification and quiet ABBA remain open. Only redaction
audit 22 is deferred. Nothing is pushed.
