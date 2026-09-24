# Contextual source-profile requirements

This profile is bounded evidence infrastructure for the XLSX source scanner.
It must remain usable while lifecycle and package APIs evolve, but it may only
bind to a frozen, callable source API. It must never guess a production API or
silently turn a failed semantic check into a timing sample.

## Scope

The source boundary under test is:

```text
borrowed drawing XML
  -> SourceDrawing::scan
  -> direct SVG owner/value projection
  -> raw fragment and NamespaceContext observations
  -> contextual SVG codec read/write
```

The following are deliberately outside the profile: OPC/ZIP open and save,
worksheet edits or commits, relationship graph mutation, media ownership,
clone/capture/attach/detach operations, one-cell/absolute lifecycle editing,
chartsheets, groups, MCE ancestry, linked-resource fetching, rendering, and
native Office validation. A source scanner result must not be presented as an
end-to-end XLSX result.

## Corpus and lanes

The deterministic regular corpus crosses picture counts `{16, 32, 128}` with
inherited root namespace declaration counts `{0, 32, 128}`. Each declaration
uses an approximately 1 KiB URI. Four operation families run at every point:

| Family | Lane prefix | Semantic receipt |
| --- | --- | --- |
| contextual read | `contextual_read_` | picture count, raw-source sum, source/context counts, binding maximum |
| standalone export/readback | `standalone_export_` | output byte sum, embedded readback count, opaque QName preservation |
| scalar reference edit | `scalar_reference_edit_` | edited relationship readback count and opaque QName preservation |
| output-cap refusal | `small_cap_refusal_` | every picture refuses at 128 bytes; no partial success |

The original corpus lane is repeated for contextual read, standalone export,
and refusal. It is exactly 149,433 bytes, 32 pictures, and 128 declarations;
its SVG fragments have no synthetic opaque QName fields. The fixture manifest
is `corpus-manifest.json`, and the harness asserts the exact original size.

## Receipt and measurement contract

The executable emits one JSON object per invocation. The runner stores one
object per fresh process and records `/usr/bin/time -v` beside it. A valid
bundle has three processes per lane, at least two warmups, and at least twenty
measured samples per process. Every process must report the same FNV fixture
hash and SHA-256 fixture hash for a lane. `corpus-sha256.tsv` repeats those
SHA-256 identities in a compact reviewable form.

Each sample must satisfy the allocator accounting equation:

```text
live_before + direct_allocated_bytes + realloc_new_bytes
  - realloc_old_bytes - deallocated_bytes == live_after
```

There must be no allocator underflow or failed allocation. Successful lanes
must have `actual_success`, `semantic_ok`, and `output_exact`; refusal lanes
must have `actual_success=false`, `semantic_ok=true`, `output_exact=true`, and
one refusal for every selected picture. Export and edit lanes must read back
every selected embedded reference with the exact expected ID (`rIdSvgN` for
export and `rIdEdited` for edit), and must report no linked reference.
Regular synthetic lanes must preserve both `q:Opaque` and `q:QName` in their
exported output. The original lane is exempt because it intentionally has no
synthetic QName fields.

`retained_raw_source_bytes` is the sum of the raw SVG fragment slices held by
the selected values. `source_none_count` is a projection diagnostic only.
`context_distinct_count` groups context handles using the public
`NamespaceContext::shares_storage` predicate, not wrapper address identity;
`shared_context_identity` is true only when every selected value has a
context and all those contexts share one immutable backing node. RSS comes
from the fresh process and is reported separately from allocator peak live
bytes.

## Before/after comparability

Before and after runs must use the same harness bytes, lockfile, compiler
settings, corpus recipes, process/sample counts, and offline dependency
resolution. The only intended production input delta is the frozen XLSX host
source snapshot recorded in each result's manifest and provenance. The shared
DrawingML codec hash must be explicitly recorded. A mutable main checkout is
not an acceptable source for either side.

The profile is exploratory source evidence. It does not set performance
thresholds, does not claim an optimization factor, and does not authorize a
lifecycle freeze. Any comparison in a report must name the lane, corpus,
source hashes, machine, build mode, and metric; medians and spread are more
useful than a single sample.

The runner pins the expected before/after commit, refuses a dirty production
`crates/` tree, refuses an existing Cargo target, validates the binary hash and
source hashes before cleanup, retains per-lane stderr, and fails a warmup
error. The frozen before snapshot must expose no contextual handles and retain
the complete source projection; the frozen after snapshot must expose one
shared context storage node per selected drawing and report
`source()==None` for every selected SVG value. A separate normalized-manifest
check admits only the two intended frozen host-file deltas.
