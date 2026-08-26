# Change 0298: CFB directory byte limit

Date: 2026-08-27

Status: Implemented deterministic bounded ingress behavior

`performance_claim: none`

## Decision

CFB ingress now carries an explicit maximum padded directory-stream size in
both `OleFileLimits` and `SharedOleFileLimits`. The default directory policy is
64 MiB, matching the existing stream-move policy, while a separate 2 GiB hard
maximum remains compatible with the existing finite CFB input ceiling. Existing
one-argument input-limit constructors retain their behavior; callers select a
different directory ceiling through a fallible builder.

Version 4 directory declarations are checked with checked arithmetic before
the exact directory chain reserves its sector vector and visited map. Version 3
directory chains use a specialized fallible walk that checks the padded logical
directory size before every sector-vector push and never reserves the configured
maximum up front. The final checked sector-count multiplication is repeated
before physical-sector claiming and directory-data allocation. A ceiling hit
returns the typed `directory bytes` `LimitExceeded` error; an exact padded
boundary remains valid.

Shared CFB views retain the selected directory ceiling and reuse it during
`validate()`. Generic CFB validation represents this configured ceiling as an
incomplete blocked ingress result. XLS validation maps only the CFB directory
byte ceiling to the same blocked-ingress status path; malformed structure,
source I/O, and source-version changes retain their existing precedence and
error/report behavior.

## Scope and claim boundary

The boundedness claim is limited to refusal before materializing an oversized
logical padded directory stream and its directly derived directory-entry
buffers. It excludes total RSS, allocator traffic, FAT/DIFAT/MiniFAT and
physical-sector-role allocations, input-source memory, physical I/O, latency,
throughput, and broad Office-format improvements.

Focused tests cover default and invalid CFB limits, exact and one-byte-under or
over v3/v4 boundaries, shared propagation and revalidation retention, generic
CFB blocked reporting, and an XLS fixture whose two-sector directory validates
at the exact ceiling but is blocked without errors at one sector.
