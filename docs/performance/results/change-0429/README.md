# 0429 PPTX provider and native image baselines

This batch extends the source-backed performance evidence with explicit byte,
positional-file, and capped-range providers. Synthetic plain/media cross-copy
keeps the existing exact output and preservation gates. Separate native image
lifecycles use unmodified POI slide-image and video-poster fixtures and retain
the returned image past package-owner drops.

The protocol specifies 32 fresh release processes, two repeats, 30 retained
samples and three warmups per process. Supplementary CPU profiles cover 100
media-rich iterations each for bytes and files; they are excluded from the
formal baseline sample count. Capture completion must be established by the
command receipts and capture/profile indexes, not inferred from this protocol.

The clocks surround named public APIs only. Their sum excludes corpus gates,
source construction, copying, reservations, checks, observers, and drops.
Rates use that API-only denominator. Fixed caller-request histogram buckets
describe logical ReadAt requests; range caps constrain returned chunks. The
delay is a tool-side sleep per delegated call, including scheduler overhead.
It is not a calibrated bandwidth model or a remote-service measurement.

File inputs are written or hash-read before samples and reopened for each
iteration. These are warm positional-file observations. Process RSS includes
setup and observers. Managed budget gauges and cache retention remain separate
scopes; unavailable owners never acquire fabricated zero diagnostics.

[`native-review.md`](native-review.md) records the inspected native cross-copy
gap and fixture provenance. Original/LibreOffice shapes were initially proposed
as positive image inputs, but source review found unsupported nested image
markup and a markup-compatibility branch. Their untouched bytes and static
payload oracles remain as refusal evidence. The initial, uncaptured proposal
is retained separately from the final protocol.

`native-oracles.py --check` independently derives the positive selected payloads
from retained ZIP/XML inputs. `--check --shapes` verifies the separate static
shapes oracles. This is not a new native-application run.

Root serializes all builds, tests, analysis, profiles and captures. Rust is
pinned to 1.98.1, four build jobs, incremental compilation disabled, and one
test thread. Failed development commands remain in the bundle. The final
warning-denied harness diagnostics must match the retained 0428 baseline;
existing lint failures are not represented as passing commands.

The global non-iWork goal remains active. This batch adds reproducible
provider baselines, attribution evidence, and a bounded ZIP correctness fix, with no
causal speedup/regression, physical-copy, allocator, cold-filesystem, general
leak, or scaling claim.

Runtime preflight exposed a central-directory short-read bug: the ZIP iterator
requested a full fixed header even when bytes were already buffered, and could
silently discard a truncated trailing header. Four generated regressions failed
before the production refill fix and passed after it. This correction enables
valid short-read providers; it is not a measured speedup.
