# Full-goal continuation after the one-paragraph append path

The complete non-iWork objective in `docs/GOAL.md` remains authoritative.
The one-paragraph tail operation addresses retained source/candidate XML and
paragraph-index ownership for its admitted existing-document closure. It is
not the very-large authored-stream requirement by itself.

The [0484 implementation design](../change-0484/design.md) maps the next
API, ownership, replay and validation work to the current code. That work must
implement the
[replayable paragraph contract](../change-0481/window-contract.md#full-goal-continuation-bounded-authored-streams):
borrowed bounded text chunks and explicit paragraph events, one authenticated
replay source shared across candidate validation, OPC publication, and durable
forward patches, and exact inverse restoration requiring the original source
provider. A one-shot iterator cannot satisfy all those passes without explicit
replay storage. Measure 64, 256, and larger authored paragraph streams while
varying existing source size independently. Generated XML must not accumulate
in a fragment proportional to the entire authored stream.

The [process profiles](profile-review.md) identify repeated XML auditing and
replay as the bounded route's CPU concern. Add phase attribution or equivalent
matching-stack evidence before changing those passes. A valid optimization must
preserve source freshness, format semantic validation, OPC publication auditing,
measure/emit authentication and error ordering; a lower heap peak alone does
not justify the measured latency and instruction increases.

Continue source-provider and publication evidence with positional filesystem
input, short-read range adapters, configurable high-latency providers,
sequential sinks, and atomic filesystem replacement. Keep cold/warm process
measurements and concurrent work separate from the present route comparison.
No package or format owner should acquire ambient networking or storage.

The wider goal still requires the complete Office CRUD scenario matrix,
reproducible baselines for important uncovered workflows, remaining legacy
CFB paths, justified data-layout changes, explicit bounded parallel execution
with measured scaling, native producer evidence, and the final
requirement-by-requirement completion audit. Keep individual latency, RSS,
allocation, I/O and scaling regressions visible throughout that work.
