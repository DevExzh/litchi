# 0450: Decode and retain a verified compressed entry in one read

Previous turn: progress, committed 0449. Initial worktree contains only the
user GOAL.md; no owned CPU job is live. Accepted ADR tree remains unchanged
from the complete prior read. ZIP owns physical verification (0010/0011/0024);
source identity, semantic validation and managed reservations stay OPC-owned.

The existing capture API requires expected logical bytes, forcing a cold decoded
read before authorization. Add an explicit ZIP entry API returning decoded bytes
and an unforgeable verified compressed token from the same bounded source capture.
Share existing Store/Deflate validation, full compressed-span consumption, actual
CRC, authoritative size/descriptor checks and immediate callback cancellation.
The new method retains compressed C plus decoded U bytes and fixed scratch; callers
must reserve both before calling. It does not mutate the normal decoded cache.

Validate old and new paths on identical deterministic Store/Deflate fixtures,
including empty, short-read, malformed/CRC/size/trailing-stream, out-of-bounds entry and
cancellation boundaries. Record exact positional read calls/bytes for cold
read-then-capture versus fused capture/decode. Require exact decoded bytes and
transfer-token equality, bounded requests and at least one compressed payload's
bytes removed for nonempty cases. This is a measured ZIP enabler, not a PPTX
end-to-end speedup. OPC/PPTX integration, managed token/cache lifetime and broader
native/cold/scaling goals remain required follow-up.
