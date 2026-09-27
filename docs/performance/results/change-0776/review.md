# 0776 source review

The independent reviewer confirmed placement after local namespace installation
and element expansion, before opaque returns, directive handling, branch
selection and output. Active offsets use the same processor. Namespace identity
plus borrowed local name is sound; spill storage is fallible, bounded and
accounted conservatively by the DOCX envelope. Sorting does not copy/hash URIs.

The reviewer requested broader QName validation for a sole prefixed attribute
on inactive opaque paths. Root's disposition: this check addresses expanded
attribute uniqueness; one valid prefixed attribute cannot collide with a valid
unprefixed attribute because its URI is nonempty. The early return preserves
existing QName handling rather than introducing acceptance. Full skipped/opaque
QName parity is not claimed by this batch; malformed-name validation outside
this check remains a follow-up audit. There is no identified pair of valid equal
expanded names missed by the fast path. The report records this scope limit.

The test writer initially placed duplicate names at the second attribute even
in large lists. Root requested late duplicates and an overflow-only duplicate,
so final tests actually exercise sorting rejection, alongside legal controls.
The tautological namespace-constant assertion was removed before source freeze.
Only root runs Cargo/native/profiling workloads. Both agents performed bounded
source/test work without native execution.

## Revised helper at 6ed76a881b

Independent review found no missed valid duplicate. Every collision shares a
local name, so stack comparisons and local-name grouping are sound. Equal-local
groups resolve post-declaration URI identities and sort them; total work is
O(n log n) for at most 4,096 source attributes. The DOCX envelope charges the
40-byte target tuple, eight slots, vector capacity and owner conservatively.

The revision removes the first candidate's incidental QName expansion for
unique-local attributes. Compared with that intermediate candidate, malformed
unique names in inactive opaque content may again follow the old behavior;
compared with the authoritative base, existing handling is preserved. Complete
malformed-name parity is not claimed. Error precedence can change on inputs
with multiple defects. The first paired capture's 24–38% synthetic regressions
are retained and motivate the revision; they are not results for revised code.

## Final evidence disposition

An independent read-only review accepts 6ed76a881b as a bounded correctness
repair with performance follow-up and found no correctness blocker. Root and
reviewer recomputed the seven final latency flags: worksheet +5.77%, low alias
+5.54%, styles refusal +5.62%, prefixed 2/8/9/32 +6.11/+11.11/+13.70/+19.70%.
No final RSS regression exceeds 5%. Ordinary allocation counts are unchanged;
the 32-name control adds 4,000 temporary spill allocations. These costs remain
visible; neither universal speedup nor full malformed-QName parity is claimed.

The reviewer initially alleged stale final analysis. Root checked without
rewriting it, and the reviewer withdrew that finding after exact in-memory
replay matched analysis.json (SHA256
b7ce42f8c5d5e36d8e38208b5a3ae1016ba829026500b3e72673e61122bcec6b).
The retained intermediate analysis-0.json has the earlier +37.57% prefixed-32
result; the final analysis has +19.70%. No raw capture was changed.

Root review corrected draft prose that incorrectly described native timings
as whole-process/startup-inclusive, and removed an unsupported assertion that
the seven new tests exercise the memory-accounting boundary. Native timings
cover the probe operation and observer; RSS and heaptrack are whole-process.

The final offline validator replays both analyses, source-bound quality logs,
all changed-document witnesses and six allocation receipts. It passes after
the three owned build targets are removed following source/binary/fixture hash
checks. The ten existing registry claims also pass structural validation; that
registry check is not evidence for the new packet's numbers.
