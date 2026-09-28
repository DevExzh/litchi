# 0801 checked short-prefix no-replay candidate

This is a source-only candidate archive rebased onto
`54aa0f3fb4967d98c923f58b3944911b9b4525ab`. The `before/` snapshots are the
five current helper copies and the shared OPC test source at that commit. The
`after/` snapshots carry the deferred short-prefix design, its candidate
tests, and the corrected fallback parity already present in the base source.
No production edit, build, native capture, or adoption claim belongs to this
archive.

The five helper owners are archived because the workspace keeps the canonical
implementation in `litchi-opc` and exact source copies in the four crates that
cannot depend on it. The copies retain their existing visibility and test
module paths.

The candidate keeps the iterator's existing three fields. Its phase is
`First`, `Second(offset)`, `ShortTwo { first_end, next_start }`, a boxed
ordered-map backend, or `Done`. The first two successful attributes use an
unchecked quick-xml iterator. The second key is recognized from its raw key
prefix and compared with the first key before quick-xml scans the second
value. A repeated key without an equals sign still goes through quick-xml and
returns `ExpectedEq`.

On the third call, `first_end` recovers the second key prefix from the tag
bytes and `next_start` addresses the third key. The candidate compares the
third key with the first two before asking the unchecked parser to read its
value. A third lexical error or end returns directly and allocates no backend.
A successful third item followed only by whitespace becomes `Done`. If more
bytes remain, the backend is seeded with the first two raw key slices, their
positions, and the third borrowed key. It checks every later key with an
ordered map. The backend repeats the key-prefix preflight before every
unchecked parser call, so a duplicate with an invalid value is refused at the
same duplicate-before-value boundary. It never reconstructs a checked
quick-xml iterator and never reparses an earlier value.

Starting the ordered map at the third item is deliberate. It keeps the short
state to two `usize` offsets and avoids an inline array or a replay phase. The
tradeoff is that tags with four or more attributes pay an ordered-map seed and
`O(log n)` comparisons from the fourth item onward, including the names of the
three already returned attributes. This remains a direct experiment for
root-owned layout and cost measurements; the source makes no claim that its
likely 128-byte phase improves on the 120-byte baseline.

The shared test copy retains the canonical differential, first-error,
recovery, clone, fused, and bounded-comparison checks. It also retains the
leading-equals and other unusual-key regression matrix, including
malformed values, exact positions, repeated clone advances, lexical controls,
and the late leading-equals case. The base fallback's `name_at` and the
candidate copy both consume the first non-whitespace key byte before scanning
for its delimiter, matching quick-xml for that raw lexer input. This is a
parity invariant; it does not validate XML Names.

Candidate-only tests cover third and later malformed transitions, duplicate
precedence after the backend starts, clone advances through short and boxed
phases, and long values before a third duplicate. They do not constitute
quality or performance evidence until the root coordinator runs the isolated
five-copy gates and measurement protocol.

The root coordinator owns the exact five-copy quality run, direct probe, the
frozen 0801 39-case protocol and seed `801080`, native captures, resource
measurements, source custody, and disposition. The 0799 protected-error policy
remains frozen; this archive does not move its advancement gates or authorize
workflow trials. The candidate remains untested and unmeasured until those
gates are independently recorded.
