# 0800 checked short-prefix candidate

This directory is a source-only candidate archive based on
`49246ebe62536f28aaee8fb80fdd2f8fa518d328`. It contains no production edit,
build receipt, workflow result, or adoption claim. The five helper owners are
archived because the workspace keeps the canonical implementation in
`litchi-opc` and exact source copies in the four crates that cannot depend on
it.

The candidate keeps the iterator's existing three fields. Its phase is
`First`, `Second(offset)`, `ShortTwo { first_end, next_start }`, a boxed
ordered-map backend, or `Done`. The first two successful attributes use an
unchecked quick-xml iterator. The second key is recognized from its raw key
prefix and compared with the first key before quick-xml scans the second
value. A repeated key without an equals sign still goes through quick-xml and
returns `ExpectedEq`.

On the third call, `first_end` lets the candidate recover the second key prefix
from the tag bytes and `next_start` addresses the third key. The candidate
compares the third key with the first two before asking the unchecked parser to
read its value. A third lexical error or end returns directly and allocates no
backend. A successful third item followed only by whitespace becomes `Done`.
If more bytes remain, the backend is seeded with the first two raw key slices,
their positions, and the third borrowed key. It then checks every later key
with an ordered map. The backend repeats the key-prefix preflight before every
unchecked parser call, so a duplicate with an invalid value is refused at the
same duplicate-before-value boundary. It never reconstructs a checked
quick-xml iterator and never reparses an earlier value.

Starting the ordered map at the third item is deliberate. It keeps the short
state to two `usize` offsets and avoids an inline array or a replay phase. The
tradeoff is that tags with four or more attributes pay an ordered-map seed and
`O(log n)` comparisons from the fourth item onward, including the names of the
three already returned attributes. This remains bounded for hostile input and
is a direct experiment for root-owned layout and cost measurements; the source
does not assume that its likely 128-byte phase is an improvement over the
120-byte baseline.

The shared test copy retains the canonical differential, first-error,
recovery, clone, fused, and bounded-comparison checks. It adds third and later
malformed transitions, duplicate precedence after the backend starts, clone
advances through the short and boxed phases, and long values before a third
duplicate. It also records a late duplicate whose key begins with `=`. The
current baseline's late fallback `name_at` treats that prefix as an empty key
in the `>32` path, while quick-xml consumes the first byte as part of the key;
the candidate's exact `key_at` preflight therefore exposes a possible
correctness correction. That behavior is intentionally disclosed for separate
review and is not presented as a performance benefit.

The root coordinator owns the isolated five-copy quality run, direct probe,
the frozen 39-case protocol and seed `800080`, native captures, resource
measurements, and disposition. The 0799 protected-error policy remains frozen;
this archive does not move its advancement gates or authorize workflow trials.
