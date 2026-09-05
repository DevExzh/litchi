# Independent review

The design review recommended one observer mutex instead of a racy active flag
and independent atomic maximum. The first counter review identified same-thread
reentry deadlock and incorrect poison-as-overflow status. Those findings were
fixed with const TLS callback-entry suppression, sticky observer invalidity and
Unavailable status for active/future invalid regions. The final review found no
remaining counter correctness blocker. Numeric-only lock scopes, System calls
outside the observer, cumulative counters, endpoint/lifetime bounds and focused
concurrency/failure tests were checked. The reviewer ran no tests or builds.

Separate aggregation review confirmed measured/unavailable/overflow propagation
and sample ordering. Its missing unavailable-envelope test was added. Comparator
review confirmed all general operation/filesystem V3 routes and historical
compatibility, then requested raw filesystem bound failures and explicit V2/V3
mismatches. Those fixtures are included in the final 98-test passing suite.
Historical dedicated ABBA wrappers intentionally keep their old policies.
