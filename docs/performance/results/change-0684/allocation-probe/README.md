# Prepared-query allocation probe

Build this identical package in the baseline and candidate checkouts. Invoke:
`xls0684-alloc owned|file q1|q2|q3|visit FILE SHEET ROW COLUMN`.
Opening and the zero/one/two preparation queries occur before counters reset.
The measured operation is one query or one visitor call. Owner and returned
value/error remain alive through gauge reads; report formatting happens after
all gauges are captured. Unlike the native probe, sources have no read-count
wrapper; owned uses core OwnedSource and file uses core FileSource.

The standalone unsafe allocator delegates to System; production adds no unsafe.
Successful realloc counts its full requested new size in allocated bytes and
only its size delta in logical live bytes. The peak excludes allocator-internal
copy overlap, headers and RSS. Explicit deallocation counts exclude realloc's
internal release. Query2 retention includes its newly published index; query3
retention is a delta above the existing index, not the total owner footprint.
Run each row in three separate processes and require identical metrics.
