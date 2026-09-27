# 0779 candidate review — bounded `read_limited` growth

Review status: **no reachable correctness blocker found**. The review is
against base `345b81ce8a` and the applied, rustfmt-checked `phys_pkg.rs`; this
review did not run Cargo, native commands, or profilers. The packet's candidate
diff is retained as the reviewed change record.

## Capacity bounds

The candidate computes `required = data.len() + read` after the existing
`read <= chunk <= maximum - data.len()` checks. Therefore `required <= maximum`
and the checked addition cannot overflow on a valid path. When growth is
needed, `bounded_reservation_target` computes a growth of
`max(current_capacity / 8, 8 KiB)`, then clamps the result to
`[required, maximum]`. `additional = target - data.len()` is consequently
positive and cannot underflow. The checked-add fallback to `maximum` also keeps
the target within the logical input ceiling if the capacity arithmetic itself
would overflow.

The initial reservation remains `min(maximum, 8 KiB)`. The patch only changes
the reservation made after a successful read; it does not increase the
8 KiB read buffer or the configured input limit. The allocator may still give a
`Vec` more physical capacity than an exact request, as it may for the existing
`try_reserve_exact` calls; the candidate bounds the requested target rather
than claiming a hard allocator-capacity guarantee.

For the 4,226,429-byte source, the candidate policy models 44 requested
capacity allocations. With the source length used as an analytical cap, the
requested-capacity sum is 40,327,094 bytes (38.459 MiB), ending at exactly the
source length. With the normal 512 MiB input ceiling, it is 40,650,096 bytes
(38.767 MiB), ending at 4,549,431 bytes, about 7.6% spare. These are request
models, not allocator copy traffic, physical memory, RSS, or live memory.

## Reader and error behavior

The candidate preserves the existing sequence and boundaries:

* every ordinary read still receives an 8 KiB buffer, reduced only by the
  remaining logical limit;
* `Interrupted` is retried before any capacity decision;
* a zero-length read returns the accumulated bytes without a growth request;
* an invalid count greater than the supplied buffer is rejected before growth;
* late reader errors are returned before growth, with the same `OpcError::IoError`
  wrapping; and
* the one-byte exact-limit probe remains in place.

The authoritative applied version retains the original sentinel expression
`maximum as u64 + 1`. `ReadLimitsBuilder::max_input_bytes` rejects values above
`usize::MAX - 1`, so this addition is representable for every builder-created
`ReadLimits`, including the largest accepted value. It would be prudent to add
a focused arithmetic assertion at that boundary only as an optional scope note;
the current tests cover ordinary exact and over-limit values and the growth
helper's `usize::MAX - 1` edge, while the existing builder bounds are unchanged
and the full limit cannot be physically exercised through `read_limited`.

Allocation failures remain mapped to `OpcError::Allocation` with resource
`"OPC package input"`. The two new checked-arithmetic failures map to
`IoError(InvalidData)`, but the invariants above make those branches unreachable
for a valid `ReadLimits` and `Read` result. No existing reachable typed error
path is changed.

The added `RecordingReader` cases cover exact full chunks, short reads, one
`Interrupted`, a late I/O error, an invalid read count, an exact limit, and an
over-limit sentinel. Their expected request lengths agree with the current
remaining-limit calculation, including the final short request in the
short-reader case.

## Retained-memory tradeoff

The one-eighth growth plus the 8 KiB floor removes the exact-reserve-per-chunk
quadratic request pattern while retaining at most roughly 12.5% growth slack
once the vector is large (and at most one 8 KiB growth increment while it is
small). The existing logical input ceiling remains the final bound. This is a
reasonable generic-`Read` tradeoff because it avoids a metadata hint and does
not change read scheduling. The direct before/after probe still needs to
measure cumulative allocation requests and peak/live/RSS metrics separately;
the arithmetic above cannot establish those physical outcomes.
