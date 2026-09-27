# 0779 source review — XLSX open allocation attribution

Status: source review complete; candidate and direct before/after phase
measurements remain pending.

Base reviewed: `345b81ce8a`. The review covers the shared OPC owned ingress used
by the ordinary XLSX path. It does not cover iWork.

## Attribution

The current ordinary XLSX path is:

```text
litchi_xlsx::Workbook::open
  -> litchi_opc::OpcPackage::open_with_limits
  -> read_owned_path_with_limits
  -> phys_pkg::read_limited
```

`read_limited` starts with an 8 KiB reservation and reads through an 8 KiB
stack buffer. After each successful read it calls `try_reserve_exact(read)`.
Once the vector is full, this requests a new exact allocation for every 8 KiB
chunk. The read boundary, the `Interrupted` retry loop, the invalid read-count
check, the late I/O error behavior, and the one-byte exact-limit probe are all
in this function.

The generated XLSX source used by the 0778 ordinary-save matrix is 4,226,429
bytes. For a model in which the reader returns 8,192 bytes for each full read,
there are 516 reads and the requested capacities sum to:

```text
8,192 + 16,384 + ... + 4,226,429 = 1,092,697,469 bytes
```

That is an allocation-request model. It is not a measurement of bytes copied
by the allocator, physical ZIP copying, resident memory, or live memory. The
model closely matches the prior 0778 generated-XLSX lifecycle observation of
1,100,160,722 cumulative requested allocation bytes, but the candidate still
requires direct phase and allocator confirmation before this is called an
attributed result. The prior phase observations also put the edit and save
paths at approximately 5.1 MiB and 1.2 MiB respectively, which is consistent
with open being the dominant lane; those phase values are evidence from the
earlier matrix, not a candidate comparison.

## Candidate growth policy

The smallest bounded change is to alter capacity growth while leaving the
8 KiB read requests and all error ordering unchanged. When a successful read
needs more capacity, the candidate should choose:

```text
required = data.len() + read
geometric = data.capacity() + data.capacity() / 8
target = min(max(required, geometric), maximum)
```

It should call `try_reserve_exact(target - data.len())` only when the current
capacity is below `required`. `maximum` is the existing checked
`max_input_bytes` value. The initial reservation remains the existing
`min(maximum, 8 KiB)` reservation.

The `max(required, geometric)` rule is deliberate. While capacity is below
64 KiB, the required 8 KiB read dominates and preserves the existing read
size. Once capacity is larger, the one-eighth growth leaves enough spare
capacity for several subsequent 8 KiB reads, so the vector grows geometrically
without changing the reader's request sizes. The target is never larger than
the existing logical input ceiling, including for a generic sequential reader
whose final length is unknown.

For the 4,226,429-byte model, using the source length as an analytical upper
bound gives 44 capacity allocations, a final capacity of 4,226,429 bytes, and
40.327 MB (38.459 MiB) of cumulative requested capacity. With the normal 512
MiB input ceiling rather than the source length as the cap, the modeled final
capacity is 4,549,431 bytes (about 7.6% spare) and cumulative requested
capacity is 40.650 MB (38.767 MiB). These are model outputs, pending direct allocator
measurements. They imply a large reduction in cumulative allocation requests
while retaining bounded spare capacity; they do not imply a latency, RSS, or
physical-copy improvement.

Plain `try_reserve(read)` would also remove the quadratic request pattern, but
can retain close to twice the input due to allocator growth. An exact file
metadata preallocation would be smaller for stable filesystem files, but it
would make an untrusted or stale metadata hint an immediate retained-memory
commit and would not solve generic `Read` ingress. A fixed 1 MiB increment has
larger slack for small inputs and still has quadratic cumulative requests near
the maximum. The bounded one-eighth policy avoids both tradeoffs without a
metadata hint, unsafe code, or a changed read schedule.

## Required confirmation

The candidate is acceptable only after a fresh before/after probe binds the
same source, binary options, and fixture identities and measures `open`,
`edit`, `save`, and `lifecycle` separately. The direct attribution control
should exercise `read_limited` with a reader returning exact 8 KiB chunks, then
repeat with short reads, one `Interrupted`, an exact-limit input, an over-limit
input, and an invalid read count. Each case must retain the output bytes or
typed error identity.

The phase probe must keep the existing source/output identity checks. Native
timing and allocator observations must remain separate processes, and no phase
peak or cumulative allocation value may be added to another phase. Any report
must distinguish cumulative allocation requests from copied bytes, allocator
reallocation traffic, RSS, and live bytes. A lower allocation-request total
alone is insufficient evidence of a physical-copy or host-memory reduction.

The candidate must preserve the exact-limit sentinel (`actual = maximum + 1`),
`Interrupted` retries, short-read completion, late reader errors, and the
invalid `Read` count refusal. The source package and ordinary XLSX output must
remain byte/semantic-identical under the existing admission and reopening
checks.
