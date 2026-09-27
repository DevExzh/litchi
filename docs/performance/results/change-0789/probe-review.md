# 0789 probe review

This is a standalone Linux C11 diagnostic probe for the 0789 memory-accounting
experiment. It has no production-code dependency and writes only its fixed
stdout transcript; the parent owns timing, `/proc` snapshots, wait status and
raw-artifact retention. The probe is intentionally not a report generator.

## Invocation and validation

The command requires exactly three option/value pairs, in any order:

```text
--mib {0,1,4,16,64} --workers {0,4,32} --touch {main,workers}
```

Values are accepted only as nonempty decimal strings whose complete text is
consumed by `strtoull`; signs, whitespace, suffixes, overflow and duplicate or
unknown options are rejected. `--touch workers --workers 0` is rejected because
there would be no payload worker. `--touch main --workers 0` is the serial
control. The probe does not support `--option=value` or implicit defaults.

The requested size is converted to bytes with checked `uint64_t` arithmetic and
then to `size_t`. The page size comes from `_SC_PAGESIZE`; the page count is the
ceil division of the requested bytes by that size. A zero MiB case has no
mapping but still executes the complete checkpoint and ACK protocol.

## Mapping and touch routes

For a nonzero case the probe uses one `mmap` with
`PROT_READ | PROT_WRITE` and `MAP_PRIVATE | MAP_ANONYMOUS`, without
`MAP_NORESERVE`. It writes exactly one volatile byte at the first byte of every
page. Page `i` receives `(i % 251) + 1`; the main thread then reads every page
back and computes the checksum before the `touched` checkpoint. A worker route
partitions page indices into disjoint contiguous ranges, so no two workers
write the same byte. The checksum is a deterministic witness that the intended
pages were faulted and that the worker writes were visible to the main thread.

`--touch main` performs all payload writes and readback on the main thread. If
workers are requested, they still run and join before `touched`, but only touch
a small volatile stack buffer. This isolates thread creation/join overhead from
payload page faults. `--touch workers` creates the requested workers, has them
touch their disjoint payload ranges, joins them, and only then performs the
main-thread readback. The fixed `touched` line therefore follows the join in
both routes; the subsequent `workers_joined` line records the same already-joined
state so that every invocation retains the six-phase transcript. Workers also
touch a small volatile stack buffer, including the zero-MiB case.

The mapping remains live through `touched` and `workers_joined`. It is unmapped
before the `unmapped` checkpoint, and the final checksum is retained as a
diagnostic value after unmapping. On an error, already-created workers are
joined and a live mapping is cleaned up (an unexpected join failure terminates
the process immediately so worker storage cannot be released while still in use), but incomplete runs deliberately do
not manufacture later checkpoints.

## Transcript and handshake

Each of the six phases is emitted in this order:

```text
startup
mapped
touched
workers_joined
unmapped
final
```

The first line is sampled with `getrusage(RUSAGE_SELF)` and has the exact form:

```text
RSS0789\tphase\tpid\tmaxrss\tminflt\tmajflt\tbytes\tchecksum\n
```

The probe writes each line through a fixed stack buffer and an EINTR-safe exact
`write` loop. It then reads exactly two bytes from stdin and accepts only `+\n`.
EOF, a short read, a wrong byte, or a write/read/resource error terminates the
run nonzero. After the ACK, it samples `RUSAGE_SELF` again and writes:

```text
ACK0789\tphase\tpid\tmaxrss\tminflt\tmajflt\n
```

The ACK line itself does not require another ACK. The parent must wait for that
line before sending the next two-byte ACK. The same handshake applies to the
`final` phase, which prevents the child from exiting before the parent has
captured its final post-ACK usage sample. The mapped byte count is zero at startup, unmapped, and final; it is the
requested size at mapped, touched, and workers_joined. `checksum` is zero before payload readback and is the
verified page-pattern sum from `touched` onward (also zero for a zero-MiB case).

The probe emits no `/proc` data, timing, heap data, JSON, allocator controls or
environment-derived behavior. `ru_maxrss` is the process-local kernel resource
counter; the parent must retain independent `/proc` snapshots and distinguish
those observations from the probe's deterministic page/checksum witness.

## Static review status

The source is written for the requested strict command:

```text
cc -std=c11 -O2 -g -Wall -Wextra -Werror -pthread probe.c -o probe
```

This review performed source inspection only. Compilation and execution remain
root-owned so that the parent can record the exact compiler, binary identity,
handshake transcript, and raw artifacts in the 0789 packet.
