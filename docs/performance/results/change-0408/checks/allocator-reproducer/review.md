# Allocator failure-test review

The unchanged optimized failure test failed because its maximal-layout request
returned a value treated as non-null. The first panic poisoned its test mutex.
Passing the maximal request size through `std::hint::black_box` restores the
real failure-path exercise; all null/counter/content/deallocation assertions
remain. Root's final single-threaded test log retains five passes. The reviewer
also ran the same five tests with default threading successfully.

The reviewer observed pre-patch `objdump -dr --demangle` output calling
`record_allocation` after loading `isize::MAX`, without a `malloc` call. The
post-patch output included a stack round trip, `malloc`, a null test and the
`record_failed_allocation` branch. The pre-patch executable was rebuilt in place;
no pre-patch assembly artifact was retained. This is a recorded review
observation, not an independently hash-bound before/after compiler proof.
The retained standalone reproducer sources returned null and incremented failure
counters; they did not themselves reproduce the unit-test failure.

Exact reviewer verification, from tools/perf-baseline:

```sh
CARGO_BUILD_JOBS=4 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=1 \
RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes' \
rustup run 1.98.1 cargo test --release --locked \
  --features allocator-metrics --bin litchi-perf-baseline-alloc \
  --target-dir /tmp/litchi-goal-fp-target -- --nocapture --test-threads=1
```

No production allocator logic changed. Root's exact final command and output
are separately retained in the checks directory.
