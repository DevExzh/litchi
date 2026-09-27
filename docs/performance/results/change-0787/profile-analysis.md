# 0787 cached Part profile analysis

This report replays retained receipts, reports, raw Callgrind files, and strace summaries offline.
Callgrind values are guest instruction counts for the exact operation owner. The trace lane is a
separate whole-child syscall census. No native wall-time claim, physical worker fraction, or
trace-time pooling is made.

- Reused parser: `../change-0784/profile_analysis.py` (Git blob `1c42922ea9ab5f3624f294f4f0e5f615775c234e`)
- Production source census: 9196 files at revision `6b47617020dd08bca196cbf3ab0b87a69aefbcb1`
- Profile lane: `profiles-compat`, six qualified regions
- Failed baseline attempt retained: SIGSEGV before owner, zero Ir, qualified `False`

## Callgrind diagnostic by width

| Width | Repeat Ir | Mean Ir | Owner self Ir | Immediate child Ir | Threading self Ir | Allocation self Ir | Channel self Ir | Cache/hash self Ir |
|---:|---|---:|---|---|---:|---:|---:|---:|
| 1 | 27877, 27877 | 27877.0 | 5, 5 | 27872, 27872 | 0.0 | 183.0 | 0.0 | 12084.0 |
| 8 | 58890, 57924 | 58407.0 | 5, 5 | 58885, 57919 | 13810.0 | 10505.0 | 4581.0 | 7245.0 |
| 32 | 96902, 98223 | 97562.5 | 5, 5 | 96897, 98218 | 35397.0 | 31289.5 | 0.0 | 7225.0 |

The owner self row plus its immediate child inclusive rows equals the owner edge and the numbered region summary.
Descendant inclusive rows are retained in each dominant path and are not added to that partition.
Worker roots may be disconnected in a threaded Callgrind graph, so owner inclusive is not treated as an all-thread total.

## Top self-Ir functions

| Width | Repeat | Top self-Ir functions (name: Ir) |
|---:|---:|---|
| 1 | 0 | litchi_opc::source_backed::batch::read_parts_ordered: 6825; litchi_opc::source_backed::SourceBackedPackage::read_part_with_observer_and_capture: 6176; core::hash::BuildHasher::hash_one: 6048; <core::hash::sip::Hasher<S> as core::hash::Hasher>::write: 5376; litchi_opc::source_backed::batch::read_serial: 1444 |
| 8 | 0 | litchi_opc::source_backed::batch::read_parallel: 9539; litchi_opc::source_backed::batch::read_parts_ordered: 7078; <std::sync::mpmc::waker::SyncWaker>::notify: 4449; <core::hash::sip::Hasher<S> as core::hash::Hasher>::write: 3584; pthread_create@@GLIBC_2.34: 3148 |
| 32 | 0 | _int_malloc: 11542; pthread_create@@GLIBC_2.34: 11037; litchi_opc::source_backed::batch::read_parallel: 7653; litchi_opc::source_backed::batch::read_parts_ordered: 7379; free: 4912 |
| 32 | 1 | _int_malloc: 12549; pthread_create@@GLIBC_2.34: 11037; litchi_opc::source_backed::batch::read_parallel: 7653; litchi_opc::source_backed::batch::read_parts_ordered: 7379; free: 4909 |
| 8 | 1 | litchi_opc::source_backed::batch::read_parallel: 9539; litchi_opc::source_backed::batch::read_parts_ordered: 7078; <std::sync::mpmc::waker::SyncWaker>::notify: 4439; <core::hash::sip::Hasher<S> as core::hash::Hasher>::write: 3584; pthread_create@@GLIBC_2.34: 3148 |
| 1 | 1 | litchi_opc::source_backed::batch::read_parts_ordered: 6825; litchi_opc::source_backed::SourceBackedPackage::read_part_with_observer_and_capture: 6176; core::hash::BuildHasher::hash_one: 6048; <core::hash::sip::Hasher<S> as core::hash::Hasher>::write: 5376; litchi_opc::source_backed::batch::read_serial: 1444 |

## Dominant paths

| Width | Repeat | Path (owner to dominant child) |
|---:|---:|---|
| 1 | 0 | cached_part_profile::cached_part_region_0787 → litchi_opc::source_backed::batch::read_parts_ordered → litchi_opc::source_backed::batch::read_serial → litchi_opc::source_backed::SourceBackedPackage::read_part_with_observer_and_capture → core::hash::BuildHasher::hash_one → <core::hash::sip::Hasher<S> as core::hash::Hasher>::write |
| 8 | 0 | cached_part_profile::cached_part_region_0787 → litchi_opc::source_backed::batch::read_parts_ordered → litchi_opc::source_backed::batch::read_parallel → <std::sys::thread::unix::Thread>::new → pthread_create@@GLIBC_2.34 → 0x0000000004a77750 → _dl_allocate_tls_init → memset |
| 32 | 0 | cached_part_profile::cached_part_region_0787 → litchi_opc::source_backed::batch::read_parts_ordered → litchi_opc::source_backed::batch::read_parallel → <std::sys::thread::unix::Thread>::new → pthread_create@@GLIBC_2.34 → 0x0000000004a77750 → _dl_allocate_tls_init → memset |
| 32 | 1 | cached_part_profile::cached_part_region_0787 → litchi_opc::source_backed::batch::read_parts_ordered → litchi_opc::source_backed::batch::read_parallel → <std::sys::thread::unix::Thread>::new → pthread_create@@GLIBC_2.34 → 0x0000000004a77750 → _dl_allocate_tls_init → memset |
| 8 | 1 | cached_part_profile::cached_part_region_0787 → litchi_opc::source_backed::batch::read_parts_ordered → litchi_opc::source_backed::batch::read_parallel → <std::sys::thread::unix::Thread>::new → pthread_create@@GLIBC_2.34 → 0x0000000004a77750 → _dl_allocate_tls_init → memset |
| 1 | 1 | cached_part_profile::cached_part_region_0787 → litchi_opc::source_backed::batch::read_parts_ordered → litchi_opc::source_backed::batch::read_serial → litchi_opc::source_backed::SourceBackedPackage::read_part_with_observer_and_capture → core::hash::BuildHasher::hash_one → <core::hash::sip::Hasher<S> as core::hash::Hasher>::write |

## Thread syscall counts

Trace wall time is deliberately absent from this table and is never pooled with profile or native results.

| Lane | Repeat | Width | State | clone attempts/errors/successes | clone3 attempts/errors/successes | futex attempts/errors/successes |
|---|---:|---:|---|---|---|---|
| before | 0 | 1 | fresh | 0/0/0 | 0/0/0 | 32/0/32 |
| before | 0 | 1 | primed | 0/0/0 | 0/0/0 | 32/0/32 |
| before | 0 | 8 | fresh | 0/0/0 | 8/0/8 | 232/17/215 |
| before | 0 | 8 | primed | 0/0/0 | 16/0/16 | 338/36/302 |
| before | 0 | 32 | fresh | 0/0/0 | 32/0/32 | 66/6/60 |
| before | 0 | 32 | primed | 0/0/0 | 64/0/64 | 66/3/63 |
| before | 1 | 32 | fresh | 0/0/0 | 32/0/32 | 57/4/53 |
| before | 1 | 32 | primed | 0/0/0 | 64/0/64 | 59/4/55 |
| before | 1 | 8 | fresh | 0/0/0 | 8/0/8 | 247/29/218 |
| before | 1 | 8 | primed | 0/0/0 | 16/0/16 | 334/39/295 |
| before | 1 | 1 | fresh | 0/0/0 | 0/0/0 | 32/0/32 |
| before | 1 | 1 | primed | 0/0/0 | 0/0/0 | 32/0/32 |
| after | 0 | 1 | fresh | 0/0/0 | 0/0/0 | 32/0/32 |
| after | 0 | 1 | primed | 0/0/0 | 0/0/0 | 32/0/32 |
| after | 0 | 8 | fresh | 0/0/0 | 8/0/8 | 226/16/210 |
| after | 0 | 8 | primed | 0/0/0 | 8/0/8 | 226/20/206 |
| after | 0 | 32 | fresh | 0/0/0 | 32/0/32 | 79/6/73 |
| after | 0 | 32 | primed | 0/0/0 | 32/0/32 | 53/4/49 |
| after | 1 | 32 | fresh | 0/0/0 | 32/0/32 | 51/3/48 |
| after | 1 | 32 | primed | 0/0/0 | 32/0/32 | 76/6/70 |
| after | 1 | 8 | fresh | 0/0/0 | 8/0/8 | 239/29/210 |
| after | 1 | 8 | primed | 0/0/0 | 8/0/8 | 246/32/214 |
| after | 1 | 1 | fresh | 0/0/0 | 0/0/0 | 32/0/32 |
| after | 1 | 1 | primed | 0/0/0 | 0/0/0 | 32/0/32 |

For the before lane, fresh one-operation clone3 successes are W (W>1) and primed preload-plus-operation successes are 2W; W=1 is zero. The optional after lane expects W in both states because only the preload creates the extra worker set.
Each fresh/primed pair retains source calls 64/0 and matching output sequence identity.

## Custody and limits

All packet artifacts are checked by recorded size and SHA-256. Profile executables are accepted only while live or when an exact verified cleanup witness binds basename, size, and hash. The initial baseline profile attempt remains disclosed and unqualified; the compatibility profile is the qualified Callgrind lane.
