# 0822 PPTX edit profile probe review

Status: static review pass. This review covers the packet-local probe source at
base revision 353aa00a7d and records the checks that must be repeated by the
root build and capture owner. No Cargo command, formatter, release build,
workload, or profiler was run while preparing this packet.

The probe is confined to docs/performance/results/change-0822/probe-src/. It
has its own manifest and a copied 0811 lockfile whose root package name is
deterministically changed to pptx-edit-profile-0822. It has no production
crate or performance-tool edits. The dependency set is the requested
litchi-pptx, serde, serde_json, and sha2 set.

The measured operation is the Pptx arm of
tools/perf-baseline/src/ordinary_save.rs at the admitted 353aa00a7d base:

    Package::open(input)                         outside the clock
      opened_presentation_transaction()
      set_shape_text(0, 0, "litchi-perf-0638-ordinary-save")
      commit()
      apply_opened_presentation_commit(commit)   inside the clock
    Package::to_bytes(), drop, hashing, reopen, full readback
                                                 outside the clock

edit_helper_0822 is the one #[inline(always)] operation body. Direct mode calls
it directly. Wrapped mode calls #[inline(never)] edit_region_0822, which
invokes the same helper and applies black_box(&result) inside the wrapper. The
helper returns Result<()>, and the returned
apply_opened_presentation_commit snapshot is discarded at the statement
semicolon inside that helper. This keeps the owner and snapshot lifetimes
aligned with the ordinary-save owner edit and keeps package destruction out of
the measured interval.

The CLI requires --input, --reference, --samples, --warmup, and
--mode direct|wrapped; --output is optional and supports a path, -, or stdout.
Samples are bounded to 1..=10000, warmups to 0..=100, unknown arguments are
rejected, and both input files are bounded to 32 MiB before reading. The input
is pinned at 68,822 bytes and
19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571. The
reference must be the sealed 0821 real PPTX default output at 68,284 bytes and
38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf.

The reference is opened by a separate owner before samples begin. Its complete
presentation text, slide count, and exact (slide 0, shape 0) text are the
semantic oracle. The target text must equal the marker. Every warmup and
measured iteration reopens the produced bytes, checks the exact output hash
and size, compares the complete output byte vector with the admitted reference,
checks the exact complete text and digest, checks slide count, and checks the
target shape text. The report carries ordered raw elapsed samples,
the input/reference/output identities, target and marker, full-text digest,
slide count, per-sample verification flags, and aggregate all_verified.

Focused tests cover CLI rejection of unknown and out-of-bound values, direct
and wrapped byte and semantic parity against the pinned real fixture, and
wrong source/reference identity rejection. The parity test intentionally uses
the sealed 0821 exported real-002-pptx/default.pptx reference and therefore
must run before any cleanup that removes that retained oracle.

Root-owned validation should first run formatting, offline locked checks,
focused tests, and the ordinary full quality gates. The root owner may then
build plain and frame-pointer release binaries and capture profiles. This
probe itself makes no save, fsync, durability, filesystem, or optimization
claim; its elapsed vector is limited to the named in-memory opened
presentation edit transaction.
