# DOCX `stylesWithEffects` bounded profile

This directory retains the matched profile captured from production source
`d1f299d00e0dd5cc5cd8ddf9811c4b1ad21d1119` with scaffold checkout
`96958498f88e263b3bc7f0a9e544f2144ff156c5`. The runner and independent root
replay both passed. The retained raw capture is 307 files and 10,855,027 bytes;
`root-retention-verification.json` records every raw-file size and SHA-256.

The bounded matrix has 31 lane/scale rows, three fresh process indexes, two
warmups, and twenty measured samples per process, for 93 JSON receipts and
1,860 samples. It covers native, 64 KiB, and 1 MiB resource classes. Cap,
malformed-input, stale-patch, signed-change, and other refusal lanes remain in
the correctness smoke and were not timed. No 8 MiB or near-limit timing was
run.

Preparation opens and serializes the package, hashes members, and constructs
resources before the operation clock. Measured phase receipts separate
capture/snapshot, staging, commit, publication, reopen, inverse reopen and
inverse, projection, opaque checks, graph accounting, readback, and validation.
The prepared package and setup remain live through elapsed and allocator
snapshots. Allocator counters describe the Rust allocation window;
`requested_alloc_bytes`, live bytes, and `peak_live_delta` are distinct from
process RSS. RSS is the maximum from `/usr/bin/time -v` for each process
invocation, which includes its warmups and samples.

The largest absolute medians were replacement at 1 MiB: 183.707 ms for main
and 188.024 ms for glossary. Publication contributed 122.336 ms and 125.259
ms, followed by reopen at 28.968 ms and 29.639 ms. Synthetic 1 MiB capture
medians were 30.865 ms (main) and 31.420 ms (glossary), dominated by capture
at 29.022 ms and 29.585 ms. From 64 KiB to 1 MiB, the observed median ratios
were 7.27x/6.33x for capture, 8.04x/7.08x for projection, and 8.44x/7.36x
for replacement (main/glossary). These ratios are observations for these
synthetic fixtures only; they are not a general or asymptotic scaling result,
production workload estimate, or optimization comparison.

Native process maximum RSS medians were roughly 6,304–7,356 KiB. At 1 MiB
they were 11,880–12,860 KiB for capture/projection and 18,060–18,684 KiB for
replace.
For main replacement, median requested allocator bytes rose from 24,061,974
at 64 KiB to 153,146,975 at 1 MiB; median peak live delta rose from 799,352
to 5,864,443 bytes. These allocator values must not be read as RSS or as a
managed process-memory cap.

The host was Ubuntu 26.04.1 on an AMD EPYC 9R45 with 32 logical CPUs and
129,447,068 KiB total memory, with affinity `0-31`. One-minute load was 1.48
before and 2.54 after capture. Physical core count, mount identity, and an
unrelated-process census were not captured. The host was shared and
unisolated; there is no idle-host or uncontended-timing claim.

Source manifests and Cargo metadata matched before and after. Key retained
receipt hashes are:

| artifact | SHA-256 |
| --- | --- |
| `verification.json` / `root-verification.json` | `e24a6edce1bba059dea50536838930fc929d09c2eaecaa8edfc488179b5adb70` |
| `source-manifest-before.txt` / `source-manifest-after.txt` | `192d3e257c914bd975b5da2b0c7c0a6e4f149f23724c53598a69992f481e8a30` |
| `metadata-before.json` / `metadata-after.json` | `6c52653b0889042afa34a97b1af5927249243c0c5e32f6c9aecdc998f85304d4` |
| `binary.sha256` / `binary-after.sha256` | `88dffb83a3cde39c414193f5d41afb39b74521dc33277a8a79a8de0fb2ee4615` |
| `generated-fixtures.json` | `5f476ca611475de61caeb21af1a93ac395612f1add31d69995d991f909ee1332` |
| `source-provenance.txt` | `624f5d36389f2adc9830fae6201e7df0eec1c219e0a3b107b20aaded8e67ded3` |

The disposable Cargo target was removed only after the matching verification
sentinel. This evidence makes no before/after speedup claim and does not claim
native Word rendering or production acceptance.

## Replay layout

The verifier requires the exact captured checkout at `96958498f` and the
recorded external results path because `generated-fixtures.json` contains
absolute generated-resource and package paths. The retained repository result
directory cannot be replayed directly. To replay, recreate
`/var/tmp/litchi-docx-styles-effects-profile-results-96958498f-20260912-a`,
copy the 307 paths listed by `root-retention-verification.json` byte-for-byte
(including `generated-fixtures/`), and preserve those absolute paths. Use the
captured clean checkout and run the verifier with all manifest/metadata inputs
and its output under that external results directory; the recorded disposable
target path may remain absent after successful cleanup. Do not rewrite the raw
receipts or substitute the retained repository path for the external results
path.

For durable captured-checkout reproduction after the original worktree is
gone, use `captured-source.bundle` in this directory (SHA-256
`508e0c487a1bb44c617e3ffcbfb6c576987c5c15691ac7d34d8829a8a6c4266d`). It is
separate from the 307 raw receipts. In a fresh clone whose history contains
the scaffold ancestor `35d069e62a2b9a72e59f6a4aa73444b928ff6195`, verify and
fetch the captured branch, then create a clean worktree:

```sh
git bundle verify /path/to/profile-clean-96958498f/captured-source.bundle
git fetch /path/to/profile-clean-96958498f/captured-source.bundle \
  refs/heads/perf/docx-styles-effects-profile-capture-20260912:refs/heads/perf/docx-styles-effects-profile-capture-20260912
git worktree add --detach /var/tmp/litchi-docx-styles-effects-profile-capture-20260912-a \
  96958498f88e263b3bc7f0a9e544f2144ff156c5
git -C /var/tmp/litchi-docx-styles-effects-profile-capture-20260912-a status --short
```

Bundle verification must report the required
`35d069e62a2b9a72e59f6a4aa73444b928ff6195` ancestor, and the fetched branch
must resolve to the captured `96958498f` commit. Replay still requires
restoring the raw receipts at the recorded external results path as described
above; the bundle does not replace those receipts.

After restoring the listed raw files and checking each size/hash against
`root-retention-verification.json`, replay without building or timing:

```sh
profile_checkout=/var/tmp/litchi-docx-styles-effects-profile-capture-20260912-a
profile_evidence="$profile_checkout/docs/report/spec-gap-validation-evidence/docx-styles-effects-performance"
profile_results=/var/tmp/litchi-docx-styles-effects-profile-results-96958498f-20260912-a
python3 -B "$profile_evidence/verify_profile.py" \
  --root "$profile_checkout" --evidence "$profile_evidence" \
  --results "$profile_results" \
  --manifest-before "$profile_results/source-manifest-before.txt" \
  --manifest-after "$profile_results/source-manifest-after.txt" \
  --metadata-before "$profile_results/metadata-before.json" \
  --metadata-after "$profile_results/metadata-after.json" \
  --output "$profile_results/replayed-verification.json"
```

The captured checkout path is also exact: Cargo metadata records absolute
source paths. Replay output is a new auxiliary receipt, not a replacement for
`verification.json` or any original raw file.
