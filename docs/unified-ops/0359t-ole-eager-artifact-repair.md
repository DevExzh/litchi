# 0359t — W5 eager XLS artifact ownership repair

- Date: 2026-10-22.
- Base: `8d1199f45e` with the W5 OLE source-owner changes already applied.
- Status: the eager XLS allocation residual is repaired; native timing remains
  unqualified.

The eager cell snapshot now keeps the captured `Arc<Vec<u8>>` as its immutable
artifact.  The packed `Arc<[u8]>` conversion allocated a second complete
artifact before the eager package publisher reopened its reader.  Retaining
the vector removes that copy while preserving the source-backed handoff,
source-version fencing, exact patch endpoints, and clone isolation.  The
source-backed path continues to pass the same vector allocation to its shared
CFB owner.

## Paired allocator receipt

The reduced receipt is outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w5-ole-source-arc-retag-20261022/counting-arcvec-current-20261022`.
Its `MANIFEST.txt` is 3,691 bytes with SHA-256
`0c81f2f4963b2b63846868bc1d63fc45941ea0cb3ea9f9899c3814be77e5eee0`.
The preserved runner SHA-256 is
`70d07e3b07bce4237536f1239be66aed23621a19b999d79b2230abd9f3c50b93`; the
probe source SHA-256 is
`ad4c52c01333fbc13c7b3cca18e2c0a0d2ccf89d5c07972e2e453fa9bd8d713a`.
The matched baseline binary is
`23a7088c3f9a6563adc3ae1de644a0d2b05fe91604824afd7517aa426c981fc0`, and
the candidate binary is
`5e189d50cb392cc1cc6f46489fedbcae02101a14ac4e44e645d70f0a000f3e32`.
The candidate was built in release mode with one Cargo job and the capped
offline workspace dependencies.

All paired routes emitted the same output extents and exact output checks. The
allocator used `--warmup 0 --samples 1 --batch 1` and retained 129 packages
and 128 patches for each history route:

| route | baseline calls / bytes | candidate calls / bytes | candidate peak bytes |
| --- | ---: | ---: | ---: |
| read | 45 / 1,192 | 45 / 1,192 | equal |
| edit | 33 / 3,523 | 33 / 3,523 | equal |
| XLS cells | 1,589 / 2,361,904 | 1,577 / 2,136,561 | 2 below baseline |
| XLS cells stage | 1,588 / 2,249,264 | 1,576 / 2,023,921 | 2 below baseline |
| DOC history | 1,609,900 / 1,830,103,867 | 1,571,008 / 1,770,417,748 | equal |
| XLS history | 387,769 / 441,765,854 | 385,453 / 398,268,772 | +56 bytes |

The XLS history retained live bytes are 24 bytes above the baseline while its
peak is 56 bytes above; requested bytes fall by 43,497,082. This is retained
source-owner metadata, not the eager artifact copy, and remains visible in the
receipt. Native timing was not run by this allocator harness.
