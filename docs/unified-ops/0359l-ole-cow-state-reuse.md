# 0359l — W5 COW state reuse correction

- Date: 2026-10-21.
- Candidate source: `a84650cd70390641f7734dc1f3362bcb9cd1e025`.
- Status: allocation correction complete; native latency qualification remains
  pending and no performance exception is accepted.

The changed `Edit::commit` path now consumes the edit and normalizes its
`Arc<PackageState>` with `Arc::make_mut`. A uniquely owned state is reused for
the published package; a state still shared with a package or cloned editor
detaches before normalization. The published state resets its original and
current bytes to the committed content, rebases the captured package as its
base, clears `changed`, and records the validated layout. The borrowed
`staged()` path remains unchanged. This preserves clone isolation, exact
forward/inverse replay, and no-op state sharing while removing one state
allocation per changed publication.

The focused object regression covers a sibling package that must remain
unchanged, forward patch replay on the original package, and a follow-up
no-op edit on the published package. The complete litchi-ole-common tests and
all-target performance-diagnostics Clippy pass after this correction.

## Allocation receipt

The paired counting-allocator receipt is outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w5-ole/counting-cow-ad990-inplace-20261021`.
Its 14,592-byte directory has manifest SHA-256
`b143a9927bdc0030ca09c62d1eb2a5b53978e06039d16969663165fb2bd6af80`.
The runner SHA-256 is
`70d07e3b07bce4237536f1239be66aed23621a19b999d79b2230abd9f3c50b93`.
The identical probe source SHA-256 is
`ad4c52c01333fbc13c7b3cca18e2c0a0d2ccf89d5c07972e2e453fa9bd8d713a`, and
the lock SHA-256 is
`bc5a133f0668142d18d637b2b196097e855c8890697cb3bbf43aeb937145f04a`.
The baseline binary SHA-256 is
`23a7088c3f9a6563adc3ae1de644a0d2b05fe91604824afd7517aa426c981fc0`; the
corrected candidate binary SHA-256 is
`9846f15d27c0942421e254379d71aa40ed2fcfb27fd24e5f6ee8bc1e1dbe43cc`.
The candidate was built with one capped Cargo job from the current source and
the matched lock; the baseline uses the immutable `baseline-0ce4-src-8f0ffa`
source snapshot. The exact build command was:

```text
source /home/zhuhe/code/litchi-worktrees/.coord/w5-ole/env.sh
export CARGO_BUILD_JOBS=1
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/w5-ole/target-alloc-cow-current3
cargo build --release --offline --locked --manifest-path /home/zhuhe/code/litchi-worktrees/.coord/w5-ole/matched-harness/current/Cargo.toml --examples --features current-api,w5-counting-allocator
```

The runner reports allocation counters only; it does not report native
latency. Against the pre-wave baseline, read and edit remain identical. XLS
cells and the stage route remain 1,589 to 1,587 calls, with requested bytes
2,361,904 to 2,362,341 (+437, +0.0185%) and peak bytes +94. XLS history
remains 387,769 to 387,383 calls, with requested bytes 441,765,854 to
441,833,100 (+67,246, +0.0152%); retained live and peak bytes are unchanged.
Compared with the prior COW candidate, the extra 152-byte state allocation per
publication and the 128-step 19,456-byte history increase are gone.

The final untimed native pair is protected outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w5-ole/final-probes-a84650cd70`.
Its manifest SHA-256 is
`92e3b027ac33e58cadbb1d5087990336cabbfd94758487475b20f09405c989b6`.
The protected current binaries are `ole_w5_probe` SHA-256
`cb4b8fdde10d30fa1c2df44c14f4e9c9a235d4dccd2ad647d71a9cf7851d51bb` and
`ole_package_clone_probe` SHA-256
`ed1a6b54e496e64951882f46f55dee429c33989839ac01ba0f8182fb4656c2b7`.
The matched baseline binaries are retained there with SHA-256 values
`03e74e97525628b281e3f9d58664301f53bdd26841fe5af3c8cc3ebf1f586328` and
`bc1eaaa0e99c5607d9a2a2486fb44a1a3442a7e5637ffbf3f73ca8d040c67627`.

Native qualification still must cover the owner clone and all eleven OLE
selectors with exact output/replay checks, host-overlap checks, and accepted
batched legs of at least 250 ms.
