# 0439: Existing ODP append lifecycle baseline

This opt-in harness addition measures owned ODP snapshot opening, one slide
append, commit, and sequential output of materialized committed bytes. It
changes no production Rust. See [measurements](measurements.md), the
[data path](data-path.md), [design](design.md), and
[validation history](validation-notes.md).

The frozen [protocol](protocol.json) fixes one after baseline: 12 reports,
360 retained samples, CPU 2, one worker, two opposing repeat orders, and
three source sizes. Two successful whole-process profiles are explicitly
selected by [lifecycle-contract.json](lifecycle-contract.json). The original
failed stat attempt remains, with a separately bound
[driver amendment](profile-driver-amendment.json).

Normal p50 is about 2.08, 85–86, and 171–175 ms for 64, 4,096, and 8,192 source
slides. The large allocator case peaks 39.9 MB above entry and retains 8.07 MB
at the endpoint because source and commit are still alive. These figures do
not establish a speedup, bounded commit memory, source-backed I/O, or streaming
save. Full-source content parsing and commit validation remain in scope.

Portable verification requires Python 3 and no original checkout, binaries,
Cargo, or profiler execution. From this directory run:

```sh
python3 -B lifecycle.py --stage final
python3 -B derive.py --check
python3 -B oracle-probes.py --check
python3 -B portable-probes.py --stage final
```

The lifecycle checks frozen oracle hashes, source and binary custody, all
formal report oracles, exact capture order and chronology, profiler commands
and artifacts, cleanup proof, deterministic compression, and the complete
SHA256 inventory. Portable probes first accept an unmodified copied bundle,
then reject 13 mutations in separate copies with refreshed inventories.
The report oracle separately rejects 14 corruptions. Replay receipts record
driver and inventory stability; adding a receipt requires resealing afterward.

Logs and profiler data are losslessly compressed with deterministic gzip.
[compression.json](compression.json) binds original and stored bytes;
[SHA256SUMS](SHA256SUMS) covers every other bundle file. The seven explicitly
owned temporary directories are removed only after passing precleanup proof.
The cleanup inventory records preserved build-cache directory identities and
the unchanged user-owned GOAL.md digest. Absolute historical paths in receipts
are identities; default portable validation does not require them to exist.

All CPU work was serialized by the root. Source-only reviewers supplied
external drafts. The new selector brings the registry to 436, while the default
36-case matrix remains unchanged. This is one owned logical-append slice of
the ongoing non-iWork goal.
