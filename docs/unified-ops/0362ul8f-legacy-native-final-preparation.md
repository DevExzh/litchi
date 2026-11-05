# 0362ul8f — final legacy native source preparation at 472b

Date: 2026-10-24.  This record covers source preparation and provenance
validation only.  No Cargo build, native execution, timing, allocation
capture, or acceptance result is claimed.

## Frozen endpoints and tooling

The production candidate is the immutable external worktree
`/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source`
at `472b6705475b073e7fe08f6e11ee17fb1e29f838`.  The repaired preparation,
runner, and validator are taken from the integration checkout at
`93bdf16bf4c7a1e28205f789cfa500e6fb946166` (which contains the
manifest-identity repair `2b01fc91d7a5fff69298f9ce058f8adff85d4a1c`):

```text
tools/native_qualification/prepare_legacy_native_inputs.py  3aad0d288aa653643d2efa2782b8ac5d96ad82fbdc22a07d57cb9fa1b47871bf
tools/native_qualification/legacy_native_validator.py        9f2d2f7a4eac2eb51b8918b3e8e8d0670cec6d5c20ad25b30bf4010f80593dc7
tools/native_qualification_runner.py                         0864679291b3a2aaac93c6ae324da5e3a9c57506328efcbf3eac39fbb64d3013
```

The fresh packet is outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/legacy-native-repaired-472b-2b01`.
It has separate absolute manifests and source-preparation receipts for RTF,
OLE-properties, and XLS-properties.  The RTF and XLS manifests initially had
an incorrect candidate-root prefix; the runner rejected that path before any
execution.  The manifests were corrected, the superseded receipts were
retained outside the repository, and the corrected receipts were regenerated.

## Prepared packet identities

The table records the final manifest and preparation receipt hashes.  Every
row passed both source legs (`before` and `after`) through the repaired runner
identity checks and the reviewed-input/preparation checks in
`legacy_native_validator.py`.

| lane | baseline / candidate heads | manifest (bytes, SHA-256) | preparation receipt (bytes, SHA-256) | source-tree SHA-256 (`before` / `after`) |
| --- | --- | --- | --- | --- |
| `legacy-rtf` | `86d708878fa537eab42fd4bd5ffbc0f79c606191` / `472b6705475b073e7fe08f6e11ee17fb1e29f838` | 4838, `6cdbd2158793d867218a9a5c85847b0140804c5c28dde6ef577125a17c5d667b` | 7414, `9c4309a7aa574ea5ee7ff58d76f91b8123ea4bf769463b39e994027eab5f8d8f` | `5089da679079ee37dc9964669e9ac726483216dd87f3faac74b260466be99551` / `9e713180f74f0515b807811a8c30ae166f43c05b4292ea9fc3824758055bb08e` |
| `legacy-ole-properties` | `604822b44107999aaddb0bffb96b705e9464f546` / `472b6705475b073e7fe08f6e11ee17fb1e29f838` | 6238, `6ef65bfa960f4869db7a511bd0d0b55604e2001df4f767c564792e7ee00349be` | 7836, `469206fa1a670fc9c962f9c1e14dd4cc127c8d4b3eef2a8a111a11e11f5e7ffa` | `476819c990821549134e7029a1adca40369bbbc8e5427c108124f1aba718316a` / `14ede31afe5866b76b7f186b2e241128b667257cb0a5b23a75d8318aff0baf0d` |
| `legacy-xls-properties` | `509bf28e7531ad75f278eaec3a732714a243915a` / `472b6705475b073e7fe08f6e11ee17fb1e29f838` | 4816, `3495667fdad6f8d9a02929357c376df4574e2448ac2b080316a4d7964f6da2bd` | 7748, `3db30953ab669a2d7606ec286f2d092d9caaee1e0aca1c86700ea6fa0f2d3b38` | `0234d770ac94bd3b5549859035617b7ce56925b901161ba847ab10674cdad602` / `9e713180f74f0515b807811a8c30ae166f43c05b4292ea9fc3824758055bb08e` |

The RTF baseline is the frozen 86d adapter worktree with the reviewed
`tools/perf-baseline` files allowed dirty; its lock is
`57c2a19ffeff395fdea95c33f295c85b08400b00e2fd9cec94d66cff56d1583f` and its
probe is `0613f279db17121711b140c273137465c23670033a4dda5392444703b01b1f81`.
The candidate lock is `eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`
and retains that probe hash.  The historical failed current-probe compile
remains external and immutable (source `70077fcec439fda2fdf9e9635f5db2f1f7c86ed31f3f972a95c5e5e3a607a585`,
log `119d4216675ee79a7e89e17579f019dd4fe3a6d43c08a8bd22d87c37622fd103`);
it is historical evidence, not a live source pin.

The OLE-properties baseline lock/probe hashes are
`cbbce0bca61d6a28776f67da264196b9d80b20f1f1e85cf958ffee032beeae4e` /
`6890d0524cf4466a96812eb9ba449d4d995b29e1db74888fe42b5a42e8a91aed`.
Its candidate lock is the 472b lock above and its reviewed consumer adapter
is `e031d0adbc00a5784658c37472693ae1fe840ad88f84a5b785b48dd77a810782`,
derived from raw fence `fdb139197d08677d6fa0e14d536e798e380e4cb1` and raw
probe hash `3b7cde72762a74796250df572704206baf9e438cca3d9ed11119af783fc8488f`.
The adapter only supplies the required `OleFileLimits::default()` argument
and preserves the reviewed owner/property feature split.

The XLS-properties baseline lock/probe hashes are
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da` /
`b840acb620c09644ca4b89cbe42927485da9d457f3bf93b578eff90fa1008173`.
The candidate lock/probe hashes are
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` /
`53b26136e915fa0bf88518fd1c05e494d55ce5e8f91012be04ac7adf476c1f10`.
The three reviewed fixture hashes are unchanged on both legs:

```text
test-data/ole/doc/FancyFoot.doc  709fe8b695a9b12be35c7e81d93beb95afc8dc36422dfc0cbd605a3057d4ddc2
test-data/ole/xls/Simple.xls     2b5ec8298bacffdca61c50f6f7607da0a4e63cf76fa011c62f203e6a66bc072a
test-data/ole/ppt/SampleShow.ppt 64275a6685b8f267827178037dcc59de927475e56385725916f4c6039c7ce7f6
```

## Source-only commands and checks

Preparation used the repaired script with the absolute packet manifests and
the immutable source endpoints:

```text
PYTHONDONTWRITEBYTECODE=1 python3 /home/zhuhe/code/litchi-worktrees/uo2-coordinator-metadata-integration/tools/native_qualification/prepare_legacy_native_inputs.py --lane legacy-rtf --manifest /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/legacy-native-repaired-472b-2b01/rtf/driver/legacy-rtf-final.json --before /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/legacy-native-repaired-472b-2b01/rtf/source-baseline --after /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source --output /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/legacy-native-repaired-472b-2b01/rtf/driver/source-preparation.json
PYTHONDONTWRITEBYTECODE=1 python3 /home/zhuhe/code/litchi-worktrees/uo2-coordinator-metadata-integration/tools/native_qualification/prepare_legacy_native_inputs.py --lane legacy-ole-properties --manifest /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/legacy-native-repaired-472b-2b01/ole/driver/legacy-ole-properties-final.json --before /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/legacy-native-repaired-472b-2b01/ole/source-baseline --after /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/legacy-native-repaired-472b-2b01/ole/source-candidate --output /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/legacy-native-repaired-472b-2b01/ole/driver/source-preparation.json
PYTHONDONTWRITEBYTECODE=1 python3 /home/zhuhe/code/litchi-worktrees/uo2-coordinator-metadata-integration/tools/native_qualification/prepare_legacy_native_inputs.py --lane legacy-xls-properties --manifest /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/legacy-native-repaired-472b-2b01/xls/driver/legacy-xls-properties-final.json --before /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/legacy-native-repaired-472b-2b01/xls/source-baseline --after /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source --output /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/legacy-native-repaired-472b-2b01/xls/driver/source-preparation.json
```

The follow-up source-only check loaded each manifest and receipt, called the
runner's source-identity and preparation-identity guards for both legs, then
called the validator's reviewed-input and preparation-receipt guards.  Result:
3/3 manifests loaded, 6/6 source legs accepted, and 3/3 preparation receipts
accepted.  All packet `output/` and `target/` directories were absent after
the check; no build helper was invoked.

When a build slot is granted, the recorded command shape is the following for
each absolute manifest, with `build` followed by `abba`:

```text
PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH
unset LITCHI_CAPPED LITCHI_BLESS
LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 \
  python3 /home/zhuhe/code/litchi-worktrees/uo2-coordinator-metadata-integration/tools/native_qualification_runner.py \
  --manifest /absolute/packet/{rtf,ole,xls}/driver/{lane}-final.json --phase build
```

The repaired runner supplies the reviewed offline/locked Cargo flags and
fixed 20,000,000,000-byte target and memory caps.  This record does not claim
those future phases were started.  Historical packets, endpoints, and failed
compile evidence remain untouched.
