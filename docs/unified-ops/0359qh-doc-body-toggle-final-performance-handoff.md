# 0359qh — DOC body-toggle final allocation/native handoff

- **Status:** source-only handoff verified; no Cargo, native timing, Valgrind,
  Callgrind, Massif, or ABBA command was run in this turn.
- **Date:** 2026-10-24.
- **Scope:** final-472b DOC body-toggle candidate and reviewed two-file
  baseline, for the queued allocation and quiet-host native qualification.
  This record does not accept performance.

## Accepted source and binary packet

The accepted correctness packet is 0359qg. Its complete external receipt is:

~~~text
(deleted-coord-artifact)/actual-receipt.txt
SHA-256 a0dfb1d1b66c16249939a702d26565466a3ca8b86d11e9bc87461bf5df414b3a
~~~

| item | candidate | baseline |
|---|---|---|
| source | `(deleted-coord-artifact)` | `(deleted-coord-artifact)` |
| commit | `472b6705475b073e7fe08f6e11ee17fb1e29f838` | same |
| tree | `ebc1cee05f124c0c63977aff07a23288c19e7d00` | same |
| Cargo.lock SHA-256 | `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` | same |
| body_text.rs SHA-256 | `db01cfce38c13c6c806382ffede990fda2b3bbe4dd53f59c4c5cb9feba9777d3` | `6dc26da8130659c212dd29f26499cdaba7c4849268583b44dfcd9413a36887a3` |
| tracked revision SHA-256 | `029721d34e7ad658e0064afd258cdb2b5a40e330110255b157aafb83f0d7b12a` | `b7155dc7c9aae6e318845b5073da4c0f275889125d38eae61df091aa2a2166df` |
| probe source SHA-256 | `b808e4dcba33dcafd4dff7dd22dc250e4170f2613839044fa8097263f5020965` | same |
| applied baseline overlay SHA-256 | `0e0bf3c606277da07e53c085fa52a2e1db99839a3fc8ca5e869154ae60ab59c0` | — |

Reuse only these accepted binaries, without rebuilding:

| leg | path under `doc-toggle-472b-build-20261024` | bytes | SHA-256 |
|---|---|---:|---|
| candidate | `target-candidate/release/examples/legacy_w5_source_probe` | 13,617,192 | `ba7042fd413c26764986b0ed9dcd79beca85e91231f0beb2f850550152e14877` |
| baseline | `target-baseline/release/examples/legacy_w5_source_probe` | 13,625,656 | `2d62dacae12b49214a086c109b506ca8eb69a33cf5e770627794efecdbcba810` |

## Route and output contract

The probe's synthetic fixture is 1,000 generated DOC paragraphs with fixed
text `Paragraph {index:05} holds a few ordinary words of body text.`.
`doc/leaf_write` toggles paragraph 500, commits, and calls `Package::write_to`.
`doc/batch1k_leaf_write` sets paragraphs 0 through 999 under a 1,000-operation
`TransactionLimits`, commits, and calls `Package::write_to`.

Use `LITCHI_PROBE_ONLY=<route>` and `LITCHI_PROBE_BATCH_ROUNDS=1` for profile
legs. The selected route computes expected bytes and asserts every call; both
expected-output blocks are inside their selector guards, so the ordinary
profile does not seed the batch route. Do not set `LITCHI_PROBE_OUTPUT_DIR`
during timing or profile runs; it is only the accepted correctness dump path.

Expected per-run JSON output hashes are:

| route | candidate | baseline |
|---|---|---|
| `doc/leaf_write` | `6bb97058e47f747c29e568395eea01db1226aa5ab8962087e1d78a1b9b82ff11` | same |
| `doc/batch1k_leaf_write` | `bd016a87f5d0015d6f8b910862eeeef9c0ce490033b4951c32ee5f403b7eebdf` | `4a5bcc2888ab819488fbb5a1ced1b088010440bd8293f10556e0903993e448d3` |

The batch hashes intentionally differ; the final CHPX oracle already passed.

## Existing evidence and why it cannot close final-472b

The previous actual deterministic packet is:

~~~text
(deleted-coord-artifact)/qualification.json
5912 bytes, SHA-256 086714a6f2190fafdf8ac7e14a22a6a7a844305e979d9ad43d48384b7c133279
~~~

It contains real Callgrind/Massif artifact hashes, but is bound to integration
`3cad59be707c7431bc1a8c115e1ff534a048dc57`, probe SHA
`9e1cd649268ad5eefe4314b76493f61cd047bf6ec4114c7e7872b2182b1af689`, and
older body/revision assemblies. Historical profile identities are retained
for provenance only:

~~~text
Callgrind ordinary candidate/baseline  4a39f0c91a5d4b07ac4899a927997ce5021e9a25991e9138b6ee5e7bca97899e / 35ca4d58257feee9851afa19c583290235383c1f7f7fb957d4d44f8dd7b28172
Callgrind batch candidate/baseline     3cdded22837d216881755a5ab928e8ee9618ca57dfa435447dd923a65ecae540 / 6989252603f52f758b2ca4ad169bb122ac8cce009e2e818a6b5d28fab6333d7a
Massif ordinary candidate/baseline     b29f0f124a7eff5574313136636b470d6362498a52754e889bcf82fe894e572a / 48c623ddc380cb93200854131f56adbd3e5bc5825a93f1b66a28342c0bf6fca2
Massif batch candidate/baseline        36384b97569813649b4c10c547c8af5928b8434ef88c31dfa89891f275c78b2b / f27144b3c9faa44baca9a64e887b03d0d3406fa8b14695c97f02bbd717e8760c
~~~

0359qe reports native `pending` because the host was not quiet. Earlier 0359qb
native receipts are loaded-host diagnostic rejects and use the older probe.
Final-472b therefore has no accepted allocation or native result yet.

## Later allocation phase (not run here)

Use a fresh external `RUN`, separate `TMPDIR` per leg/route, and the binaries
above. This exact function shape is to be run only after a profile slot is
granted:

~~~sh
source (deleted-coord-artifact)
ROOT=/home/zhuhe/code/litchi-worktrees
SOURCE_SHORT=472b
RUN=(deleted-coord-artifact)
BASE_RUN=(deleted-coord-artifact)
CAND_BIN=$BASE_RUN/target-candidate/release/examples/legacy_w5_source_probe
BASE_BIN=$BASE_RUN/target-baseline/release/examples/legacy_w5_source_probe
mkdir -p "$RUN"/allocation/{candidate,baseline} "$RUN"/tmp/{candidate,baseline}
profile() {
  leg=$1; bin=$2; route=$3; name=${route//\//_}
  mkdir -p "$RUN/tmp/$leg/$name"
  env LITCHI_PROBE_ONLY="$route" LITCHI_PROBE_BATCH_ROUNDS=1 \
    TMPDIR="$RUN/tmp/$leg/$name" valgrind --tool=callgrind --quiet \
    --callgrind-out-file="$RUN/allocation/$leg/$name.callgrind" "$bin" 1 \
    >"$RUN/allocation/$leg/$name.json" \
    2>"$RUN/allocation/$leg/$name.callgrind.log"
  env LITCHI_PROBE_ONLY="$route" LITCHI_PROBE_BATCH_ROUNDS=1 \
    TMPDIR="$RUN/tmp/$leg/$name" valgrind --tool=massif --quiet \
    --pages-as-heap=no --massif-out-file="$RUN/allocation/$leg/$name.massif" "$bin" 1 \
    >"$RUN/allocation/$leg/$name.massif.json" \
    2>"$RUN/allocation/$leg/$name.massif.log"
  callgrind_annotate --auto=no --inclusive=yes --tree=calling \
    --show-percs=no --threshold=99 "$RUN/allocation/$leg/$name.callgrind" \
    >"$RUN/allocation/$leg/$name.annotate"
  ms_print "$RUN/allocation/$leg/$name.massif" \
    >"$RUN/allocation/$leg/$name.massif.txt"
}
profile candidate "$CAND_BIN" doc/leaf_write
profile candidate "$CAND_BIN" doc/batch1k_leaf_write
profile baseline "$BASE_BIN" doc/leaf_write
profile baseline "$BASE_BIN" doc/batch1k_leaf_write
sha256sum "$RUN"/allocation/*/*.json "$RUN"/allocation/*/*callgrind \
  "$RUN"/allocation/*/*massif >"$RUN/allocation.sha256"
~~~

The review must report route-local attribution for both routes, bind every
JSON/profile hash to the source/binary/lock identities, and never mix the old
3cad profile hashes into the final packet.

## Later quiet-host native ABBA phase (not run here)

After a quiet host is granted, set and record one quiet CPU in `CPU`; do not
assume historical CPU 29 is quiet. Run baseline–candidate–candidate–baseline
independently for each route with fresh output/TMP:

~~~sh
CPU="${CPU:?set the quiet CPU recorded in the receipt}"
mkdir -p "$RUN/native" "$RUN/native-tmp"
native() {
  route=$1; name=${route//\//_}; ordinal=$2; leg=$3; bin=$4
  mkdir -p "$RUN/native-tmp/$name-$ordinal"
  systemd-run --user --scope --quiet --collect \
    -p OOMPolicy=continue -p MemoryMax=20G -p MemorySwapMax=0 -- \
    /usr/bin/time -f "$leg $name run=$ordinal maxrss_kb=%M wall=%e" \
    taskset -c "$CPU" env LITCHI_PROBE_ONLY="$route" \
    LITCHI_PROBE_BATCH_ROUNDS=3 TMPDIR="$RUN/native-tmp/$name-$ordinal" "$bin" 1 \
    >"$RUN/native/$name-$ordinal-$leg.json" \
    2>"$RUN/native/$name-$ordinal-$leg.time.log"
}
native doc/leaf_write 1 baseline "$BASE_BIN"
native doc/leaf_write 2 candidate "$CAND_BIN"
native doc/leaf_write 3 candidate "$CAND_BIN"
native doc/leaf_write 4 baseline "$BASE_BIN"
native doc/batch1k_leaf_write 1 baseline "$BASE_BIN"
native doc/batch1k_leaf_write 2 candidate "$CAND_BIN"
native doc/batch1k_leaf_write 3 candidate "$CAND_BIN"
native doc/batch1k_leaf_write 4 baseline "$BASE_BIN"
sha256sum "$RUN"/native/*.json "$RUN"/native/*.log >"$RUN/native.sha256"
~~~

Every JSON must carry its expected route hash. The final receipt must bind
quiet-host evidence, CPU, cgroup memory settings, exact order, host-load
observation, and all JSON/log hashes. A loaded-host run is a reject, not an
acceptance result.

## Concrete gaps

1. Final-472b has no Callgrind/Massif capture bound to its accepted source and
   binary hashes; 0359qe cannot close that gap.
2. Final-472b has no quiet-host native ABBA for either route; old native
   attempts remain diagnostic rejects.
3. The queued-toggle requirement is `doc/batch1k_leaf_write`, not only the
   ordinary one-toggle control. Both routes need the same fixture, lock,
   output assertions, and fresh artifact roots.
4. No result may claim batch raw-byte equality: the accepted correctness
   packet records different compact/baseline OLE bytes and a passing CHPX
   reachability audit.
