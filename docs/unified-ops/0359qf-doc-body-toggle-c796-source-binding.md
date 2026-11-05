# 0359qf — DOC body-toggle c796 provenance and final-production rebind

- **Status:** c796 source assembly preserved; final-production 472b source
  assembly and qualification protocol prepared. No Cargo, native timing, or
  allocation run was performed in this turn.
- **Date:** 2026-10-24.
- **Scope:** the private DOC body-toggle leaf and the existing W5 probe. This
  receipt does not change the public API, the body implementation, or the
  earlier 0359qe evidence.

## Source binding

The read-only acceptance candidate was
(deleted-coord-artifact).
Its exact source identity is:

| item | identity |
|---|---|
| commit | c796d0acc24c7c6e6fd08980b9e31af6bb86b18b |
| tree | 6fe4f6f68046d2ee99dc7ca982004c406e484336 |
| Cargo.lock SHA-256 | 4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1 |
| rust-toolchain.toml SHA-256 | e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310 |
| .cargo/config.toml SHA-256 | 50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d |

The qualification copies are source-bound to that commit:

- candidate: /home/zhuhe/code/litchi-worktrees/uo2-doc-body-toggle-final,
  branch uo2/doc-body-toggle-final-qualification;
- baseline: (deleted-coord-artifact),
  detached at the same commit, with the overlay below.

The candidate body and revision files have SHA-256
db01cfce38c13c6c806382ffede990fda2b3bbe4dd53f59c4c5cb9feba9777d3 and
029721d34e7ad658e0064afd258cdb2b5a40e330110255b157aafb83f0d7b12a.
The c796 baseline overlay has SHA-256
6dc26da8130659c212dd29f26499cdaba7c4849268583b44dfcd9413a36887a3 and
b7155dc7c9aae6e318845b5073da4c0f275889125d38eae61df091aa2a2166df.
The copies contain the same probe, lock, toolchain, and Cargo configuration as
the read-only candidate.

The baseline is deliberately not pristine commit 81323113be. It is the same
c796 tree with only the reviewed body-publication changes from 160a00559b,
6456aa3223, and ae094d7d055c reversed in:

~~~text
crates/litchi-doc/src/body_text.rs
crates/litchi-doc/src/tracked_revision/package.rs
~~~

The reviewed overlay stream, regenerated from the prior final source copy, has
SHA-256
10765febb64c4e099f5c3ef560366124dc1d28d540d138768ae0f7cf782da2ab.
Applying that stream to c796 produced the c796-bound two-file diff with
SHA-256
0e0bf3c606277da07e53c085fa52a2e1db99839a3fc8ca5e869154ae60ab59c0.
After the identical probe-only output-dump overlay described below, the only
candidate-versus-baseline source difference is the two body paths above, and
git diff --check is clean.

The reviewed 0359qe probe was tied to the old 3cad source and had source SHA
9e1cd649268ad5eefe4314b76493f61cd047bf6ec4114c7e7872b2182b1af689.
It must not be reused for this candidate. The unmodified c796 probe source
was SHA-256
83b5eb8bf31cee397aaffd3cd784d3da1596a4cc99eb7fdca259a85ff3256def. This
packet adds the same opt-in dump-only harness change to both legs; the final
probe source SHA-256 in each copy is
b808e4dcba33dcafd4dff7dd22dc250e4170f2613839044fa8097263f5020965.
The c796-to-probe-only diff has SHA-256
d28c3bac33f8b2a8f0b8f4cbcdfb731e5b47e7635c067b34b599efead56e67aa.
With no LITCHI_PROBE_OUTPUT_DIR, this change does not alter the route or its
timing path. With that variable set, it writes the selected output bytes so
the fresh batch can be structurally audited.
The old 0359qe binaries and profiles remain outside the checkout and are not
presented as c796 measurements.

The overlay can be independently rechecked without Cargo:

~~~sh
REVIEWED=(deleted-coord-artifact)
BASE=(deleted-coord-artifact)
PACKET=/home/zhuhe/code/litchi-worktrees/uo2-doc-body-toggle-final
git -C "$PACKET" diff c796d0acc24c7c6e6fd08980b9e31af6bb86b18b --binary -- crates/litchi-ops/examples/legacy_w5_source_probe.rs |
  git -C "$BASE" apply --check -
git -C "$PACKET" diff c796d0acc24c7c6e6fd08980b9e31af6bb86b18b --binary -- crates/litchi-ops/examples/legacy_w5_source_probe.rs |
  git -C "$BASE" apply -
git -C "$REVIEWED" diff --cached --binary -- crates/litchi-doc/src/body_text.rs crates/litchi-doc/src/tracked_revision/package.rs | sha256sum
git -C "$REVIEWED" diff --cached --binary -- crates/litchi-doc/src/body_text.rs crates/litchi-doc/src/tracked_revision/package.rs |
  git -C "$BASE" apply --check -
git -C "$BASE" diff --check HEAD -- crates/litchi-doc/src/body_text.rs crates/litchi-doc/src/tracked_revision/package.rs
sha256sum "$BASE/crates/litchi-doc/src/body_text.rs" "$BASE/crates/litchi-doc/src/tracked_revision/package.rs"
~~~

## Final production endpoint rebind

The c796 source binding above is retained as the reviewed provenance. The
executable qualification endpoint is now the final production capture:

| item | final-production identity |
|---|---|
| read-only candidate | (deleted-coord-artifact) |
| commit | 472b6705475b073e7fe08f6e11ee17fb1e29f838 |
| tree | ebc1cee05f124c0c63977aff07a23288c19e7d00 |
| Cargo.lock SHA-256 | 4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1 |
| rust-toolchain.toml SHA-256 | e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310 |
| .cargo/config.toml SHA-256 | 50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d |
| prepared baseline | (deleted-coord-artifact) |
| baseline base commit/tree | 472b6705475b073e7fe08f6e11ee17fb1e29f838 / ebc1cee05f124c0c63977aff07a23288c19e7d00 |

The prepared baseline is the same 472b tree with exactly the reviewed
body-publication reversal in body_text.rs and tracked_revision/package.rs.
The reviewed recipe remains SHA-256
10765febb64c4e099f5c3ef560366124dc1d28d540d138768ae0f7cf782da2ab; the
472b-bound applied two-file diff is SHA-256
0e0bf3c606277da07e53c085fa52a2e1db99839a3fc8ca5e869154ae60ab59c0.
The final candidate body/revision hashes are
db01cfce38c13c6c806382ffede990fda2b3bbe4dd53f59c4c5cb9feba9777d3 and
029721d34e7ad658e0064afd258cdb2b5a40e330110255b157aafb83f0d7b12a;
the prepared baseline hashes are
6dc26da8130659c212dd29f26499cdaba7c4849268583b44dfcd9413a36887a3 and
b7155dc7c9aae6e318845b5073da4c0f275889125d38eae61df091aa2a2166df.
Both legs use the actual final-production probe SHA-256
b808e4dcba33dcafd4dff7dd22dc250e4170f2613839044fa8097263f5020965.
The read-only candidate is clean; the prepared baseline has only the two
expected body files modified (plus its ignored copied Cargo.lock).

Verify the final source-only preparation without Cargo:

~~~sh
FINAL=(deleted-coord-artifact)
BASE=(deleted-coord-artifact)
git -C "$FINAL" rev-parse HEAD HEAD^{tree}
git -C "$BASE" rev-parse HEAD HEAD^{tree}
git -C "$BASE" diff --name-only HEAD -- crates/litchi-doc/src/body_text.rs crates/litchi-doc/src/tracked_revision/package.rs
git -C "$BASE" diff --check HEAD -- crates/litchi-doc/src/body_text.rs crates/litchi-doc/src/tracked_revision/package.rs
git -C "$BASE" diff --binary HEAD -- crates/litchi-doc/src/body_text.rs crates/litchi-doc/src/tracked_revision/package.rs | sha256sum
sha256sum "$FINAL/crates/litchi-ops/examples/legacy_w5_source_probe.rs" "$BASE/crates/litchi-ops/examples/legacy_w5_source_probe.rs"
~~~

The CHPX oracle is the final endpoint's
tools/audit_doc_chpx_history.py. It must consume only the two fresh
472b-bound dump files produced by the opt-in probe; the old c796/3cad OLE
artifacts remain historical evidence and are never reused as final results.

## Required qualification

0359a-F1 records the queued residual: the leaf body-toggle batch costs
1.1–1.7 ms per staged toggle at 1,000 paragraphs. The final check therefore
has two routes, both using the same source-bound probe and both legs of the
source assembly:

- doc/leaf_write: one body toggle through commit and write_to;
- doc/batch1k_leaf_write: 1,000 body toggles through commit and write_to.

The following commands are the exact pending run protocol. They use the
20-GiB capped environment, one Cargo job, offline locked dependencies, and a
single pinned CPU. They are recorded here, not claimed as run results.

~~~sh
source (deleted-coord-artifact)
ROOT=/home/zhuhe/code/litchi-worktrees
CAND=(deleted-coord-artifact)
BASE=(deleted-coord-artifact)
RUN=(deleted-coord-artifact)
CPU=29
mkdir -p "$RUN/candidate" "$RUN/baseline" "$RUN/abba" "$RUN/allocation" "$RUN/tmp"

for leg in candidate baseline; do
  if test "$leg" = candidate; then src=$CAND; else src=$BASE; fi
  mkdir -p "$RUN/target-$leg" "$RUN/tmp/$leg"
  (cd "$src" &&
    CARGO_TARGET_DIR="$RUN/target-$leg" CARGO_BUILD_JOBS=1 \
    CARGO_INCREMENTAL=0 TMPDIR="$RUN/tmp/$leg" \
    cargo build --offline --locked --release -p litchi-ops \
      --features doc,xls,rtf --example legacy_w5_source_probe)
done

for leg in candidate baseline; do
  bin="$RUN/target-$leg/release/examples/legacy_w5_source_probe"
  LITCHI_PROBE_ONLY=doc/leaf_write LITCHI_PROBE_BATCH_ROUNDS=1 \
    LITCHI_PROBE_OUTPUT_DIR="$RUN/$leg" \
    taskset -c "$CPU" "$bin" 1 \
    >"$RUN/$leg/doc_leaf_write.json" 2>"$RUN/$leg/doc_leaf_write.log"
  LITCHI_PROBE_ONLY=doc/batch1k_leaf_write LITCHI_PROBE_BATCH_ROUNDS=1 \
    LITCHI_PROBE_OUTPUT_DIR="$RUN/$leg" \
    taskset -c "$CPU" "$bin" 1 \
    >"$RUN/$leg/doc_batch1k_leaf_write.json" 2>"$RUN/$leg/doc_batch1k_leaf_write.log"
done

# Ordinary output must be byte-identical.
test "$(jq -r '.[0].output_sha256' "$RUN/candidate/doc_leaf_write.json")" =
     "$(jq -r '.[0].output_sha256' "$RUN/baseline/doc_leaf_write.json")"

# The batch route is not required, and is not expected, to have equal OLE
# bytes. Preserve both output artifacts from the route harness and run the
# structural CHPX audit; never replace this with a raw-byte equality claim.
python3 "$CAND/tools/audit_doc_chpx_history.py" \
  --candidate "$RUN/candidate/doc_batch1k_leaf_write.bin" \
  --baseline "$RUN/baseline/doc_batch1k_leaf_write.bin"
~~~

The batch audit must be accompanied by the same semantic/fidelity checks used
by 0359qe: the candidate and baseline reopen successfully, have equal
paragraph text and page/text rendering, and the audit compares the
FIB-selected PlcBteChpx ranges and referenced ChpxFkp payloads. The previous
audited result is intentionally retained: the old candidate was 182,784 bytes,
the old overlay baseline 7,715,328 bytes, and their rendered 18-page/text
outputs matched while their raw OLE hashes differed. Those are historical 3cad
values, not c796 results.

The focused semantic checks are also required on the fresh candidate and
overlay builds:

~~~sh
(cd "$CAND" && cargo test --offline --locked -p litchi-doc --lib \
  deferred_body_formatting_flushes_before_table_edits_and_keeps_failures_atomic)
(cd "$CAND" && cargo test --offline --locked -p litchi-doc --lib \
  retained_render_is_cleared_by_resource_reopen_and_formatting_flushes_at_commit)
(cd "$BASE" && cargo test --offline --locked -p litchi-doc --lib \
  retained_render_is_cleared_by_resource_reopen_and_reenabled_for_formatting)
~~~

## Quiet-host ABBA and allocation protocol

Correctness outputs are collected before any profiles. Native timing is only
valid on a quiet host; a loaded-host run is diagnostic and cannot close the
requirement. The two routes must each run in binary-level baseline–candidate–
candidate–baseline order under the 20-GiB scope:

~~~sh
for leg in baseline candidate candidate baseline; do
  bin="$RUN/target-$leg/release/examples/legacy_w5_source_probe"
  for route in doc/leaf_write doc/batch1k_leaf_write; do
    if test "$route" = doc/leaf_write; then name=doc_leaf_write
    else name=doc_batch1k_leaf_write; fi
    systemd-run --user --scope --quiet --collect \
      -p OOMPolicy=continue -p MemoryMax=20G -p MemorySwapMax=0 -- \
      /usr/bin/time -f "$leg $name maxrss_kb=%M wall=%e" \
      taskset -c "$CPU" env LITCHI_PROBE_ONLY="$route" \
        LITCHI_PROBE_BATCH_ROUNDS=1 "$bin" 1 \
      >"$RUN/abba/$leg-$name.json" 2>>"$RUN/abba/abba.log"
  done
done
~~~

Callgrind and Massif are the allocation/instruction diagnostics for the same
two routes. They must retain each leg's JSON output SHA and separate profile
files; no old 0359qe profile is a 472b result:

~~~sh
for leg in baseline candidate; do
  bin="$RUN/target-$leg/release/examples/legacy_w5_source_probe"
  for route in doc/leaf_write doc/batch1k_leaf_write; do
    if test "$route" = doc/leaf_write; then name=doc_leaf_write
    else name=doc_batch1k_leaf_write; fi
    taskset -c "$CPU" valgrind --tool=callgrind --quiet \
      --callgrind-out-file="$RUN/allocation/$leg-$name.callgrind" \
      env LITCHI_PROBE_ONLY="$route" LITCHI_PROBE_BATCH_ROUNDS=1 \
      "$bin" 1 >"$RUN/allocation/$leg-$name.json" \
      2>"$RUN/allocation/$leg-$name.callgrind.log"
    taskset -c "$CPU" valgrind --tool=massif --quiet \
      --massif-out-file="$RUN/allocation/$leg-$name.massif" \
      env LITCHI_PROBE_ONLY="$route" LITCHI_PROBE_BATCH_ROUNDS=1 \
      "$bin" 1 >"$RUN/allocation/$leg-$name.massif.json" \
      2>"$RUN/allocation/$leg-$name.massif.log"
    ms_print "$RUN/allocation/$leg-$name.massif" \
      >"$RUN/allocation/$leg-$name.massif.txt"
  done
done
~~~

The final report must include source/tree/lock identities, binary SHA-256
values, route output hashes, the explicit batch raw-byte difference, the
semantic/CHPX audit, and the quiet-host ABBA and allocation summaries. Until
those commands run on the 472b-bound copies, this receipt makes no performance
or native acceptance claim.
