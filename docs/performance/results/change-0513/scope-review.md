# 0513 save-profile scope review

The candidate Callgrind artifacts establish the declared bounded profile scope. The raw file [`after/profile.out`](after/profile.out) has one `summary` and `totals` value of **14,458,431,895 Ir**. Independently reading its function blocks gives the following direct edges:

| Direct edge | Calls | Inclusive Ir attributed to the callee |
| --- | ---: | ---: |
| `run_xlsx_update_commit_save` → `xlsx_commit_save_operation` | 3 | 14,458,431,895 |
| `xlsx_commit_save_operation` → `Edit::commit` | 3 | 11,911,855,692 |
| `xlsx_commit_save_operation` → `PackageWriter::write_to_stream` | 3 | 2,546,575,570 |

The raw records are the `cfn=(2406) ... calls=3` edge in the runner record and the `cfn=(1268) ... calls=3` and `cfn=(2408) ... calls=3` edges in the helper record. The inclusive and exclusive annotations independently show the same three direct calls in [`profile-inclusive.txt`](after/profile-inclusive.txt) and [`profile-exclusive.txt`](after/profile-exclusive.txt).

The helper’s exclusive cost is **171 Ir**, and its two direct `__memcpy_avx_unaligned_erms` edges contribute **462 Ir** in total (six calls, 231 Ir per three-call edge). The accounting closes exactly:

```text
11,911,855,692  Edit::commit
 2,546,575,570  PackageWriter::write_to_stream
              171  helper exclusive instructions
              462  direct memcpy instructions
----------------
14,458,431,895  raw profile total
```

The commit and writer portions are 82.3869129% and 17.6130827% of the raw total, respectively. Their sum leaves 633 Ir, exactly the helper’s 171 Ir plus direct memcpy’s 462 Ir.

## Why the oracle is excluded

The profile command in [`profile-receipt.json`](after/profile-receipt.json) uses `--collect-atstart=no` and toggles collection on `*litchi_perf_baseline::xlsx_commit_save_operation`. In the harness, `xlsx_expected_output(corpus, updates)` is evaluated before the per-iteration loop; that oracle performs its own `prepare_xlsx_updates(...).commit()` and `to_bytes()` before any helper entry. Per-iteration workbook loading, update staging, sink construction and reservation also precede the helper. Collection therefore starts only when the helper is entered.

The helper performs the existing `edit.commit()` and then the existing workbook `write_to(sink)`, and returns the `Commit`. The caller takes the elapsed time, finishes allocation observation, checks byte equality, reopens and verifies the output, and only then black-boxes the commit. Those caller-side checks and the eventual commit drop are outside the toggled region. The source boundary is [`lib.rs`](../../../../tools/perf-baseline/src/lib.rs#L40840) and the loop is [`lib.rs`](../../../../tools/perf-baseline/src/lib.rs#L40852).

This conclusion is about collected instructions, not merely the clock: the raw runner block attributes the active collection total to three helper entries, while the source order and toggle prevent the pre-loop expected-output commit and write from entering that collection window.

## Annotation interpretation

The helper’s 100% inclusive row is the actual inclusive cost of the three collected helper invocations, including its commit, writer, and direct helper instructions. The runner’s 100% row is an ancestor aggregate of that active collection window; it is not the cost of the complete runner including fixture generation, oracle construction, verification, or teardown. Because collection starts at the helper, both rows can show the same active-collection total without making the runner’s full invocation part of the measured scope.

Only `*` function records and their following `>` direct-callee rows should be used for the direct-edge claim. `<` rows are caller-context aggregates. Descendants can have larger call counts than their direct parent: for example, the annotation shows `Worksheet::store` and worksheet parse descendants at 9 calls near three direct commits, and deeper parser paths at still higher counts. Callgrind call metadata can retain calls made while event collection is off, including setup or readback paths, so these descendant counts cannot be interpreted as per-commit-only collected work. They do not establish additional helper invocations or change the direct-edge result; the raw direct `calls=3` records are the authoritative count for the three edges above.

The writer symbol retained by this build is `PackageWriter::write_to_stream`, reached by the helper’s generic `write_to` call. The profile does not need to freeze the generic method’s mangled/inlined symbol to establish the source-level boundary; the raw retained package-writer edge plus the helper source body establishes that the write remains inside the helper.

## Warning and limits

[`profile.log`](after/profile.log) reports `brk segment overflow in thread #1: can't grow to 0x7924000`, followed by `Collected : 14458431895`; `/usr/bin/time` records exit status 0. The raw profile is complete enough to contain matching `summary` and `totals` records, and both annotations were generated and hash-bound by the receipt. I therefore find no blocker to recording this as bounded direct-edge diagnostic evidence, but the Valgrind warning must remain attached to the artifact and prevents stronger warning-free-profile language.

This is a three-sample, zero-warmup dense-wide one-percent commit/save profile and reports synthetic Callgrind `Ir`, not wall-clock latency, hardware cycles, allocator cost, or a before/after speedup. The total is the sum of the three selected helper calls; it provides no per-sample variance. The profile excludes all setup, expected-output/oracle work, post-operation checks, and caller-side destruction by design. It should not be generalized to the full runner, other XLSX cases, OLE2, ODF, or iWork.

The direct attribution is trustworthy for the declared helper region and its three commit/write calls, subject to the `brk` warning and normal Callgrind instrumentation limits. It supports the 0513 evidence-only boundary and no optimization claim.
