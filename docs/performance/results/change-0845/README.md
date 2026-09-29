# 0845 matched direct-entry iteration experiment

Base: `d6619c101f`. The only intended production changes are in
`crates/soapberry-zip/src/office.rs` and
`crates/litchi-opc/src/source_backed.rs`: preserve existing opaque ZIP entry IDs
through the physical-member lookup and topology-limit scans. The applied patch,
source copies and full source inventories bind the eventual candidate.

The probe and independent ZIP/Heaptrack readers are byte-identical to 0844;
therefore their report schema intentionally remains 0844. New 0845 drivers,
plan and analysis identify this comparison. Both release legs use identical
probe, dependencies, compiler settings and profile owner, on CPU 12.

Four shapes × four counts × two modes are retained. Six matched blocks balance
before-first and after-first order for every case. Parts 4/64 use 1,000 samples;
Parts 1,024/4,096 use 100, all after three warmups. The main matrix has 64 archive
qualifications, 384 native reports and 64 one-sample Heaptrack reports:
512 reports / 211,328 samples. A separately planned 16-report / 16-sample
Callgrind comparison checks instruction work at 64/4,096 Parts for shared and
deep-unique publication. Two tiny Heaptrack preflights remain separate from both
matrices. Native and instrumented observations are never pooled.

Independent qualification reconstructs ZIP payloads, CRCs, logical digests,
untouched raw spans and changed-target stable local/central metadata. Exact
source/output bytes must agree across legs and modes for each shape/count.
Native results bind those qualification identities. Original reports and ZIPs
remain retained; derived summaries omit duplicated decoded manifests and
sample vectors only after every reader check runs. Replay repeats those checks.
Heaptrack summaries keep operation-owner ancestry totals and bind raw traces;
no individual allocation-site, copied-byte or peak-live-memory claim follows.

Replay after cleanup:

```sh
python3 -B docs/performance/results/change-0845/reader_tests.py
python3 -B docs/performance/results/change-0845/heap_reader_tests.py
python3 -B docs/performance/results/change-0845/close.py verify
```

Commands, start/end receipts, actual environment and source/binary hashes are
retained. Reproduction needs a fresh checkout/packet at the recorded base and
fresh owned roots with matching markers. Build and qualify the baseline before
applying the candidate, then build/qualify the candidate and run all quality
gates. Admission freezes both legs before any native/Heaptrack formal capture.
Evidence paths are exclusive; never overwrite a retained run. Cleanup verifies
binary hashes and removes only roots bearing this packet's ownership markers.

The four pure shapes are synthetic generic OPC workloads. They do not establish
mixed real-catalog or semantic Office performance, cold or remote behavior,
filesystem durability, concurrency scaling, hardware cycles, peak live memory
or physical copy volume. The broader goal remains active, with iWork excluded
and ODF deferred. See `selection.md` and `priority-review.md` for scope and the
next broader experiment.
