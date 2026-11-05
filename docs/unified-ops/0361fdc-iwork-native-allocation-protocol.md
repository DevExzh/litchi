# 0361fdc — iWork allocation and peak packet

- Scope: the final iWork constructor and owner native lanes.
- Production source under qualification: constructor candidate
  `c796d0acc24c7c6e6fd08980b9e31af6bb86b18b`; owner candidate is the same
  integrated endpoint. The constructor baseline is `ddf2ac8ae2117cd78829ef94bcb610d6cbdfa0e2`;
  the owner baseline is `a18593385cf3162911f00ebd3cbb1379855a8eba`.
- Status: protocol/probe repair prepared; no Cargo build, allocation run, or
  native timing claim was made by this source-only change.

## Gap and repair

The c796 runner already built and correctness-validated both `native` and
`allocation` binaries, but its only execution phase invoked `native`. The
constructor probe emitted `peak`, `retained`, `allocated`, and `allocations`;
the owner probe emitted none, so an owner counting binary could not provide
the required operation-allocation evidence.

The repaired tracked owner probe uses the same counting allocator and
operation boundary as the constructor probe. Fixture parsing, correctness
proofs, and final output serialization remain outside the measured closure;
the counter is reset immediately before each route's edit loop and read
immediately after it. Native binaries remain uninstrumented. The runner adds
an `allocation` phase that invokes the protected counting binaries only, in
`ABBA` order, for every route, with the existing CPU affinity and
before/during/after host-overlap guards. It requires non-negative typed
metrics, `peak >= retained`, and for owner routes a non-zero allocation
count/byte total; repeated A and B legs must retain identical metrics. This
packet makes no latency claim and is independent of the fixed native timing
phase.

The route sets are unchanged: all 15 constructor routes and all four owner
routes. The fixed allocation protocol is 100 operations per route, order
`ABBA`, CPU 29 (or another coordinator-approved allowed CPU), and the same
120-second child bound as native. A fresh external attempt directory is
required; no evidence or binary belongs in this checkout.

## Exact preparation and execution commands

The stable source endpoint used to prepare these commands is:

```text
(deleted-coord-artifact)
HEAD: c796d0acc24c7c6e6fd08980b9e31af6bb86b18b
Cargo.lock: 4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1
```

Run from the full candidate worktree after this protocol/probe repair has
been applied, with all generated directories outside the checkout:

```sh
export PATH=(deleted-coord-artifact):$PATH
export CARGO_BUILD_JOBS=1
export RUST_TEST_THREADS=1
export LITCHI_MEM_MAX=20G
export LITCHI_IWA_WORKTREE=(deleted-coord-artifact)
export IWA_ROOT_LOCK=$LITCHI_IWA_WORKTREE/Cargo.lock
export IWA_FIXTURES=(deleted-coord-artifact)
export IWA_CONSTRUCTOR=(deleted-coord-artifact)
export IWA_OWNER=(deleted-coord-artifact)

python3 docs/unified-ops/results/0361fda/iwa_native_qualification.py \
  prepare-fixtures "$IWA_FIXTURES"

python3 docs/unified-ops/results/0361fda/iwa_native_qualification.py \
  build constructor "$IWA_CONSTRUCTOR" \
  --base ddf2ac8ae2117cd78829ef94bcb610d6cbdfa0e2 \
  --candidate c796d0acc24c7c6e6fd08980b9e31af6bb86b18b \
  --lock "$IWA_ROOT_LOCK"
python3 docs/unified-ops/results/0361fda/iwa_native_qualification.py \
  validate constructor "$IWA_CONSTRUCTOR"
python3 docs/unified-ops/results/0361fda/iwa_native_qualification.py \
  allocation constructor "$IWA_CONSTRUCTOR" --attempt allocation-a1 --cpu 29

python3 docs/unified-ops/results/0361fda/iwa_native_qualification.py \
  build owner "$IWA_OWNER" \
  --base a18593385cf3162911f00ebd3cbb1379855a8eba \
  --candidate c796d0acc24c7c6e6fd08980b9e31af6bb86b18b \
  --lock "$IWA_ROOT_LOCK"
python3 docs/unified-ops/results/0361fda/iwa_native_qualification.py \
  validate owner "$IWA_OWNER" --fixtures "$IWA_FIXTURES"
python3 docs/unified-ops/results/0361fda/iwa_native_qualification.py \
  allocation owner "$IWA_OWNER" --attempt allocation-a1 --cpu 29 \
  --fixtures "$IWA_FIXTURES"
```

Only after both allocation packets complete and the coordinator confirms a
quiet host may the fixed native ABBA phases start:

```sh
python3 docs/unified-ops/results/0361fda/iwa_native_qualification.py \
  run constructor "$IWA_CONSTRUCTOR" --attempt native-a1 --cpu 29
python3 docs/unified-ops/results/0361fda/iwa_native_qualification.py \
  run owner "$IWA_OWNER" --attempt native-a1 --cpu 29 \
  --fixtures "$IWA_FIXTURES"
```

The native command remains the only timing acceptance phase: five complete
ABBA blocks, two warmups, 500 ms calibration target, 150 ms minimum leg, and
the fixed route lists. Allocation samples are not substituted for timing
blocks or treated as a relaxed performance budget.

## Input identities

At the stable endpoint, the tracked qualification inputs are:

| input | bytes | SHA-256 |
| --- | ---: | --- |
| stable c796 `results/0361fda/iwa_native_qualification.py` | 98,169 | `588a66600bcf03f39fb21ca286af6bacdd224971f976611ebbb3e7df8ecb8917` |
| repaired `results/0361fda/iwa_native_qualification.py` | 109,594 | `c6d0cd178b5c21a05ff20f9eb12fb93fcb040cd783b322f9bec8899717e325ab` |
| `results/0361fda/constructor-probe.rs` (unchanged) | 22,799 | `cc322c2189e8e3822472627e61c45b8e463f866adcba32858c88cc237bd17a3e` |
| stable c796 `results/0361fda/owner_probe.rs` | 22,737 | `f8995a3c38784c406c298271dd6a8b27462bde51b9956975e8f26328a4200ff0` |
| repaired `results/0361fda/owner_probe.rs` | 27,230 | `c580de1c52c2c37c1cebbc7a0e6324cd03cd827e01de09d11ca7f2d9de4b50a1` |
| `results/0361fda/fixture-provenance.json` | 2,326 | `72e067ae44f5017f84cf8c883db72c0d4ec97ba79b9048f4c9e74b5d4e2ea86f` |
| root `Cargo.lock` | 139,606 | `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` |
| `tools/perf-baseline/Cargo.lock` | 57,227 | `eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c` |

The owner fixture manifest is generated only by `prepare-fixtures`; its
provenance binds `basic.numbers`, `basic.pages`,
`media-replacement-native.key`, `tone.wav`, and the generated
`footnote-fixture.pages` to tracked sources. The runner replays that generator
and compares bytes before correctness and allocation collection.

Validation performed without Cargo/native execution on this repair leaf:

```text
python3 -m unittest discover -s docs/unified-ops/results/0361fda -p 'test_*.py'  # 23 passed
python3 -m py_compile docs/unified-ops/results/0361fda/iwa_native_qualification.py \
  docs/unified-ops/results/0361fda/test_iwa_native_qualification.py
```

The stable c796 identities are retained here to bind the production source
and lock inputs; the repaired runner/probe identities bind the actual
qualification inputs that must be recorded in the later build manifest.
