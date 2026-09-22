# 0735 rejected PPT record-staging evidence

Rejected and reverted: primary paired p50 −3.53%, secondary +6.34% with
eight of nine secondary pairs above 5%. Allocated bytes fall 696,415/331,274;
peak live bytes are unchanged. The entire production file is restored to the
accepted baseline. Candidate code and tests are archived only.

This packet measures eliminating the full private Editor clone from validated
record staging. Public slide-index-1 removal on `45543.ppt` and `41246-1.ppt`
exercises record replacement. Record insertion has the same clone pattern and
receives focused correctness coverage; this matrix makes no insertion-speed
or multi-record scaling claim.

Before measurements precede the live source edit. The baseline includes the
retained 0734 owned writer handoff. Both binaries use identical probe source
and flags; `baseline-build.json` and `candidate-build.json` retain the full
source census, probe census, quality manifests and exact binary identities.
`base.json`, `source-archive` and `candidate` bind the single changed module.

The prospective hypothesis and fixed schedule use nine native process pairs
per case (three cycles × three rounds), 50 samples and three warmups, plus
three allocation pairs per case with one sample and no warmup. CPU 12 is
pinned and execution is serialized. All samples, tails and absolute >5% flags
are retained. Bootstrap intervals use nine process pairs, 10,000 resamples
and seed 7335; within-process samples are not independent process repetitions.

Both cases require the full prior preservation oracle: exact bytes and stream
inventory, normalized raw CFB directory metadata, live record/slide identities,
survivor payloads, text, outlines, list data, notes and comments. The primary
remains pinned to the sealed 0731/0728 reference. Every process records eight
rejected corruption controls outside the measured owner interval.

To reproduce without overwriting retained evidence, use a new packet and
scratch directory with the archived baseline source. Run `build.py baseline`
and `qualify.py baseline`; apply the archived candidate; run `quality.py`,
`build.py candidate`, and `qualify.py candidate`. After source review, finalize
the environment and prospective plan, run `run.py freeze`, `preflight-run.py`,
and `run.py capture`. Then run `analyze.py`, `audit.py` and
`negative-checks.py`. The synthetic preflight is schema evidence only.

Before production reversion, post-cleanup `analyze.py`, `audit.py`, the source
guard and all 19 corruption controls passed using exact deleted-binary identity
receipts. After reversion, use `replay-rejected.py`: it first verifies the real
workspace equals the baseline, then replays the frozen candidate validators
against archived candidate source in an isolated temporary root. It does not
modify the real source or run native binaries. Direct candidate validators
intentionally reject the restored production file. `artifact-seal.py --check`
verifies every packet file and rejects omissions or additions. Peak live bytes
are boundary-relative allocator ownership, not RSS. No cold-I/O, instruction,
concurrency, broad-producer or broad-CRUD claim follows from this packet.

Both validators reject all 19 evidence-corruption controls. The failed first
synthetic preflight and its original review receipt remain in
`preflight-attempt-0`; only a boolean field name was corrected before the second
preflight and all real measurements. See [the report](../../0735-ppt-record-staging-without-editor-clone.md),
`disposition.json`, `reversion.json` and `results-review.md`.
