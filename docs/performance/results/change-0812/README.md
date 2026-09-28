# 0812 — scanner instruction localization

The [report](../../0812-pptx-scanner-instruction-localization.md) records a
failed exact reconstruction followed by a separately frozen fresh-sample
amendment. No production source changed and no optimization is adopted.

`rebuild.py` preserves the attempted reconstruction and complete binary
mismatch. `offset_audit.py` independently counts historical offsets without
mapping them to assembly. `fresh.py` owns two new perf captures and decodes;
its post-capture NameError is retained, and `finish_assembly.py` records the
assembly-only continuation. `line_mapping.py` retains DWARF inline attribution.
`instruction_analysis.py` joins only fresh samples to their exact binary;
`fresh_offset_audit.py` independently checks fresh counts and output oracles.

All execution artifacts are immutable. Drivers refuse existing output paths;
do not rerun them over this sealed packet. The original absent 0811 target was
temporarily recreated by this batch and is removed after reader validation.
No older evidence file was modified. The owned target is not needed for replay:

```sh
python3 -B docs/performance/results/change-0812/validate.py --final
```

The original and amended plans remain separate. The primary evidence now
supports an explicit event success/error match as a bounded next experiment.
Sampling skid and frame-pointer perturbation prohibit causal cost or
ordinary-build phase-share claims. Production adoption needs its own fresh
correctness, workflow, tail, and resource qualification.
