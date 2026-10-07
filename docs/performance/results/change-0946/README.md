# 0946 bounded XML reader attribution evidence

See [the report](../../0946-xml-reader-attribution.md) for the finding and
constraints. This packet decomposes four retained 0944 Callgrind part-6 dumps;
it executes no benchmark, build, or profiler and establishes no native speedup.
The profiles predate 0945. Current XML auditor source identity does not make
them current caller measurements.

The committed files are review copies: final analysis, reducer, retained
reader, bound metadata, source audit, independent controls, and archive receipts.
Raw dumps and historical/redundant outputs remain in the external packet.
Run the reducer from the complete extracted packet, preserving its layout and
`tools/reduce.py` filename. A fresh output path is required.

## External artifact

- File: `/home/zhuhe/litchi-performance-0946-archive/litchi-performance-0946-v1.tar.gz`
- Size: **402,597 bytes**
- SHA-256: `447b116dabb7576e25a452a3d33059c01dbfe34af9ba6a05ba190a65203f3b9b`
- Contains 33 regular files; exact member sizes and hashes are in
  [inventory-v1.json](inventory-v1.json).
- Original retained packet: `/home/zhuhe/litchi-performance-0946`.

To regenerate the archive from the retained packet into a fresh external directory:

```sh
python3 -B docs/performance/results/change-0946/archive_packet.py \
  --packet /home/zhuhe/litchi-performance-0946 \
  --destination /tmp/litchi-0946-new-archive
```

To replay from the archive using fresh destinations:

```sh
mkdir /tmp/litchi-0946-replay
tar -xzf /home/zhuhe/litchi-performance-0946-archive/litchi-performance-0946-v1.tar.gz \
  -C /tmp/litchi-0946-replay
python3 -B /tmp/litchi-0946-replay/litchi-performance-0946/tools/reduce.py \
  --output /tmp/litchi-0946-recomputed.json
cmp /tmp/litchi-0946-recomputed.json \
  /tmp/litchi-0946-replay/litchi-performance-0946/attribution-summary-v2.json
```

Final output: **353,314 bytes**, SHA-256
`e0ca6ecd41879a1bdfcb0530d76dfae8a92cfe7b87365e91cf7e5ea64dd27a4b`.

## Verification and limitations

Independent original and moved copied-input replays matched final v2 byte for
byte. Six actual-helper negative controls rejected corrupt Callgrind input,
duplicate schedule index, traversal path, descriptor corruption, symlink input,
and output overwrite. The root independently replayed the final helper, verified
every archived member after extraction, and replayed the extracted helper with
identical output. See [controls](review/independent-controls-v1.json),
[archive verification](archive-verified-v1.json), and
[root verification](root-verification-v1.json).

The archive initially refused an intentional symlink in disposable negative-control
scratch. Owned replay/control trees were inventoried and removed before archiving;
the control results and script remain. The owned verification extraction was
removed after replay. No retained source input or archive was deleted.

[Preparation history](preparation-history.json) explicitly records two missing
drafts: the original reducer and a provisional plan were overwritten during
preparation and cannot be independently rehashed. Plan v1 and output v1 remain
in the external packet, but only final v2 is claimed reproducible. This was an
offline exploratory analysis, not a prospectively frozen benchmark experiment.

Nested function rows overlap and must not be summed. Call counts are provenance,
not token counts. Aggregated edges do not independently identify original versus
replacement audit cost. Source identities use the focused hash-bound receipts;
this is not a new full source-tree or raw 0944 JSON requalification. No Rust
source changed, so Rust compilation/tests were not rerun for this batch.
