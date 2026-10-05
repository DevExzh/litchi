# Candidate source manifest regeneration

`candidate-source.json` records the external source archive's filename, size,
SHA-256 and exact `git archive` arguments. Extract that archive into a fresh
external directory, apply `candidate.patch` inside `tree-v2`, and copy the
retained `Cargo.lock.workspace-generated-copy` into `tree-v2/Cargo.lock`.
The source tree then has 10,898 regular files. Build commands and executable
descriptors are in `candidate-build-final.json`; this rejected candidate must
not be applied to the production checkout.

The three complete manifests are larger than 1 MiB and remain external.
Their descriptors are in `candidate-source.json`, `candidate-build-final.json`
and `postcapture-source.json`. The following command, run inside the restored
`tree-v2`, regenerates their exact historical bytes in its parent directory.
The historical header times identify the retained witnesses; they are not
timestamps of the regeneration. Exclusive output creation prevents replacing
existing evidence.

```sh
python3 -B - <<'PY'
from pathlib import Path
import hashlib
root = Path('.').resolve()
rows = ''.join(
    f'{p.relative_to(root).as_posix()}\t{p.stat().st_size}\t'
    f'{hashlib.sha256(p.read_bytes()).hexdigest()}\n'
    for p in sorted(root.rglob('*'), key=lambda p: p.relative_to(root).as_posix())
    if p.is_file()
)
versions = [
    ('source-before-final.manifest.tsv', 'source-manifest',
     '2026-10-05T18:28:39,457429655+00:00'),
    ('source-after-final.manifest.tsv', 'source-manifest',
     '2026-10-05T18:31:54,727230056+00:00'),
    ('source-postcapture.manifest.tsv', 'postcapture-source-manifest',
     '2026-10-05T18:55:30.568828+00:00'),
]
for name, schema, timestamp in versions:
    header = (
        f'# schema=litchi.performance.0918.{schema}.v1\n'
        '# revision=f5485b3ca8582051247d31946b1c5f45c8a561ad\n'
        '# root=/home/zhuhe/litchi-performance-0918/candidate/tree-v2\n'
        f'# generated_utc={timestamp}\n'
        '# columns=relative_path<TAB>bytes<TAB>sha256\n'
    )
    data = (header + rows).encode()
    with (root.parent / (name + '.regenerated')).open('xb') as output:
        output.write(data)
    print(name, len(data), hashlib.sha256(data).hexdigest())
PY
```

The before/after/postcapture SHA-256 values must be, respectively:

```text
41a5d78bd26493575ca2f6a9b129532d1d46971fe9cfce7f9c612bc4a2d5cb9e
33242e08d85e380f3434891b369f588690ccb0e708725790bdebd4cf58c736a4
2108aa1fd620470684e0624be9a783f0ebd2a2beda64213686a6dece801e20e5
```
