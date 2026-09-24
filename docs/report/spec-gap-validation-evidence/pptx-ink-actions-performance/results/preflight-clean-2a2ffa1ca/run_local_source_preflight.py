from pathlib import Path
import subprocess

root = Path('/var/tmp/litchi-pptx-ink-actions-profiler-2a2ffa1c')
out = Path('/var/tmp/litchi-pptx-ink-actions-profiler-2a2ffa1c-preflight')
here = root / 'docs/report/spec-gap-validation-evidence/pptx-ink-actions-performance'
harness = here / 'harness'
extras = [
    root / 'Cargo.toml', root / 'rust-toolchain.toml', root / '.cargo/config.toml',
    root / 'rustfmt.toml', root / 'clippy.toml', root / 'deny.toml',
    root / 'docs/adr/0001-priorities-and-api-layers.md',
    root / 'docs/adr/0003-snapshots-edits-and-patches.md',
    root / 'docs/adr/0005-io-memory-and-performance.md',
    root / 'docs/adr/0006-validation-security-and-compatibility.md',
    root / 'docs/report/spec-gap-validation-evidence/pptx-ink-actions-design.md',
    harness / 'Cargo.toml', harness / 'Cargo.lock', harness / 'main.rs', harness / 'adapter.rs',
    harness / 'support.rs', here / 'source_manifest.py', here / 'committed_inputs.py', here / 'verify.py',
    here / 'cleanup_target.sh', here / 'run_profile.sh', here / 'README.md', here / 'PLAN.md',
    here / 'report.md', here / 'requirements.md', here / 'corpus-manifest.json', here / 'source-contract.json',
    here / 'scaffold_tests.py', here / 'root-review.md',
]
cmd = ['python3', str(here/'source_manifest.py'), '--metadata', str(out/'cargo-metadata.json'), '--root', str(root), '--output', str(out/'source-manifest-local.txt'), '--git-commit', '2a2ffa1cae4e6b7070082768ce84483e5d411dc8']
for extra in extras:
    cmd.extend(['--extra', str(extra)])
result = subprocess.run(cmd, cwd=root, text=True, capture_output=True)
(out/'source-manifest-local.stdout').write_text(result.stdout)
(out/'source-manifest-local.stderr').write_text(result.stderr)
(out/'source-manifest-local.exit').write_text(f'{result.returncode}\n')
print(result.returncode)
