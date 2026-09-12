"""Exercise smoke target ownership with disposable source and build fixtures."""

import os
from pathlib import Path
import re
import shutil
import subprocess
from tempfile import TemporaryDirectory
import unittest


class SmokeTargetOwnership(unittest.TestCase):
    def test_normalized_existing_target_is_not_owned_or_deleted(self):
        here = Path(__file__).resolve().parent
        repository = here.parents[3]
        script = (here / "smoke.sh").read_text()
        with TemporaryDirectory(prefix="litchi-root-smoke-target-") as temporary:
            root = Path(temporary)
            evidence = root / here.relative_to(repository)
            evidence.mkdir(parents=True)
            (evidence / "smoke.sh").write_text(script)
            for relative in re.findall(r'\["\$ROOT/([^"\n]+)"\]=[0-9a-f]{64}', script):
                destination = root / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(repository / relative, destination)

            commands = root / "commands"
            commands.mkdir()
            (commands / "git").write_text(
                "#!/bin/sh\n"
                'if [ "$1" = -C ]; then shift 2; fi\n'
                'case "$1" in\n'
                "rev-parse) echo 079cbbcbfc00c8d2412a38586a9688ae7eb0009e ;;\n"
                "merge-base|status) exit 0 ;;\n"
                "*) exit 81 ;;\nesac\n"
            )
            (commands / "cargo").write_text("#!/bin/sh\nexit 81\n")
            for command in commands.iterdir():
                command.chmod(0o755)
            target = root / "existing-target"
            target.mkdir()
            sentinel = target / "keep"
            sentinel.write_bytes(b"not owned by runner")
            environment = os.environ.copy()
            for key in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_BOOTSTRAP", "RUSTDOCFLAGS", "ALLOW_EXISTING_TARGET"):
                environment.pop(key, None)
            environment.update(
                PROFILE_FROZEN="1",
                CLEAN_TARGET="1",
                CARGO_TARGET_DIR=str(target / "missing-parent" / ".."),
                TMPDIR=str(root),
                PATH=str(commands) + os.pathsep + environment.get("PATH", ""),
            )
            result = subprocess.run(["bash", str(evidence / "smoke.sh")], env=environment, capture_output=True, text=True, timeout=20)
            self.assertTrue(sentinel.is_file(), "runner deleted a pre-existing normalized target")
            self.assertEqual(result.returncode, 2, result.stderr)
            self.assertIn("existing smoke target", result.stderr)
            self.assertEqual(sentinel.read_bytes(), b"not owned by runner")


if __name__ == "__main__":
    unittest.main()
