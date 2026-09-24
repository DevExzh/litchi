"""Exercise smoke target ownership with disposable source and build fixtures."""

import hashlib
import os
from pathlib import Path
import shutil
import subprocess
from tempfile import TemporaryDirectory
import unittest

import profile_pins


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
            (evidence / "profile_pins.py").write_text((here / "profile_pins.py").read_text())

            def digest(path: Path) -> str:
                return hashlib.sha256(path.read_bytes()).hexdigest()

            arm = next(
                arm
                for arm, hashes in profile_pins.SOURCE_HASHES.items()
                if digest(repository / "crates/litchi-drawingml/src/ink/actions.rs")
                == hashes["crates/litchi-drawingml/src/ink/actions.rs"]
            )
            for relative in profile_pins.SOURCE_HASHES[arm]:
                destination = root / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(repository / relative, destination)

            commands = root / "commands"
            commands.mkdir()
            (commands / "git").write_text(
                "#!/bin/sh\n"
                'if [ "$1" = -C ]; then shift 2; fi\n'
                'case "$1" in\n'
                f"rev-parse) echo {profile_pins.SOURCE_PINS[arm]} ;;\n"
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
                PROFILE_ARM=arm,
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
