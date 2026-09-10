#!/usr/bin/env python3
"""Fixture-only tests for the 0493 bundle verifier."""

from __future__ import annotations

import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import measure
import verify_bundle as bundle


def _manifest() -> dict[str, object]:
    commands = {
        f'final-gate-{index + 1}': list(command)
        for index, command in enumerate(bundle._expected_final_commands())
    }
    return {
        'schema': bundle.FINAL_GATES_SCHEMA,
        'version': bundle.FINAL_GATES_VERSION,
        'commands': commands,
    }


class VerifyBundleFixtures(unittest.TestCase):
    def test_exact_command_set_accepts_attempt_suffixed_labels(self) -> None:
        pairs = bundle._validate_final_gates(_manifest())
        self.assertEqual(len(pairs), 12)

    def test_exact_command_set_rejects_missing_or_arbitrary_gate(self) -> None:
        missing = _manifest()
        missing['commands'].pop('final-gate-1')  # type: ignore[union-attr]
        with self.assertRaises(measure.ProviderMatrixError):
            bundle._validate_final_gates(missing)

        arbitrary = _manifest()
        arbitrary['commands']['final-gate-1'] = ['true']  # type: ignore[index]
        with self.assertRaises(measure.ProviderMatrixError):
            bundle._validate_final_gates(arbitrary)

    def test_gate_receipt_label_is_bound_to_filename(self) -> None:
        with tempfile.TemporaryDirectory(prefix='litchi-bundle-0493-') as raw:
            path = Path(raw) / 'final-gate-1.json'
            path.write_text('{"label":"other-gate"}\n', encoding='utf-8')
            with self.assertRaises(measure.ProviderMatrixError):
                bundle._validate_gate_receipt('final-gate-1', ['true'], path, {})

    def test_inventory_rejects_special_files(self) -> None:
        with tempfile.TemporaryDirectory(prefix='litchi-bundle-0493-') as raw:
            fifo = Path(raw) / 'unbound.fifo'
            os.mkfifo(fifo)
            with self.assertRaises(RuntimeError):
                bundle.inventory(Path(raw))

    def test_inventory_excludes_regular_seal_after_creation(self) -> None:
        with tempfile.TemporaryDirectory(prefix='litchi-bundle-0493-') as raw:
            root = Path(raw)
            (root / 'evidence.json').write_text('{}\n', encoding='utf-8')
            (root / 'seal.json').write_text('{}\n', encoding='utf-8')
            self.assertEqual(bundle.inventory(root), {'evidence.json': bundle.meta(root / 'evidence.json')})

    def test_cleanup_uses_canonical_verifier(self) -> None:
        proof = {'schema': 'cleanup-proof', 'status': 'pass'}
        with patch.object(bundle.cleanup_driver, 'verify', return_value=proof) as verifier:
            self.assertIs(bundle._verify_cleanup(), proof)
        verifier.assert_called_once_with(
            root=bundle.ROOT, temp=bundle.TEMP, target=bundle.TARGET_DIR
        )

    def test_cleanup_failure_is_rejected(self) -> None:
        with patch.object(
            bundle.cleanup_driver, 'verify', return_value={'status': 'failed'}
        ):
            with self.assertRaises(measure.ProviderMatrixError):
                bundle._verify_cleanup()


if __name__ == '__main__':
    raise SystemExit(unittest.main())
