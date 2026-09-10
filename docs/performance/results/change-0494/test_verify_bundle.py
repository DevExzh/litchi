#!/usr/bin/env python3
"""Fixture-only tests for the 0494 bundle verifier."""

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


def _allocation() -> dict[str, object]:
    return {
        'status': 'measured',
        'scope': measure.SAMPLE_ALLOCATION_SCOPE,
        'allocation_calls': 4,
        'deallocation_calls': 3,
        'reallocation_calls': 1,
        'failed_allocation_calls': 0,
        'allocated_bytes': 100,
        'deallocated_bytes': 90,
        'live_bytes_before': 10,
        'live_bytes_after': 20,
        'peak_live_bytes_before': 20,
        'peak_live_bytes_after': 40,
        'region_peak_live_bytes': 30,
    }


def _entries(role: str, rows: int) -> list[dict[str, object]]:
    return [{
        'spec': {'role': role},
        'rows': [{'allocation': _allocation()} for _ in range(rows)],
    }]


class VerifyBundleFixtures(unittest.TestCase):
    def test_exact_command_set_accepts_attempt_suffixed_labels(self) -> None:
        pairs = bundle._validate_final_gates(_manifest())
        self.assertEqual(len(pairs), 8)
        format_commands = [command for _, command in pairs if command and command[0] == 'rustfmt']
        self.assertEqual(len(format_commands), 1)
        self.assertIn('tools/perf-baseline/src/docx_edit_provider/cold.rs', format_commands[0])

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
        with tempfile.TemporaryDirectory(prefix='litchi-bundle-0494-') as raw:
            path = Path(raw) / 'final-gate-1.json'
            path.write_text('{"label":"other-gate"}\n', encoding='utf-8')
            with self.assertRaises(measure.ProviderMatrixError):
                bundle._validate_gate_receipt('final-gate-1', ['true'], path, {})

    def test_python_helper_gate_uses_its_own_source_receipt(self) -> None:
        command = list(bundle._expected_final_commands()[5])
        with tempfile.TemporaryDirectory(prefix='litchi-bundle-0494-') as raw:
            path = Path(raw) / 'final-gate-6.json'
            path.write_text('{"label":"final-gate-6"}\n', encoding='utf-8')
            binding = {'source': {'helper': 'current'}, 'argv': command}
            with patch.object(bundle.measure, '_gate_binding', return_value=binding):
                self.assertEqual(
                    bundle._validate_gate_receipt(
                        'final-gate-6', command, path, {'rust': 'frozen'}
                    ),
                    binding,
                )

    def test_rust_gate_still_requires_frozen_source_receipt(self) -> None:
        command = list(bundle._expected_final_commands()[0])
        with tempfile.TemporaryDirectory(prefix='litchi-bundle-0494-') as raw:
            path = Path(raw) / 'final-gate-1.json'
            path.write_text('{"label":"final-gate-1"}\n', encoding='utf-8')
            binding = {'source': {'current': 'drifted'}, 'argv': command}
            with patch.object(bundle.measure, '_gate_binding', return_value=binding):
                with self.assertRaises(measure.ProviderMatrixError):
                    bundle._validate_gate_receipt(
                        'final-gate-1', command, path, {'frozen': 'source'}
                    )

    def test_recovery_helpers_are_in_custody_inventory(self) -> None:
        self.assertTrue(
            {
                'recover_profile.py', 'test_recover_profile.py',
                'recover_observer_counters.py',
                'test_recover_observer_counters.py',
                'test_profile_strace_status.py',
            }
            <= bundle._expected_helpers()
        )

    def test_postprocessing_attempt_rank_uses_parent_directory(self) -> None:
        self.assertGreater(
            bundle._attempt_rank(Path('/tmp/profiling-counter-recovery-r5/observer-correction.json')),
            bundle._attempt_rank(Path('/tmp/profiling-counter-recovery-r3/observer-correction.json')),
        )

    def test_inventory_rejects_special_files(self) -> None:
        with tempfile.TemporaryDirectory(prefix='litchi-bundle-0494-') as raw:
            fifo = Path(raw) / 'unbound.fifo'
            os.mkfifo(fifo)
            with self.assertRaises(RuntimeError):
                bundle.inventory(Path(raw))

    def test_inventory_excludes_regular_seal_after_creation(self) -> None:
        with tempfile.TemporaryDirectory(prefix='litchi-bundle-0494-') as raw:
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

    def test_cold_allocator_conservation_and_inventory(self) -> None:
        summary = bundle._validate_cold_allocator_entries(
            _entries('allocator', 60), _entries('allocator', 3)
        )
        self.assertEqual(summary, {
            'formal_allocator_rows': 60,
            'pilot_allocator_rows': 3,
            'allocator_rows': 63,
        })

    def test_cold_allocator_rejects_live_balance_tamper(self) -> None:
        allocation = _allocation()
        allocation['live_bytes_after'] = 21
        with self.assertRaises(measure.ProviderMatrixError):
            bundle._validate_cold_allocator_row(allocation, 'fixture.allocation')

    def test_cold_allocator_rejects_peak_and_call_tamper(self) -> None:
        for field, value in (
            ('peak_live_bytes_before', 9),
            ('region_peak_live_bytes', 9),
            ('allocation_calls', 0),
        ):
            allocation = _allocation()
            allocation[field] = value
            with self.subTest(field=field), self.assertRaises(measure.ProviderMatrixError):
                bundle._validate_cold_allocator_row(allocation, 'fixture.allocation')

    def test_cold_allocator_requires_strict_unsigned_integers(self) -> None:
        allocation = _allocation()
        allocation['allocated_bytes'] = True
        with self.assertRaises(measure.ProviderMatrixError):
            bundle._validate_cold_allocator_row(allocation, 'fixture.allocation')

    def test_cold_allocator_does_not_require_deallocation_for_reallocation(self) -> None:
        allocation = _allocation()
        allocation.update({
            'deallocation_calls': 0,
            'allocated_bytes': 100,
            'deallocated_bytes': 0,
            'live_bytes_after': 110,
            'peak_live_bytes_after': 120,
            'region_peak_live_bytes': 115,
        })
        bundle._validate_cold_allocator_row(allocation, 'fixture.allocation')

    def test_cold_allocator_rejects_wrong_formal_inventory(self) -> None:
        with self.assertRaises(measure.ProviderMatrixError):
            bundle._validate_cold_allocator_entries(
                _entries('allocator', 59), _entries('allocator', 3)
            )


if __name__ == '__main__':
    raise SystemExit(unittest.main())
