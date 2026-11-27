import argparse
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from tools.gate import chunked_cargo_tests as runner
from tools.gate import target_headroom


def target(kind, name, *, test=True, doctest=False, requires=()):
    return {
        'kind': [kind], 'name': name, 'test': test, 'doctest': doctest,
        'required-features': list(requires),
    }


class CargoTestBatches(unittest.TestCase):
    def test_batches_preserve_feature_unification_and_select_each_name_once(self):
        packages = [
            {'name': 'a', 'targets': [
                target('lib', 'a', doctest=True), target('test', 'shared'),
                target('example', 'optional', test=False, requires=['formats']),
            ]},
            {'name': 'b', 'targets': [
                target('test', 'shared'), target('test', 'another'),
                target('example', 'demo', test=False, requires=['formats']),
            ]},
        ]
        selection = ['-p', 'a', '-p', 'b', '--features', 'b/formats']
        enabled = {'a': set(), 'b': {'formats'}}
        tests = [
            ('test', [*selection, '--lib']),
            ('test', [*selection, '--test', 'another']),
            ('test', [*selection, '--test', 'shared']),
        ]
        self.assertEqual(runner.plan(packages, selection, enabled, 1, True), tests)
        self.assertEqual(runner.plan(packages, selection, enabled, 1, False), [
            *tests,
            ('build', [*selection, '--profile', 'test', '--example', 'demo']),
            ('test', [*selection, '--doc']),
        ])

    def workspace_plan(self):
        packages = [
            {'name': 'a', 'targets': [
                target('lib', 'a', doctest=True), target('bin', 'tool'),
                *(target('test', f'a{number}') for number in range(9)),
                target('example', 'harness'), target('example', 'demo', test=False),
            ]},
            {'name': 'b', 'targets': [
                target('lib', 'b'), *(target('test', f'b{number}') for number in range(7)),
            ]},
        ]
        selection = ['--workspace', '--all-features']
        enabled = {'a': set(), 'b': set()}
        return selection, runner.plan(packages, selection, enabled, 2, False)

    def test_partitions_are_disjoint_cover_the_plan_and_keep_the_selection(self):
        selection, commands = self.workspace_plan()
        tests = [command for command in commands if not runner.is_lib_command(command)]
        self.assertGreater(len(tests), 6)
        for count in range(1, 7):
            with self.subTest(count=count):
                partitions = [runner.select(commands, 'tests', (index, count))
                              for index in range(1, count + 1)]
                selected = [command for partition in partitions for command in partition]
                self.assertEqual(len(selected), len(tests))
                self.assertEqual(sorted(selected), sorted(tests))
                for partition in partitions:
                    self.assertTrue(partition)
                    for _verb, arguments in partition:
                        self.assertEqual(arguments[:len(selection)], selection)

    def test_lib_and_tests_parts_make_up_the_whole_plan(self):
        _selection, commands = self.workspace_plan()
        lib = runner.select(commands, 'lib', (1, 1))
        self.assertEqual(lib, [commands[0]])
        self.assertTrue(runner.is_lib_command(commands[0]))
        self.assertEqual(runner.select(commands, 'all', (1, 1)), commands)
        self.assertEqual([*lib, *runner.select(commands, 'tests', (1, 1))], commands)
        # A partition does not split the --lib command off its part.
        self.assertEqual(runner.select(commands, 'lib', (2, 3)), lib)

    def test_partition_argument_requires_index_within_count(self):
        self.assertEqual(runner.parse_partition('2/3'), (2, 3))
        self.assertEqual(runner.parse_partition('1/1'), (1, 1))
        for value in ('0/3', '4/3', '1/0', '3', '1/2/3', 'a/b', ''):
            with self.subTest(value=value), self.assertRaises(argparse.ArgumentTypeError):
                runner.parse_partition(value)

    def test_disk_usage_without_unix_allocated_block_counts(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'artifact').write_bytes(b'abc')
            with patch.object(target_headroom.os, 'lstat', return_value=SimpleNamespace(st_size=3)):
                self.assertEqual(target_headroom.target_bytes(root), 3)
