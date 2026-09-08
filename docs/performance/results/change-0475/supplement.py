#!/usr/bin/env python3
"""Correct the empty demangled Heaptrack filter without altering frozen attempts."""
import json
from pathlib import Path
import capture

ROOT = Path(__file__).resolve().parent
TOKEN = '21pptx_streaming_create3run'


def main():
    declaration = json.loads((ROOT / 'supplement-protocol.json').read_text())
    assert declaration['driver_sha256'] == capture.sha(Path(__file__))
    assert declaration['filter_token'] == TOKEN
    capture.check()
    records = []
    for lane in ('heap-H1', 'heap-H2'):
        out = ROOT / lane
        argv = ['heaptrack_print', '-f', str(out / 'heaptrack.zst'), '-t', '0', '-m', '0',
            '-p', '0', '-T', '0', '-l', '0', '--filter-bt-function', TOKEN, '-n', '100',
            '-F', str(out / 'runner-mangled-stacks.txt'), '--flamegraph-cost-type', 'allocations']
        assert capture.command(argv, out / 'runner-mangled',
            supplemental_protocol_sha256=capture.sha(ROOT / 'supplement-protocol.json'),
            supplemental_driver_sha256=capture.sha(Path(__file__))) == 0
        records += [capture.compress(out / 'runner-mangled.stdout'),
                    capture.compress(out / 'runner-mangled-stacks.txt')]
    capture.write(ROOT / 'supplement-compression.json', dict(
        schema='litchi-0475-compression-v1', artifacts=records))
    capture.check()


if __name__ == '__main__':
    main()
