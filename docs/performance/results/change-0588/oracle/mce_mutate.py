"""Adversarial differential for change 0588: refusal identity under mutation.

Real fixtures never refuse, so the corpus oracle cannot compare refusals. This
takes MCE-bearing seeds, applies single-byte substitutions, deletions,
insertions and truncations, and requires the BEFORE and AFTER codecs to agree
on the canonical form of every accepted mutant and on the exact Debug identity
and Display message of every refused one.

usage: mce_mutate.py <before-bin> <after-bin> <seed-dir> <count> <report.tsv>
"""

import concurrent.futures as cf
import os
import random
import subprocess
import sys

BEFORE, AFTER, SEEDS, COUNT, REPORT = (
    sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4]), sys.argv[5])
MODE = sys.argv[6] if len(sys.argv) > 6 else "canon"
INTERESTING = b'<>/&";\':= \tmcxIgnorableAlternateContentChoiceFallback0\xff\x00'


def canon(binary, data):
    proc = subprocess.run([binary, MODE, "-"], input=data,
                          stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=600)
    if proc.returncode != 0:
        return b"PROBE-FAILED\t%d" % proc.returncode
    return proc.stdout


def mutate(rng, seed):
    data = bytearray(seed)
    for _ in range(rng.randint(1, 3)):
        if not data:
            break
        kind = rng.randrange(4)
        at = rng.randrange(len(data))
        if kind == 0:
            data[at] = rng.choice(INTERESTING)
        elif kind == 1:
            del data[at]
        elif kind == 2:
            data.insert(at, rng.choice(INTERESTING))
        else:
            del data[at:]
    return bytes(data)


def job(args):
    index, path, seed = args
    rng = random.Random(index)
    data = mutate(rng, seed)
    before, after = canon(BEFORE, data), canon(AFTER, data)
    refused = before.startswith(b"ERR\t")
    if before != after:
        return ("MISMATCH", path, str(index),
                "before=%r after=%r" % (before[:200], after[:200])), refused
    return None, refused


def main():
    seeds = []
    for name in sorted(os.listdir(SEEDS)):
        full = os.path.join(SEEDS, name)
        if os.path.isfile(full):
            seeds.append((full, open(full, "rb").read()))
    assert seeds, "no seeds"
    work = [(i, seeds[i % len(seeds)][0], seeds[i % len(seeds)][1]) for i in range(COUNT)]
    findings, refusals = [], 0
    with cf.ThreadPoolExecutor(max_workers=8) as pool:
        for finding, refused in pool.map(job, work):
            if finding:
                findings.append(finding)
            refusals += int(refused)
    with open(REPORT, "w") as out:
        out.write("kind\tseed\tmutant\tdetail\n")
        for row in findings:
            out.write("\t".join(row) + "\n")
        out.write("#summary\tseeds=%d\tmutants=%d\trefusals=%d\tmismatches=%d\n"
                  % (len(seeds), COUNT, refusals, len(findings)))
    print("seeds=%d mutants=%d refusals=%d mismatches=%d"
          % (len(seeds), COUNT, refusals, len(findings)))


main()
