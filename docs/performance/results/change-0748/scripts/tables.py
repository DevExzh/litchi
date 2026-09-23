#!/usr/bin/env python3
"""Render an abba.py summary as Markdown tables and list every >5% regression.

usage: tables.py SUMMARY_JSON [ISOLATION_JSON]
"""
import json
import sys


def ms(value):
    return f"{value / 1e6:.4f}" if value < 1e7 else f"{value / 1e6:.2f}"


def main():
    summary = json.load(open(sys.argv[1]))
    print('| case | before p50 (ms) | after p50 (ms) | paired p50 ratio | 95% CI | before p95 | after p95 | mean ratio |')
    print('| --- | ---: | ---: | ---: | --- | ---: | ---: | ---: |')
    flags = []
    for record in summary:
        lo, hi = record['bootstrap_95ci_p50_ratio']
        print(f"| `{record['case']}` | {ms(record['before_median_p50'])} | {ms(record['after_median_p50'])} | "
              f"{record['median_paired_p50_ratio']:.3f} | [{lo:.3f}, {hi:.3f}] | "
              f"{ms(record['before_median_p95'])} | {ms(record['after_median_p95'])} | "
              f"{record['median_paired_mean_ratio']:.3f} |")
        for process_pair in range(4):
            ratio = record['paired_p50_ratios'][process_pair]
            if ratio > 1.05:
                flags.append((record['case'], 'pair', process_pair, 'p50', ratio))
            mean_ratio = record['paired_mean_ratios'][process_pair]
            if mean_ratio > 1.05:
                flags.append((record['case'], 'pair', process_pair, 'mean', mean_ratio))
        if record['median_paired_p50_ratio'] > 1.05:
            flags.append((record['case'], 'median', None, 'p50', record['median_paired_p50_ratio']))
        p95_ratio = record['after_median_p95'] / record['before_median_p95']
        if p95_ratio > 1.05:
            flags.append((record['case'], 'median', None, 'p95', p95_ratio))
    print()
    print('Regression flags (>5%):' if flags else 'Regression flags (>5%): none')
    for flag in flags:
        print('  ', flag)
    if len(sys.argv) > 2:
        isolation = json.load(open(sys.argv[2]))
        print()
        print('| case | before instructions (M) | after (M) | ratio | before cycles (M) | after (M) | ratio |')
        print('| --- | ---: | ---: | ---: | ---: | ---: | ---: |')
        for record in isolation['results']:
            print(f"| `{record['case']}` | {record['before_instructions'] / 1e6:.3f} | "
                  f"{record['after_instructions'] / 1e6:.3f} | {record['instructions_ratio']:.3f} | "
                  f"{record['before_cycles'] / 1e6:.3f} | {record['after_cycles'] / 1e6:.3f} | "
                  f"{record['cycles_ratio']:.3f} |")


if __name__ == '__main__':
    main()
