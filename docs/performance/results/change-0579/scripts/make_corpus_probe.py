#!/usr/bin/env python3
"""Inject a temporary corpus probe into `crates/litchi-xls/tests/source_backed.rs`.

The probe opens every `.xls`/`.xlt` fixture under `test-data` source-backed
through the file's existing `CountingSource`, and prints the read count, the
byte count, the source-version observation count and a digest of the exact
positional read ranges -- or, for a fixture that is refused, the refusal text.

Running it against two trees that differ only in
`crates/litchi-xls/src/workbook/source.rs` is the corpus-wide control for
change 0579: every field must be identical on both legs.

usage: make_corpus_probe.py <source_backed.rs>   (edits in place)
"""

import sys

PROBE = r'''
#[test]
fn zz_corpus_globals_reads() {
    fn collect(dir: &std::path::Path, files: &mut Vec<std::path::PathBuf>) {
        let mut entries: Vec<_> = std::fs::read_dir(dir)
            .unwrap()
            .map(|entry| entry.unwrap().path())
            .collect();
        entries.sort();
        for path in entries {
            if path.is_dir() {
                collect(&path, files);
            } else if path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    extension.eq_ignore_ascii_case("xls") || extension.eq_ignore_ascii_case("xlt")
                })
            {
                files.push(path);
            }
        }
    }

    fn digest(ranges: &[(u64, usize)]) -> u64 {
        let mut hash = 0xcbf2_9ce4_8422_2325_u64;
        for (offset, length) in ranges {
            for byte in offset.to_le_bytes().iter().chain((*length as u64).to_le_bytes().iter()) {
                hash ^= u64::from(*byte);
                hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
            }
        }
        hash
    }

    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    let mut files = Vec::new();
    collect(&root, &mut files);
    files.sort();

    let mut opened = 0usize;
    let mut total_reads = 0usize;
    let mut total_bytes = 0usize;
    for path in &files {
        let bytes = std::fs::read(path).unwrap();
        let source = Arc::new(CountingSource::new(bytes));
        let relative = path.strip_prefix(&root).unwrap_or(path);
        match SourceBackedWorkbook::from_read_at(source.clone()) {
            Ok(_workbook) => {
                let ranges = source.ranges();
                opened += 1;
                total_reads += ranges.len();
                total_bytes += source.bytes_read();
                println!(
                    "CORPUS {} ok reads={} bytes={} versions={} digest={:016x}",
                    relative.display(),
                    ranges.len(),
                    source.bytes_read(),
                    source.version_calls(),
                    digest(&ranges)
                );
            },
            Err(error) => {
                let ranges = source.ranges();
                total_reads += ranges.len();
                total_bytes += source.bytes_read();
                println!(
                    "CORPUS {} refused reads={} bytes={} versions={} digest={:016x} error={}",
                    relative.display(),
                    ranges.len(),
                    source.bytes_read(),
                    source.version_calls(),
                    digest(&ranges),
                    error
                );
            },
        }
    }
    println!(
        "CORPUSTOTAL files={} opened={} reads={} bytes={}",
        files.len(),
        opened,
        total_reads,
        total_bytes
    );
}
'''

ANCHOR = "#[test]\nfn retained_metadata_queries_observe_the_source_once() {"


def main() -> int:
    target = sys.argv[1]
    with open(target, encoding="utf-8") as handle:
        source = handle.read()
    if "fn zz_corpus_globals_reads()" in source:
        print("probe already present", file=sys.stderr)
        return 1
    if source.count(ANCHOR) != 1:
        print("anchor not found exactly once", file=sys.stderr)
        return 1
    source = source.replace(ANCHOR, PROBE.strip() + "\n\n" + ANCHOR, 1)
    with open(target, "w", encoding="utf-8") as handle:
        handle.write(source)
    print(f"probe injected into {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
