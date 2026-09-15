#!/usr/bin/env python3
"""Insert a throwaway target-scoped strict-layout read counter into soapberry-zip.

Counts the positional reads `IndexedArchive::strict_layout_for` alone issues,
per scenario, on the three fixtures change 0575 froze a prediction for, plus a
corpus-wide accept/refuse census used to check convergence.

The probe calls only APIs that exist both before and after change 0580, so the
same file measures both trees.

Usage: make_strict_scope_probe.py <path to crates/soapberry-zip/src/office.rs>
Run:   cargo test -p soapberry-zip --lib zz_target_scoped -- --nocapture --test-threads=1
"""

import sys

PROBE = r'''
    #[derive(Debug)]
    struct ScopeProbeReaderAt {
        bytes: Vec<u8>,
        reads: Mutex<Vec<(u64, usize, usize)>>,
    }

    impl ScopeProbeReaderAt {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                reads: Mutex::new(Vec::new()),
            }
        }

        fn clear(&self) {
            self.reads.lock().unwrap().clear();
        }

        /// (calls, bytes returned)
        fn totals(&self) -> (usize, usize) {
            let reads = self.reads.lock().unwrap();
            (reads.len(), reads.iter().map(|&(_, _, got)| got).sum())
        }
    }

    impl ReaderAt for ScopeProbeReaderAt {
        fn read_at(&self, buf: &mut [u8], offset: u64) -> io::Result<usize> {
            let start = usize::try_from(offset).unwrap_or(self.bytes.len());
            let count = if start >= self.bytes.len() {
                0
            } else {
                let count = buf.len().min(self.bytes.len() - start);
                buf[..count].copy_from_slice(&self.bytes[start..start + count]);
                count
            };
            self.reads.lock().unwrap().push((offset, buf.len(), count));
            Ok(count)
        }
    }

    fn zz_scope_open(path: &std::path::Path) -> Option<IndexedArchive<ScopeProbeReaderAt>> {
        let bytes = std::fs::read(path).ok()?;
        let length = bytes.len() as u64;
        IndexedArchive::from_reader(ScopeProbeReaderAt::new(bytes), length).ok()
    }

    fn zz_scope_wayfinder(
        indexed: &IndexedArchive<ScopeProbeReaderAt>,
        name: &str,
    ) -> Option<crate::ZipArchiveEntryWayfinder> {
        let entry_id = indexed.entry_id(name)?;
        indexed.indexed_entry(entry_id).ok().map(|e| e.info.wayfinder)
    }

    /// Reads and bytes `strict_layout_for` issues for one ordered scenario,
    /// with the reader's own memoisation in force across the whole scenario.
    fn zz_scope_scenario(
        path: &std::path::Path,
        names: &[&str],
    ) -> Option<(usize, usize, usize, usize)> {
        let indexed = zz_scope_open(path)?;
        let targets: Vec<_> = names.iter().filter_map(|n| zz_scope_wayfinder(&indexed, n)).collect();
        let matched = targets.len();
        indexed.archive.get_ref().clear();
        let mut ok = 0usize;
        for target in targets {
            if indexed.strict_layout_for(target).is_ok() {
                ok += 1;
            }
        }
        let (reads, bytes) = indexed.archive.get_ref().totals();
        Some((reads, bytes, matched, ok))
    }

    /// The same, over every central record in physical order (or reversed).
    fn zz_scope_all(path: &std::path::Path, reverse: bool) -> Option<(usize, usize, usize, usize)> {
        let indexed = zz_scope_open(path)?;
        let mut targets: Vec<_> = indexed.layout.iter().map(|e| e.wayfinder).collect();
        if reverse {
            targets.reverse();
        }
        let matched = targets.len();
        indexed.archive.get_ref().clear();
        let mut ok = 0usize;
        for target in targets {
            if indexed.strict_layout_for(target).is_ok() {
                ok += 1;
            }
        }
        let (reads, bytes) = indexed.archive.get_ref().totals();
        Some((reads, bytes, matched, ok))
    }

    /// Reads issued between construction and the first strict-layout call.
    fn zz_scope_open_and_list(path: &std::path::Path) -> Option<(usize, usize, usize)> {
        let indexed = zz_scope_open(path)?;
        indexed.archive.get_ref().clear();
        let listed = indexed.file_names().count();
        let (reads, bytes) = indexed.archive.get_ref().totals();
        Some((reads, bytes, listed))
    }

    const ZZ_XLSX_ONE: &[&str] = &["xl/worksheets/sheet1.xml"];
    const ZZ_XLSX_CLOSURE: &[&str] = &[
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/_rels/workbook.xml.rels",
        "xl/workbook.xml",
        "xl/worksheets/sheet1.xml",
        "xl/sharedStrings.xml",
        "xl/styles.xml",
    ];
    const ZZ_PPTX_ONE: &[&str] = &["ppt/slides/slide1.xml"];
    const ZZ_PPTX_CLOSURE: &[&str] = &[
        "[Content_Types].xml",
        "_rels/.rels",
        "ppt/_rels/presentation.xml.rels",
        "ppt/presentation.xml",
        "ppt/slides/_rels/slide1.xml.rels",
        "ppt/slides/slide1.xml",
    ];

    #[test]
    fn zz_target_scoped_strict_layout_scenarios() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
        let root = root.canonicalize().unwrap();
        let cases: &[(&str, &[&str], &[&str])] = &[
            ("ooxml/xlsx/sheet-names.xlsx", ZZ_XLSX_ONE, ZZ_XLSX_CLOSURE),
            (
                "ooxml/xlsx/ConditionalFormattingSamples.xlsx",
                ZZ_XLSX_ONE,
                ZZ_XLSX_CLOSURE,
            ),
            ("ooxml/pptx/shapes.pptx", ZZ_PPTX_ONE, ZZ_PPTX_CLOSURE),
            (
                "ooxml/xlsx/universal-content.xlsx",
                ZZ_XLSX_ONE,
                ZZ_XLSX_CLOSURE,
            ),
            ("ooxml/docx/comment.docx", &["word/document.xml"], &["word/document.xml"]),
            (
                "ooxml/pptx/shape-soft-edges.pptx",
                ZZ_PPTX_ONE,
                ZZ_PPTX_CLOSURE,
            ),
        ];
        for (relative, one, closure) in cases {
            let path = root.join(relative);
            if !path.exists() {
                println!("SCOPE {relative} MISSING");
                continue;
            }
            let (r, b, listed) = zz_scope_open_and_list(&path).unwrap();
            println!("SCOPE {relative} scenario=open_list reads={r} bytes={b} members={listed} ok=0");
            for (label, names) in [("one_member", *one), ("closure", *closure)] {
                let (r, b, m, ok) = zz_scope_scenario(&path, names).unwrap();
                println!("SCOPE {relative} scenario={label} reads={r} bytes={b} members={m} ok={ok}");
            }
            let (r, b, m, ok) = zz_scope_all(&path, false).unwrap();
            println!("SCOPE {relative} scenario=all_forward reads={r} bytes={b} members={m} ok={ok}");
            let (r, b, m, ok) = zz_scope_all(&path, true).unwrap();
            println!("SCOPE {relative} scenario=all_reverse reads={r} bytes={b} members={m} ok={ok}");
        }
    }

    #[test]
    fn zz_target_scoped_strict_layout_corpus_census() {
        fn collect(dir: &std::path::Path, files: &mut Vec<std::path::PathBuf>) {
            let Ok(entries) = std::fs::read_dir(dir) else {
                return;
            };
            for entry in entries {
                let path = entry.unwrap().path();
                if path.is_dir() {
                    collect(&path, files);
                } else if path
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| {
                        extension.eq_ignore_ascii_case("docx")
                            || extension.eq_ignore_ascii_case("pptx")
                            || extension.eq_ignore_ascii_case("xlsx")
                    })
                {
                    files.push(path);
                }
            }
        }

        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ooxml");
        let root = root.canonicalize().unwrap();
        let mut files = Vec::new();
        collect(&root, &mut files);
        files.sort();

        let (mut total_reads, mut total_bytes, mut total_members, mut total_ok) = (0, 0, 0, 0);
        let mut fixtures = 0usize;
        for path in &files {
            let Some((reads, bytes, members, ok)) = zz_scope_all(path, false) else {
                continue;
            };
            let relative = path.strip_prefix(&root).unwrap_or(path);
            // A per-target verdict string, so two trees can be diffed member by
            // member rather than only in aggregate.
            let indexed = zz_scope_open(path).unwrap();
            let verdicts: String = indexed
                .layout
                .iter()
                .map(|e| {
                    if indexed.strict_layout_for(e.wayfinder).is_ok() {
                        'A'
                    } else {
                        'R'
                    }
                })
                .collect();
            println!(
                "CENSUS {} members={} reads={} bytes={} accepted={} verdicts={}",
                relative.display(),
                members,
                reads,
                bytes,
                ok,
                verdicts
            );
            fixtures += 1;
            total_reads += reads;
            total_bytes += bytes;
            total_members += members;
            total_ok += ok;
        }
        println!(
            "CENSUSTOTAL fixtures={fixtures} members={total_members} reads={total_reads} \
             bytes={total_bytes} accepted={total_ok}"
        );
    }
'''

ANCHOR = "    #[derive(Debug)]\n    struct ShortWriter {"


def main() -> int:
    target = sys.argv[1]
    with open(target, encoding="utf-8") as handle:
        source = handle.read()
    if "zz_target_scoped_strict_layout_scenarios" in source:
        print("probe already present", file=sys.stderr)
        return 1
    if source.count(ANCHOR) != 1:
        print(f"anchor found {source.count(ANCHOR)} times, expected 1", file=sys.stderr)
        return 1
    source = source.replace(ANCHOR, PROBE.lstrip("\n") + "\n" + ANCHOR, 1)
    with open(target, "w", encoding="utf-8") as handle:
        handle.write(source)
    print("probe inserted into", target)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
