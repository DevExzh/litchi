#!/usr/bin/env python3
"""Insert a throwaway strict-layout-proof read counter into soapberry-zip.

Usage: make_proof_probe.py <path to crates/soapberry-zip/src/office.rs>

The probe walks every DOCX/PPTX/XLSX fixture under `test-data/ooxml`, indexes
it over a positional reader that records every `read_at`, resets the counters,
builds one whole-archive strict layout proof through the private
`IndexedArchive::build_strict_layout_proof`, and prints the reads and
bytes that proof alone issued. It also reports how many members needed the
heap fallback, measured from the local headers themselves.

Run with:  cargo test -p soapberry-zip --lib zz_strict_layout_proof -- --nocapture
"""

import sys

PROBE = r'''
    #[derive(Debug)]
    struct ProofCountingReaderAt {
        bytes: Vec<u8>,
        reads: Mutex<Vec<(u64, usize, usize)>>,
    }

    impl ProofCountingReaderAt {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                reads: Mutex::new(Vec::new()),
            }
        }

        fn clear(&self) {
            self.reads.lock().unwrap().clear();
        }

        /// Returns (calls, bytes returned, bytes requested).
        fn totals(&self) -> (usize, usize, usize) {
            let reads = self.reads.lock().unwrap();
            (
                reads.len(),
                reads.iter().map(|&(_, _, got)| got).sum(),
                reads.iter().map(|&(_, want, _)| want).sum(),
            )
        }
    }

    impl ReaderAt for ProofCountingReaderAt {
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

    /// The local variable-region length of every member, read straight out of
    /// the raw bytes so the fallback count does not depend on the code under
    /// measurement.
    fn local_variable_lengths(bytes: &[u8]) -> Vec<usize> {
        let mut out = Vec::new();
        let eocd = match bytes
            .windows(4)
            .rposition(|window| window == [0x50, 0x4b, 0x05, 0x06])
        {
            Some(position) => position,
            None => return out,
        };
        let total = u16::from_le_bytes([bytes[eocd + 10], bytes[eocd + 11]]) as usize;
        let mut cursor = u32::from_le_bytes([
            bytes[eocd + 16],
            bytes[eocd + 17],
            bytes[eocd + 18],
            bytes[eocd + 19],
        ]) as usize;
        for _ in 0..total {
            if bytes.get(cursor..cursor + 4) != Some(&[0x50, 0x4b, 0x01, 0x02]) {
                return out;
            }
            let name = u16::from_le_bytes([bytes[cursor + 28], bytes[cursor + 29]]) as usize;
            let extra = u16::from_le_bytes([bytes[cursor + 30], bytes[cursor + 31]]) as usize;
            let comment = u16::from_le_bytes([bytes[cursor + 32], bytes[cursor + 33]]) as usize;
            let local = u32::from_le_bytes([
                bytes[cursor + 42],
                bytes[cursor + 43],
                bytes[cursor + 44],
                bytes[cursor + 45],
            ]) as usize;
            cursor += 46 + name + extra + comment;
            if bytes.get(local..local + 4) != Some(&[0x50, 0x4b, 0x03, 0x04]) {
                return out;
            }
            let local_name = u16::from_le_bytes([bytes[local + 26], bytes[local + 27]]) as usize;
            let local_extra = u16::from_le_bytes([bytes[local + 28], bytes[local + 29]]) as usize;
            out.push(local_name + local_extra);
        }
        out
    }

    #[test]
    fn zz_strict_layout_proof_corpus_reads() {
        fn collect(dir: &std::path::Path, files: &mut Vec<std::path::PathBuf>) {
            for entry in std::fs::read_dir(dir).unwrap() {
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

        let (mut total_reads, mut total_bytes, mut total_requested) = (0usize, 0usize, 0usize);
        let (mut total_members, mut total_spilled, mut proven) = (0usize, 0usize, 0usize);
        for path in &files {
            let bytes = std::fs::read(path).unwrap();
            let length = bytes.len() as u64;
            let variable_lengths = local_variable_lengths(&bytes);
            let reader = ProofCountingReaderAt::new(bytes);
            let Ok(indexed) = IndexedArchive::from_reader(reader, length) else {
                continue;
            };
            let Some(first) = indexed.layout.first().map(|entry| entry.wayfinder) else {
                continue;
            };
            indexed.archive.get_ref().clear();
            let Ok(_proof) = indexed.build_strict_layout_proof(first) else {
                continue;
            };
            let (reads, got, requested) = indexed.archive.get_ref().totals();
            let members = indexed.layout.len();
            let spilled = variable_lengths
                .iter()
                .filter(|&&variable| 30 + variable > 640)
                .count();
            proven += 1;
            total_reads += reads;
            total_bytes += got;
            total_requested += requested;
            total_members += members;
            total_spilled += spilled;
            let relative = path.strip_prefix(&root).unwrap_or(path);
            println!(
                "PROOF {} members={} reads={} bytes={} requested={} spill_candidates={}",
                relative.display(),
                members,
                reads,
                got,
                requested,
                spilled
            );
        }
        println!(
            "PROOFTOTAL files={} proven={} members={} reads={} bytes={} requested={} \
             spill_candidates={}",
            files.len(),
            proven,
            total_members,
            total_reads,
            total_bytes,
            total_requested,
            total_spilled
        );
    }
'''

ANCHOR = "    #[derive(Debug)]\n    struct ShortWriter {"


def main() -> int:
    target = sys.argv[1]
    with open(target, encoding="utf-8") as handle:
        source = handle.read()
    if "zz_strict_layout_proof_corpus_reads" in source:
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
