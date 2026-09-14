S = "/tmp/claude-1001/-home-zhuhe-code-litchi/4b20edf7-6b1c-40a9-9e79-3ca6a15219f0/scratchpad/fat"
import sys

base, out = sys.argv[1], sys.argv[2]
src = open(f"{S}/{base}", encoding="utf-8").read()
probe = '''
    #[test]
    fn zz_corpus_open_reads() {
        fn collect(dir: &std::path::Path, files: &mut Vec<std::path::PathBuf>) {
            for entry in std::fs::read_dir(dir).unwrap() {
                let path = entry.unwrap().path();
                if path.is_dir() {
                    collect(&path, files);
                } else if path
                    .extension()
                    .and_then(|e| e.to_str())
                    .is_some_and(|e| {
                        e.eq_ignore_ascii_case("doc")
                            || e.eq_ignore_ascii_case("xls")
                            || e.eq_ignore_ascii_case("ppt")
                    })
                {
                    files.push(path);
                }
            }
        }
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole");
        let mut files = Vec::new();
        collect(&root, &mut files);
        files.sort();

        let mut total_reads = 0usize;
        let mut total_bytes = 0usize;
        let mut opened = 0usize;
        for path in &files {
            let bytes = std::fs::read(path).unwrap();
            let reader = TrackingReader::new(bytes);
            let Ok(file) = OleFile::open(reader) else {
                continue;
            };
            opened += 1;
            let reads = file.reader.reads.len();
            let read_bytes: usize = file.reader.reads.iter().map(|&(_, l)| l).sum();
            total_reads += reads;
            total_bytes += read_bytes;
            let rel = path.strip_prefix(&root).unwrap_or(path);
            println!("CORPUS {} reads={} bytes={}", rel.display(), reads, read_bytes);
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
anchor = "    #[test]\n    fn batched_sector_reads_place_noncontiguous_sectors_in_order()"
assert src.count(anchor) == 1
src = src.replace(anchor, probe + "\n" + anchor, 1)
open(f"{S}/{out}", "w", encoding="utf-8").write(src)
print("wrote", out)
