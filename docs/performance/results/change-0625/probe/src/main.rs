//! Storage-order determinism probe for `litchi_cfb::writer::OleWriter`.
//!
//! Change 0617 showed that `OleWriter::write_to` iterated a `HashSet` of
//! explicitly created storage paths, so the directory SIDs — and therefore the
//! whole file — depended on the process's hash seed. This probe measures the
//! same property over the repository's real OLE2 fixture corpus, and is built
//! twice: once from the before checkout and once from the fixed worktree.
//!
//! Modes:
//!   determinism <storages> <repeats>
//!       Rebuilds change 0617's synthetic document (N sibling storages, one
//!       stream each) `repeats` times in one process and prints the distinct
//!       digests. Every `OleWriter::new()` constructs a fresh `RandomState`,
//!       whose seed advances per construction, so in-process repeats vary the
//!       hash order exactly as separate processes do.
//!   corpus <root> <repeats>
//!       Finds every CFB artifact under `root`, reads its complete logical model
//!       through the public parser, rebuilds it through `OleWriter` feeding the
//!       storages in source directory order, and prints one JSON line per
//!       fixture with the distinct output digests over `repeats` rebuilds.
//!   rebuild <file> <iterations>
//!       Rebuilds one fixture `iterations` times and prints the last digest.
//!       Two runs at different iteration counts form a callgrind isolation pair.

use litchi_cfb::consts::{STGTY_ROOT, STGTY_STORAGE, STGTY_STREAM};
use litchi_cfb::writer::OleWriter;
use litchi_cfb::{OleError, OleFile, is_ole_file};
use std::collections::BTreeSet;
use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};

/// FNV-1a with the standard 64-bit prime.
fn fnv1a(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}

/// The multiplier change 0617's probe actually used: sixteen times the FNV-1a
/// prime. Recomputing it here lets this probe's digests be matched against
/// `docs/performance/results/change-0617/determinism.txt`.
fn fnv1a_0617(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x1000_0000_01b3);
    }
    hash
}

fn refs(path: &[String]) -> Vec<&str> {
    path.iter().map(String::as_str).collect()
}

/// The logical model of one compound file: its sector size, every storage path
/// in source directory order, and every stream path with its bytes.
struct Model {
    sector_size: usize,
    storages: Vec<Vec<String>>,
    streams: Vec<(Vec<String>, Vec<u8>)>,
}

impl Model {
    /// Two storage paths are incomparable when neither is a prefix of the other.
    /// A document with no incomparable pair cannot be reordered by the hash
    /// seed, because `add_storage_path` creates ancestors on the way down.
    fn has_incomparable_storages(&self) -> bool {
        for (index, left) in self.storages.iter().enumerate() {
            for right in &self.storages[index + 1..] {
                if !left.starts_with(right.as_slice()) && !right.starts_with(left.as_slice()) {
                    return true;
                }
            }
        }
        false
    }

    fn rebuild(&self) -> Result<Vec<u8>, OleError> {
        let mut writer = OleWriter::with_sector_size(self.sector_size)?;
        for storage in &self.storages {
            writer.create_storage(&refs(storage))?;
        }
        for (path, data) in &self.streams {
            writer.create_stream(&refs(path), data)?;
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output)?;
        Ok(output.into_inner())
    }
}

fn collect(
    file: &OleFile<Cursor<Vec<u8>>>,
    path: &mut Vec<String>,
    storages: &mut Vec<Vec<String>>,
    streams: &mut Vec<Vec<String>>,
) -> Result<(), OleError> {
    let children: Vec<(String, u8)> = file
        .list_directory_entries(&refs(path))?
        .into_iter()
        .map(|entry| (entry.name.clone(), entry.entry_type))
        .collect();
    for (name, entry_type) in children {
        path.push(name);
        if entry_type == STGTY_STORAGE || entry_type == STGTY_ROOT {
            storages.push(path.clone());
            collect(file, path, storages, streams)?;
        } else if entry_type == STGTY_STREAM {
            streams.push(path.clone());
        }
        path.pop();
    }
    Ok(())
}

fn model(bytes: Vec<u8>) -> Result<Model, OleError> {
    let mut file = OleFile::open(Cursor::new(bytes))?;
    let sector_size = file.sector_size();
    let mut storages = Vec::new();
    let mut stream_paths = Vec::new();
    collect(&file, &mut Vec::new(), &mut storages, &mut stream_paths)?;
    let mut streams = Vec::new();
    for path in stream_paths {
        let data = file.open_stream(&refs(&path))?;
        streams.push((path, data));
    }
    Ok(Model {
        sector_size,
        storages,
        streams,
    })
}

fn walk(root: &Path, found: &mut Vec<PathBuf>) {
    let Ok(entries) = fs::read_dir(root) else {
        return;
    };
    let mut paths: Vec<PathBuf> = entries.flatten().map(|entry| entry.path()).collect();
    paths.sort();
    for path in paths {
        if path.is_dir() {
            walk(&path, found);
        } else if path.is_file() {
            found.push(path);
        }
    }
}

fn escape(value: &str) -> String {
    value.replace('\\', "\\\\").replace('"', "\\\"")
}

fn corpus(root: &Path, repeats: usize) {
    let mut files = Vec::new();
    walk(root, &mut files);
    let mut considered = 0usize;
    let mut opened = 0usize;
    let mut refused = 0usize;
    for path in files {
        let Ok(bytes) = fs::read(&path) else {
            continue;
        };
        if !is_ole_file(&bytes) {
            continue;
        }
        considered += 1;
        let relative = path
            .strip_prefix(root)
            .unwrap_or(&path)
            .display()
            .to_string();
        let model = match model(bytes) {
            Ok(model) => model,
            Err(error) => {
                refused += 1;
                println!(
                    "{{\"file\":\"{}\",\"status\":\"parse-refused\",\"error\":\"{}\"}}",
                    escape(&relative),
                    escape(&format!("{error:?}"))
                );
                continue;
            },
        };
        let mut digests = BTreeSet::new();
        let mut lengths = BTreeSet::new();
        let mut failure = None;
        for _ in 0..repeats {
            match model.rebuild() {
                Ok(output) => {
                    lengths.insert(output.len());
                    digests.insert(format!("{:016x}", fnv1a(&output)));
                },
                Err(error) => {
                    failure = Some(format!("{error:?}"));
                    break;
                },
            }
        }
        if let Some(error) = failure {
            refused += 1;
            println!(
                "{{\"file\":\"{}\",\"status\":\"rebuild-refused\",\"storages\":{},\"streams\":{},\"error\":\"{}\"}}",
                escape(&relative),
                model.storages.len(),
                model.streams.len(),
                escape(&error)
            );
            continue;
        }
        opened += 1;
        let digest_list = digests
            .iter()
            .map(|digest| format!("\"{digest}\""))
            .collect::<Vec<_>>()
            .join(",");
        let length_list = lengths
            .iter()
            .map(usize::to_string)
            .collect::<Vec<_>>()
            .join(",");
        println!(
            "{{\"file\":\"{}\",\"status\":\"ok\",\"sector_size\":{},\"storages\":{},\"streams\":{},\"incomparable\":{},\"lengths\":[{}],\"distinct\":{},\"digests\":[{}]}}",
            escape(&relative),
            model.sector_size,
            model.storages.len(),
            model.streams.len(),
            model.has_incomparable_storages(),
            length_list,
            digests.len(),
            digest_list
        );
    }
    eprintln!("considered={considered} rebuilt={opened} refused={refused} repeats={repeats}");
}

fn determinism(count: usize, repeats: usize) {
    let mut digests = BTreeSet::new();
    for _ in 0..repeats {
        let mut writer = OleWriter::new();
        for index in 0..count {
            let name = format!("Storage{index:02}");
            writer.create_storage(&[name.as_str()]).expect("storage");
            writer
                .create_stream(
                    &[name.as_str(), "Payload"],
                    format!("payload-{index}").as_bytes(),
                )
                .expect("stream");
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).expect("write");
        let output = output.into_inner();
        digests.insert((
            output.len(),
            format!("{:016x}", fnv1a(&output)),
            format!("{:016x}", fnv1a_0617(&output)),
        ));
    }
    for (bytes, digest, digest_0617) in &digests {
        println!(
            "{{\"storages\":{count},\"repeats\":{repeats},\"bytes\":{bytes},\"fnv1a\":\"{digest}\",\"fnv1a_0617_multiplier\":\"{digest_0617}\"}}"
        );
    }
    println!("{{\"storages\":{count},\"repeats\":{repeats},\"distinct\":{}}}", digests.len());
}

fn rebuild(path: &Path, iterations: usize) {
    let bytes = fs::read(path).expect("fixture");
    let model = model(bytes).expect("model");
    let mut digest = 0u64;
    let mut length = 0usize;
    for _ in 0..iterations {
        let output = model.rebuild().expect("rebuild");
        length = output.len();
        digest = fnv1a(&output);
    }
    println!(
        "{{\"file\":\"{}\",\"iterations\":{iterations},\"bytes\":{length},\"fnv1a\":\"{digest:016x}\"}}",
        escape(&path.display().to_string())
    );
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    match args.get(1).map(String::as_str) {
        Some("determinism") => {
            let count = args[2].parse().expect("storage count");
            let repeats = args[3].parse().expect("repeats");
            determinism(count, repeats);
        },
        Some("corpus") => {
            let repeats = args[3].parse().expect("repeats");
            corpus(Path::new(&args[2]), repeats);
        },
        Some("rebuild") => {
            let iterations = args[3].parse().expect("iterations");
            rebuild(Path::new(&args[2]), iterations);
        },
        _ => {
            eprintln!("usage: cfb_storage_order_probe <determinism N R|corpus ROOT R|rebuild FILE N>");
            std::process::exit(2);
        },
    }
}
