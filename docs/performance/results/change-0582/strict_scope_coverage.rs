//! Coverage probe for change 0582.
//!
//! This is **not** the differential build.  It is the `after` tree with two
//! atomic counters added to `strict_layout_for` and to the proof builder, so
//! the harness can report how many (input, member, API) triples actually enter
//! the strict-layout proof.  A large corpus that never reaches the changed code
//! would not be evidence, and this is what separates the two.
//!
//! Usage: <corpus-dir> <report-path>

use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::Ordering;

use soapberry_zip::ReaderAt;
use soapberry_zip::office::{
    ArchiveLimits, ArchiveReader, IndexedArchive, PROBE_STRICT_BUILDS, PROBE_STRICT_TARGETS,
};

fn fuzz_limits() -> ArchiveLimits {
    ArchiveLimits {
        max_files: 256,
        max_member_name_bytes: 4 << 10,
        max_metadata_bytes: 64 << 10,
        max_compressed_size: 1 << 20,
        max_entry_size: 1 << 20,
        max_total_size: 1 << 20,
    }
}

fn wide_limits() -> ArchiveLimits {
    ArchiveLimits {
        max_files: 65_535,
        max_member_name_bytes: 64 << 10,
        max_metadata_bytes: 16 << 20,
        max_compressed_size: 256 << 20,
        max_entry_size: 256 << 20,
        max_total_size: 1 << 30,
    }
}

struct NullSink;

impl Write for NullSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        Ok(input.len())
    }
    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// The same four strict-layout entry points the differential harness sweeps,
/// in the same order, on the same fresh-reader schedule.
fn sweep(data: &[u8], limits: ArchiveLimits) {
    if let Ok(reader) = ArchiveReader::new_with_limits(data, limits) {
        let mut names: Vec<String> = reader.file_names().map(str::to_string).collect();
        names.sort();
        names.dedup();
        for name in &names {
            let _ = reader.read_to(name, &mut NullSink);
        }
        if let Ok(fresh) = ArchiveReader::new_with_limits(data, limits) {
            for name in names.iter().rev() {
                let _ = fresh.read_to(name, &mut NullSink);
            }
        }
        for name in &names {
            let _ = reader.read_to(name, &mut NullSink);
        }
    }
    let Ok(archive) = IndexedArchive::from_reader_with_limits(data, data.len() as u64, limits)
    else {
        return;
    };
    let mut names: Vec<String> = archive.file_names().map(str::to_string).collect();
    names.sort();
    names.dedup();
    for name in &names {
        if let Some(entry_id) = archive.entry_id(name) {
            let _ = archive.read_entry_to(entry_id, &mut NullSink);
        }
    }
    if let Ok(fresh) = IndexedArchive::from_reader_with_limits(data, data.len() as u64, limits) {
        for name in names.iter().rev() {
            if let Some(entry_id) = fresh.entry_id(name) {
                let _ = fresh.read_entry_to(entry_id, &mut NullSink);
            }
        }
    }
    for name in &names {
        if let Some(entry_id) = archive.entry_id(name) {
            let _ = archive.read_entry_to(entry_id, &mut NullSink);
        }
    }
    if let Ok(verified) = IndexedArchive::from_reader_with_limits(data, data.len() as u64, limits) {
        for name in &names {
            if let Some(entry_id) = verified.entry_id(name) {
                let _ = verified.with_verified_entry_reader(entry_id, |reader| {
                    io::copy(reader, &mut NullSink)
                });
            }
        }
    }
    if let Ok(pre) = IndexedArchive::from_reader_with_limits(data, data.len() as u64, limits) {
        for name in &names {
            if let Some(entry_id) = pre.entry_id(name) {
                let mut events = 0usize;
                let _ = pre.read_entry_precompressed_and_decoded_with_progress(entry_id, |_| {
                    events += 1;
                    if events > 8 { Err(()) } else { Ok(()) }
                });
            }
        }
    }
}

fn collect(root: &Path) -> Vec<PathBuf> {
    let mut found = std::collections::BTreeSet::new();
    let mut stack = vec![root.to_path_buf()];
    while let Some(dir) = stack.pop() {
        let Ok(entries) = std::fs::read_dir(&dir) else { continue };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
            } else if path.is_file() {
                found.insert(path);
            }
        }
    }
    found.into_iter().collect()
}

fn main() {
    let mut args = std::env::args_os().skip(1);
    let corpus = PathBuf::from(args.next().expect("usage: <corpus-dir> <report-path>"));
    let report = PathBuf::from(args.next().expect("usage: <corpus-dir> <report-path>"));
    let mut out = String::new();
    let mut inputs_with_targets = 0u64;
    let mut grand_targets = 0u64;
    let mut grand_builds = 0u64;
    let paths = collect(&corpus);
    for path in &paths {
        let relative = path
            .strip_prefix(&corpus)
            .unwrap_or(path)
            .to_string_lossy()
            .replace('\\', "/");
        let Ok(data) = std::fs::read(path) else { continue };
        PROBE_STRICT_TARGETS.store(0, Ordering::Relaxed);
        PROBE_STRICT_BUILDS.store(0, Ordering::Relaxed);
        sweep(&data, fuzz_limits());
        let fuzz_targets = PROBE_STRICT_TARGETS.swap(0, Ordering::Relaxed);
        let fuzz_builds = PROBE_STRICT_BUILDS.swap(0, Ordering::Relaxed);
        sweep(&data, wide_limits());
        let wide_targets = PROBE_STRICT_TARGETS.swap(0, Ordering::Relaxed);
        let wide_builds = PROBE_STRICT_BUILDS.swap(0, Ordering::Relaxed);
        out.push_str(&format!(
            "{relative}\t{fuzz_targets}\t{fuzz_builds}\t{wide_targets}\t{wide_builds}\n"
        ));
        if fuzz_targets + wide_targets > 0 {
            inputs_with_targets += 1;
        }
        grand_targets += fuzz_targets + wide_targets;
        grand_builds += fuzz_builds + wide_builds;
    }
    out.push_str(&format!(
        "TOTAL\tinputs={}\tinputs_entering_proof={}\tproof_entries={}\tproof_builds={}\n",
        paths.len(),
        inputs_with_targets,
        grand_targets,
        grand_builds
    ));
    std::fs::write(&report, out).expect("report write");
    eprintln!(
        "inputs={} entering_proof={} entries={} builds={}",
        paths.len(),
        inputs_with_targets,
        grand_targets,
        grand_builds
    );
}
