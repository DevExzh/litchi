//! Stream-by-stream differential of the two sector-layout policies over every
//! OLE2 fixture in `test-data/`.
//!
//! For each fixture and each of three edit shapes — an exact no-op, a
//! same-length edit and a length-changing edit that forces the append
//! fallback — the default policy's output must re-open with every stream
//! byte-identical to the opt-in rewrite policy's output, every directory
//! entry's name, type, hierarchy and class identifier must agree, and the
//! validation walk must admit both.

use litchi_cfb::{
    DirectoryEntry, OleError, OleFile, OleWriter, SectorLayoutPolicy, validate_source,
};
use litchi_core::OwnedSource;
use std::collections::BTreeSet;
use std::hint::black_box;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::Arc;
use std::time::Instant;

const MAGIC: &[u8; 8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";
const MAX_FIXTURE_BYTES: u64 = 48 * 1024 * 1024;
const DETERMINISM_SOURCE_ENV: &str = "LITCHI_CFB_LAYOUT_DETERMINISM_SOURCE";
const ENDOFCHAIN: u32 = 0xFFFF_FFFE;
const HEADER_DIFAT_OFFSET: usize = 0x4C;
const HEADER_DIFAT_ENTRIES: usize = 109;
const DIRECTORY_ENTRY_SIZE: usize = 128;
const ENTRY_START_SECTOR_OFFSET: usize = 0x74;
const ENTRY_STREAM_SIZE_OFFSET: usize = 0x78;

fn repository_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(Path::parent)
        .expect("crate lives two levels below the repository root")
        .to_path_buf()
}

fn collect(directory: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(directory) else {
        return;
    };
    let mut sorted: Vec<PathBuf> = entries.flatten().map(|entry| entry.path()).collect();
    sorted.sort();
    for path in sorted {
        if path.is_dir() {
            collect(&path, out);
        } else if path.is_file() {
            let Ok(metadata) = std::fs::metadata(&path) else {
                continue;
            };
            if metadata.len() < 512 || metadata.len() > MAX_FIXTURE_BYTES {
                continue;
            }
            let Ok(bytes) = std::fs::read(&path) else {
                continue;
            };
            if bytes.get(..8) == Some(MAGIC) {
                out.push(path);
            }
        }
    }
}

struct Model {
    streams: Vec<(Vec<String>, Vec<u8>)>,
    storages: Vec<Vec<String>>,
    sector_size: usize,
}

fn walk_storages(
    ole: &OleFile<Cursor<Vec<u8>>>,
    prefix: &[String],
    out: &mut Vec<Vec<String>>,
) -> Result<(), OleError> {
    let refs: Vec<&str> = prefix.iter().map(String::as_str).collect();
    let entries: Vec<(String, u8)> = ole
        .list_directory_entries(&refs)?
        .into_iter()
        .map(|entry: &DirectoryEntry| (entry.name.clone(), entry.entry_type))
        .collect();
    for (name, entry_type) in entries {
        if entry_type != 1 {
            continue;
        }
        let mut path = prefix.to_vec();
        path.push(name);
        out.push(path.clone());
        walk_storages(ole, &path, out)?;
    }
    Ok(())
}

fn capture(bytes: &[u8]) -> Result<Model, OleError> {
    let mut ole = OleFile::open(Cursor::new(bytes.to_vec()))?;
    let sector_size = ole.sector_size();
    let mut storages = Vec::new();
    walk_storages(&ole, &[], &mut storages)?;
    let paths = ole.list_streams();
    let mut streams = Vec::new();
    for path in paths {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        let data = ole.open_stream(&refs)?;
        streams.push((path, data));
    }
    Ok(Model {
        streams,
        storages,
        sector_size,
    })
}

fn serialize(model: &Model, source: Option<&[u8]>, policy: SectorLayoutPolicy) -> (Vec<u8>, bool) {
    let mut writer = OleWriter::with_sector_size(model.sector_size).expect("sector size");
    writer.set_sector_layout_policy(policy);
    if let Some(source) = source {
        writer.adopt_source_layout(source).expect("adoption");
    }
    let mut ordered = model.storages.clone();
    ordered.sort_by(|left, right| left.len().cmp(&right.len()).then_with(|| left.cmp(right)));
    for path in &ordered {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        writer.create_storage(&refs).expect("storage");
    }
    for (path, data) in &model.streams {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        writer.create_stream(&refs, data).expect("stream");
    }
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).expect("serialize");
    let reused = writer
        .last_sector_layout()
        .expect("report")
        .reused_source_layout();
    (out.into_inner(), reused)
}

fn entries_of(bytes: &[u8]) -> Vec<(Vec<String>, u8, String)> {
    let ole = OleFile::open(Cursor::new(bytes.to_vec())).expect("reopen");
    let mut out = Vec::new();
    let mut pending: Vec<Vec<String>> = vec![Vec::new()];
    while let Some(prefix) = pending.pop() {
        let refs: Vec<&str> = prefix.iter().map(String::as_str).collect();
        let entries: Vec<(String, u8, String)> = ole
            .list_directory_entries(&refs)
            .expect("entries")
            .into_iter()
            .map(|entry| (entry.name.clone(), entry.entry_type, entry.clsid.clone()))
            .collect();
        for (name, entry_type, clsid) in entries {
            let mut path = prefix.clone();
            path.push(name);
            out.push((path.clone(), entry_type, clsid));
            if entry_type == 1 {
                pending.push(path);
            }
        }
    }
    out.sort();
    out
}

fn u32_at(bytes: &[u8], offset: usize) -> u32 {
    let raw: [u8; 4] = bytes[offset..offset + 4].try_into().expect("u32 field");
    u32::from_le_bytes(raw)
}

fn directory_catalog(
    bytes: &[u8],
) -> std::collections::BTreeMap<Vec<String>, (u32, u8, u64, bool)> {
    let ole = OleFile::open(Cursor::new(bytes.to_vec())).expect("directory catalog");
    let mut out = std::collections::BTreeMap::new();
    let root = ole.root_entry().expect("root entry");
    out.insert(
        Vec::new(),
        (root.sid, root.entry_type, root.size, root.is_minifat),
    );
    let mut pending: Vec<Vec<String>> = vec![Vec::new()];
    while let Some(prefix) = pending.pop() {
        let refs: Vec<&str> = prefix.iter().map(String::as_str).collect();
        for entry in ole
            .list_directory_entries(&refs)
            .expect("directory entries")
        {
            let mut path = prefix.clone();
            path.push(entry.name.clone());
            out.insert(
                path.clone(),
                (entry.sid, entry.entry_type, entry.size, entry.is_minifat),
            );
            if entry.entry_type == 1 {
                pending.push(path);
            }
        }
    }
    out
}

/// Returns the raw directory image with only planner-owned allocation fields
/// blanked. Sibling links, names, node colours, CLSIDs, state bits and both
/// timestamps remain byte-for-byte compared with the source. Version-3 files
/// also have an unused high size word; that word is normalized separately so
/// the low size and starting sector remain part of the comparison.
fn normalized_directory_image(
    bytes: &[u8],
    allocation_sids: &BTreeSet<u32>,
    v3_size_sids: &BTreeSet<u32>,
) -> Vec<u8> {
    let sector_shift = u16::from_le_bytes(bytes[0x1E..0x20].try_into().expect("sector shift"));
    let sector_size = 1usize << sector_shift;
    assert!(sector_size == 512 || sector_size == 4096, "sector size");
    let header = bytes.get(..sector_size).expect("header");
    let fat_count = u32_at(header, 0x2C) as usize;
    assert!(u32_at(header, 0x48) == 0, "test helper expects no DIFAT");
    assert!(
        fat_count <= HEADER_DIFAT_ENTRIES,
        "test helper expects header FAT"
    );
    let mut fat = Vec::with_capacity(fat_count * (sector_size / 4));
    for index in 0..fat_count {
        let sector = u32_at(header, HEADER_DIFAT_OFFSET + index * 4) as usize;
        let start = (sector + 1) * sector_size;
        let end = start + sector_size;
        for word in bytes[start..end].chunks_exact(4) {
            fat.push(u32::from_le_bytes(word.try_into().expect("FAT word")));
        }
    }
    let mut image = Vec::new();
    let mut sector = u32_at(header, 0x30);
    let mut seen = BTreeSet::new();
    while sector != ENDOFCHAIN {
        assert!(seen.insert(sector), "directory chain cycle");
        let start = (sector as usize + 1) * sector_size;
        image.extend_from_slice(&bytes[start..start + sector_size]);
        sector = fat[sector as usize];
    }
    assert_eq!(image.len() % DIRECTORY_ENTRY_SIZE, 0);
    for (sid, entry) in image.chunks_exact_mut(DIRECTORY_ENTRY_SIZE).enumerate() {
        if allocation_sids.contains(&(sid as u32)) {
            entry[ENTRY_START_SECTOR_OFFSET..ENTRY_STREAM_SIZE_OFFSET + 8].fill(0);
        } else if v3_size_sids.contains(&(sid as u32)) {
            entry[ENTRY_STREAM_SIZE_OFFSET + 4..ENTRY_STREAM_SIZE_OFFSET + 8].fill(0);
        }
    }
    image
}

fn compare(
    label: &str,
    source: &[u8],
    reused: &[u8],
    rewritten: &[u8],
    expected: &Model,
    reused_layout: bool,
) {
    let left = capture(reused).unwrap_or_else(|error| panic!("{label}: reuse reopen: {error}"));
    let right =
        capture(rewritten).unwrap_or_else(|error| panic!("{label}: rewrite reopen: {error}"));
    let mut left_streams = left.streams.clone();
    let mut right_streams = right.streams.clone();
    let mut want = expected.streams.clone();
    left_streams.sort_by(|a, b| a.0.cmp(&b.0));
    right_streams.sort_by(|a, b| a.0.cmp(&b.0));
    want.sort_by(|a, b| a.0.cmp(&b.0));
    assert_eq!(left_streams.len(), want.len(), "{label}: stream count");
    for ((path, got), (wanted_path, wanted)) in left_streams.iter().zip(want.iter()) {
        assert_eq!(path, wanted_path, "{label}: stream path");
        assert_eq!(got, wanted, "{label}: stream {path:?} bytes");
    }
    assert_eq!(left_streams, right_streams, "{label}: policy differential");
    let source_entries = entries_of(source);
    let reused_entries = entries_of(reused);
    assert_eq!(
        reused_entries, source_entries,
        "{label}: reused directory metadata"
    );
    if reused_layout {
        let source_catalog = directory_catalog(source);
        let sector_shift = u16::from_le_bytes(source[0x1E..0x20].try_into().unwrap());
        let sector_size = 1usize << sector_shift;
        let cutoff = u64::from(u32_at(source, 0x38));
        let mut allocation_sids = BTreeSet::new();
        // The root allocation is always serialized by the planner. Version-3
        // files also require its unused high size word to be canonicalized.
        allocation_sids.insert(0);
        let mut v3_size_sids = BTreeSet::new();
        for (path, bytes) in &expected.streams {
            let Some((sid, _kind, old_size, old_is_mini)) = source_catalog.get(path) else {
                continue;
            };
            let new_size = u64::try_from(bytes.len()).expect("stream size");
            let new_is_mini = new_size > 0 && new_size < cutoff;
            if *old_size != new_size || *old_is_mini != new_is_mini {
                allocation_sids.insert(*sid);
            }
            if sector_size == 512 {
                v3_size_sids.insert(*sid);
            }
        }
        assert_eq!(
            normalized_directory_image(reused, &allocation_sids, &v3_size_sids),
            normalized_directory_image(source, &allocation_sids, &v3_size_sids),
            "{label}: reused directory bytes changed outside stream allocation fields"
        );
    }
    assert_eq!(
        reused_entries
            .iter()
            .map(|(path, kind, _)| (path.clone(), *kind))
            .collect::<Vec<_>>(),
        entries_of(rewritten)
            .into_iter()
            .map(|(path, kind, _)| (path, kind))
            .collect::<Vec<_>>(),
        "{label}: directory shape"
    );
    let report = validate_source(Arc::new(OwnedSource::new(reused.to_vec())))
        .unwrap_or_else(|error| panic!("{label}: validation walk: {error}"));
    assert!(
        report.is_complete(),
        "{label}: validation walk did not complete on the reused layout: {report:?}"
    );
    assert!(
        !report.has_errors(),
        "{label}: validation walk rejected the reused layout: {report:?}"
    );
}

#[test]
fn every_ole2_fixture_agrees_stream_by_stream_under_both_policies() {
    let mut fixtures = Vec::new();
    collect(&repository_root().join("test-data"), &mut fixtures);
    assert!(
        fixtures.len() > 150,
        "expected the OLE2 corpus: {}",
        fixtures.len()
    );

    let mut examined = 0usize;
    let mut skipped = 0usize;
    let mut reused_noop = 0usize;
    let mut reused_same = 0usize;
    let mut reused_grow = 0usize;
    let mut appended = 0usize;
    for path in &fixtures {
        let bytes = std::fs::read(path).expect("fixture");
        let Ok(model) = capture(&bytes) else {
            skipped += 1;
            continue;
        };
        if model.streams.is_empty() {
            skipped += 1;
            continue;
        }
        examined += 1;
        let label = path.display().to_string();

        // 1. exact no-op
        let (reuse, did_reuse) = serialize(&model, Some(&bytes), SectorLayoutPolicy::Reuse);
        let (rewrite, _) = serialize(&model, Some(&bytes), SectorLayoutPolicy::Rewrite);
        reused_noop += usize::from(did_reuse);
        compare(
            &format!("{label} [no-op]"),
            &bytes,
            &reuse,
            &rewrite,
            &model,
            did_reuse,
        );

        // 2. same-length edit of the largest stream
        let largest = model
            .streams
            .iter()
            .enumerate()
            .max_by_key(|(_, (_, data))| data.len())
            .map(|(index, _)| index)
            .expect("a stream");
        if model.streams[largest].1.is_empty() {
            continue;
        }
        let mut same = Model {
            streams: model.streams.clone(),
            storages: model.storages.clone(),
            sector_size: model.sector_size,
        };
        let last = same.streams[largest].1.len() - 1;
        same.streams[largest].1[last] ^= 0xFF;
        let (reuse, did_reuse) = serialize(&same, Some(&bytes), SectorLayoutPolicy::Reuse);
        let (rewrite, _) = serialize(&same, Some(&bytes), SectorLayoutPolicy::Rewrite);
        reused_same += usize::from(did_reuse);
        compare(
            &format!("{label} [same-length]"),
            &bytes,
            &reuse,
            &rewrite,
            &same,
            did_reuse,
        );

        // 3. length-changing edit that outgrows the existing allocation
        let mut grown = Model {
            streams: model.streams.clone(),
            storages: model.storages.clone(),
            sector_size: model.sector_size,
        };
        let padding = model.sector_size * 3 + 7;
        grown.streams[largest]
            .1
            .extend(std::iter::repeat_n(0xA5u8, padding));
        let mut writer = OleWriter::with_sector_size(model.sector_size).expect("sector size");
        writer.adopt_source_layout(&bytes).expect("adoption");
        let mut ordered = grown.storages.clone();
        ordered.sort_by(|left, right| left.len().cmp(&right.len()).then_with(|| left.cmp(right)));
        for storage in &ordered {
            let refs: Vec<&str> = storage.iter().map(String::as_str).collect();
            writer.create_storage(&refs).expect("storage");
        }
        for (stream, data) in &grown.streams {
            let refs: Vec<&str> = stream.iter().map(String::as_str).collect();
            writer.create_stream(&refs, data).expect("stream");
        }
        let mut out = Cursor::new(Vec::new());
        writer.write_to(&mut out).expect("serialize");
        let report = writer.last_sector_layout().expect("report");
        reused_grow += usize::from(report.reused_source_layout());
        appended += usize::from(report.appended_sectors() > 0);
        let reuse = out.into_inner();
        let (rewrite, _) = serialize(&grown, Some(&bytes), SectorLayoutPolicy::Rewrite);
        compare(
            &format!("{label} [length-changing]"),
            &bytes,
            &reuse,
            &rewrite,
            &grown,
            report.reused_source_layout(),
        );
    }

    println!(
        "fixtures={} examined={examined} skipped={skipped} reused_noop={reused_noop} \
         reused_same_length={reused_same} reused_length_changing={reused_grow} appended={appended}",
        fixtures.len()
    );
    assert!(examined > 150, "examined {examined} fixtures");
    assert!(
        reused_noop * 100 >= examined * 90,
        "the default policy reused only {reused_noop} of {examined} no-op saves"
    );
    assert!(
        appended * 100 >= examined * 80,
        "the append fallback fired on only {appended} of {examined} length-changing saves"
    );
}

/// Emits one JSON line of deterministic counts per fixture and edit shape.
///
/// Run with `--nocapture` to retain the evidence; the assertions are the same
/// stream-by-stream identity the differential above proves, so this case is a
/// counting harness rather than an independent gate.
#[test]
fn sector_layout_counts_over_the_corpus() {
    let mut fixtures = Vec::new();
    collect(&repository_root().join("test-data"), &mut fixtures);
    let root = repository_root();
    for path in &fixtures {
        let bytes = std::fs::read(path).expect("fixture");
        let Ok(model) = capture(&bytes) else { continue };
        if model.streams.is_empty() {
            continue;
        }
        let largest = model
            .streams
            .iter()
            .enumerate()
            .max_by_key(|(_, (_, data))| data.len())
            .map(|(index, _)| index)
            .expect("a stream");
        for shape in ["noop", "same-length", "length-changing"] {
            let mut edited = Model {
                streams: model.streams.clone(),
                storages: model.storages.clone(),
                sector_size: model.sector_size,
            };
            match shape {
                "same-length" => {
                    if edited.streams[largest].1.is_empty() {
                        continue;
                    }
                    let last = edited.streams[largest].1.len() - 1;
                    edited.streams[largest].1[last] ^= 0xFF;
                },
                "length-changing" => {
                    let padding = model.sector_size * 3 + 7;
                    edited.streams[largest]
                        .1
                        .extend(std::iter::repeat_n(0xA5u8, padding));
                },
                _ => {},
            }
            let mut writer = OleWriter::with_sector_size(edited.sector_size).expect("sector size");
            writer.adopt_source_layout(&bytes).expect("adoption");
            let mut ordered = edited.storages.clone();
            ordered
                .sort_by(|left, right| left.len().cmp(&right.len()).then_with(|| left.cmp(right)));
            for storage in &ordered {
                let refs: Vec<&str> = storage.iter().map(String::as_str).collect();
                writer.create_storage(&refs).expect("storage");
            }
            for (stream, data) in &edited.streams {
                let refs: Vec<&str> = stream.iter().map(String::as_str).collect();
                writer.create_stream(&refs, data).expect("stream");
            }
            let mut out = Cursor::new(Vec::new());
            writer.write_to(&mut out).expect("serialize");
            let report = writer.last_sector_layout().expect("report");
            let reuse_bytes = out.into_inner().len();
            let (rewrite, _) = serialize(&edited, Some(&bytes), SectorLayoutPolicy::Rewrite);
            let relative = path.strip_prefix(&root).unwrap_or(path).display();
            println!(
                "LAYOUT {{\"fixture\":\"{relative}\",\"shape\":\"{shape}\",\
\"sector_size\":{},\"source_bytes\":{},\"reuse_bytes\":{},\"rewrite_bytes\":{},\
\"reused\":{},\"fallback\":\"{:?}\",\"output_sectors\":{},\"kept_sectors\":{},\
\"rewritten_sectors\":{},\"appended_sectors\":{},\"reclaimed_sectors\":{},\
\"free_sectors\":{},\"streams\":{}}}",
                edited.sector_size,
                bytes.len(),
                reuse_bytes,
                rewrite.len(),
                report.reused_source_layout(),
                report.fallback(),
                report.output_sectors(),
                report.kept_sectors(),
                report.rewritten_sectors(),
                report.appended_sectors(),
                report.reclaimed_sectors(),
                report.free_sectors(),
                edited.streams.len(),
            );
        }
    }
}

/// The same model and the same adopted source produce byte-identical output,
/// within one process and across the process boundary the sibling binary test
/// crosses.
#[test]
fn a_reused_layout_is_deterministic() {
    let mut fixtures = Vec::new();
    collect(&repository_root().join("test-data"), &mut fixtures);
    let mut checked = 0usize;
    for path in fixtures.iter().take(40) {
        let bytes = std::fs::read(path).expect("fixture");
        let Ok(model) = capture(&bytes) else { continue };
        if model.streams.is_empty() {
            continue;
        }
        let (first, _) = serialize(&model, Some(&bytes), SectorLayoutPolicy::Reuse);
        let (second, _) = serialize(&model, Some(&bytes), SectorLayoutPolicy::Reuse);
        assert_eq!(
            first,
            second,
            "{}: reused layout is not deterministic",
            path.display()
        );
        checked += 1;
    }
    assert!(checked > 20, "checked {checked} fixtures");
}

fn digest(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

/// Child half of [`reused_layout_is_deterministic_across_processes`]. Keeping
/// this as a normal integration-test entry point lets the test harness launch
/// it with the same binary and a fresh Rust hash-map seed.
#[test]
fn reused_layout_determinism_child() {
    let Some(path) = std::env::var_os(DETERMINISM_SOURCE_ENV) else {
        return;
    };
    let bytes = std::fs::read(path).expect("determinism source");
    let model = capture(&bytes).expect("determinism source should parse");
    let (output, reused) = serialize(&model, Some(&bytes), SectorLayoutPolicy::Reuse);
    assert!(reused, "determinism source should admit reuse");
    println!("LAYOUT_DIGEST={:016x}", digest(&output));
}

#[test]
fn reused_layout_is_deterministic_across_processes() {
    let mut fixtures = Vec::new();
    collect(&repository_root().join("test-data"), &mut fixtures);
    let path = fixtures
        .iter()
        .find(|path| capture(&std::fs::read(path).expect("fixture")).is_ok())
        .expect("a parseable OLE2 fixture");
    let mut digests = Vec::new();
    for _ in 0..3 {
        let output = Command::new(std::env::current_exe().expect("test executable"))
            .args(["--exact", "reused_layout_determinism_child", "--nocapture"])
            .env(DETERMINISM_SOURCE_ENV, path)
            .output()
            .expect("child determinism test should start");
        assert!(
            output.status.success(),
            "child determinism test failed: {}{}",
            String::from_utf8_lossy(&output.stdout),
            String::from_utf8_lossy(&output.stderr)
        );
        let digest = String::from_utf8_lossy(&output.stdout)
            .lines()
            .find_map(|line| line.strip_prefix("LAYOUT_DIGEST="))
            .expect("child should report a layout digest")
            .to_owned();
        digests.push(digest);
    }
    assert!(
        digests.windows(2).all(|pair| pair[0] == pair[1]),
        "cross-process layout digests differ: {digests:?}"
    );
}

fn median(values: &mut [u128]) -> u128 {
    values.sort_unstable();
    values[values.len() / 2]
}

fn paired_floor_percent(first: &[u128], second: &[u128]) -> f64 {
    let mut deltas = first
        .iter()
        .zip(second)
        .map(|(left, right)| {
            let difference = left.abs_diff(*right) as f64;
            let floor = (*left).min(*right) as f64;
            difference * 100.0 / floor
        })
        .collect::<Vec<_>>();
    deltas.sort_by(f64::total_cmp);
    deltas[deltas.len() / 2]
}

/// Times the source-adopted route against the source-adopted from-scratch
/// route with an in-window A/A floor. This is deliberately an evidence test,
/// not a speed assertion: the current planner re-emits every payload sector,
/// so `kept_sectors` describes physical placement retention and does not claim
/// that sink writes or payload copies were avoided.
#[test]
fn source_layout_measurement_with_aa_floor() {
    let mut fixtures = Vec::new();
    collect(&repository_root().join("test-data"), &mut fixtures);
    let path = fixtures
        .iter()
        .find(|path| path.to_string_lossy().ends_with("ole/doc/picture.doc"))
        .expect("the representative DOC fixture");
    let source = std::fs::read(path).expect("measurement fixture");
    let model = capture(&source).expect("measurement fixture should parse");
    let largest = model
        .streams
        .iter()
        .enumerate()
        .max_by_key(|(_, (_, data))| data.len())
        .expect("measurement stream")
        .0;
    let mut edited = Model {
        streams: model.streams.clone(),
        storages: model.storages.clone(),
        sector_size: model.sector_size,
    };
    if edited.streams[largest].1.is_empty() {
        return;
    }
    let last = edited.streams[largest].1.len() - 1;
    edited.streams[largest].1[last] ^= 0xFF;

    for _ in 0..3 {
        let (reuse, _) = serialize(&edited, Some(&source), SectorLayoutPolicy::Reuse);
        let (rewrite, _) = serialize(&edited, Some(&source), SectorLayoutPolicy::Rewrite);
        assert_eq!(
            capture(&reuse).expect("reuse output").streams,
            capture(&rewrite).expect("rewrite output").streams
        );
        black_box((reuse.len(), rewrite.len()));
    }

    let mut reuse_first = Vec::new();
    let mut reuse_second = Vec::new();
    let mut rewrite_first = Vec::new();
    let mut rewrite_second = Vec::new();
    for _ in 0..24 {
        let start = Instant::now();
        let (reuse_a, reused_a) = serialize(&edited, Some(&source), SectorLayoutPolicy::Reuse);
        let reuse_a_ns = start.elapsed().as_nanos();
        let start = Instant::now();
        let (rewrite_a, reused_rewrite_a) =
            serialize(&edited, Some(&source), SectorLayoutPolicy::Rewrite);
        let rewrite_a_ns = start.elapsed().as_nanos();
        let start = Instant::now();
        let (rewrite_b, reused_rewrite_b) =
            serialize(&edited, Some(&source), SectorLayoutPolicy::Rewrite);
        let rewrite_b_ns = start.elapsed().as_nanos();
        let start = Instant::now();
        let (reuse_b, reused_b) = serialize(&edited, Some(&source), SectorLayoutPolicy::Reuse);
        let reuse_b_ns = start.elapsed().as_nanos();
        assert!(reused_a && reused_b);
        assert!(!reused_rewrite_a && !reused_rewrite_b);
        assert_eq!(reuse_a, reuse_b);
        assert_eq!(rewrite_a, rewrite_b);
        black_box((
            reuse_a.len(),
            rewrite_a.len(),
            reuse_b.len(),
            rewrite_b.len(),
        ));
        reuse_first.push(reuse_a_ns);
        reuse_second.push(reuse_b_ns);
        rewrite_first.push(rewrite_a_ns);
        rewrite_second.push(rewrite_b_ns);
    }
    let reuse_p50 = median(
        &mut reuse_first
            .iter()
            .zip(&reuse_second)
            .map(|(left, right)| (*left + *right) / 2)
            .collect::<Vec<_>>(),
    );
    let rewrite_p50 = median(
        &mut rewrite_first
            .iter()
            .zip(&rewrite_second)
            .map(|(left, right)| (*left + *right) / 2)
            .collect::<Vec<_>>(),
    );
    let reuse_floor = paired_floor_percent(&reuse_first, &reuse_second);
    let rewrite_floor = paired_floor_percent(&rewrite_first, &rewrite_second);
    println!(
        "LAYOUT_TIMING fixture={} samples=24 source_bytes={} reuse_p50_ns={} rewrite_p50_ns={} reuse_aa_floor_p50_pct={reuse_floor:.2} rewrite_aa_floor_p50_pct={rewrite_floor:.2}",
        path.display(),
        source.len(),
        reuse_p50,
        rewrite_p50,
    );
    assert!(reuse_p50 > 0 && rewrite_p50 > 0);
}

/// A source the parser rejects, or whose layout has been tampered with, never
/// steers the output: the writer declines to adopt it and serializes exactly
/// the bytes the from-scratch policy would have produced.
#[test]
fn a_mutated_source_never_steers_the_layout() {
    let mut fixtures = Vec::new();
    collect(&repository_root().join("test-data"), &mut fixtures);
    let path = fixtures
        .iter()
        .find(|path| path.to_string_lossy().contains("/ole/doc/"))
        .cloned()
        .expect("a DOC fixture");
    let bytes = std::fs::read(&path).expect("fixture");
    let model = capture(&bytes).expect("capture");
    let (baseline, reused) = serialize(&model, Some(&bytes), SectorLayoutPolicy::Reuse);
    assert!(reused, "the unmutated source must be adopted");
    let (from_scratch, _) = serialize(&model, None, SectorLayoutPolicy::Reuse);

    let sector_size = model.sector_size as u64;
    let mut declined = 0usize;
    let mut admitted = 0usize;
    // Header FAT count, header directory pointer, header MiniFAT pointer, the
    // header DIFAT array, the first FAT sector, and the directory image.
    let offsets: Vec<usize> = vec![
        0x2C,
        0x30,
        0x3C,
        0x40,
        0x44,
        0x48,
        0x4C,
        0x50,
        0x1E,
        sector_size as usize,
        sector_size as usize + 4,
        sector_size as usize * 2 + 0x74,
        sector_size as usize * 2 + 0x78,
        sector_size as usize * 2 + 0x42,
    ];
    for offset in offsets {
        for pattern in [0x00u8, 0xFF, 0xA5] {
            let mut mutated = bytes.clone();
            if offset + 4 > mutated.len() {
                continue;
            }
            for slot in &mut mutated[offset..offset + 4] {
                *slot = pattern;
            }
            let mut writer = OleWriter::with_sector_size(model.sector_size).expect("sector size");
            let adopted = writer.adopt_source_layout(&mutated).expect("adoption");
            let mut ordered = model.storages.clone();
            ordered
                .sort_by(|left, right| left.len().cmp(&right.len()).then_with(|| left.cmp(right)));
            for storage in &ordered {
                let refs: Vec<&str> = storage.iter().map(String::as_str).collect();
                writer.create_storage(&refs).expect("storage");
            }
            for (stream, data) in &model.streams {
                let refs: Vec<&str> = stream.iter().map(String::as_str).collect();
                writer.create_stream(&refs, data).expect("stream");
            }
            let mut out = Cursor::new(Vec::new());
            writer.write_to(&mut out).expect("serialize");
            let output = out.into_inner();
            let report = writer.last_sector_layout().expect("report");
            if adopted && report.reused_source_layout() {
                admitted += 1;
                // An admitted mutation must still publish every stream intact.
                let reopened = capture(&output).expect("reopen");
                let mut got = reopened.streams.clone();
                let mut want = model.streams.clone();
                got.sort_by(|a, b| a.0.cmp(&b.0));
                want.sort_by(|a, b| a.0.cmp(&b.0));
                assert_eq!(got, want, "mutation at {offset:#x}/{pattern:#x}");
            } else {
                declined += 1;
                assert_eq!(
                    output, from_scratch,
                    "a declined mutation at {offset:#x}/{pattern:#x} must serialize from scratch"
                );
            }
        }
    }
    assert!(declined > 0, "no mutation was declined");
    assert_eq!(baseline.len() % model.sector_size, 0);
    println!("mutations declined={declined} admitted={admitted}");
}
