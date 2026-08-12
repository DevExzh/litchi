use crate::{
    OleFile, OleWriter, OutputProgress, OverlayError, OverlayLimits, SameLengthStreamOverlay,
    SharedOleFile,
};
use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use std::io::{self, Cursor, Write};
use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};
use std::sync::{Arc, Mutex};

fn sample_bytes() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["MiniTiny"], &[0x09; 70]).unwrap();
    writer
        .create_stream(&["MiniNeighbor"], &[0x0a; 83])
        .unwrap();
    writer
        .create_stream(&["Mini4095"], &vec![0x11; 4_095])
        .unwrap();
    writer
        .create_stream(&["Fat4096"], &vec![0x22; 4_096])
        .unwrap();
    writer
        .create_stream(&["Fat4097"], &vec![0x33; 4_097])
        .unwrap();
    writer
        .create_stream(&["Opaque"], &vec![0x44; 5_003])
        .unwrap();
    writer
        .create_stream(&["LargeOpaque"], &vec![0x45; 130_123])
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn version_four_bytes() -> Vec<u8> {
    let mut writer = OleWriter::with_sector_size(4_096).unwrap();
    writer.create_stream(&["Mini"], &vec![0x18; 1_003]).unwrap();
    writer.create_stream(&["Fat"], &vec![0x28; 5_003]).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn small_bytes() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Fragment"], &vec![0x31; 4_097])
        .unwrap();
    writer
        .create_stream(&["Other"], &vec![0x41; 5_003])
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn shared(bytes: Vec<u8>) -> SharedOleFile {
    SharedOleFile::open(Arc::new(OwnedSource::new(bytes))).unwrap()
}

fn limits() -> OverlayLimits {
    OverlayLimits::new(8, 4_096, 1_000_000).unwrap()
}

fn replacement(path: &str, byte: u8, length: usize) -> SameLengthStreamOverlay {
    SameLengthStreamOverlay::new(vec![path.to_string()], Arc::from(vec![byte; length]))
}

#[test]
fn overlays_minifat_and_fat_cutover_and_preserves_unselected_bytes() {
    let source = sample_bytes();
    let plan = shared(source.clone())
        .plan_same_length_stream_overlays(
            vec![
                replacement("Mini4095", 0xa1, 4_095),
                replacement("Fat4096", 0xa2, 4_096),
                replacement("Fat4097", 0xa3, 4_097),
            ],
            limits(),
        )
        .unwrap();
    assert!(!plan.is_noop());
    assert!(plan.changed_spans() >= 3);
    let mut output = Vec::new();
    let report = plan.write_to(&mut output).unwrap();
    assert_eq!(report.bytes(), source.len() as u64);
    assert_eq!(output.len(), source.len());

    let mut reopened = OleFile::open(Cursor::new(output)).unwrap();
    assert_eq!(
        reopened.open_stream(&["Mini4095"]).unwrap(),
        vec![0xa1; 4_095]
    );
    assert_eq!(
        reopened.open_stream(&["Fat4096"]).unwrap(),
        vec![0xa2; 4_096]
    );
    assert_eq!(
        reopened.open_stream(&["Fat4097"]).unwrap(),
        vec![0xa3; 4_097]
    );
    assert_eq!(
        reopened.open_stream(&["Opaque"]).unwrap(),
        vec![0x44; 5_003]
    );
}

#[test]
fn exact_and_empty_noops_copy_the_source_byte_for_byte() {
    let source = sample_bytes();
    for overlays in [Vec::new(), vec![replacement("Fat4096", 0x22, 4_096)]] {
        let plan = shared(source.clone())
            .plan_same_length_stream_overlays(overlays, limits())
            .unwrap();
        assert!(plan.is_noop());
        assert_eq!(plan.source_fingerprint(), plan.target_fingerprint());
        let mut output = Vec::new();
        let report = plan.write_to(&mut output).unwrap();
        assert_eq!(report.changed_spans(), 0);
        assert_eq!(output, source);
    }
}

#[test]
fn mini_overlay_preserves_a_neighbor_in_the_same_host_sector() {
    let plan = shared(sample_bytes())
        .plan_same_length_stream_overlays(vec![replacement("MiniTiny", 0xb1, 70)], limits())
        .unwrap();
    let mut output = Vec::new();
    plan.write_to(&mut output).unwrap();
    let mut reopened = OleFile::open(Cursor::new(output)).unwrap();
    assert_eq!(reopened.open_stream(&["MiniTiny"]).unwrap(), vec![0xb1; 70]);
    assert_eq!(
        reopened.open_stream(&["MiniNeighbor"]).unwrap(),
        vec![0x0a; 83]
    );
}

#[test]
fn version_four_sector_geometry_publishes_mini_and_fat_overlays() {
    let plan = shared(version_four_bytes())
        .plan_same_length_stream_overlays(
            vec![
                replacement("Mini", 0xc1, 1_003),
                replacement("Fat", 0xc2, 5_003),
            ],
            limits(),
        )
        .unwrap();
    let mut output = Vec::new();
    plan.write_to(&mut output).unwrap();
    let mut reopened = OleFile::open(Cursor::new(output)).unwrap();
    assert_eq!(reopened.sector_size(), 4_096);
    assert_eq!(reopened.open_stream(&["Mini"]).unwrap(), vec![0xc1; 1_003]);
    assert_eq!(reopened.open_stream(&["Fat"]).unwrap(), vec![0xc2; 5_003]);
}

#[test]
fn duplicate_length_and_limit_failures_are_typed() {
    let file = shared(sample_bytes());
    assert!(matches!(
        file.plan_same_length_stream_overlays(vec![replacement("Fat4096", 7, 4_095)], limits()),
        Err(OverlayError::Unavailable { .. })
    ));
    assert!(matches!(
        file.plan_same_length_stream_overlays(
            vec![
                replacement("Fat4096", 7, 4_096),
                replacement("Fat4096", 8, 4_096)
            ],
            limits()
        ),
        Err(OverlayError::Unavailable { .. })
    ));
    let one = OverlayLimits::new(1, 32, 8_192).unwrap();
    assert!(matches!(
        file.plan_same_length_stream_overlays(
            vec![
                replacement("Fat4096", 7, 4_096),
                replacement("Fat4097", 8, 4_097)
            ],
            one
        ),
        Err(OverlayError::Unavailable { .. })
    ));
    let short = OverlayLimits::new(2, 32, 4_095).unwrap();
    assert!(matches!(
        file.plan_same_length_stream_overlays(vec![replacement("Fat4096", 7, 4_096)], short),
        Err(OverlayError::Unavailable { .. })
    ));
}

struct ShortSink {
    bytes: Vec<u8>,
    maximum: usize,
    interrupt: bool,
}

impl Write for ShortSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.interrupt {
            self.interrupt = false;
            return Err(io::ErrorKind::Interrupted.into());
        }
        let count = self.maximum.min(bytes.len());
        self.bytes.extend_from_slice(&bytes[..count]);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn short_and_interrupted_sink_writes_complete_exactly() {
    let source = sample_bytes();
    let plan = shared(source.clone())
        .plan_same_length_stream_overlays(vec![replacement("Fat4097", 0x91, 4_097)], limits())
        .unwrap();
    let mut sink = ShortSink {
        bytes: Vec::new(),
        maximum: 7,
        interrupt: true,
    };
    let report = plan.write_to(&mut sink).unwrap();
    assert_eq!(report.bytes(), source.len() as u64);
    let mut reopened = OleFile::open(Cursor::new(sink.bytes)).unwrap();
    assert_eq!(
        reopened.open_stream(&["Fat4097"]).unwrap(),
        vec![0x91; 4_097]
    );
}

struct FailingSink {
    bytes: Vec<u8>,
    remaining: usize,
    overreport: bool,
    zero: bool,
    fail_flush: bool,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.overreport {
            return Ok(bytes.len() + 1);
        }
        if self.zero {
            return Ok(0);
        }
        if self.remaining == 0 {
            return Err(io::Error::other("injected sink failure"));
        }
        let count = bytes.len().min(self.remaining);
        self.bytes.extend_from_slice(&bytes[..count]);
        self.remaining -= count;
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        if self.fail_flush {
            Err(io::Error::other("injected flush failure"))
        } else {
            Ok(())
        }
    }
}

#[test]
fn hostile_sink_progress_is_typed() {
    let source = sample_bytes();
    let plan = shared(source.clone())
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x81, 4_096)], limits())
        .unwrap();
    let mut zero = FailingSink {
        bytes: Vec::new(),
        remaining: usize::MAX,
        overreport: false,
        zero: true,
        fail_flush: false,
    };
    assert!(matches!(plan.write_to(&mut zero), Err(OverlayError::Io(_))));
    assert!(zero.bytes.is_empty());
    let mut partial = FailingSink {
        bytes: Vec::new(),
        remaining: 1_003,
        overreport: false,
        zero: false,
        fail_flush: false,
    };
    assert!(matches!(
        plan.write_to(&mut partial),
        Err(OverlayError::IncompleteOutput {
            progress: OutputProgress::Prefix {
                accepted: 1_003,
                ..
            },
            ..
        })
    ));
    assert_eq!(partial.bytes.len(), 1_003);
    let mut overreport = FailingSink {
        bytes: Vec::new(),
        remaining: usize::MAX,
        overreport: true,
        zero: false,
        fail_flush: false,
    };
    assert!(matches!(
        plan.write_to(&mut overreport),
        Err(OverlayError::IncompleteOutput {
            progress: OutputProgress::Indeterminate { accepted_before: 0 },
            ..
        })
    ));
    let mut flush = FailingSink {
        bytes: Vec::new(),
        remaining: usize::MAX,
        overreport: false,
        zero: false,
        fail_flush: true,
    };
    assert!(matches!(
        plan.write_to(&mut flush),
        Err(OverlayError::IncompleteOutput {
            progress: OutputProgress::CompleteUnflushed { .. },
            ..
        })
    ));
    assert_eq!(flush.bytes.len(), source.len());
}

#[test]
fn fragmented_fat_chain_yields_sorted_nonoverlapping_publication() {
    let mut bytes = small_bytes();
    let parsed = shared(bytes.clone());
    let entry = parsed.find_entry(&["Fragment"]).unwrap();
    let sector_size = parsed.index.sector_size;
    let count = 4_097usize.div_ceil(sector_size);
    let mut chain = Vec::new();
    let mut sector = entry.start_sector;
    for _ in 0..count {
        chain.push(sector);
        sector = parsed.index.fat[sector as usize];
    }
    assert!(chain.len() >= 4);

    let fat_sector = u32::from_le_bytes(bytes[0x4c..0x50].try_into().unwrap());
    let fat_offset = (fat_sector as usize + 1) * sector_size;
    let [first, second, third, fourth] = [chain[0], chain[1], chain[2], chain[3]];
    for (current, next) in [(first, third), (third, second), (second, fourth)] {
        let offset = fat_offset + current as usize * 4;
        bytes[offset..offset + 4].copy_from_slice(&next.to_le_bytes());
    }
    let second_offset = (second as usize + 1) * sector_size;
    let third_offset = (third as usize + 1) * sector_size;
    for index in 0..sector_size {
        bytes.swap(second_offset + index, third_offset + index);
    }

    let plan = shared(bytes)
        .plan_same_length_stream_overlays(vec![replacement("Fragment", 0x77, 4_097)], limits())
        .unwrap();
    assert!(plan.changed_spans() >= 3);
    let mut output = Vec::new();
    plan.write_to(&mut output).unwrap();
    let mut reopened = OleFile::open(Cursor::new(output)).unwrap();
    assert_eq!(
        reopened.open_stream(&["Fragment"]).unwrap(),
        vec![0x77; 4_097]
    );
    assert_eq!(reopened.open_stream(&["Other"]).unwrap(), vec![0x41; 5_003]);
}

struct MutableSource {
    bytes: Mutex<Vec<u8>>,
    revision: AtomicU64,
    reads: AtomicUsize,
    fail_read: AtomicUsize,
    mutate_after_read: AtomicUsize,
    overreport: AtomicBool,
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Mutex::new(bytes),
            revision: AtomicU64::new(0),
            reads: AtomicUsize::new(0),
            fail_read: AtomicUsize::new(usize::MAX),
            mutate_after_read: AtomicUsize::new(usize::MAX),
            overreport: AtomicBool::new(false),
        }
    }

    fn change_version(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }

    fn change_bytes_without_version(&self) {
        self.bytes.lock().unwrap()[700] ^= 0xff;
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.lock().unwrap().len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let call = self.reads.fetch_add(1, Ordering::SeqCst) + 1;
        if call == self.fail_read.load(Ordering::SeqCst) {
            if self.overreport.load(Ordering::SeqCst) {
                return Ok(output.len() + 1);
            }
            return Ok(0);
        }
        let mut bytes = self.bytes.lock().unwrap();
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset"))?;
        let Some(available) = bytes.get(offset..) else {
            return Ok(0);
        };
        let count = output.len().min(available.len());
        output[..count].copy_from_slice(&available[..count]);
        if call == self.mutate_after_read.load(Ordering::SeqCst) {
            bytes[700] ^= 0xff;
        }
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0xfeed,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

#[test]
fn version_and_stable_token_byte_changes_are_caught_before_output() {
    let source = Arc::new(MutableSource::new(sample_bytes()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x71, 4_096)], limits())
        .unwrap();
    source.change_version();
    let mut output = Vec::new();
    assert!(matches!(
        plan.write_to(&mut output),
        Err(OverlayError::SourceChanged { .. })
    ));
    assert!(output.is_empty());

    let source = Arc::new(MutableSource::new(sample_bytes()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x72, 4_096)], limits())
        .unwrap();
    source.change_bytes_without_version();
    let mut output = Vec::new();
    assert!(matches!(
        plan.write_to(&mut output),
        Err(OverlayError::SourceFingerprintChanged { .. })
    ));
    assert!(output.is_empty());
}

#[test]
fn hostile_read_during_emission_reports_exact_sink_prefix() {
    let source = Arc::new(MutableSource::new(sample_bytes()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x61, 4_096)], limits())
        .unwrap();
    source.reads.store(0, Ordering::SeqCst);
    let chunks = usize::try_from(file.file_size()).unwrap().div_ceil(65_536);
    source.fail_read.store(chunks + 2, Ordering::SeqCst);
    let mut output = Vec::new();
    assert!(matches!(
        plan.write_to(&mut output),
        Err(OverlayError::IncompleteOutput {
            progress: OutputProgress::Prefix {
                accepted: 65_536,
                ..
            },
            ..
        })
    ));
    assert_eq!(output.len(), 65_536);
}

#[test]
fn hostile_read_overreport_during_emission_reports_the_exact_sink_prefix() {
    let source = Arc::new(MutableSource::new(sample_bytes()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x62, 4_096)], limits())
        .unwrap();
    source.reads.store(0, Ordering::SeqCst);
    let chunks = usize::try_from(file.file_size()).unwrap().div_ceil(65_536);
    source.overreport.store(true, Ordering::SeqCst);
    source.fail_read.store(chunks + 2, Ordering::SeqCst);
    let mut output = Vec::new();
    assert!(matches!(
        plan.write_to(&mut output),
        Err(OverlayError::IncompleteOutput {
            progress: OutputProgress::Prefix {
                accepted: 65_536,
                ..
            },
            ..
        })
    ));
    assert_eq!(output.len(), 65_536);
}

#[test]
fn stable_token_mutation_of_an_emitted_chunk_is_caught_before_success() {
    let source = Arc::new(MutableSource::new(sample_bytes()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x63, 4_096)], limits())
        .unwrap();
    source.reads.store(0, Ordering::SeqCst);
    let length = file.file_size();
    let chunks = usize::try_from(length).unwrap().div_ceil(65_536);
    // The first `chunks` reads are the preflight. Mutate the first source
    // chunk only after its bytes have been copied for emission.
    source.mutate_after_read.store(chunks + 1, Ordering::SeqCst);

    let mut output = Vec::new();
    assert!(matches!(
        plan.write_to(&mut output),
        Err(OverlayError::IncompleteOutput {
            progress: OutputProgress::CompleteUnflushed { bytes },
            source,
        }) if bytes == length && matches!(*source, OverlayError::SourceFingerprintChanged { .. })
    ));
    assert_eq!(output.len() as u64, length);
}

#[test]
fn atomic_path_publication_replaces_after_complete_staging() {
    static NEXT: AtomicU64 = AtomicU64::new(0);
    let directory = std::env::temp_dir().join(format!(
        "litchi-cfb-overlay-{}-{}",
        std::process::id(),
        NEXT.fetch_add(1, Ordering::Relaxed)
    ));
    std::fs::create_dir(&directory).unwrap();
    let destination = directory.join("document.ole");
    std::fs::write(&destination, b"old destination").unwrap();
    let plan = shared(sample_bytes())
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x51, 4_096)], limits())
        .unwrap();
    let report = plan.save(&destination).unwrap();
    let output = std::fs::read(&destination).unwrap();
    assert_eq!(report.bytes(), output.len() as u64);
    let mut reopened = OleFile::open(Cursor::new(output)).unwrap();
    assert_eq!(
        reopened.open_stream(&["Fat4096"]).unwrap(),
        vec![0x51; 4_096]
    );
    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 1);
    std::fs::remove_file(destination).unwrap();
    std::fs::remove_dir(directory).unwrap();
}

#[test]
fn atomic_path_preflight_failure_leaves_destination_and_directory_unchanged() {
    static NEXT: AtomicU64 = AtomicU64::new(0);
    let directory = std::env::temp_dir().join(format!(
        "litchi-cfb-overlay-failure-{}-{}",
        std::process::id(),
        NEXT.fetch_add(1, Ordering::Relaxed)
    ));
    std::fs::create_dir(&directory).unwrap();
    let destination = directory.join("document.ole");
    std::fs::write(&destination, b"old destination").unwrap();
    let source = Arc::new(MutableSource::new(sample_bytes()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x52, 4_096)], limits())
        .unwrap();
    source.change_version();
    assert!(matches!(
        plan.save(&destination),
        Err(OverlayError::SourceChanged { .. })
    ));
    assert_eq!(std::fs::read(&destination).unwrap(), b"old destination");
    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 1);
    std::fs::remove_file(destination).unwrap();
    std::fs::remove_dir(directory).unwrap();
}

#[test]
fn atomic_path_late_stable_token_mutation_leaves_destination_unchanged() {
    static NEXT: AtomicU64 = AtomicU64::new(0);
    let directory = std::env::temp_dir().join(format!(
        "litchi-cfb-overlay-late-mutation-{}-{}",
        std::process::id(),
        NEXT.fetch_add(1, Ordering::Relaxed)
    ));
    std::fs::create_dir(&directory).unwrap();
    let destination = directory.join("document.ole");
    std::fs::write(&destination, b"old destination").unwrap();
    let source = Arc::new(MutableSource::new(sample_bytes()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x53, 4_096)], limits())
        .unwrap();
    source.reads.store(0, Ordering::SeqCst);
    let chunks = usize::try_from(file.file_size()).unwrap().div_ceil(65_536);
    source.mutate_after_read.store(chunks + 1, Ordering::SeqCst);

    assert!(matches!(
        plan.save(&destination),
        Err(OverlayError::SourceFingerprintChanged { .. })
    ));
    assert_eq!(std::fs::read(&destination).unwrap(), b"old destination");
    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 1);
    std::fs::remove_file(destination).unwrap();
    std::fs::remove_dir(directory).unwrap();
}
