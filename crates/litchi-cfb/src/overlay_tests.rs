use crate::{
    OleFile, OleWriter, OutputProgress, OverlayError, OverlayLimits, OverlaySourceMode,
    SameLengthStreamOverlay, SharedOleFile,
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

fn publication_chunks(file_size: u64) -> usize {
    usize::try_from(file_size).unwrap().div_ceil(65_536)
}

fn fingerprint_chunks(file_size: u64) -> usize {
    usize::try_from(file_size).unwrap().div_ceil(1024 * 1024)
}

fn assert_operation_shape(
    shape: crate::OverlayOperationShape,
    source_mode: OverlaySourceMode,
    source_bytes: u64,
    is_noop: bool,
) {
    let fingerprint_chunks = source_bytes.div_ceil(1024 * 1024);
    let publication_chunks = source_bytes.div_ceil(65_536);
    // An effective plan hashes source and composed target; a no-op plan
    // composes the source, so one digest is both identities.
    let fingerprint_bytes = source_bytes * if is_noop { 1 } else { 2 };
    let fenced = u64::from(source_mode == OverlaySourceMode::GenericReadAt);
    assert_eq!(
        shape.counter_scope,
        "validated overlay logical pass shape; no runtime, allocator, or syscall counters"
    );
    assert_eq!(shape.source_mode, source_mode);
    assert_eq!(shape.source_bytes, source_bytes);
    assert_eq!(shape.fingerprint_chunk_bytes, 1024 * 1024);
    assert_eq!(shape.publication_chunk_bytes, 65_536);
    assert_eq!(
        shape.fingerprint_buffer_bytes,
        source_bytes.min(1024 * 1024)
    );
    assert_eq!(shape.publication_buffer_bytes, 65_536);
    assert_eq!(shape.planning_fingerprint_scans, 1 + fenced);
    assert_eq!(
        shape.planning_fingerprint_bytes,
        fingerprint_bytes * (1 + fenced)
    );
    assert_eq!(
        shape.planning_fingerprint_chunks,
        fingerprint_chunks * (1 + fenced)
    );
    assert_eq!(shape.composed_source_preflight_scans, 1);
    assert_eq!(shape.composed_source_preflight_bytes, fingerprint_bytes);
    assert_eq!(shape.composed_source_preflight_chunks, fingerprint_chunks);
    assert_eq!(shape.target_materialization_write_pre_scans, fenced);
    assert_eq!(
        shape.target_materialization_write_pre_bytes,
        fingerprint_bytes * fenced
    );
    assert_eq!(
        shape.target_materialization_write_pre_chunks,
        fingerprint_chunks * fenced
    );
    assert_eq!(shape.target_materialization_emission_scans, 1);
    assert_eq!(
        shape.target_materialization_emission_bytes,
        fingerprint_bytes
    );
    assert_eq!(
        shape.target_materialization_emission_chunks,
        publication_chunks
    );
    assert_eq!(shape.target_materialization_write_post_scans, fenced);
    assert_eq!(
        shape.target_materialization_write_post_bytes,
        fingerprint_bytes * fenced
    );
    assert_eq!(
        shape.target_materialization_write_post_chunks,
        fingerprint_chunks * fenced
    );
    assert_eq!(shape.direct_write_pre_scans, fenced);
    assert_eq!(shape.direct_write_pre_bytes, fingerprint_bytes * fenced);
    assert_eq!(shape.direct_write_pre_chunks, fingerprint_chunks * fenced);
    assert_eq!(shape.direct_emission_scans, 1);
    assert_eq!(shape.direct_emission_bytes, fingerprint_bytes);
    assert_eq!(shape.direct_emission_chunks, publication_chunks);
    assert_eq!(shape.direct_write_post_scans, fenced);
    assert_eq!(shape.direct_write_post_bytes, fingerprint_bytes * fenced);
    assert_eq!(shape.direct_write_post_chunks, fingerprint_chunks * fenced);
    assert_eq!(shape.atomic_save_pre_temp_scans, fenced);
    assert_eq!(shape.atomic_save_pre_temp_bytes, fingerprint_bytes * fenced);
    assert_eq!(
        shape.atomic_save_pre_temp_chunks,
        fingerprint_chunks * fenced
    );
    assert_eq!(shape.atomic_save_emission_scans, 1);
    assert_eq!(shape.atomic_save_emission_bytes, fingerprint_bytes);
    assert_eq!(shape.atomic_save_emission_chunks, publication_chunks);
    assert_eq!(shape.atomic_save_pre_rename_scans, fenced);
    assert_eq!(
        shape.atomic_save_pre_rename_bytes,
        fingerprint_bytes * fenced
    );
    assert_eq!(
        shape.atomic_save_pre_rename_chunks,
        fingerprint_chunks * fenced
    );
    assert_eq!(
        shape.atomic_save_event_scope,
        "public save durability contract; temporary-file, flush, fsync, rename, and parent-sync stage counters are not observable"
    );
}

#[test]
fn operation_shape_matches_generic_and_owned_overlay_policy() {
    let bytes = sample_bytes();
    let generic = shared(bytes.clone())
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x66, 4_096)], limits())
        .unwrap();
    let owned =
        SharedOleFile::open_owned(Arc::from(bytes.clone()), SourceVersion::new(0xcafe_0261, 0))
            .unwrap()
            .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x66, 4_096)], limits())
            .unwrap();

    assert_operation_shape(
        generic.operation_shape(),
        OverlaySourceMode::GenericReadAt,
        bytes.len() as u64,
        false,
    );
    assert_operation_shape(
        owned.operation_shape(),
        OverlaySourceMode::OwnedImmutableArc,
        bytes.len() as u64,
        false,
    );

    // An exact byte no-op keeps every pass and chunk count and halves only the
    // logical bytes hashed, because one digest is both identities.
    let noop = shared(bytes.clone())
        .plan_same_length_stream_overlays(
            vec![SameLengthStreamOverlay::new(
                vec!["Fat4096".to_string()],
                Arc::from(vec![0x22; 4_096]),
            )],
            limits(),
        )
        .unwrap();
    assert!(noop.is_noop());
    assert_operation_shape(
        noop.operation_shape(),
        OverlaySourceMode::GenericReadAt,
        bytes.len() as u64,
        true,
    );
    assert_eq!(
        noop.operation_shape().planning_fingerprint_scans,
        generic.operation_shape().planning_fingerprint_scans
    );
    assert_eq!(
        noop.operation_shape().planning_fingerprint_chunks,
        generic.operation_shape().planning_fingerprint_chunks
    );
    assert_eq!(noop.source_fingerprint(), noop.target_fingerprint());
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
    request_sizes: Mutex<Vec<usize>>,
    fail_read: AtomicUsize,
    mutate_after_read: AtomicUsize,
    /// Byte flipped by the hostile mutation; 700 lies in the CFB header/FAT
    /// region, so tests that need the reopen to succeed move it into an
    /// unselected stream payload instead.
    mutate_offset: AtomicUsize,
    overreport: AtomicBool,
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Mutex::new(bytes),
            revision: AtomicU64::new(0),
            reads: AtomicUsize::new(0),
            request_sizes: Mutex::new(Vec::new()),
            fail_read: AtomicUsize::new(usize::MAX),
            mutate_after_read: AtomicUsize::new(usize::MAX),
            mutate_offset: AtomicUsize::new(700),
            overreport: AtomicBool::new(false),
        }
    }

    fn change_version(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }

    fn change_bytes_without_version(&self) {
        let offset = self.mutate_offset.load(Ordering::SeqCst);
        self.bytes.lock().unwrap()[offset] ^= 0xff;
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.lock().unwrap().len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let call = self.reads.fetch_add(1, Ordering::SeqCst) + 1;
        self.request_sizes.lock().unwrap().push(output.len());
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
            let target = self.mutate_offset.load(Ordering::SeqCst);
            bytes[target] ^= 0xff;
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
fn direct_write_to_retains_three_complete_source_scans() {
    let source = Arc::new(MutableSource::new(sample_bytes()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x64, 4_096)], limits())
        .unwrap();
    source.reads.store(0, Ordering::SeqCst);

    let mut output = Vec::new();
    plan.write_to(&mut output).unwrap();

    // Direct sequential publication retains its initial fingerprint preflight,
    // 64 KiB emission scan, and post-emission fingerprint preflight. The
    // output-time source hash is part of emission and does not add a read.
    assert_eq!(
        source.reads.load(Ordering::SeqCst),
        fingerprint_chunks(file.file_size()) * 2 + publication_chunks(file.file_size())
    );
    assert_eq!(output.len() as u64, file.file_size());
}

#[test]
fn owned_direct_write_uses_only_the_hashed_emission_scan() {
    let source: Arc<[u8]> = Arc::from(sample_bytes());
    let reads = Arc::new(AtomicUsize::new(0));
    let file =
        SharedOleFile::open_owned_arc_source_for_test(source.clone(), reads.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x67, 4_096)], limits())
        .unwrap();

    reads.store(0, Ordering::SeqCst);
    let candidate = plan.composed_source().unwrap();
    assert_eq!(
        reads.load(Ordering::SeqCst),
        fingerprint_chunks(file.file_size())
    );
    drop(candidate);

    reads.store(0, Ordering::SeqCst);
    let mut output = Vec::new();
    let report = plan.write_to(&mut output).unwrap();
    assert_eq!(
        reads.load(Ordering::SeqCst),
        publication_chunks(file.file_size())
    );
    assert_eq!(report.bytes(), file.file_size());
    assert_eq!(output.len() as u64, file.file_size());
    let mut reopened = OleFile::open(Cursor::new(output)).unwrap();
    assert_eq!(
        reopened.open_stream(&["Fat4096"]).unwrap(),
        vec![0x67; 4_096]
    );
}

#[test]
fn public_owned_arc_source_publishes_exactly() {
    let source: Arc<[u8]> = Arc::from(sample_bytes());
    let file =
        SharedOleFile::open_owned(Arc::clone(&source), SourceVersion::new(0xcafe_0181, 0)).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![replacement("Fat4097", 0x68, 4_097)], limits())
        .unwrap();
    let mut output = Vec::new();
    let report = plan.write_to(&mut output).unwrap();
    assert_eq!(report.bytes(), source.len() as u64);
    let mut reopened = OleFile::open(Cursor::new(output)).unwrap();
    assert_eq!(
        reopened.open_stream(&["Fat4097"]).unwrap(),
        vec![0x68; 4_097]
    );
}

#[test]
fn owned_overlay_planning_elides_only_the_final_complete_fingerprint() {
    let bytes = sample_bytes();
    let generic_source = Arc::new(MutableSource::new(bytes.clone()));
    let generic_file = SharedOleFile::open(generic_source.clone()).unwrap();
    generic_source.reads.store(0, Ordering::SeqCst);
    let generic_plan = generic_file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x6a, 4_096)], limits())
        .unwrap();
    let generic_reads = generic_source.reads.load(Ordering::SeqCst);

    let owned_reads = Arc::new(AtomicUsize::new(0));
    let owned_file =
        SharedOleFile::open_owned_arc_source_for_test(Arc::from(bytes), owned_reads.clone())
            .unwrap();
    owned_reads.store(0, Ordering::SeqCst);
    let owned_plan = owned_file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x6a, 4_096)], limits())
        .unwrap();
    let owned_planning_reads = owned_reads.load(Ordering::SeqCst);

    assert_eq!(
        generic_plan.source_fingerprint(),
        owned_plan.source_fingerprint()
    );
    assert_eq!(
        generic_plan.target_fingerprint(),
        owned_plan.target_fingerprint()
    );
    assert_eq!(
        generic_reads,
        owned_planning_reads + fingerprint_chunks(generic_file.file_size())
    );
}

#[test]
fn fingerprint_requests_are_coalesced_without_widening_publication_chunks() {
    let source = Arc::new(MutableSource::new(sample_bytes()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x66, 4_096)], limits())
        .unwrap();
    source.reads.store(0, Ordering::SeqCst);
    source.request_sizes.lock().unwrap().clear();

    let mut output = Vec::new();
    plan.write_to(&mut output).unwrap();

    let requests = source.request_sizes.lock().unwrap().clone();
    let file_size = usize::try_from(file.file_size()).unwrap();
    let fingerprint_request = file_size.min(1024 * 1024);
    let fingerprint_requests = fingerprint_chunks(file.file_size());
    let publication_requests = publication_chunks(file.file_size());
    assert_eq!(
        requests.len(),
        fingerprint_requests * 2 + publication_requests
    );
    let before = &requests[..fingerprint_requests];
    let emission = &requests[fingerprint_requests..fingerprint_requests + publication_requests];
    let after = &requests[fingerprint_requests + publication_requests..];
    assert_eq!(before, after);
    assert_eq!(before.first(), Some(&fingerprint_request));
    assert!(before.iter().all(|request| *request <= 1024 * 1024));
    assert!(emission.iter().all(|request| *request <= 65_536));
    assert!(emission.contains(&65_536));
    assert_eq!(output.len(), file_size);
}

#[test]
fn atomic_save_skips_only_the_duplicate_post_emission_source_scan() {
    static NEXT: AtomicU64 = AtomicU64::new(0);
    let directory = std::env::temp_dir().join(format!(
        "litchi-cfb-overlay-scan-count-{}-{}",
        std::process::id(),
        NEXT.fetch_add(1, Ordering::Relaxed)
    ));
    std::fs::create_dir(&directory).unwrap();
    let destination = directory.join("document.ole");
    std::fs::write(&destination, b"old destination").unwrap();

    let source = Arc::new(MutableSource::new(sample_bytes()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x65, 4_096)], limits())
        .unwrap();
    source.reads.store(0, Ordering::SeqCst);

    let report = plan.save(&destination).unwrap();

    // Atomic save retains the initial fingerprint preflight, 64 KiB emission
    // scan, and mandatory post-flush/fsync pre-rename fingerprint preflight.
    assert_eq!(
        source.reads.load(Ordering::SeqCst),
        fingerprint_chunks(file.file_size()) * 2 + publication_chunks(file.file_size())
    );
    assert_eq!(report.bytes(), file.file_size());
    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 1);

    std::fs::remove_file(destination).unwrap();
    std::fs::remove_dir(directory).unwrap();
}

#[test]
fn owned_atomic_save_uses_only_hashed_emission_scan() {
    static NEXT: AtomicU64 = AtomicU64::new(0);
    let directory = std::env::temp_dir().join(format!(
        "litchi-cfb-owned-overlay-scan-count-{}-{}",
        std::process::id(),
        NEXT.fetch_add(1, Ordering::Relaxed)
    ));
    std::fs::create_dir(&directory).unwrap();
    let destination = directory.join("document.ole");
    std::fs::write(&destination, b"old destination").unwrap();

    let source: Arc<[u8]> = Arc::from(sample_bytes());
    let reads = Arc::new(AtomicUsize::new(0));
    let file = SharedOleFile::open_owned_arc_source_for_test(source, reads.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x69, 4_096)], limits())
        .unwrap();
    reads.store(0, Ordering::SeqCst);

    let report = plan.save(&destination).unwrap();
    assert_eq!(
        reads.load(Ordering::SeqCst),
        publication_chunks(file.file_size())
    );
    assert_eq!(report.bytes(), file.file_size());
    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 1);

    std::fs::remove_file(destination).unwrap();
    std::fs::remove_dir(directory).unwrap();
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
    let preflight_chunks = fingerprint_chunks(file.file_size());
    source
        .fail_read
        .store(preflight_chunks + 2, Ordering::SeqCst);
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
    let preflight_chunks = fingerprint_chunks(file.file_size());
    source.overreport.store(true, Ordering::SeqCst);
    source
        .fail_read
        .store(preflight_chunks + 2, Ordering::SeqCst);
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
    let preflight_chunks = fingerprint_chunks(length);
    // The first reads are the coalesced preflight. Mutate the first source
    // chunk only after its bytes have been copied for emission.
    source
        .mutate_after_read
        .store(preflight_chunks + 1, Ordering::SeqCst);

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
    let fingerprint_reads = fingerprint_chunks(file.file_size());
    let publication_reads = publication_chunks(file.file_size());
    // Mutate only after the last emission read so the output-time source hash
    // has already accepted the original bytes; the mandatory final preflight
    // must then observe the changed byte before rename.
    source
        .mutate_after_read
        .store(fingerprint_reads + publication_reads, Ordering::SeqCst);

    assert!(matches!(
        plan.save(&destination),
        Err(OverlayError::SourceFingerprintChanged { .. })
    ));
    assert_eq!(
        source.reads.load(Ordering::SeqCst),
        fingerprint_reads * 2 + publication_reads
    );
    assert_eq!(std::fs::read(&destination).unwrap(), b"old destination");
    assert_eq!(std::fs::read_dir(&directory).unwrap().count(), 1);
    std::fs::remove_file(destination).unwrap();
    std::fs::remove_dir(directory).unwrap();
}

// ---------------------------------------------------------------------------
// Change 0589: an exact no-op composes the source, so one digest is both
// identities. These tests pin the value identity and every fence point that
// the elided second hasher must not move.
// ---------------------------------------------------------------------------

fn noop_overlay() -> SameLengthStreamOverlay {
    // `sample_bytes` stores `Fat4096` as 4,096 bytes of 0x22, so this
    // replacement is an exact byte no-op and plans to zero physical spans.
    SameLengthStreamOverlay::new(vec!["Fat4096".to_string()], Arc::from(vec![0x22; 4_096]))
}

/// Offset of a byte inside the unselected `LargeOpaque` payload.
///
/// `sample_bytes` fills that stream with 0x45 and no other stream or CFB
/// structure uses that value in a 4 KiB run, so a flip here leaves the
/// directory, FAT and the selected `Fat4096` stream intact.
fn unselected_payload_offset(bytes: &[u8]) -> usize {
    let run = bytes
        .windows(4_096)
        .position(|window| window.iter().all(|byte| *byte == 0x45))
        .expect("an unselected 0x45 payload run");
    run + 2_048
}

fn sha256_of(bytes: &[u8]) -> [u8; 32] {
    use sha2::{Digest as _, Sha256};
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    hasher.finalize().into()
}

/// Read ordinal of the last chunk of the first complete fingerprint pass.
///
/// A fingerprint pass requests `min(length, 1 MiB)`; the precondition
/// comparison requests one sector and emission requests 64 KiB, so on this
/// fixture the request size identifies the pass unambiguously.
fn first_fingerprint_read(file: &SharedOleFile, source: &MutableSource) -> usize {
    assert!(file.file_size() < 1024 * 1024);
    let sizes = source.request_sizes.lock().unwrap();
    sizes
        .iter()
        .position(|size| *size as u64 == file.file_size())
        .expect("a complete fingerprint chunk request")
        + 1
}

#[test]
fn noop_plan_reports_the_source_digest_as_both_identities() {
    let bytes = sample_bytes();
    let expected = sha256_of(&bytes);

    let plan = shared(bytes.clone())
        .plan_same_length_stream_overlays(vec![noop_overlay()], limits())
        .unwrap();
    assert!(plan.is_noop());
    assert_eq!(plan.changed_spans(), 0);
    assert_eq!(plan.source_fingerprint().as_bytes(), &expected);
    assert_eq!(plan.target_fingerprint().as_bytes(), &expected);

    // An effective plan over the same artifact keeps two distinct digests, and
    // its source digest is still the complete source artifact digest.
    let effective = shared(bytes.clone())
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x67, 4_096)], limits())
        .unwrap();
    assert!(!effective.is_noop());
    assert_eq!(effective.source_fingerprint().as_bytes(), &expected);
    assert_ne!(
        effective.source_fingerprint(),
        effective.target_fingerprint()
    );
}

#[test]
fn noop_plan_retains_every_complete_source_scan() {
    let bytes = sample_bytes();
    let source = Arc::new(MutableSource::new(bytes.clone()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let chunks = fingerprint_chunks(file.file_size());
    source.reads.store(0, Ordering::SeqCst);
    source.request_sizes.lock().unwrap().clear();
    let noop = file
        .plan_same_length_stream_overlays(vec![noop_overlay()], limits())
        .unwrap();
    assert!(noop.is_noop());
    let noop_fingerprint_reads = source
        .request_sizes
        .lock()
        .unwrap()
        .iter()
        .filter(|size| **size as u64 == file.file_size())
        .count();

    let effective_source = Arc::new(MutableSource::new(bytes.clone()));
    let effective_file = SharedOleFile::open(effective_source.clone()).unwrap();
    effective_source.reads.store(0, Ordering::SeqCst);
    effective_source.request_sizes.lock().unwrap().clear();
    let effective = effective_file
        .plan_same_length_stream_overlays(vec![replacement("Fat4096", 0x68, 4_096)], limits())
        .unwrap();
    assert!(!effective.is_noop());
    let effective_fingerprint_reads = effective_source
        .request_sizes
        .lock()
        .unwrap()
        .iter()
        .filter(|size| **size as u64 == effective_file.file_size())
        .count();

    // Both plans keep the generic-source read-twice-compare bracket: the
    // planning preflight and the post-reopen fence each read the complete
    // artifact once. Only the digest work differs.
    assert_eq!(noop_fingerprint_reads, chunks * 2);
    assert_eq!(noop_fingerprint_reads, effective_fingerprint_reads);
    assert_eq!(
        noop.operation_shape().planning_fingerprint_scans,
        effective.operation_shape().planning_fingerprint_scans
    );
    assert_eq!(
        noop.operation_shape().planning_fingerprint_chunks,
        effective.operation_shape().planning_fingerprint_chunks
    );
    assert_eq!(
        noop.operation_shape().planning_fingerprint_bytes * 2,
        effective.operation_shape().planning_fingerprint_bytes
    );
}

#[test]
fn noop_plan_catches_a_stable_token_mutation_between_planning_scans() {
    let bytes = sample_bytes();
    let probe = Arc::new(MutableSource::new(bytes.clone()));
    let probe_file = SharedOleFile::open(probe.clone()).unwrap();
    probe.request_sizes.lock().unwrap().clear();
    probe_file
        .plan_same_length_stream_overlays(vec![noop_overlay()], limits())
        .unwrap();
    let boundary = first_fingerprint_read(&probe_file, &probe);

    let source = Arc::new(MutableSource::new(bytes.clone()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    source.request_sizes.lock().unwrap().clear();
    source.reads.store(0, Ordering::SeqCst);
    // Mutate immediately after the first complete planning fingerprint. The
    // post-reopen fence reads the artifact again and must observe a different
    // source digest even though only one hasher now runs per pass.
    source.mutate_after_read.store(boundary, Ordering::SeqCst);

    // Flip a byte of the unselected `LargeOpaque` payload so the composed CFB
    // reopen and the `Fat4096` precondition both still succeed; only the
    // second complete fingerprint can detect this mutation.
    let unselected = unselected_payload_offset(&bytes);
    source.mutate_offset.store(unselected, Ordering::SeqCst);

    assert!(matches!(
        file.plan_same_length_stream_overlays(vec![noop_overlay()], limits()),
        Err(OverlayError::SourceFingerprintChanged { .. })
    ));
}

#[test]
fn noop_direct_write_catches_a_late_stable_token_mutation() {
    let bytes = sample_bytes();
    let source = Arc::new(MutableSource::new(bytes.clone()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![noop_overlay()], limits())
        .unwrap();
    assert!(plan.is_noop());
    let length = file.file_size();
    let fingerprint_reads = fingerprint_chunks(length);
    let publication_reads = publication_chunks(length);
    source.reads.store(0, Ordering::SeqCst);
    // Mutate after the last emission read: the emission-time hash accepted the
    // original bytes, so only the mandatory post-emission preflight sees it.
    source
        .mutate_after_read
        .store(fingerprint_reads + publication_reads, Ordering::SeqCst);

    let mut output = Vec::new();
    assert!(matches!(
        plan.write_to(&mut output),
        Err(OverlayError::IncompleteOutput {
            progress: OutputProgress::CompleteUnflushed { bytes },
            source,
        }) if bytes == length && matches!(*source, OverlayError::SourceFingerprintChanged { .. })
    ));
    assert_eq!(output.len() as u64, length);
    assert_eq!(
        source.reads.load(Ordering::SeqCst),
        fingerprint_reads * 2 + publication_reads
    );
}

#[test]
fn noop_direct_write_catches_a_mutation_inside_the_emission_scan() {
    let bytes = sample_bytes();
    let source = Arc::new(MutableSource::new(bytes.clone()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let plan = file
        .plan_same_length_stream_overlays(vec![noop_overlay()], limits())
        .unwrap();
    assert!(plan.is_noop());
    source.reads.store(0, Ordering::SeqCst);
    // Mutate after the write preflight but during emission: the single
    // emission hasher must still diverge from the retained source identity.
    source
        .mutate_after_read
        .store(fingerprint_chunks(file.file_size()) + 1, Ordering::SeqCst);

    let mut output = Vec::new();
    assert!(matches!(
        plan.write_to(&mut output),
        Err(OverlayError::IncompleteOutput { source, .. })
            if matches!(*source, OverlayError::SourceFingerprintChanged { .. })
    ));
}

#[test]
fn noop_plan_publishes_the_exact_source_bytes() {
    let bytes = sample_bytes();
    let plan = shared(bytes.clone())
        .plan_same_length_stream_overlays(vec![noop_overlay()], limits())
        .unwrap();
    let mut output = Vec::new();
    let report = plan.write_to(&mut output).unwrap();
    assert_eq!(output, bytes);
    assert_eq!(report.changed_spans(), 0);
    assert_eq!(report.source_fingerprint(), report.target_fingerprint());
    assert_eq!(report.source_fingerprint().as_bytes(), &sha256_of(&bytes));
}

// ---------------------------------------------------------------------------
// Change 0659: the two identity entry points variant B2 of change 0644 needs.
// ---------------------------------------------------------------------------

/// A generic positional source whose complete reads can be counted, so a test
/// can assert how many complete artifact scans an entry point takes.
struct ScanCountingSource {
    bytes: Vec<u8>,
    reads: AtomicUsize,
    complete_reads: AtomicUsize,
}

impl ScanCountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            reads: AtomicUsize::new(0),
            complete_reads: AtomicUsize::new(0),
        }
    }
}

impl ReadAt for ScanCountingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
        if offset == 0 && output.len() >= self.bytes.len() {
            self.complete_reads.fetch_add(1, Ordering::SeqCst);
        }
        let start = usize::try_from(offset).map_err(io::Error::other)?;
        let Some(input) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let count = input.len().min(output.len());
        output[..count].copy_from_slice(&input[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x0659, 0))
    }
}

#[test]
fn caller_bracketed_identity_is_the_empty_splice_digest_in_one_scan() {
    let bytes = sample_bytes();
    let expected = sha256_of(&bytes);

    let counted = Arc::new(ScanCountingSource::new(bytes.clone()));
    let file = SharedOleFile::open(counted.clone()).unwrap();
    counted.complete_reads.store(0, Ordering::SeqCst);
    let single = file.caller_bracketed_identity_fingerprint().unwrap();
    let single_scans = counted.complete_reads.load(Ordering::SeqCst);

    let counted = Arc::new(ScanCountingSource::new(bytes));
    let file = SharedOleFile::open(counted.clone()).unwrap();
    counted.complete_reads.store(0, Ordering::SeqCst);
    let plan = file
        .plan_same_length_stream_splices(Vec::new(), crate::StreamSpliceLimits::default())
        .unwrap();
    let paired_scans = counted.complete_reads.load(Ordering::SeqCst);

    assert_eq!(single.as_bytes(), &expected);
    assert_eq!(single, plan.source_fingerprint());
    assert_eq!(single_scans, 1, "the caller owns the read-twice-compare");
    assert_eq!(
        paired_scans, 2,
        "the self-bracketing entry point is unchanged"
    );
}

#[test]
fn caller_bracketed_identity_still_reopens_the_composed_candidate() {
    // The composed reopen is ADR 0003's proof that the candidate parses. A
    // mutation of the CFB header after the planning scan must still be
    // reported by it, as `Ole`, and not silently accepted.
    let source = Arc::new(MutableSource::new(sample_bytes()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    // Byte 30 is the sector shift: flipping it makes every later parse fail.
    source.mutate_offset.store(30, Ordering::SeqCst);
    let after_planning_scan = source.reads.load(Ordering::SeqCst) + 1;
    source
        .mutate_after_read
        .store(after_planning_scan, Ordering::SeqCst);

    let error = file.caller_bracketed_identity_fingerprint().unwrap_err();
    assert!(
        matches!(error, OverlayError::Ole(_)),
        "the composed reopen must still report a structural failure: {error:?}"
    );
}

#[test]
fn unbracketed_source_fingerprint_is_the_same_digest_with_no_reopen() {
    let bytes = sample_bytes();
    let expected = sha256_of(&bytes);

    let counted = Arc::new(ScanCountingSource::new(bytes));
    let file = SharedOleFile::open(counted.clone()).unwrap();
    counted.reads.store(0, Ordering::SeqCst);
    counted.complete_reads.store(0, Ordering::SeqCst);
    let digest = file.unbracketed_source_fingerprint().unwrap();

    assert_eq!(digest.as_bytes(), &expected);
    assert_eq!(counted.complete_reads.load(Ordering::SeqCst), 1);
    assert_eq!(
        counted.reads.load(Ordering::SeqCst),
        1,
        "a digest-only path parses no index and reopens nothing"
    );
}

#[test]
fn only_the_unbracketed_digest_survives_an_index_damaged_under_the_read() {
    // This is the property `litchi-doc`'s relocated error precedence depends
    // on: once the artifact's own CFB metadata has been corrupted under the
    // read, every entry point that reopens the composed candidate fails on the
    // damage, and only the digest-only path can still say whether the bytes
    // moved — which is what makes the caller's writer, not the file, the thing
    // the reported error blames.
    let source = Arc::new(MutableSource::new(sample_bytes()));
    let file = SharedOleFile::open(source.clone()).unwrap();
    let clean = file.unbracketed_source_fingerprint().unwrap();
    assert_eq!(clean.as_bytes(), &sha256_of(&sample_bytes()));

    source.mutate_offset.store(30, Ordering::SeqCst);
    source.change_bytes_without_version();

    let paired = file
        .plan_same_length_stream_splices(Vec::new(), crate::StreamSpliceLimits::default())
        .unwrap_err();
    assert!(
        matches!(paired, OverlayError::Ole(_)),
        "the self-bracketing entry point fails on the damage: {paired:?}"
    );
    let single = file.caller_bracketed_identity_fingerprint().unwrap_err();
    assert!(
        matches!(single, OverlayError::Ole(_)),
        "so does the caller-bracketed one: {single:?}"
    );

    let observed = file.unbracketed_source_fingerprint().unwrap();
    assert_ne!(observed, clean, "the digest-only path reports the movement");
}
