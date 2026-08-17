use crate::{
    OleFile, OleWriter, OverlayError, SameLengthStreamSplice, SharedOleFile, StreamSpliceLimits,
};
use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use std::io::{self, Cursor};
use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};
use std::sync::{Arc, Mutex};

fn sequence(length: usize, seed: u8) -> Vec<u8> {
    (0..length)
        .map(|index| seed.wrapping_add(index as u8))
        .collect()
}

fn sample_bytes() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["Mini"], &sequence(200, 3)).unwrap();
    writer
        .create_stream(&["Fat"], &sequence(6_000, 17))
        .unwrap();
    writer
        .create_stream(&["Other"], &sequence(5_003, 29))
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn shared(bytes: Vec<u8>) -> SharedOleFile {
    SharedOleFile::open(Arc::new(OwnedSource::new(bytes))).unwrap()
}

fn splice(path: &str, offset: usize, source: &[u8], replacement: &[u8]) -> SameLengthStreamSplice {
    SameLengthStreamSplice::new(
        vec![path.to_string()],
        offset as u64,
        Arc::from(source[offset..offset + replacement.len()].to_vec()),
        Arc::from(replacement.to_vec()),
    )
}

fn limits() -> StreamSpliceLimits {
    StreamSpliceLimits::new(8, 64, 1_000_000, 1_000, 4_096).unwrap()
}

#[test]
fn first_middle_last_and_multi_stream_splices_preserve_every_other_byte() {
    let source_bytes = sample_bytes();
    let mini = sequence(200, 3);
    let fat = sequence(6_000, 17);
    let other = sequence(5_003, 29);
    let edits = vec![
        splice("Mini", 0, &mini, &[0xa0, 0xa1, 0xa2]),
        splice("Mini", 63, &mini, &[0xb0, 0xb1, 0xb2]),
        splice("Mini", 198, &mini, &[0xc0, 0xc1]),
        splice("Fat", 0, &fat, &[0xd0, 0xd1]),
        splice("Fat", 510, &fat, &[0xe0, 0xe1, 0xe2, 0xe3]),
        splice("Fat", 5_998, &fat, &[0xf0, 0xf1]),
    ];
    let plan = shared(source_bytes.clone())
        .plan_same_length_stream_splices(edits, limits())
        .unwrap();
    let view = plan.composed_source().unwrap();
    let candidate = SharedOleFile::open(Arc::new(view)).unwrap();

    let mut expected_mini = mini;
    expected_mini[0..3].copy_from_slice(&[0xa0, 0xa1, 0xa2]);
    expected_mini[63..66].copy_from_slice(&[0xb0, 0xb1, 0xb2]);
    expected_mini[198..200].copy_from_slice(&[0xc0, 0xc1]);
    let mut expected_fat = fat;
    expected_fat[0..2].copy_from_slice(&[0xd0, 0xd1]);
    expected_fat[510..514].copy_from_slice(&[0xe0, 0xe1, 0xe2, 0xe3]);
    expected_fat[5_998..6_000].copy_from_slice(&[0xf0, 0xf1]);
    assert_eq!(candidate.open_stream(&["Mini"]).unwrap(), expected_mini);
    assert_eq!(candidate.open_stream(&["Fat"]).unwrap(), expected_fat);
    assert_eq!(candidate.open_stream(&["Other"]).unwrap(), other);

    let mut output = Vec::new();
    plan.write_to(&mut output).unwrap();
    assert_eq!(output.len(), source_bytes.len());
    assert_eq!(
        source_bytes
            .iter()
            .zip(&output)
            .filter(|(left, right)| left != right)
            .count(),
        16,
        "only selected stream payload bytes may change"
    );
}

#[test]
fn every_finite_limit_accepts_exactly_the_bound_and_rejects_above_it() {
    let mini = sequence(200, 3);
    let fat = sequence(6_000, 17);
    let edits = || {
        vec![
            splice("Fat", 0, &fat, &[0x90, 0x91]),
            splice("Fat", 20, &fat, &[0x92, 0x93]),
            splice("Mini", 0, &mini, &[0x94, 0x95]),
        ]
    };
    let exact = StreamSpliceLimits::new(2, 3, 6, 3, 10).unwrap();
    shared(sample_bytes())
        .plan_same_length_stream_splices(edits(), exact)
        .unwrap();
    shared(sample_bytes())
        .plan_same_length_stream_splices(
            vec![
                splice("Fat", 0, &fat, &[0x90, 0x91]),
                splice("Mini", 0, &mini, &[0x94, 0x95]),
            ],
            exact,
        )
        .unwrap();

    for too_small in [
        StreamSpliceLimits::new(1, 3, 6, 3, 10).unwrap(),
        StreamSpliceLimits::new(2, 2, 6, 3, 10).unwrap(),
        StreamSpliceLimits::new(2, 3, 5, 3, 10).unwrap(),
        StreamSpliceLimits::new(2, 3, 6, 2, 10).unwrap(),
        StreamSpliceLimits::new(2, 3, 6, 3, 9).unwrap(),
    ] {
        assert!(matches!(
            shared(sample_bytes()).plan_same_length_stream_splices(edits(), too_small),
            Err(OverlayError::Unavailable { .. })
        ));
    }
    assert!(StreamSpliceLimits::new(0, 1, 1, 1, 1).is_err());
}

#[test]
fn length_range_overlap_and_expected_preconditions_are_checked() {
    let fat = sequence(6_000, 17);
    let file = shared(sample_bytes());
    let changed_length = SameLengthStreamSplice::new(
        vec!["Fat".to_string()],
        10,
        Arc::from(vec![fat[10], fat[11]]),
        Arc::from(vec![1]),
    );
    assert!(matches!(
        file.plan_same_length_stream_splices(vec![changed_length], limits()),
        Err(OverlayError::Unavailable { .. })
    ));
    assert!(matches!(
        file.plan_same_length_stream_splices(
            vec![
                splice("Fat", 10, &fat, &[1, 2]),
                splice("Fat", 11, &fat, &[3, 4]),
            ],
            limits(),
        ),
        Err(OverlayError::Unavailable { .. })
    ));
    assert!(matches!(
        file.plan_same_length_stream_splices(
            vec![SameLengthStreamSplice::new(
                vec!["Fat".to_string()],
                6_000,
                Arc::from(vec![0]),
                Arc::from(vec![1]),
            )],
            limits(),
        ),
        Err(OverlayError::Unavailable { .. })
    ));

    let foreign = SameLengthStreamSplice::new(
        vec!["Fat".to_string()],
        100,
        Arc::from(vec![0, 0, 0]),
        Arc::from(vec![1, 2, 3]),
    );
    assert!(matches!(
        file.plan_same_length_stream_splices(vec![foreign], limits()),
        Err(OverlayError::PreconditionFailed { offset: 100, .. })
    ));
}

#[derive(Debug)]
struct ShortInterruptedSource {
    bytes: Arc<Vec<u8>>,
    calls: AtomicUsize,
    maximum: usize,
    version: SourceVersion,
}

impl ReadAt for ShortInterruptedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let call = self.calls.fetch_add(1, Ordering::Relaxed);
        if call % 7 == 0 {
            return Err(io::ErrorKind::Interrupted.into());
        }
        let Ok(start) = usize::try_from(offset) else {
            return Ok(0);
        };
        let Some(input) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let count = input.len().min(output.len()).min(self.maximum);
        output[..count].copy_from_slice(&input[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

#[test]
fn composed_view_handles_short_and_interrupted_reads_and_is_concurrent() {
    let source_bytes = sample_bytes();
    let fat = sequence(6_000, 17);
    let source: Arc<dyn ReadAt> = Arc::new(ShortInterruptedSource {
        bytes: Arc::new(source_bytes.clone()),
        calls: AtomicUsize::new(0),
        maximum: 37,
        version: SourceVersion::new(0x51_1ce, 0),
    });
    let file = SharedOleFile::open(source).unwrap();
    let plan = file
        .plan_same_length_stream_splices(
            vec![splice("Fat", 509, &fat, &[0xa1, 0xa2, 0xa3, 0xa4])],
            limits(),
        )
        .unwrap();
    let mut expected = Vec::new();
    plan.write_to(&mut expected).unwrap();
    let view = Arc::new(plan.composed_source().unwrap());

    let mut handles = Vec::new();
    for _ in 0..8 {
        let view = Arc::clone(&view);
        let expected = expected.clone();
        handles.push(std::thread::spawn(move || {
            let mut observed = vec![0; expected.len()];
            view.read_exact_at(0, &mut observed).unwrap();
            assert_eq!(observed, expected);
        }));
    }
    for handle in handles {
        handle.join().unwrap();
    }

    let mut tail = [0_u8; 64];
    let read = loop {
        match view.read_at(source_bytes.len() as u64 - 3, &mut tail) {
            Ok(read) => break read,
            Err(error) if error.kind() == io::ErrorKind::Interrupted => {},
            Err(error) => panic!("unexpected composed-view read failure: {error}"),
        }
    };
    assert!(read <= 3);
    let mut eof = [0_u8; 1];
    assert_eq!(
        view.read_at(source_bytes.len() as u64, &mut eof).unwrap(),
        0
    );
}

#[derive(Debug)]
struct MutableSource {
    bytes: Mutex<Vec<u8>>,
    revision: AtomicU64,
    reads: AtomicUsize,
    lie: bool,
}

impl MutableSource {
    fn replace(&self, bytes: Vec<u8>) {
        *self.bytes.lock().unwrap() = bytes;
        if !self.lie {
            self.revision.fetch_add(1, Ordering::SeqCst);
        }
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.lock().unwrap().len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
        let bytes = self.bytes.lock().unwrap();
        let Ok(start) = usize::try_from(offset) else {
            return Ok(0);
        };
        let Some(input) = bytes.get(start..) else {
            return Ok(0);
        };
        let count = input.len().min(output.len());
        output[..count].copy_from_slice(&input[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            if self.lie { 0x1_1e } else { 0x5_7a1e },
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

#[derive(Debug)]
enum OwnerValidationError {
    Overlay(OverlayError),
    Rejected(&'static str),
}

impl std::fmt::Display for OwnerValidationError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Overlay(error) => write!(formatter, "{error}"),
            Self::Rejected(reason) => formatter.write_str(reason),
        }
    }
}

impl std::error::Error for OwnerValidationError {}

impl From<OverlayError> for OwnerValidationError {
    fn from(error: OverlayError) -> Self {
        Self::Overlay(error)
    }
}

#[test]
fn owner_validation_shares_the_fenced_target_and_preserves_typed_results() {
    let source_bytes = sample_bytes();
    let fat = sequence(6_000, 17);
    let callback_calls = AtomicUsize::new(0);
    let (plan, owner_version) = shared(source_bytes)
        .plan_same_length_stream_splices_with_owner(
            vec![splice("Fat", 510, &fat, &[0xa1, 0xa2, 0xa3, 0xa4])],
            limits(),
            |candidate| {
                callback_calls.fetch_add(1, Ordering::SeqCst);
                let reopened = SharedOleFile::open(Arc::new(candidate.clone()))
                    .map_err(OverlayError::from)
                    .map_err(OwnerValidationError::from)?;
                let mut observed = [0_u8; 4];
                reopened
                    .read_stream_range(&["Fat"], 510, &mut observed)
                    .map_err(OverlayError::from)
                    .map_err(OwnerValidationError::from)?;
                assert_eq!(observed, [0xa1, 0xa2, 0xa3, 0xa4]);
                candidate
                    .version()
                    .map_err(OverlayError::from)
                    .map_err(OwnerValidationError::from)
            },
        )
        .unwrap();
    assert_eq!(callback_calls.load(Ordering::SeqCst), 1);
    assert_eq!(
        owner_version,
        Some(plan.composed_source().unwrap().version().unwrap())
    );

    let callback_calls = AtomicUsize::new(0);
    let (noop, owner) = shared(sample_bytes())
        .plan_same_length_stream_splices_with_owner(
            vec![splice("Fat", 510, &fat, &fat[510..514])],
            limits(),
            |_candidate| {
                callback_calls.fetch_add(1, Ordering::SeqCst);
                Ok::<(), OwnerValidationError>(())
            },
        )
        .unwrap();
    assert!(noop.is_noop());
    assert!(owner.is_none());
    assert_eq!(callback_calls.load(Ordering::SeqCst), 0);
}

#[test]
fn owned_owner_validated_planning_elides_only_the_final_complete_fingerprint() {
    let bytes = sample_bytes();
    let fat = sequence(6_000, 17);
    let replacement = [0xa1, 0xa2, 0xa3, 0xa4];

    let generic_source = Arc::new(MutableSource {
        bytes: Mutex::new(bytes.clone()),
        revision: AtomicU64::new(0),
        reads: AtomicUsize::new(0),
        lie: false,
    });
    let generic_file = SharedOleFile::open(generic_source.clone()).unwrap();
    generic_source.reads.store(0, Ordering::SeqCst);
    let (generic_plan, generic_owner) = generic_file
        .plan_same_length_stream_splices_with_owner(
            vec![splice("Fat", 510, &fat, &replacement)],
            limits(),
            |_candidate| Ok::<_, OverlayError>("generic"),
        )
        .unwrap();
    let generic_reads = generic_source.reads.load(Ordering::SeqCst);

    let owned_reads = Arc::new(AtomicUsize::new(0));
    let owned_file =
        SharedOleFile::open_owned_arc_source_for_test(Arc::from(bytes), owned_reads.clone())
            .unwrap();
    owned_reads.store(0, Ordering::SeqCst);
    let (owned_plan, owned_owner) = owned_file
        .plan_same_length_stream_splices_with_owner(
            vec![splice("Fat", 510, &fat, &replacement)],
            limits(),
            |_candidate| Ok::<_, OverlayError>("owned"),
        )
        .unwrap();
    let owned_planning_reads = owned_reads.load(Ordering::SeqCst);

    assert_eq!(generic_owner, Some("generic"));
    assert_eq!(owned_owner, Some("owned"));
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
        owned_planning_reads
            + usize::try_from(generic_file.file_size())
                .unwrap()
                .div_ceil(1024 * 1024)
    );
}

#[test]
fn owner_validation_error_returns_no_plan_and_keeps_its_native_type() {
    let fat = sequence(6_000, 17);
    let result = shared(sample_bytes()).plan_same_length_stream_splices_with_owner(
        vec![splice("Fat", 510, &fat, &[0xa1, 0xa2, 0xa3, 0xa4])],
        limits(),
        |_candidate| Err::<(), _>(OwnerValidationError::Rejected("semantic refusal")),
    );
    assert!(matches!(
        result,
        Err(OwnerValidationError::Rejected("semantic refusal"))
    ));
}

#[test]
fn owner_validation_is_closed_by_the_final_stable_token_fingerprint() {
    let source_bytes = sample_bytes();
    let fat = sequence(6_000, 17);
    let mutable = Arc::new(MutableSource {
        bytes: Mutex::new(source_bytes.clone()),
        revision: AtomicU64::new(0),
        reads: AtomicUsize::new(0),
        lie: true,
    });
    let source: Arc<dyn ReadAt> = mutable.clone();
    let file = SharedOleFile::open(source).unwrap();
    let mut changed = source_bytes;
    changed[100] ^= 0xff;
    let result = file.plan_same_length_stream_splices_with_owner(
        vec![splice("Fat", 510, &fat, &[0xa1, 0xa2, 0xa3, 0xa4])],
        limits(),
        |_candidate| {
            mutable.replace(changed);
            Ok::<(), OverlayError>(())
        },
    );
    assert!(matches!(
        result,
        Err(OverlayError::SourceFingerprintChanged { .. })
    ));
}

#[test]
fn no_op_inverse_and_stale_source_contracts_are_exact() {
    let source_bytes = sample_bytes();
    let fat = sequence(6_000, 17);
    let no_op = shared(source_bytes.clone())
        .plan_same_length_stream_splices(vec![splice("Fat", 100, &fat, &fat[100..104])], limits())
        .unwrap();
    assert!(no_op.is_noop());
    let mut unchanged = Vec::new();
    no_op.write_to(&mut unchanged).unwrap();
    assert_eq!(unchanged, source_bytes);

    let replacement = [0xa0, 0xa1, 0xa2, 0xa3];
    let forward = shared(source_bytes.clone())
        .plan_same_length_stream_splices(vec![splice("Fat", 510, &fat, &replacement)], limits())
        .unwrap();
    let mut target = Vec::new();
    forward.write_to(&mut target).unwrap();
    let inverse = shared(target)
        .plan_same_length_stream_splices(
            vec![SameLengthStreamSplice::new(
                vec!["Fat".to_string()],
                510,
                Arc::from(replacement.to_vec()),
                Arc::from(fat[510..514].to_vec()),
            )],
            limits(),
        )
        .unwrap();
    let mut restored = Vec::new();
    inverse.write_to(&mut restored).unwrap();
    assert_eq!(restored, source_bytes);

    for lie in [false, true] {
        let mutable = Arc::new(MutableSource {
            bytes: Mutex::new(source_bytes.clone()),
            revision: AtomicU64::new(0),
            reads: AtomicUsize::new(0),
            lie,
        });
        let source: Arc<dyn ReadAt> = mutable.clone();
        let plan = SharedOleFile::open(source)
            .unwrap()
            .plan_same_length_stream_splices(vec![splice("Fat", 510, &fat, &replacement)], limits())
            .unwrap();
        let mut changed = source_bytes.clone();
        changed[100] ^= 0xff;
        mutable.replace(changed);
        let mut sink = Vec::new();
        assert!(matches!(
            plan.write_to(&mut sink),
            Err(OverlayError::SourceChanged { .. })
                | Err(OverlayError::SourceFingerprintChanged { .. })
        ));
        assert!(sink.is_empty());
        assert!(matches!(
            plan.composed_source(),
            Err(OverlayError::SourceChanged { .. })
                | Err(OverlayError::SourceFingerprintChanged { .. })
        ));
    }
}

#[derive(Debug)]
struct MutateAfterRangeSource {
    bytes: Mutex<Vec<u8>>,
    trigger_offset: u64,
    trigger_length: usize,
    armed: AtomicBool,
    fired: AtomicBool,
}

impl ReadAt for MutateAfterRangeSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.lock().unwrap().len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let mut bytes = self.bytes.lock().unwrap();
        let Ok(start) = usize::try_from(offset) else {
            return Ok(0);
        };
        let Some(input) = bytes.get(start..) else {
            return Ok(0);
        };
        let count = input.len().min(output.len());
        output[..count].copy_from_slice(&input[..count]);
        if self.armed.load(Ordering::SeqCst)
            && offset == self.trigger_offset
            && output.len() == self.trigger_length
            && !self.fired.swap(true, Ordering::SeqCst)
        {
            bytes[start] ^= 0xff;
        }
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x70c_700, 0))
    }
}

#[test]
fn stable_token_mutation_after_initial_precondition_read_is_rejected() {
    let bytes = sample_bytes();
    let parsed = shared(bytes.clone());
    let entry = parsed.find_entry(&["Fat"]).unwrap();
    let logical_offset = 100_usize;
    let physical_offset = (u64::from(entry.start_sector) + 1) * parsed.index.sector_size as u64
        + logical_offset as u64;
    let hostile = Arc::new(MutateAfterRangeSource {
        bytes: Mutex::new(bytes),
        trigger_offset: physical_offset,
        trigger_length: 4,
        armed: AtomicBool::new(false),
        fired: AtomicBool::new(false),
    });
    let source: Arc<dyn ReadAt> = hostile.clone();
    let file = SharedOleFile::open(source).unwrap();
    hostile.armed.store(true, Ordering::SeqCst);
    let fat = sequence(6_000, 17);
    assert!(matches!(
        file.plan_same_length_stream_splices(
            vec![splice(
                "Fat",
                logical_offset,
                &fat,
                &[0xa1, 0xa2, 0xa3, 0xa4]
            )],
            limits(),
        ),
        Err(OverlayError::PreconditionFailed { offset: 100, .. })
    ));
    assert!(hostile.fired.load(Ordering::SeqCst));
}

#[test]
fn mini_splice_planning_does_not_materialize_the_root_mini_stream() {
    let mini = sequence(200, 3);
    let file = shared(sample_bytes());
    assert!(!file.mini_stream_is_materialized());
    file.plan_same_length_stream_splices(
        vec![splice("Mini", 63, &mini, &[0xb1, 0xb2, 0xb3])],
        limits(),
    )
    .unwrap();
    assert!(!file.mini_stream_is_materialized());
}

#[test]
fn invalid_fat_and_minifat_topology_is_rejected_at_validated_ingress() {
    let source = sample_bytes();
    let mut invalid_fat = source.clone();
    invalid_fat[76..80].copy_from_slice(&u32::MAX.to_le_bytes());
    assert!(SharedOleFile::open(Arc::new(OwnedSource::new(invalid_fat))).is_err());

    let mut invalid_minifat = source;
    invalid_minifat[60..64].copy_from_slice(&0_u32.to_le_bytes());
    assert!(SharedOleFile::open(Arc::new(OwnedSource::new(invalid_minifat))).is_err());
}

#[test]
fn sequential_publication_reopens_with_the_existing_reader() {
    let source = sample_bytes();
    let fat = sequence(6_000, 17);
    let plan = shared(source)
        .plan_same_length_stream_splices(
            vec![splice("Fat", 511, &fat, &[0x71, 0x72, 0x73])],
            limits(),
        )
        .unwrap();
    let mut output = Vec::new();
    plan.write_to(&mut output).unwrap();
    let mut reopened = OleFile::open(Cursor::new(output)).unwrap();
    let stream = reopened.open_stream(&["Fat"]).unwrap();
    assert_eq!(&stream[511..514], &[0x71, 0x72, 0x73]);
}

#[test]
fn fragmented_fat_stream_ranges_are_mapped_in_logical_order() {
    let mut bytes = sample_bytes();
    let parsed = shared(bytes.clone());
    let entry = parsed.find_entry(&["Fat"]).unwrap();
    let sector_size = parsed.index.sector_size;
    let mut chain = Vec::new();
    let mut sector = entry.start_sector;
    for _ in 0..6_000usize.div_ceil(sector_size) {
        chain.push(sector);
        sector = parsed.index.fat[sector as usize];
    }
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

    let fat = sequence(6_000, 17);
    let plan = shared(bytes)
        .plan_same_length_stream_splices(
            vec![splice(
                "Fat",
                sector_size + 17,
                &fat,
                &[0x81, 0x82, 0x83, 0x84],
            )],
            limits(),
        )
        .unwrap();
    let mut output = Vec::new();
    plan.write_to(&mut output).unwrap();
    let mut reopened = OleFile::open(Cursor::new(output)).unwrap();
    let stream = reopened.open_stream(&["Fat"]).unwrap();
    assert_eq!(
        &stream[sector_size + 17..sector_size + 21],
        &[0x81, 0x82, 0x83, 0x84]
    );
}
