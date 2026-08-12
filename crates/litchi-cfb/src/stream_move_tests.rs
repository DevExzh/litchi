use crate::overlay::collect_chain_exact;
use crate::{
    ExistingStreamMove, OleFile, OleWriter, OverlayError, SharedOleFile, StreamMoveLimits,
};
use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use std::io::{self, Cursor};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};

fn sample_bytes() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_storage(&["Alpha"]).unwrap();
    writer.create_storage(&["Beta"]).unwrap();
    writer
        .create_stream(&["Alpha", "Mini"], &[0x31; 137])
        .unwrap();
    writer
        .create_stream(&["Alpha", "Fat"], &vec![0x41; 7_003])
        .unwrap();
    writer
        .create_stream(&["Beta", "Other"], &vec![0x51; 4_097])
        .unwrap();
    writer.create_stream(&["RootStream"], &[0x61; 91]).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn shared(bytes: Vec<u8>) -> SharedOleFile {
    SharedOleFile::open(Arc::new(OwnedSource::new(bytes))).unwrap()
}

fn request(source: &[&str], destination: &[&str]) -> ExistingStreamMove {
    ExistingStreamMove::new(
        source.iter().map(|value| (*value).to_string()).collect(),
        destination
            .iter()
            .map(|value| (*value).to_string())
            .collect(),
    )
}

#[test]
fn moves_and_renames_a_batch_without_touching_payload_or_allocation_sectors() {
    let source = sample_bytes();
    let parsed = shared(source.clone());
    let directory_sectors = collect_chain_exact(
        &parsed.index.fat,
        parsed.index.first_dir_sector,
        parsed.index.dir_entries.len() * crate::consts::DIRENTRY_SIZE / parsed.index.sector_size,
        "directory",
    )
    .unwrap();
    let plan = parsed
        .plan_stream_moves(
            vec![
                request(&["Alpha", "Mini"], &["Beta", "RenamedMini"]),
                request(&["Beta", "Other"], &["Alpha", "Other"]),
                request(&["RootStream"], &["rOOTsTREAM"]),
            ],
            StreamMoveLimits::default(),
        )
        .unwrap();

    let mut target = Vec::new();
    plan.write_to(&mut target).unwrap();
    let mut reopened = OleFile::open(Cursor::new(target.clone())).unwrap();
    assert_eq!(
        reopened.open_stream(&["Beta", "RenamedMini"]).unwrap(),
        vec![0x31; 137]
    );
    assert_eq!(
        reopened.open_stream(&["Alpha", "Other"]).unwrap(),
        vec![0x51; 4_097]
    );
    assert_eq!(
        reopened.open_stream(&["Alpha", "Fat"]).unwrap(),
        vec![0x41; 7_003]
    );
    assert_eq!(
        reopened.open_stream(&["rOOTsTREAM"]).unwrap(),
        vec![0x61; 91]
    );

    for (offset, (before, after)) in source.iter().zip(&target).enumerate() {
        if before != after {
            assert!(directory_sectors.iter().any(|sector| {
                let start = (*sector as usize + 1) * parsed.index.sector_size;
                (start..start + parsed.index.sector_size).contains(&offset)
            }));
        }
    }

    let mut restored = Vec::new();
    plan.inverse().write_to(&mut restored).unwrap();
    assert_eq!(restored, source);
}

#[test]
fn evaluates_final_casefold_collisions_atomically_and_allows_swaps() {
    let source = sample_bytes();
    let swap = shared(source.clone())
        .plan_stream_moves(
            vec![
                request(&["Alpha", "Mini"], &["Beta", "Other"]),
                request(&["Beta", "Other"], &["Alpha", "Mini"]),
            ],
            StreamMoveLimits::default(),
        )
        .unwrap();
    let moved = SharedOleFile::open(Arc::new(swap.forward().composed_source().unwrap())).unwrap();
    assert_eq!(
        moved.open_stream(&["Beta", "Other"]).unwrap(),
        vec![0x31; 137]
    );
    assert_eq!(
        moved.open_stream(&["Alpha", "Mini"]).unwrap(),
        vec![0x51; 4_097]
    );

    assert!(matches!(
        shared(source).plan_stream_moves(
            vec![request(&["Alpha", "Mini"], &["Beta", "oTHER"])],
            StreamMoveLimits::default(),
        ),
        Err(OverlayError::Unavailable { .. })
    ));
}

#[test]
fn rejects_storage_moves_duplicate_sources_missing_parents_and_invalid_names() {
    let source = sample_bytes();
    for moves in [
        vec![request(&["Alpha"], &["Moved"])],
        vec![
            request(&["Alpha", "Mini"], &["Beta", "One"]),
            request(&["aLPHA", "mINI"], &["Beta", "Two"]),
        ],
        vec![request(&["Alpha", "Mini"], &["Missing", "Moved"])],
        vec![request(&["Alpha", "Mini"], &["Beta", "bad/name"])],
    ] {
        assert!(
            shared(source.clone())
                .plan_stream_moves(moves, StreamMoveLimits::default())
                .is_err()
        );
    }
}

#[test]
fn no_op_and_every_finite_limit_are_exact() {
    let source = sample_bytes();
    let empty = shared(source.clone())
        .plan_stream_moves(Vec::new(), StreamMoveLimits::default())
        .unwrap();
    let mut output = Vec::new();
    empty.write_to(&mut output).unwrap();
    assert_eq!(output, source);
    let exact_noop = shared(source.clone())
        .plan_stream_moves(
            vec![request(&["Alpha", "Mini"], &["Alpha", "Mini"])],
            StreamMoveLimits::default(),
        )
        .unwrap();
    let mut output = Vec::new();
    exact_noop.write_to(&mut output).unwrap();
    assert_eq!(output, source);

    let one = vec![request(&["Alpha", "Mini"], &["Beta", "Moved"])];
    let parsed = shared(source.clone());
    let entries = parsed.index.dir_entries.len();
    let bytes = entries * crate::consts::DIRENTRY_SIZE;
    let exact = StreamMoveLimits::new(1, 4, 18, entries, bytes, 2).unwrap();
    parsed.plan_stream_moves(one.clone(), exact).unwrap();
    for too_small in [
        StreamMoveLimits::new(1, 3, 18, entries, bytes, 2).unwrap(),
        StreamMoveLimits::new(1, 4, 17, entries, bytes, 2).unwrap(),
        StreamMoveLimits::new(1, 4, 18, entries - 1, bytes, 2).unwrap(),
        StreamMoveLimits::new(1, 4, 18, entries, bytes - 1, 2).unwrap(),
        StreamMoveLimits::new(1, 4, 18, entries, bytes, 1).unwrap(),
    ] {
        assert!(matches!(
            shared(source.clone()).plan_stream_moves(one.clone(), too_small),
            Err(OverlayError::Unavailable { .. })
        ));
    }
    assert!(StreamMoveLimits::new(0, 1, 1, 1, 1, 1).is_err());
}

#[test]
fn version_four_directory_fragments_are_bounded_before_collection() {
    let mut writer = OleWriter::with_sector_size(4_096).unwrap();
    for index in 0..40 {
        writer
            .create_stream(&[&format!("Stream{index:02}")], &[index as u8; 17])
            .unwrap();
    }
    let mut serialized = Cursor::new(Vec::new());
    writer.write_to(&mut serialized).unwrap();
    let bytes = serialized.into_inner();
    let parsed = shared(bytes.clone());
    let directory_bytes = parsed.index.dir_entries.len() * crate::consts::DIRENTRY_SIZE;
    let one_fragment =
        StreamMoveLimits::new(1, 2, 64, parsed.index.dir_entries.len(), directory_bytes, 1)
            .unwrap();
    assert!(matches!(
        parsed.plan_stream_moves(vec![request(&["Stream39"], &["Moved39"])], one_fragment,),
        Err(OverlayError::Unavailable { .. })
    ));

    let plan = shared(bytes)
        .plan_stream_moves(
            vec![request(&["Stream39"], &["Moved39"])],
            StreamMoveLimits::default(),
        )
        .unwrap();
    let moved = SharedOleFile::open(Arc::new(plan.forward().composed_source().unwrap())).unwrap();
    assert_eq!(moved.open_stream(&["Moved39"]).unwrap(), vec![39; 17]);
}

struct MutableSource {
    bytes: Mutex<Vec<u8>>,
    revision: AtomicU64,
}

impl MutableSource {
    fn mutate(&self) {
        self.bytes.lock().unwrap()[0x22] ^= 1;
        self.revision.fetch_add(1, Ordering::Release);
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.lock().unwrap().len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let bytes = self.bytes.lock().unwrap();
        let start = usize::try_from(offset).unwrap();
        if start >= bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(bytes.len() - start);
        output[..count].copy_from_slice(&bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            81,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

#[test]
fn forward_and_inverse_refuse_a_stale_source() {
    let source = Arc::new(MutableSource {
        bytes: Mutex::new(sample_bytes()),
        revision: AtomicU64::new(0),
    });
    let opened = SharedOleFile::open(source.clone()).unwrap();
    let plan = opened
        .plan_stream_moves(
            vec![request(&["Alpha", "Mini"], &["Beta", "Moved"])],
            StreamMoveLimits::default(),
        )
        .unwrap();
    source.mutate();
    assert!(matches!(
        plan.write_to(&mut Vec::new()),
        Err(OverlayError::SourceChanged { .. })
            | Err(OverlayError::SourceFingerprintChanged { .. })
    ));
    assert!(plan.inverse().write_to(&mut Vec::new()).is_err());
}
