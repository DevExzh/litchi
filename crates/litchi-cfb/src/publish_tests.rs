//! Change 0761: every CFB save route publishes through one shared
//! `publish_staged`, and each caller-chosen [`Durability`] makes exactly its
//! synchronizations while keeping the bytes, the checks and the typed errors.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused publication tests use panic-on-failure assertions"
)]

use crate::writer::PublishSteps;
use crate::writer::testing::CountingSteps;
use crate::{
    OleError, OleWriter, OverlayError, OverlayLimits, SameLengthStreamOverlay, SequentialOleWriter,
    SequentialWriteError, SequentialWriteProgress, SharedOleFile, ValidatedOverlayPlan,
};
use litchi_core::{Durability, OwnedSource, ReadAt, SourceVersion};
use std::ffi::OsString;
use std::fs::{self, File};
use std::io::{self, Cursor};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};
use std::sync::{Arc, Mutex};

const LEVELS: [Durability; 3] = [Durability::Full, Durability::FileOnly, Durability::NoSync];

/// `(file syncs, parent syncs, replacements)` per entry of [`LEVELS`].
const EXPECTED_COUNTS: [(usize, usize, usize); 3] = [(1, 1, 1), (1, 0, 1), (0, 0, 1)];

const DESTINATION: &str = "document.ole";
const OLD_DESTINATION: &[u8] = b"old destination";
const STREAM: &str = "Document";
const STREAM_BYTES: usize = 5_003;

/// A private directory per test, removed on drop.
struct Scratch(PathBuf);

impl Scratch {
    fn new(label: &str) -> Self {
        static NEXT: AtomicU64 = AtomicU64::new(0);
        let path = std::env::temp_dir().join(format!(
            "litchi-cfb-0761-{label}-{}-{}",
            std::process::id(),
            NEXT.fetch_add(1, Ordering::Relaxed)
        ));
        fs::create_dir(&path).unwrap();
        Self(path)
    }

    fn join(&self, name: &str) -> PathBuf {
        self.0.join(name)
    }

    /// A destination that already holds [`OLD_DESTINATION`].
    fn seeded_destination(&self) -> PathBuf {
        let destination = self.join(DESTINATION);
        fs::write(&destination, OLD_DESTINATION).unwrap();
        destination
    }

    fn entries(&self) -> Vec<OsString> {
        let mut names: Vec<_> = fs::read_dir(&self.0)
            .unwrap()
            .map(|entry| entry.unwrap().file_name())
            .collect();
        names.sort();
        names
    }

    fn only_destination(&self) -> Vec<OsString> {
        vec![OsString::from(DESTINATION)]
    }
}

impl Drop for Scratch {
    fn drop(&mut self) {
        drop(fs::remove_dir_all(&self.0));
    }
}

fn payload(byte: u8) -> Vec<u8> {
    vec![byte; STREAM_BYTES]
}

fn ole_writer() -> OleWriter {
    let mut writer = OleWriter::new();
    writer.create_stream(&[STREAM], &payload(0x21)).unwrap();
    writer.create_stream(&["Mini"], &[0x07; 97]).unwrap();
    writer
}

fn sequential_writer() -> SequentialOleWriter<'static> {
    static PAYLOAD: [u8; STREAM_BYTES] = [0x31; STREAM_BYTES];
    let mut writer = SequentialOleWriter::new();
    writer
        .add_stream(&[STREAM], STREAM_BYTES as u64, Cursor::new(&PAYLOAD[..]))
        .unwrap();
    writer
}

fn source_bytes() -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    ole_writer().write_to(&mut output).unwrap();
    output.into_inner()
}

fn overlay_plan(file: &SharedOleFile) -> ValidatedOverlayPlan {
    file.plan_same_length_stream_overlays(
        vec![SameLengthStreamOverlay::new(
            vec![STREAM.to_owned()],
            Arc::from(payload(0x5a)),
        )],
        OverlayLimits::default(),
    )
    .unwrap()
}

/// The three CFB save routes; the overlay route with both source kinds.
#[derive(Clone, Copy, Debug)]
enum Route {
    Writer,
    Sequential,
    OwnedOverlay,
    GenericOverlay,
}

const ROUTES: [Route; 4] = [
    Route::Writer,
    Route::Sequential,
    Route::OwnedOverlay,
    Route::GenericOverlay,
];

#[derive(Debug)]
enum RouteError {
    Ole(OleError),
    Sequential(SequentialWriteError),
    Overlay(OverlayError),
}

impl RouteError {
    /// The I/O failure a pre-replacement filesystem step reported, however
    /// the route types it.
    fn staging_io(&self) -> Option<&io::Error> {
        match self {
            Self::Ole(OleError::Io(source))
            | Self::Overlay(OverlayError::Io(source))
            | Self::Sequential(SequentialWriteError::Stage {
                source,
                progress: SequentialWriteProgress::Complete { .. },
                ..
            }) => Some(source),
            _ => None,
        }
    }

    fn is_committed(&self) -> bool {
        matches!(
            self,
            Self::Ole(OleError::Committed { .. })
                | Self::Overlay(OverlayError::Committed { .. })
                | Self::Sequential(SequentialWriteError::Committed {
                    progress: SequentialWriteProgress::Complete { .. },
                    ..
                })
        )
    }
}

fn owned_file() -> SharedOleFile {
    SharedOleFile::open_owned(Arc::from(source_bytes()), SourceVersion::new(0x761, 0)).unwrap()
}

fn generic_file() -> SharedOleFile {
    SharedOleFile::open(Arc::new(OwnedSource::new(source_bytes()))).unwrap()
}

/// Publishes `route` through the crate-private steps seam.
fn publish_with_steps<S: PublishSteps>(
    route: Route,
    destination: &Path,
    durability: Durability,
    steps: &mut S,
) -> Result<(), RouteError> {
    match route {
        Route::Writer => ole_writer()
            .save_with_steps(destination, durability, steps)
            .map_err(RouteError::Ole),
        Route::Sequential => sequential_writer()
            .save_with_steps(destination, durability, steps)
            .map(drop)
            .map_err(RouteError::Sequential),
        Route::OwnedOverlay => overlay_plan(&owned_file())
            .save_with_steps(destination, durability, steps)
            .map(drop)
            .map_err(RouteError::Overlay),
        Route::GenericOverlay => overlay_plan(&generic_file())
            .save_with_steps(destination, durability, steps)
            .map(drop)
            .map_err(RouteError::Overlay),
    }
}

/// Publishes `route` through its public API: `save` for `None`,
/// `save_with_durability` otherwise.
fn publish_public(
    route: Route,
    destination: &Path,
    durability: Option<Durability>,
) -> Result<(), RouteError> {
    match (route, durability) {
        (Route::Writer, None) => ole_writer().save(destination).map_err(RouteError::Ole),
        (Route::Writer, Some(level)) => ole_writer()
            .save_with_durability(destination, level)
            .map_err(RouteError::Ole),
        (Route::Sequential, None) => sequential_writer()
            .save(destination)
            .map(drop)
            .map_err(RouteError::Sequential),
        (Route::Sequential, Some(level)) => sequential_writer()
            .save_with_durability(destination, level)
            .map(drop)
            .map_err(RouteError::Sequential),
        (Route::OwnedOverlay | Route::GenericOverlay, level) => {
            let file = match route {
                Route::OwnedOverlay => owned_file(),
                _ => generic_file(),
            };
            let plan = overlay_plan(&file);
            match level {
                None => plan.save(destination),
                Some(level) => plan.save_with_durability(destination, level),
            }
            .map(drop)
            .map_err(RouteError::Overlay)
        },
    }
}

#[test]
fn every_route_makes_exactly_its_level_synchronizations_and_identical_bytes() {
    for route in ROUTES {
        let scratch = Scratch::new("counts");
        let reference_path = scratch.join("reference.ole");
        publish_public(route, &reference_path, None).unwrap();
        let reference = fs::read(&reference_path).unwrap();
        fs::remove_file(&reference_path).unwrap();

        for (durability, expected) in LEVELS.into_iter().zip(EXPECTED_COUNTS) {
            let destination = scratch.seeded_destination();
            let mut steps = CountingSteps::default();
            publish_with_steps(route, &destination, durability, &mut steps).unwrap();
            assert_eq!(steps.counts(), expected, "{route:?} {durability:?}");
            assert_eq!(
                fs::read(&destination).unwrap(),
                reference,
                "{route:?} {durability:?} changed the published bytes"
            );
            assert_eq!(scratch.entries(), scratch.only_destination());

            // The public entry point at the same level publishes the same bytes.
            let destination = scratch.seeded_destination();
            publish_public(route, &destination, Some(durability)).unwrap();
            assert_eq!(fs::read(&destination).unwrap(), reference);
            assert_eq!(scratch.entries(), scratch.only_destination());
        }
    }
}

#[test]
fn file_sync_failure_leaves_the_destination_and_removes_the_temporary() {
    for route in ROUTES {
        for durability in [Durability::Full, Durability::FileOnly] {
            let scratch = Scratch::new("file-sync");
            let destination = scratch.seeded_destination();
            let mut steps = CountingSteps {
                fail_file_sync: true,
                ..CountingSteps::default()
            };
            let error = publish_with_steps(route, &destination, durability, &mut steps)
                .expect_err("an injected file-sync failure must fail the save");
            assert_eq!(
                error.staging_io().map(ToString::to_string).as_deref(),
                Some("injected file sync failure"),
                "{route:?} {durability:?}: {error:?}"
            );
            assert_eq!(steps.counts(), (1, 0, 0), "{route:?} {durability:?}");
            assert_eq!(fs::read(&destination).unwrap(), OLD_DESTINATION);
            assert_eq!(
                scratch.entries(),
                scratch.only_destination(),
                "{route:?} {durability:?} left a temporary file behind"
            );
        }
    }
}

#[test]
fn no_sync_never_attempts_a_synchronization_it_skips() {
    for route in ROUTES {
        let scratch = Scratch::new("no-sync");
        let destination = scratch.seeded_destination();
        let mut steps = CountingSteps {
            fail_file_sync: true,
            fail_parent_sync: true,
            ..CountingSteps::default()
        };
        publish_with_steps(route, &destination, Durability::NoSync, &mut steps)
            .unwrap_or_else(|error| panic!("{route:?}: {error:?}"));
        assert_eq!(steps.counts(), (0, 0, 1), "{route:?}");
        assert_ne!(fs::read(&destination).unwrap(), OLD_DESTINATION);
        assert_eq!(scratch.entries(), scratch.only_destination());
    }
}

#[test]
fn parent_sync_failure_is_committed_only_under_full() {
    for route in ROUTES {
        for durability in LEVELS {
            let scratch = Scratch::new("parent-sync");
            let destination = scratch.seeded_destination();
            let mut steps = CountingSteps {
                fail_parent_sync: true,
                ..CountingSteps::default()
            };
            let result = publish_with_steps(route, &destination, durability, &mut steps);
            if durability == Durability::Full {
                let error = result.expect_err("a failed parent sync under Full is reported");
                assert!(error.is_committed(), "{route:?}: {error:?}");
                assert_eq!(steps.parent_syncs, 1, "{route:?}");
            } else {
                result.unwrap_or_else(|error| panic!("{route:?} {durability:?}: {error:?}"));
                assert_eq!(steps.parent_syncs, 0, "{route:?} {durability:?}");
            }
            // The destination was replaced at every level, and nothing was
            // unlinked after the replacement.
            assert_ne!(fs::read(&destination).unwrap(), OLD_DESTINATION);
            assert_eq!(scratch.entries(), scratch.only_destination());
        }
    }
}

#[test]
fn committed_converts_into_the_typed_core_variant() {
    let error: litchi_core::Error = OleError::Committed {
        source: io::Error::other("directory sync failed"),
    }
    .into();
    assert!(
        matches!(&error, litchi_core::Error::Committed(source) if source.to_string() == "directory sync failed"),
        "{error:?}"
    );
}

/// A positional source whose bytes change, under a stable version token,
/// after a chosen number of reads.
struct LateMutation {
    bytes: Mutex<Vec<u8>>,
    reads: AtomicUsize,
    mutate_after: AtomicUsize,
}

impl LateMutation {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Mutex::new(bytes),
            reads: AtomicUsize::new(0),
            mutate_after: AtomicUsize::new(usize::MAX),
        }
    }
}

impl ReadAt for LateMutation {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.lock().unwrap().len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let call = self.reads.fetch_add(1, Ordering::SeqCst) + 1;
        let mut bytes = self.bytes.lock().unwrap();
        let offset = usize::try_from(offset).unwrap();
        let Some(available) = bytes.get(offset..) else {
            return Ok(0);
        };
        let count = output.len().min(available.len());
        output[..count].copy_from_slice(&available[..count]);
        if call == self.mutate_after.load(Ordering::SeqCst) {
            // A byte near the end of the source artifact.
            let target = bytes.len() - 7;
            bytes[target] ^= 0xff;
        }
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x761, 0))
    }
}

#[test]
fn a_late_generic_source_change_is_refused_before_rename_at_every_level() {
    for durability in LEVELS {
        let scratch = Scratch::new("late-change");
        let destination = scratch.seeded_destination();
        let source = Arc::new(LateMutation::new(source_bytes()));
        let file = SharedOleFile::open(source.clone()).unwrap();
        let plan = overlay_plan(&file);
        // Small source: one read per complete fingerprint pass and one per
        // emission pass. Mutate after the emission, so only the mandatory
        // pre-rename recheck can observe the change.
        assert!(file.file_size() < 64 * 1024);
        source.reads.store(0, Ordering::SeqCst);
        source.mutate_after.store(2, Ordering::SeqCst);

        let result = plan.save_with_durability(&destination, durability);

        assert!(
            matches!(result, Err(OverlayError::SourceFingerprintChanged { .. })),
            "{durability:?}: {result:?}"
        );
        assert_eq!(source.reads.load(Ordering::SeqCst), 3, "{durability:?}");
        assert_eq!(fs::read(&destination).unwrap(), OLD_DESTINATION);
        assert_eq!(scratch.entries(), scratch.only_destination());
    }
}

/// Real steps whose replacement substitutes the staged temporary path, the
/// hostile-directory shape the sequential route's identity cleanup defends.
struct SubstitutingSteps {
    attacker_path: Option<PathBuf>,
    displaced: PathBuf,
}

impl PublishSteps for SubstitutingSteps {
    fn sync_file(&mut self, file: &File) -> io::Result<()> {
        file.sync_all()
    }

    fn replace(&mut self, temporary: &Path, _destination: &Path) -> io::Result<()> {
        self.attacker_path = Some(temporary.to_path_buf());
        fs::rename(temporary, &self.displaced)?;
        fs::write(temporary, b"attacker replacement")?;
        Err(io::Error::other("injected replacement failure"))
    }

    fn sync_parent(&mut self, _parent: &Path) -> io::Result<()> {
        Err(io::Error::other("unreachable parent sync"))
    }
}

#[test]
fn sequential_identity_cleanup_holds_at_every_level() {
    for durability in LEVELS {
        let scratch = Scratch::new("identity");
        let destination = scratch.seeded_destination();
        let mut steps = SubstitutingSteps {
            attacker_path: None,
            displaced: scratch.join("displaced.tmp"),
        };
        let error = sequential_writer()
            .save_with_steps(&destination, durability, &mut steps)
            .expect_err("the injected replacement fails");
        assert!(
            matches!(error, SequentialWriteError::Stage { .. }),
            "{durability:?}: {error:?}"
        );
        let attacker_path = steps.attacker_path.unwrap();
        assert_eq!(fs::read(&attacker_path).unwrap(), b"attacker replacement");
        assert_eq!(fs::read(&destination).unwrap(), OLD_DESTINATION);
    }
}
