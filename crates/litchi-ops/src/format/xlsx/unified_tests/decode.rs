//! Decode counts (record 0010 §4.10, §9.2): what each read takes from the
//! package. A recording source maps every positional read to the archive
//! member it starts in, and the leaf's semantic-cache diagnostics count the
//! worksheet grids built.

use std::io;
use std::ops::Range;
use std::sync::{Arc, Mutex, PoisonError};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::SourceBackedPackage;

use super::{Mode, Options, get, kitchen, query_paths, view};
use crate::format::xlsx::{Nav, Xlsx, XlsxOwner};
use crate::options::OpenOptions;
use crate::query::walk::{PlanRef, Walk};
use crate::query::{Plan, Query};

/// A source over package bytes that records the offset of every read.
struct Recording {
    bytes: Arc<Vec<u8>>,
    offsets: Arc<Mutex<Vec<(u64, usize)>>>,
}

impl ReadAt for Recording {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        // The archive reader ends each member with an empty read at its
        // end, which is the next member's start; it reads nothing.
        if !output.is_empty() {
            self.offsets
                .lock()
                .unwrap_or_else(PoisonError::into_inner)
                .push((offset, output.len()));
        }
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        if output.is_empty() || start >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x584c_5358_5245_4144, 0))
    }
}

/// A source-backed workbook over a recording source.
struct Probe {
    file: Xlsx,
    offsets: Arc<Mutex<Vec<(u64, usize)>>>,
    /// Each archive member and the bytes it occupies: its local header,
    /// name and data, up to the next member or the central directory.
    members: Vec<(String, Range<u64>, Range<u64>)>,
}

impl Probe {
    /// Opens `bytes` as the unified opener's source route does, then forgets
    /// the reads of the open.
    fn open(bytes: Vec<u8>) -> Self {
        let mut archive = zip::ZipArchive::new(io::Cursor::new(&bytes)).expect("an archive");
        let mut starts = Vec::new();
        let mut directory = u64::MAX;
        for index in 0..archive.len() {
            let member = archive.by_index_raw(index).expect("a member");
            starts.push((
                member.name().expect("valid ZIP member name").into_owned(),
                member.header_start(),
                member.data_start().expect("local header resolved")
                    ..member.data_start().expect("local header resolved")
                        + member.compressed_size(),
            ));
            directory = directory.min(member.central_header_start());
        }
        starts.sort_by_key(|(_, start, _)| *start);
        let members = starts
            .iter()
            .enumerate()
            .map(|(index, (name, start, payload))| {
                let end = starts
                    .get(index + 1)
                    .map_or(directory, |(_, next, _)| *next);
                (name.clone(), *start..end, payload.clone())
            })
            .collect();
        let offsets = Arc::new(Mutex::new(Vec::new()));
        let source: Arc<dyn ReadAt> = Arc::new(Recording {
            bytes: Arc::new(bytes),
            offsets: Arc::clone(&offsets),
        });
        let package = SourceBackedPackage::from_read_at(
            source,
            litchi_opc::SourceLimits::default(),
            litchi_opc::SourceCacheLimits::default(),
            litchi_opc::SourceReadPolicy::exact(),
            None,
        )
        .expect("indexes");
        let file = Xlsx::from_source_package(package, &OpenOptions::default()).expect("opens");
        let probe = Self {
            file,
            offsets,
            members,
        };
        probe
            .offsets
            .lock()
            .unwrap_or_else(PoisonError::into_inner)
            .clear();
        probe
    }

    /// Members whose payload bytes were read since the previous observation.
    /// Cold selected reads may also validate predecessor ZIP local headers.
    /// Those probes must be exactly the fixed 30-byte header and never touch
    /// a predecessor's payload; assert that separately rather than counting
    /// container framing reads as document-part decoding.
    fn touched(&self) -> Vec<&str> {
        let reads =
            std::mem::take(&mut *self.offsets.lock().unwrap_or_else(PoisonError::into_inner));
        let mut payloads = Vec::new();
        for (name, member, payload) in &self.members {
            let touches_payload = reads.iter().any(|(offset, size)| {
                *offset < payload.end && offset.saturating_add(*size as u64) > payload.start
            });
            if touches_payload {
                payloads.push(name.as_str());
            } else {
                for &(offset, size) in &reads {
                    if member.contains(&offset) {
                        assert_eq!(
                            (offset, size),
                            (member.start, 30),
                            "unexpected non-payload read of {name}"
                        );
                    }
                }
            }
        }
        payloads
    }

    /// Worksheet grids built from the source so far.
    fn builds(&self) -> u64 {
        match &self.file.0.owner {
            XlsxOwner::Source(workbook) => workbook.semantic_cache_diagnostics().builds,
            XlsxOwner::Eager(_) => unreachable!("the probe opens the source owner"),
        }
    }
}

const SHEET1: &str = "xl/worksheets/sheet1.xml";
const SHEET2: &str = "xl/worksheets/sheet2.xml";
const SHEET3: &str = "xl/worksheets/sheet3.xml";
const SHEET4: &str = "xl/worksheets/sheet4.xml";
const STRINGS: &str = "xl/sharedStrings.xml";
const STYLES: &str = "xl/styles.xml";
const TABLE1: &str = "xl/tables/table1.xml";

/// `get` decodes only the payloads its target needs; cold selected streaming
/// additionally validates predecessor ZIP headers. The catalog
/// paths read nothing past the open (the package index read every
/// relationship part), a cell, row, column or range reads its own worksheet
/// and the shared parts the leaf types its values with, and a table reads
/// its table part.
#[test]
fn get_reads_only_the_target_parts() {
    let cases: &[(&str, &[&str])] = &[
        ("/", &[]),
        ("/Data", &[]),
        ("/sheet[last()]", &[]),
        ("/namedrange[2]", &[]),
        ("/Notes/C3", &[STRINGS, SHEET2]),
        ("/My Data/A1", &[STRINGS, SHEET3]),
        ("/Data/A1", &[STRINGS, STYLES, SHEET1]),
        ("/Data/Z99", &[STRINGS, STYLES, SHEET1]),
        ("Data!E2", &[STRINGS, STYLES, SHEET1]),
        ("/Data/row[3]", &[STRINGS, STYLES, SHEET1]),
        ("/Data/col[C]", &[STRINGS, STYLES, SHEET1]),
        ("/Data/A1:B2", &[STRINGS, STYLES, SHEET1]),
        ("/Data/table[1]", &[TABLE1]),
        ("/table[@name=Sales]", &[TABLE1]),
    ];
    for (path, expected) in cases {
        let probe = Probe::open(kitchen());
        get(&probe.file, path).unwrap_or_else(|error| panic!("{path}: {error}"));
        assert_eq!(probe.touched(), *expected, "{path}");
    }
    // A sheet's node view lists its rows, which reads that sheet only.
    let probe = Probe::open(kitchen());
    let sheet = get(&probe.file, "/sheet[last()]").expect("a sheet");
    let mut sink = Vec::new();
    sheet
        .view()
        .serialize_into(&mut serde_json::Serializer::new(&mut sink))
        .expect("serializes");
    assert_eq!(probe.touched(), [SHEET4]);
}

/// A walk builds each worksheet's grid once; a second query over the same
/// sheet reads nothing, and a query scoped to one sheet reads that sheet.
#[test]
fn queries_build_each_grid_once() {
    let probe = Probe::open(kitchen());
    assert_eq!(
        query_paths(&probe.file, "Notes!cell[value>80]"),
        ["/sheet[2]/C3"]
    );
    assert_eq!(probe.touched(), [STRINGS, SHEET2]);
    assert_eq!(probe.builds(), 1);
    assert_eq!(
        query_paths(&probe.file, "Notes!cell[value<80]"),
        ["/sheet[2]/C4"]
    );
    assert!(probe.touched().is_empty());
    assert_eq!(probe.builds(), 1);
    // A format key reads the styles part once for the whole grid.
    assert_eq!(query_paths(&probe.file, "Data!cell[bold=true]").len(), 3);
    assert_eq!(probe.touched(), [STYLES, SHEET1]);
    assert_eq!(probe.builds(), 2);
    assert_eq!(query_paths(&probe.file, "Data!cell[fill=yellow]").len(), 2);
    assert!(probe.touched().is_empty());
}

/// A limit stops the walk in the first worksheet; the later worksheets stay
/// unread.
#[test]
fn a_limited_walk_reads_only_what_it_reaches() {
    let probe = Probe::open(kitchen());
    let query = Query::parse("cell").expect("parses");
    let plan = Plan::compile(&query, litchi_core::FileFormat::Xlsx).expect("compiles");
    let nav = Nav::new(&probe.file);
    let first = Walk::new(&nav, PlanRef::Borrowed(&plan))
        .expect("runs")
        .first()
        .expect("reads");
    assert!(first.is_some());
    assert_eq!(probe.touched(), [STRINGS, STYLES, SHEET1]);
    // `Data` stores 25 records and `Notes` 6: the 32nd match is the first
    // record of `My Data`, and `Ärger` stays unread.
    let walk = Walk::new(&nav, PlanRef::Borrowed(&plan))
        .expect("runs")
        .limit(32);
    assert_eq!(walk.count(), 32);
    assert_eq!(probe.touched(), [SHEET2, SHEET3]);
}

/// A bounded view stops at its last line: the text of the first rows reads
/// the first worksheet only.
#[test]
fn bounded_views_stop_early() {
    let probe = Probe::open(kitchen());
    let options = Options {
        max_lines: Some(2),
        ..Options::default()
    };
    let text = view(&probe.file, Mode::Text, &options);
    assert!(
        text.ends_with("... (showed 2 rows, 7 total in sheet, use --start/--end to view more)")
    );
    assert_eq!(probe.touched(), [STRINGS, STYLES, SHEET1]);
    let options = Options {
        range: Some(super::SheetRange::parse("'My Data'!A1:B2").expect("a range")),
        ..Options::default()
    };
    view(&probe.file, Mode::Text, &options);
    assert_eq!(probe.touched(), [SHEET3]);
    // The outline counts every sheet's records and tables; the grids read
    // so far are still cached.
    view(&probe.file, Mode::Outline, &Options::default());
    assert_eq!(probe.touched(), [TABLE1, SHEET2, SHEET4]);
}
