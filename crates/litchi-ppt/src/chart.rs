//! Neutral, inert chart views for legacy PowerPoint OLE objects.
//!
//! PPT owns only external-object discovery, frame attribution, storage
//! decompression, and host metadata. `[MS-OGRAPH]` Workbook validation and
//! chart traversal belong to [`litchi_ograph`]. Linked objects are never opened.

use std::io::{Cursor, Read};
use std::num::{NonZeroU32, NonZeroUsize};

use litchi_cfb::consts::STGTY_STREAM;
use litchi_ograph::chart::{Book, Refs};
use litchi_ograph::{Limits, Package as GraphPackage, PackageRef};

use super::embedded::storage::{Compression, Kind as StorageKind, Storage};
use super::ole_object::{
    PowerPointOleContainerKind, PowerPointOleExternalObject, PowerPointOleObjectCollection,
    PowerPointOleObjectSubtype,
};
use super::package::{PptError, Result};
use super::presentation::Presentation;
use super::shapes::{PictureFrameKind, ShapeEnum};
use litchi_cfb::OleFile;

/// Maximum chart-bearing OLE objects enumerated per presentation.
pub(crate) const MAX_CHART_OBJECTS: usize = 512;
/// Maximum direct children inspected in an Excel-hosted compound package.
const MAX_EXCEL_ROOT_ENTRIES: usize = 64;

const WORKBOOK: &str = "Workbook";
const BOOK: &str = "Book";

/// The producer grammar and host topology behind a chart object.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Kind {
    /// Standalone Microsoft Graph package (`MSGraph.Chart`).
    Graph,
    /// Excel compound file containing one or more chart substreams.
    Excel,
}

/// Slide shape displaying an embedded chart object.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Frame {
    slide: NonZeroUsize,
    shape: NonZeroU32,
}

impl Frame {
    /// Creates an exact slide-and-shape selector, rejecting zero identifiers.
    pub const fn new(slide: usize, shape: u32) -> Option<Self> {
        match (NonZeroUsize::new(slide), NonZeroU32::new(shape)) {
            (Some(slide), Some(shape)) => Some(Self { slide, shape }),
            _ => None,
        }
    }

    /// One-based slide number.
    pub const fn slide(self) -> usize {
        self.slide.get()
    }

    /// OfficeArt shape identifier of the OLE frame.
    pub const fn shape(self) -> u32 {
        self.shape.get()
    }
}

/// PPT metadata associated with one embedded chart object.
#[derive(Debug)]
pub struct Info {
    object_id: u32,
    persist_id: u32,
    program: Option<String>,
    frame: Option<Frame>,
}

impl Info {
    /// Declared ProgID, when present.
    pub fn program(&self) -> Option<&str> {
        self.program.as_deref()
    }

    /// Slide frame displaying this object, when attribution succeeded.
    pub const fn frame(&self) -> Option<Frame> {
        self.frame
    }

    /// Low-level `ExOleObjAtom` identifier.
    pub const fn object_id(&self) -> u32 {
        self.object_id
    }

    /// Low-level persist identifier of the `ExOleObjStg` record.
    pub const fn persist_id(&self) -> u32 {
        self.persist_id
    }
}

/// Validated standalone Microsoft Graph object.
#[derive(Debug)]
pub struct Graph {
    info: Info,
    package: Box<GraphPackage>,
    book: Book,
}

impl Graph {
    /// PPT host metadata.
    pub const fn info(&self) -> &Info {
        &self.info
    }

    /// Borrowed strict standalone OGraph package view.
    pub fn package(&self) -> PackageRef<'_> {
        self.package.as_ref().as_ref()
    }

    /// Validated Workbook chart inventory.
    pub const fn book(&self) -> &Book {
        &self.book
    }
}

/// Validated Excel-hosted chart object.
#[derive(Debug)]
pub struct Excel {
    info: Info,
    book: Book,
}

impl Excel {
    /// PPT host metadata.
    pub const fn info(&self) -> &Info {
        &self.info
    }

    /// Validated Workbook chart inventory.
    pub const fn book(&self) -> &Book {
        &self.book
    }
}

/// One embedded chart object, distinguished by its host topology.
#[derive(Debug)]
pub enum Chart {
    /// Strict standalone Microsoft Graph compound package.
    Graph(Graph),
    /// Excel compound package with a neutral validated Workbook inventory.
    Excel(Excel),
}

impl Chart {
    /// Producer and host kind.
    pub const fn kind(&self) -> Kind {
        match self {
            Self::Graph(_) => Kind::Graph,
            Self::Excel(_) => Kind::Excel,
        }
    }

    /// PPT host metadata.
    pub const fn info(&self) -> &Info {
        match self {
            Self::Graph(chart) => chart.info(),
            Self::Excel(chart) => chart.info(),
        }
    }

    /// Traverse every neutral chart substream without allocation.
    pub fn charts(&self) -> Refs<'_> {
        match self {
            Self::Graph(chart) => chart.book.charts(),
            Self::Excel(chart) => chart.book.charts(),
        }
    }

    /// Typed standalone Graph view, when applicable.
    pub const fn as_graph(&self) -> Option<&Graph> {
        match self {
            Self::Graph(chart) => Some(chart),
            Self::Excel(_) => None,
        }
    }

    /// Typed Excel-hosted view, when applicable.
    pub const fn as_excel(&self) -> Option<&Excel> {
        match self {
            Self::Excel(chart) => Some(chart),
            Self::Graph(_) => None,
        }
    }
}

/// One chart object that could not be decoded or validated.
#[derive(Debug)]
pub struct Failure {
    info: Info,
    kind: Kind,
    error: PptError,
}

impl Failure {
    /// Expected producer and host kind.
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// PPT host metadata.
    pub const fn info(&self) -> &Info {
        &self.info
    }

    /// Decode or validation failure.
    pub const fn error(&self) -> &PptError {
        &self.error
    }
}

#[derive(Debug)]
enum Entry {
    Chart(Chart),
    Failure(Failure),
}

/// Read-only inventory of embedded native chart objects.
#[derive(Debug, Default)]
pub struct Inventory {
    entries: Vec<Entry>,
}

impl Inventory {
    /// Whether no embedded chart object was discovered.
    pub fn is_empty(&self) -> bool {
        self.charts().next().is_none()
    }

    /// Number of successfully validated charts addressable by [`Self::get`].
    pub fn len(&self) -> usize {
        self.charts().count()
    }

    /// Total chart objects seen, including isolated validation failures.
    pub fn seen(&self) -> usize {
        self.entries.len()
    }

    /// Successfully validated chart objects in external-object order.
    pub fn charts(&self) -> impl Iterator<Item = &Chart> {
        self.entries.iter().filter_map(|entry| match entry {
            Entry::Chart(chart) => Some(chart),
            Entry::Failure(_) => None,
        })
    }

    /// Failed chart objects in external-object order.
    pub fn failures(&self) -> impl Iterator<Item = &Failure> {
        self.entries.iter().filter_map(|entry| match entry {
            Entry::Failure(failure) => Some(failure),
            Entry::Chart(_) => None,
        })
    }

    /// Successful chart by its semantic inventory position.
    pub fn get(&self, index: usize) -> Option<&Chart> {
        self.charts().nth(index)
    }

    /// First successful chart displayed by an exact slide frame.
    pub fn at(&self, frame: Frame) -> Option<&Chart> {
        self.charts()
            .find(|chart| chart.info().frame() == Some(frame))
    }

    /// Successful charts displayed on a one-based slide number.
    pub fn on_slide(&self, slide: usize) -> impl Iterator<Item = &Chart> {
        self.charts().filter(move |chart| {
            chart
                .info()
                .frame()
                .is_some_and(|frame| frame.slide() == slide)
        })
    }
}

fn classify(subtype: PowerPointOleObjectSubtype, program: Option<&str>) -> Option<Kind> {
    match subtype {
        PowerPointOleObjectSubtype::Graph => return Some(Kind::Graph),
        PowerPointOleObjectSubtype::ExcelChart => return Some(Kind::Excel),
        _ => {},
    }
    let base = program_base(program?);
    if base.eq_ignore_ascii_case("MSGraph.Chart") || base.eq_ignore_ascii_case("MSGraph") {
        Some(Kind::Graph)
    } else if base.eq_ignore_ascii_case("Excel.Chart") {
        Some(Kind::Excel)
    } else {
        None
    }
}

fn program_base(program: &str) -> &str {
    match program.rsplit_once('.') {
        Some((base, version))
            if !version.is_empty() && version.bytes().all(|byte| byte.is_ascii_digit()) =>
        {
            base
        },
        _ => program,
    }
}

fn decode(storage: Storage, limits: Limits) -> Result<Vec<u8>> {
    if storage.kind() != StorageKind::OleObject {
        return corrupted("chart persist ID does not reference an OLE object storage");
    }
    match storage.compression() {
        Compression::Uncompressed => {
            check_limit(
                "chart package bytes",
                storage.stored_payload_len(),
                limits.max_package_bytes,
            )?;
            Ok(storage.into_stored_bytes())
        },
        Compression::Zlib => {
            let uncompressed_len = storage.declared_uncompressed_len().ok_or_else(|| {
                PptError::Corrupted("compressed chart storage is missing its size".into())
            })?;
            let declared = usize::try_from(uncompressed_len)
                .map_err(|_| PptError::Corrupted("chart storage size exceeds usize".into()))?;
            check_limit("chart package bytes", declared, limits.max_package_bytes)?;
            let capacity = declared
                .checked_add(1)
                .ok_or_else(|| PptError::Corrupted("chart storage size overflows usize".into()))?;
            let mut bytes = Vec::new();
            bytes
                .try_reserve_exact(capacity)
                .map_err(|_| litchi_ograph::Error::Allocation {
                    resource: "PPT chart package bytes",
                })?;
            flate2::read::ZlibDecoder::new(storage.stored_bytes())
                .take(u64::from(uncompressed_len).saturating_add(1))
                .read_to_end(&mut bytes)?;
            if bytes.len() != declared {
                return corrupted("compressed chart storage size mismatch");
            }
            Ok(bytes)
        },
    }
}

/// Enumerate embedded native charts, degrading malformed payloads per object.
pub(crate) fn enumerate(presentation: &Presentation, limits: Limits) -> Result<Inventory> {
    let limits = limits.validate()?;
    let document = presentation.live_document_record()?;
    let Some(objects) = PowerPointOleObjectCollection::parse(&document)? else {
        return Ok(Inventory::default());
    };
    objects.validate_persist_mapping(&presentation.persist_mapping)?;

    let mut wanted = Vec::new();
    wanted
        .try_reserve(objects.objects.len().min(MAX_CHART_OBJECTS))
        .map_err(|_| litchi_ograph::Error::Allocation {
            resource: "PPT chart object identifiers",
        })?;
    for object in &objects.objects {
        let PowerPointOleExternalObject::Object(definition) = object else {
            continue;
        };
        if matches!(definition.kind, PowerPointOleContainerKind::Embedded(_))
            && classify(definition.object.subtype, definition.program_id.as_deref()).is_some()
        {
            if wanted.len() >= MAX_CHART_OBJECTS {
                return corrupted(format!(
                    "presentation exceeds {MAX_CHART_OBJECTS} chart objects"
                ));
            }
            wanted.push(definition.object.id);
        }
    }

    let frames = chart_frames(presentation, &wanted)?;
    let mut inventory = Inventory::default();
    inventory
        .entries
        .try_reserve(wanted.len())
        .map_err(|_| litchi_ograph::Error::Allocation {
            resource: "PPT chart inventory",
        })?;
    for object in objects.objects {
        let PowerPointOleExternalObject::Object(definition) = object else {
            continue;
        };
        if !matches!(definition.kind, PowerPointOleContainerKind::Embedded(_)) {
            continue;
        }
        let Some(kind) = classify(definition.object.subtype, definition.program_id.as_deref())
        else {
            continue;
        };
        let info = Info {
            object_id: definition.object.id,
            persist_id: definition.object.persist_id,
            program: definition.program_id,
            frame: frames
                .iter()
                .find_map(|(object, frame)| (*object == definition.object.id).then_some(*frame)),
        };
        match parse(presentation, info.persist_id, kind, limits) {
            Ok(Parsed::Graph { package, book }) => {
                inventory.entries.push(Entry::Chart(Chart::Graph(Graph {
                    info,
                    package,
                    book,
                })));
            },
            Ok(Parsed::Excel { book }) => {
                inventory
                    .entries
                    .push(Entry::Chart(Chart::Excel(Excel { info, book })));
            },
            Err(error) => inventory
                .entries
                .push(Entry::Failure(Failure { info, kind, error })),
        }
    }
    Ok(inventory)
}

enum Parsed {
    Graph {
        package: Box<GraphPackage>,
        book: Book,
    },
    Excel {
        book: Book,
    },
}

fn parse(
    presentation: &Presentation,
    persist_id: u32,
    kind: Kind,
    limits: Limits,
) -> Result<Parsed> {
    let storage = presentation.ole_storage(persist_id)?.ok_or_else(|| {
        PptError::Corrupted(format!(
            "chart object persist ID {persist_id} has no storage"
        ))
    })?;
    let package_bytes = decode(storage, limits)?;
    match kind {
        Kind::Graph => {
            let package = GraphPackage::with_limits(package_bytes, limits)?;
            let workbook = package.workbook()?.into_bytes();
            let book = Book::with_limits(workbook, limits)?;
            ensure_kind(&book, litchi_ograph::chart::Kind::Graph)?;
            Ok(Parsed::Graph {
                package: Box::new(package),
                book,
            })
        },
        Kind::Excel => {
            let workbook = extract_excel_workbook(package_bytes, limits)?;
            let book = Book::with_limits(workbook, limits)?;
            ensure_kind(&book, litchi_ograph::chart::Kind::Excel)?;
            Ok(Parsed::Excel { book })
        },
    }
}

fn ensure_kind(book: &Book, expected: litchi_ograph::chart::Kind) -> Result<()> {
    for chart in book.charts() {
        if chart?.kind() != expected {
            return corrupted("chart Workbook grammar conflicts with its PPT object kind");
        }
    }
    Ok(())
}

fn extract_excel_workbook(package: Vec<u8>, limits: Limits) -> Result<Vec<u8>> {
    let mut cfb = OleFile::open(Cursor::new(package))?;
    let stream = {
        let entries = cfb.list_directory_entries(&[])?;
        check_limit(
            "chart package root entries",
            entries.len(),
            MAX_EXCEL_ROOT_ENTRIES,
        )?;
        let mut stream = None;
        for entry in entries {
            if entry.size > as_u64(limits.max_stream_bytes) {
                return Err(limit_error(
                    "chart package stream bytes",
                    entry.size,
                    as_u64(limits.max_stream_bytes),
                ));
            }
            let candidate =
                entry.name.eq_ignore_ascii_case(WORKBOOK) || entry.name.eq_ignore_ascii_case(BOOK);
            if !candidate {
                continue;
            }
            if entry.entry_type != STGTY_STREAM {
                return corrupted("Excel chart Workbook entry is not a stream");
            }
            if stream.replace(entry.name.clone()).is_some() {
                return corrupted("Excel chart package has multiple Workbook streams");
            }
            if entry.size > as_u64(limits.max_workbook_bytes) {
                return Err(limit_error(
                    "Workbook bytes",
                    entry.size,
                    as_u64(limits.max_workbook_bytes),
                ));
            }
        }
        stream.ok_or_else(|| PptError::Corrupted("chart package has no Workbook stream".into()))?
    };
    let workbook = cfb.open_stream(&[stream.as_str()])?;
    check_limit("Workbook bytes", workbook.len(), limits.max_workbook_bytes)?;
    Ok(workbook)
}

fn chart_frames(presentation: &Presentation, wanted: &[u32]) -> Result<Vec<(u32, Frame)>> {
    let mut frames = Vec::new();
    frames
        .try_reserve(wanted.len())
        .map_err(|_| litchi_ograph::Error::Allocation {
            resource: "PPT chart frames",
        })?;
    let Ok(slides) = presentation.slides() else {
        return Ok(frames);
    };
    for slide in &slides {
        let Ok(shapes) = slide.shapes() else {
            continue;
        };
        collect_frames(shapes, slide.slide_number(), wanted, &mut frames);
    }
    Ok(frames)
}

fn collect_frames(
    shapes: &[ShapeEnum<'_>],
    slide: usize,
    wanted: &[u32],
    frames: &mut Vec<(u32, Frame)>,
) {
    for shape in shapes {
        match shape {
            ShapeEnum::Picture(picture) => {
                if picture.frame_kind() == PictureFrameKind::OleObject
                    && let Some(object) = picture.external_object_id()
                    && wanted.contains(&object)
                    && !frames.iter().any(|(existing, _)| *existing == object)
                    && let Some(frame) = Frame::new(slide, picture.properties.id)
                {
                    frames.push((object, frame));
                }
            },
            ShapeEnum::Group(group) => collect_frames(group.children(), slide, wanted, frames),
            _ => {},
        }
    }
}

fn check_limit(resource: &'static str, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(limit_error(resource, as_u64(observed), as_u64(maximum)));
    }
    Ok(())
}

fn limit_error(resource: &'static str, observed: u64, maximum: u64) -> PptError {
    litchi_ograph::Error::LimitExceeded {
        resource,
        observed,
        maximum,
    }
    .into()
}

fn as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shapes::PictureShape;
    use std::io::Write;

    #[test]
    fn subtype_and_program_identify_chart_objects() {
        assert_eq!(
            classify(PowerPointOleObjectSubtype::Graph, None),
            Some(Kind::Graph)
        );
        assert_eq!(
            classify(PowerPointOleObjectSubtype::ExcelChart, None),
            Some(Kind::Excel)
        );
        for (program, kind) in [
            ("MSGraph.Chart.8", Kind::Graph),
            ("MSGraph.Chart", Kind::Graph),
            ("MSGraph", Kind::Graph),
            ("msgraph.chart.8", Kind::Graph),
            ("Excel.Chart.8", Kind::Excel),
            ("Excel.Chart", Kind::Excel),
            ("EXCEL.CHART.8", Kind::Excel),
        ] {
            assert_eq!(
                classify(PowerPointOleObjectSubtype::Default, Some(program)),
                Some(kind),
                "{program}"
            );
        }
        for program in [
            "Excel.Sheet.8",
            "Excel.SheetMacroEnabled.12",
            "Word.Document.8",
            "PowerPoint.Show.8",
            "Equation.3",
            "Excel.ChartTool.8",
        ] {
            assert_eq!(
                classify(PowerPointOleObjectSubtype::Default, Some(program)),
                None,
                "{program}"
            );
        }
    }

    #[test]
    fn program_base_strips_only_numeric_versions() {
        assert_eq!(program_base("Excel.Chart.8"), "Excel.Chart");
        assert_eq!(program_base("Excel.Chart"), "Excel.Chart");
        assert_eq!(program_base("MSGraph"), "MSGraph");
        assert_eq!(program_base("Excel.Chart."), "Excel.Chart.");
    }

    fn storage(compression: Compression, declared: u32, data: Vec<u8>) -> Storage {
        match compression {
            Compression::Uncompressed => Storage::uncompressed(StorageKind::OleObject, data),
            Compression::Zlib => Storage::compressed(StorageKind::OleObject, declared, data),
        }
        .unwrap()
    }

    #[test]
    fn payload_decoding_is_move_first_and_bounded() {
        let bytes = b"compound".to_vec();
        let pointer = bytes.as_ptr();
        let decoded = decode(
            storage(Compression::Uncompressed, 0, bytes),
            Limits::default(),
        )
        .expect("uncompressed payload");
        assert_eq!(decoded.as_ptr(), pointer);

        let raw = vec![7u8; 65_536];
        let mut encoder =
            flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
        encoder.write_all(&raw).expect("compress");
        let compressed = encoder.finish().expect("finish compression");
        let decoded = decode(
            storage(Compression::Zlib, raw.len() as u32, compressed),
            Limits::default(),
        )
        .expect("compressed payload");
        assert_eq!(decoded, raw);

        let limits = Limits {
            max_package_bytes: 1024,
            ..Limits::default()
        };
        let bomb = storage(Compression::Zlib, 4096, vec![0x78, 0x9c]);
        assert!(decode(bomb, limits).is_err());
    }

    #[test]
    fn ole_frames_are_semantically_attributed_and_first_wins() {
        let mut chart = PictureShape::new(7);
        chart.set_frame_kind(PictureFrameKind::OleObject);
        chart.set_external_object_id(42);
        let mut nested = PictureShape::new(9);
        nested.set_frame_kind(PictureFrameKind::OleObject);
        nested.set_external_object_id(77);
        let mut group = crate::shapes::shape_enum::GroupShape::new(10);
        group.add_child(ShapeEnum::Picture(nested));
        let mut duplicate = PictureShape::new(11);
        duplicate.set_frame_kind(PictureFrameKind::OleObject);
        duplicate.set_external_object_id(42);
        let shapes = vec![
            ShapeEnum::Picture(chart),
            ShapeEnum::Group(group),
            ShapeEnum::Picture(duplicate),
        ];
        let mut frames = Vec::new();
        collect_frames(&shapes, 3, &[42, 77], &mut frames);
        assert_eq!(
            frames,
            vec![
                (42, Frame::new(3, 7).expect("valid frame")),
                (77, Frame::new(3, 9).expect("valid frame"))
            ]
        );
    }
}
