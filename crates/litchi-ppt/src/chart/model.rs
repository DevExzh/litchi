//! Contextual chart views and the read-only semantic inventory.

use std::iter::FusedIterator;
use std::num::{NonZeroU32, NonZeroUsize};

use litchi_ograph::chart::{Book, Context, Refs};
use litchi_ograph::{Package as GraphPackage, PackageRef};

use super::codec::encode_storage;
use super::package::classify;
use super::transaction::PackageEditor;
use crate::embedded::object::{ContainerKind, Editor as ObjectEditor, ExternalObject};
use crate::embedded::storage::Compression;
use crate::package::{Error, Result as PackageResult};

/// Owned host-neutral semantic view of one validated chart substream.
///
/// This is the MS-OGRAPH model from `litchi-ograph`, exposed from the PPT
/// chart host with the producer context selected from [`Kind`]. Parsed values
/// retain an exact bounded source stream and can replay it unchanged; parsed
/// mutation and fresh emission remain subject to the neutral model's explicit
/// safety errors.
pub type SemanticChart = litchi_ograph::chart::Chart;

/// Fallible semantic views over every chart substream in one PPT chart object.
///
/// Raw chart discovery remains allocation-free. Each successful item owns a
/// bounded semantic copy because the neutral semantic model must retain the
/// source stream for lossless pristine replay.
#[derive(Debug)]
pub struct SemanticCharts<'a> {
    charts: Refs<'a>,
    context: Context,
    limits: litchi_ograph::Limits,
}

impl Iterator for SemanticCharts<'_> {
    type Item = PackageResult<SemanticChart>;

    fn next(&mut self) -> Option<Self::Item> {
        let chart = self.charts.next()?;
        Some(match chart {
            Ok(chart) => {
                SemanticChart::parse_with(chart, self.context, self.limits).map_err(Into::into)
            },
            Err(error) => Err(error.into()),
        })
    }
}

impl FusedIterator for SemanticCharts<'_> {}

fn semantic_context(kind: Kind) -> Context {
    match kind {
        Kind::Graph => Context::graph(),
        Kind::Excel => Context::excel(),
    }
}

fn semantic_charts(book: &Book, kind: Kind) -> SemanticCharts<'_> {
    SemanticCharts {
        charts: book.charts(),
        context: semantic_context(kind),
        limits: book.limits(),
    }
}

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

    pub(super) fn new(
        object_id: u32,
        persist_id: u32,
        program: Option<String>,
        frame: Option<Frame>,
    ) -> Self {
        Self {
            object_id,
            persist_id,
            program,
            frame,
        }
    }
}

/// Validated standalone Microsoft Graph object.
#[derive(Debug)]
pub struct Graph {
    info: Info,
    package: Box<GraphPackage>,
    book: Book,
    compression: Compression,
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

    /// Iterate the single chart-sheet substream through the bounded semantic
    /// model, retaining source bytes for exact pristine replay.
    pub fn semantic_charts(&self) -> SemanticCharts<'_> {
        semantic_charts(&self.book, Kind::Graph)
    }

    /// Parse the single chart-sheet substream into the bounded semantic model.
    ///
    /// The raw [`Self::book`] view remains available for zero-copy traversal.
    /// This method is the opt-in semantic validation boundary and does not
    /// change the inventory's lossless raw parsing behavior.
    pub fn semantic_chart(&self) -> PackageResult<SemanticChart> {
        let mut charts = self.semantic_charts();
        let chart = charts
            .next()
            .ok_or_else(|| Error::Corrupted("Graph Workbook has no chart stream".to_string()))??;
        if let Some(chart) = charts.next() {
            let _ = chart?;
            return Err(Error::Corrupted(
                "Graph Workbook has more than one chart stream".to_string(),
            ));
        }
        Ok(chart)
    }

    /// Open a copy-on-write transaction over this chart's standalone OLE2
    /// package.
    ///
    /// The read model remains unchanged. The returned editor can replace the
    /// validated Graph chart stream and commit a new package, while preserving
    /// the original package when no edit is staged.
    pub fn edit_package(&self) -> PackageResult<PackageEditor> {
        let package = self.package.as_ref().as_ref();
        PackageEditor::with_limits(package.as_bytes().to_vec(), package.limits())
    }

    /// Move this Graph object into a package transaction without cloning its
    /// validated OLE2 allocation.
    pub fn into_package_editor(self) -> PackageResult<PackageEditor> {
        let Self { package, .. } = self;
        let package = *package;
        let limits = package.as_ref().limits();
        PackageEditor::with_limits(package.finish().into_bytes(), limits)
    }

    /// Stage a committed Graph package back into its owning PPT object.
    ///
    /// The host editor is checked against this chart's `[MS-PPT]`
    /// `ExOleObjAtom`: the external-object ID, persist ID, embedded-object
    /// container, and Graph subtype/ProgID must all still agree. The supplied
    /// [`PackageEditor`] is then committed and stored as one inert
    /// `ExOleObjStg` payload. No slide record, OfficeArt frame, or
    /// `ExObjRefAtom` is rewritten, so the existing `[MS-ODRAW]` anchor remains
    /// attached to the same external object.
    ///
    /// This is the host-side half of the typed replacement path. The package
    /// transaction still owns `[MS-OGRAPH]` chart-stream validation and
    /// replacement; this method only bridges its resulting OLE2 bytes into the
    /// PPT editor's failure-atomic storage transaction.
    pub fn replace_package(
        &self,
        editor: &mut ObjectEditor,
        package: PackageEditor,
    ) -> PackageResult<()> {
        validate_replacement_target(editor, &self.info)?;
        let bytes = package.finish()?;
        let storage = encode_storage(bytes, self.compression)?;
        editor.replace_storage(self.info.persist_id(), storage)
    }

    pub(super) fn new(
        info: Info,
        package: Box<GraphPackage>,
        book: Book,
        compression: Compression,
    ) -> Self {
        Self {
            info,
            package,
            book,
            compression,
        }
    }
}

fn validate_replacement_target(editor: &ObjectEditor, info: &Info) -> PackageResult<()> {
    let object = editor
        .objects()
        .objects
        .iter()
        .find(|object| object.id() == info.object_id())
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart external-object ID {} is not present in the host editor",
                info.object_id()
            ))
        })?;
    let ExternalObject::Object(definition) = object else {
        return Err(Error::InvalidFormat(
            "chart replacement target is not an OLE object definition".to_string(),
        ));
    };
    if !matches!(definition.kind, ContainerKind::Embedded(_)) {
        return Err(Error::InvalidFormat(
            "chart replacement target is not an embedded OLE object".to_string(),
        ));
    }
    if definition.object.persist_id != info.persist_id() {
        return Err(Error::InvalidFormat(format!(
            "chart persist ID {} does not match host object persist ID {}",
            info.persist_id(),
            definition.object.persist_id
        )));
    }
    if classify(definition.object.subtype, definition.program_id.as_deref()) != Some(Kind::Graph) {
        return Err(Error::InvalidFormat(
            "chart replacement target is not a Graph OLE object".to_string(),
        ));
    }
    Ok(())
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

    /// Iterate the Workbook's chart substreams through the bounded semantic
    /// model, retaining source bytes for exact pristine replay.
    pub fn semantic_charts(&self) -> SemanticCharts<'_> {
        semantic_charts(&self.book, Kind::Excel)
    }

    /// Parse one Excel-hosted chart by its zero-based Workbook order.
    pub fn semantic_chart(&self, index: usize) -> PackageResult<Option<SemanticChart>> {
        self.semantic_charts().nth(index).transpose()
    }

    pub(super) fn new(info: Info, book: Book) -> Self {
        Self { info, book }
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

    /// Iterate every chart substream through the bounded semantic model.
    ///
    /// This is opt-in; [`Self::charts`] continues to expose the allocation-free
    /// raw view and therefore preserves the existing permissive inventory
    /// contract for callers that only need framed records.
    pub fn semantic_charts(&self) -> SemanticCharts<'_> {
        match self {
            Self::Graph(chart) => semantic_charts(&chart.book, Kind::Graph),
            Self::Excel(chart) => semantic_charts(&chart.book, Kind::Excel),
        }
    }

    /// Parse one chart substream through the semantic model by its zero-based
    /// chart order.
    pub fn semantic_chart(&self, index: usize) -> PackageResult<Option<SemanticChart>> {
        self.semantic_charts().nth(index).transpose()
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

    /// Open a package transaction for a standalone Graph chart.
    ///
    /// Excel-hosted charts expose their validated Workbook inventory through
    /// [`Excel::book`], but their surrounding workbook package belongs to the
    /// Excel host and is intentionally not rewritten here.
    pub fn edit_package(&self) -> PackageResult<PackageEditor> {
        match self {
            Self::Graph(chart) => chart.edit_package(),
            Self::Excel(_) => Err(Error::InvalidFormat(
                "only standalone Graph chart packages support this transaction".to_string(),
            )),
        }
    }
}

/// One chart object that could not be decoded or validated.
#[derive(Debug)]
pub struct Failure {
    info: Info,
    kind: Kind,
    error: Error,
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
    pub const fn error(&self) -> &Error {
        &self.error
    }

    pub(super) fn new(info: Info, kind: Kind, error: Error) -> Self {
        Self { info, kind, error }
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

    pub(super) fn try_reserve(
        &mut self,
        additional: usize,
    ) -> Result<(), std::collections::TryReserveError> {
        self.entries.try_reserve(additional)
    }

    pub(super) fn push_chart(&mut self, chart: Chart) {
        self.entries.push(Entry::Chart(chart));
    }

    pub(super) fn push_failure(&mut self, failure: Failure) {
        self.entries.push(Entry::Failure(failure));
    }
}
