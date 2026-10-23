//! Contextual chart views and the read-only semantic inventory.

use std::iter::FusedIterator;
use std::num::{NonZeroU32, NonZeroUsize};

use litchi_ograph::chart::{Book, Context, Model, Refs};
use litchi_ograph::{Package as OgraphPackage, PackageRef};

use super::codec::encode_storage;
use super::package::classify;
use crate::builder::graph::Draft;
use crate::embedded::object::{self, Collection, ContainerKind, ExternalObject};
use crate::embedded::storage::Compression;
use crate::package::{Error, Result as PackageResult};

/// Fallible semantic views over every chart substream in one PPT chart object.
///
/// Raw chart discovery remains allocation-free. Each successful item owns a
/// bounded semantic copy because the neutral semantic model must retain the
/// source stream for lossless pristine replay.
#[derive(Debug)]
pub struct Models<'a> {
    charts: Refs<'a>,
    context: Context,
    limits: litchi_ograph::Limits,
}

impl Iterator for Models<'_> {
    type Item = PackageResult<Model>;

    fn next(&mut self) -> Option<Self::Item> {
        let next_chart = self.charts.next()?;
        Some(match next_chart {
            Ok(chart_ref) => {
                Model::parse_with(chart_ref, self.context, self.limits).map_err(Into::into)
            },
            Err(error) => Err(error.into()),
        })
    }
}

impl FusedIterator for Models<'_> {}

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
    #[must_use]
    pub const fn new(slide: usize, shape: u32) -> Option<Self> {
        match (NonZeroUsize::new(slide), NonZeroU32::new(shape)) {
            (Some(nonzero_slide), Some(nonzero_shape)) => Some(Self {
                slide: nonzero_slide,
                shape: nonzero_shape,
            }),
            _ => None,
        }
    }

    /// One-based slide number.
    #[must_use]
    pub const fn slide(self) -> usize {
        self.slide.get()
    }

    /// `OfficeArt` shape identifier of the OLE frame.
    #[must_use]
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
    /// Declared `ProgID`, when present.
    #[must_use]
    pub fn program(&self) -> Option<&str> {
        self.program.as_deref()
    }

    /// Slide frame displaying this object, when attribution succeeded.
    #[must_use]
    pub const fn frame(&self) -> Option<Frame> {
        self.frame
    }

    /// Low-level `ExOleObjAtom` identifier.
    #[must_use]
    pub const fn object_id(&self) -> u32 {
        self.object_id
    }

    /// Low-level persist identifier of the `ExOleObjStg` record.
    #[must_use]
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
    package: Box<OgraphPackage>,
    book: Book,
    compression: Compression,
}

impl Graph {
    /// PPT host metadata.
    #[must_use]
    pub const fn info(&self) -> &Info {
        &self.info
    }

    /// Borrowed strict standalone `OGraph` package view.
    #[must_use]
    pub fn package(&self) -> PackageRef<'_> {
        self.package.as_ref().as_ref()
    }

    /// Validated Workbook chart inventory.
    #[must_use]
    pub const fn book(&self) -> &Book {
        &self.book
    }

    /// Iterate the single chart-sheet substream through the bounded semantic
    /// model, retaining source bytes for exact pristine replay.
    pub fn semantic_charts(&self) -> PackageResult<Models<'_>> {
        semantic_charts(&self.book, Kind::Graph)
    }

    /// Parse the single chart-sheet substream into the bounded semantic model.
    ///
    /// The raw [`Self::book`] view remains available for zero-copy traversal.
    /// This method is the opt-in semantic validation boundary and does not
    /// change the inventory's lossless raw parsing behavior.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn semantic_chart(&self) -> PackageResult<Model> {
        let mut charts = self.semantic_charts()?;
        let chart = charts
            .next()
            .ok_or_else(|| Error::Corrupted("Graph Workbook has no chart stream".to_string()))??;
        if let Some(extra) = charts.next() {
            let _ = extra?;
            return Err(Error::Corrupted(
                "Graph Workbook has more than one chart stream".to_string(),
            ));
        }
        Ok(chart)
    }

    /// A copy-on-write draft of this chart's standalone OLE2 package.
    ///
    /// The read model remains unchanged. The draft can replace the validated
    /// Graph chart stream; it keeps the original package bytes when no
    /// replacement is staged.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn draft_package(&self) -> PackageResult<Draft> {
        let package = self.package.as_ref().as_ref();
        Draft::open_with_options(
            package.as_bytes().to_vec(),
            litchi_ograph::OpenOptions::new().limits(package.limits()),
        )
    }

    /// Move this Graph object into a package draft without cloning its
    /// validated OLE2 allocation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn into_draft_package(self) -> PackageResult<Draft> {
        let Self { package, .. } = self;
        let graph_package = *package;
        let limits = graph_package.as_ref().limits();
        Draft::open_with_options(
            graph_package.finish().into_bytes(),
            litchi_ograph::OpenOptions::new().limits(limits),
        )
    }

    /// Stage a Graph package draft back into its owning PPT object on a
    /// package-level embedded-object edit.
    ///
    /// The staged objects of `edit` are checked against this chart's
    /// `[MS-PPT]` `ExOleObjAtom`: the external-object ID, persist ID,
    /// embedded-object container, and Graph subtype/ProgID must all still
    /// agree. The draft's package bytes are then stored as one inert
    /// `ExOleObjStg` payload. No slide record, `OfficeArt` frame, or
    /// `ExObjRefAtom` is rewritten, so the existing `[MS-ODRAW]` anchor remains
    /// attached to the same external object. [`object::Edit::commit`]
    /// publishes the new presentation.
    ///
    /// # Errors
    ///
    /// Returns [`object::Error::Refused`] with
    /// [`litchi_core::edit::RefusalKind::NotFound`] when `edit` holds no
    /// object with this chart's id, with
    /// [`litchi_core::edit::RefusalKind::UnsupportedTarget`] when that object
    /// is no longer this embedded Graph object, and [`object::Error::Package`]
    /// when the draft cannot be encoded as a storage.
    pub fn replace_package(&self, edit: &mut object::Edit, package: Draft) -> object::Result<()> {
        validate_replacement_target(edit.objects()?, &self.info)?;
        let bytes = package.into_bytes()?;
        let storage = encode_storage(bytes, self.compression)?;
        edit.replace_storage(self.info.persist_id(), storage)
    }

    pub(super) fn new(
        info: Info,
        package: Box<OgraphPackage>,
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

/// Validated Excel-hosted chart object.
#[derive(Debug)]
pub struct Excel {
    info: Info,
    book: Book,
}

impl Excel {
    /// PPT host metadata.
    #[must_use]
    pub const fn info(&self) -> &Info {
        &self.info
    }

    /// Validated Workbook chart inventory.
    #[must_use]
    pub const fn book(&self) -> &Book {
        &self.book
    }

    /// Iterate the Workbook's chart substreams through the bounded semantic
    /// model, retaining source bytes for exact pristine replay.
    pub fn semantic_charts(&self) -> PackageResult<Models<'_>> {
        semantic_charts(&self.book, Kind::Excel)
    }

    /// Parse one Excel-hosted chart by its zero-based Workbook order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn semantic_chart(&self, index: usize) -> PackageResult<Option<Model>> {
        self.semantic_charts()?.nth(index).transpose()
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
    #[must_use]
    pub const fn kind(&self) -> Kind {
        match self {
            Self::Graph(_) => Kind::Graph,
            Self::Excel(_) => Kind::Excel,
        }
    }

    /// PPT host metadata.
    #[must_use]
    pub const fn info(&self) -> &Info {
        match self {
            Self::Graph(chart) => chart.info(),
            Self::Excel(chart) => chart.info(),
        }
    }

    /// Traverse every neutral chart substream without allocation.
    pub fn charts(&self) -> PackageResult<Refs<'_>> {
        Ok(match self {
            Self::Graph(chart) => chart.book.charts(),
            Self::Excel(chart) => chart.book.charts(),
        }?)
    }

    /// Iterate every chart substream through the bounded semantic model.
    ///
    /// This is opt-in; [`Self::charts`] continues to expose the allocation-free
    /// raw view and therefore preserves the existing permissive inventory
    /// contract for callers that only need framed records.
    pub fn semantic_charts(&self) -> PackageResult<Models<'_>> {
        match self {
            Self::Graph(chart) => semantic_charts(&chart.book, Kind::Graph),
            Self::Excel(chart) => semantic_charts(&chart.book, Kind::Excel),
        }
    }

    /// Parse one chart substream through the semantic model by its zero-based
    /// chart order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn semantic_chart(&self, index: usize) -> PackageResult<Option<Model>> {
        self.semantic_charts()?.nth(index).transpose()
    }

    /// Typed standalone Graph view, when applicable.
    #[must_use]
    pub const fn as_graph(&self) -> Option<&Graph> {
        match self {
            Self::Graph(chart) => Some(chart),
            Self::Excel(_) => None,
        }
    }

    /// Typed Excel-hosted view, when applicable.
    #[must_use]
    pub const fn as_excel(&self) -> Option<&Excel> {
        match self {
            Self::Excel(chart) => Some(chart),
            Self::Graph(_) => None,
        }
    }

    /// A package draft of a standalone Graph chart.
    ///
    /// Excel-hosted charts expose their validated Workbook inventory through
    /// [`Excel::book`], but their surrounding workbook package belongs to the
    /// Excel host and is intentionally not rewritten here.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn draft_package(&self) -> PackageResult<Draft> {
        match self {
            Self::Graph(chart) => chart.draft_package(),
            Self::Excel(_) => Err(Error::InvalidFormat(
                "only standalone Graph chart packages support a package draft".to_string(),
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
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// PPT host metadata.
    #[must_use]
    pub const fn info(&self) -> &Info {
        &self.info
    }

    /// Decode or validation failure.
    #[must_use]
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
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.charts().next().is_none()
    }

    /// Number of successfully validated charts addressable by [`Self::get`].
    #[must_use]
    pub fn len(&self) -> usize {
        self.charts().count()
    }

    /// Total chart objects seen, including isolated validation failures.
    #[must_use]
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
    #[must_use]
    pub fn get(&self, index: usize) -> Option<&Chart> {
        self.charts().nth(index)
    }

    /// First successful chart displayed by an exact slide frame.
    #[must_use]
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

fn semantic_context(kind: Kind) -> Context {
    match kind {
        Kind::Graph => Context::graph(),
        Kind::Excel => Context::excel(),
    }
}

fn semantic_charts(book: &Book, kind: Kind) -> PackageResult<Models<'_>> {
    Ok(Models {
        charts: book.charts()?,
        context: semantic_context(kind),
        limits: book.limits(),
    })
}

fn validate_replacement_target(objects: &Collection, info: &Info) -> object::Result<()> {
    const OPERATION: &str = "chart.replace_package";
    let target = || -> object::Error {
        crate::scoped::unsupported_target(
            OPERATION,
            "the chart's external object is no longer an embedded Graph object",
        )
        .into()
    };
    let found = objects
        .get(info.object_id())
        .ok_or_else(|| -> object::Error {
            crate::scoped::not_found(OPERATION, "the edit holds no object with the chart's id")
                .into()
        })?;
    let ExternalObject::Object(definition) = found else {
        return Err(target());
    };
    if !matches!(definition.kind, ContainerKind::Embedded(_))
        || definition.object.persist_id != info.persist_id()
        || classify(definition.object.subtype, definition.program_id.as_deref())
            != Some(Kind::Graph)
    {
        return Err(target());
    }
    Ok(())
}
