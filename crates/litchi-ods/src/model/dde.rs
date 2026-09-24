//! Bounded, inert inspection and editing of ODF Dynamic Data Exchange metadata.
//!
//! This module models declarations, typed cache authoring, and exact XML patches. It contains no DDE
//! client, refresh path, resolver, process launcher, or ambient I/O.

use core::fmt;
use litchi_core::{
    Budget, CancellationSource, Error as CoreError, ExecutionContext, ExecutionError,
    ExecutionLimits, Profile, Reservation, Resource,
};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{
    mem::size_of,
    num::{NonZeroU64, NonZeroUsize},
    ops::Range,
    sync::Arc,
};

mod lexical;
mod source_backed;
mod transaction;

pub use source_backed::{SourceCommit, SourceEdit, SourcePatch, SourceSnapshot};
pub use transaction::{
    CachedCell, CachedRow, CachedTable, CachedValue, Commit, Edit, LinkSpec, Patch, SheetSelector,
};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const ODF_SPREADSHEET_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.spreadsheet";
const MAX_INPUT_BYTES: usize = 256 * 1024 * 1024;
const MAX_LINKS: usize = 65_536;
const MAX_SHEET_SOURCES: usize = 65_536;
const MAX_TEXT_BYTES: usize = 65_536;
const MAX_CACHED_TABLE_BYTES: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 1_024;
const MAX_ATTRIBUTES: usize = 256;
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;
const UTF8_BOM: &[u8] = b"\xEF\xBB\xBF";
const NAMESPACE_RESOLVER_BASE_BYTES: usize = 512;
const NAMESPACE_BINDING_BYTES: usize = 128;
const PARSER_CAPACITY_FACTOR: usize = 2;
const INITIAL_SCOPE_CAPACITY: usize = 8;

/// A DDE metadata inspection result.
pub type Result<T> = std::result::Result<T, Error>;

/// Errors produced while inspecting inert DDE metadata.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A configured or hard resource limit was exceeded.
    ResourceLimit {
        /// The bounded resource.
        resource: &'static str,
        /// The observed or configured value.
        actual: usize,
        /// The maximum accepted value.
        maximum: usize,
    },
    /// A resource refusal returned by the caller's execution budget.
    ///
    /// Keep the core diagnostic intact so callers can inspect the exact
    /// resource, observed value, limit, and budget scope that rejected the
    /// operation.
    ExecutionLimit(litchi_core::ResourceLimit),
    /// The XML stream could not be decoded.
    InvalidXml(String),
    /// The document has invalid DDE structure or content.
    InvalidStructure(String),
    /// An XML byte position cannot be represented on this platform.
    PositionOverflow,
    /// The caller's cooperative cancellation token was cancelled.
    Cancelled,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ResourceLimit {
                resource,
                actual,
                maximum,
            } => write!(
                formatter,
                "{resource} limit exceeded: observed {actual}, maximum {maximum}"
            ),
            Self::ExecutionLimit(limit) => write!(formatter, "{limit}"),
            Self::InvalidXml(message) => write!(formatter, "invalid XML: {message}"),
            Self::InvalidStructure(message) => formatter.write_str(message),
            Self::PositionOverflow => formatter.write_str("XML byte position exceeds usize"),
            Self::Cancelled => formatter.write_str("ODS DDE operation cancelled"),
        }
    }
}

impl std::error::Error for Error {}

impl From<Error> for CoreError {
    fn from(error: Error) -> Self {
        match error {
            Error::ResourceLimit {
                resource,
                actual,
                maximum,
            } => CoreError::ResourceLimit(litchi_core::ResourceLimit {
                resource: match resource {
                    "input bytes" => Resource::InputBytes,
                    "output bytes" => Resource::OutputBytes,
                    "DDE links" | "links" | "sheet sources" | "DDE objects" => Resource::Objects,
                    "DDE XML depth" | "XML depth" => Resource::Depth,
                    "DDE work" => Resource::Work,
                    _ => Resource::Memory,
                },
                observed: actual as u64,
                limit: maximum as u64,
                scope: Arc::from("ods-dde"),
            }),
            Error::ExecutionLimit(limit) => CoreError::ResourceLimit(limit),
            Error::Cancelled => CoreError::Unsupported("ODS DDE operation cancelled".to_string()),
            Error::InvalidXml(message) => CoreError::InvalidFormat(message),
            Error::InvalidStructure(message) => CoreError::InvalidFormat(message),
            Error::PositionOverflow => {
                CoreError::InvalidFormat("XML byte position exceeds usize".to_string())
            },
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Table,
    MarkupCompatibility,
    Other,
}

/// Resource limits for inert DDE inspection.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    input_bytes: usize,
    links: usize,
    sheet_sources: usize,
    text_bytes: usize,
    cached_table_bytes: usize,
    depth: usize,
    output_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            input_bytes: MAX_INPUT_BYTES,
            links: MAX_LINKS,
            sheet_sources: MAX_SHEET_SOURCES,
            text_bytes: MAX_TEXT_BYTES,
            cached_table_bytes: MAX_CACHED_TABLE_BYTES,
            depth: MAX_DEPTH,
            output_bytes: MAX_OUTPUT_BYTES,
        }
    }
}

impl Limits {
    #[must_use]
    pub const fn with_input_bytes(mut self, value: usize) -> Self {
        self.input_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_links(mut self, value: usize) -> Self {
        self.links = value;
        self
    }

    #[must_use]
    pub const fn with_sheet_sources(mut self, value: usize) -> Self {
        self.sheet_sources = value;
        self
    }

    #[must_use]
    pub const fn with_text_bytes(mut self, value: usize) -> Self {
        self.text_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_cached_table_bytes(mut self, value: usize) -> Self {
        self.cached_table_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_depth(mut self, value: usize) -> Self {
        self.depth = value;
        self
    }

    /// Sets the maximum candidate `content.xml` size produced by a transaction.
    #[must_use]
    pub const fn with_output_bytes(mut self, value: usize) -> Self {
        self.output_bytes = value;
        self
    }

    /// Returns the maximum candidate `content.xml` size produced by a transaction.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    pub(crate) const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    pub(crate) const fn depth(self) -> usize {
        self.depth
    }

    fn validate(self) -> Result<Self> {
        for (name, value, ceiling) in [
            ("input bytes", self.input_bytes, MAX_INPUT_BYTES),
            ("links", self.links, MAX_LINKS),
            ("sheet sources", self.sheet_sources, MAX_SHEET_SOURCES),
            ("text bytes", self.text_bytes, MAX_TEXT_BYTES),
            (
                "cached table bytes",
                self.cached_table_bytes,
                MAX_CACHED_TABLE_BYTES,
            ),
            ("XML depth", self.depth, MAX_DEPTH),
            ("output bytes", self.output_bytes, MAX_OUTPUT_BYTES),
        ] {
            if value > ceiling {
                return Err(Error::ResourceLimit {
                    resource: name,
                    actual: value,
                    maximum: ceiling,
                });
            }
        }
        Ok(self)
    }
}

/// ODF conversion policy retained without applying it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ConversionMode {
    /// No conversion mode was specified.
    Unspecified,
    IntoDefaultStyleDataStyle,
    IntoEnglishNumber,
    KeepText,
}

/// Whether a DDE source requests automatic updates.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum AutomaticUpdate {
    /// The source did not specify an update policy.
    #[default]
    Unspecified,
    /// Automatic updates were requested by the source document.
    Enabled,
    /// Automatic updates were explicitly disabled.
    Disabled,
}

impl From<Option<bool>> for AutomaticUpdate {
    fn from(value: Option<bool>) -> Self {
        match value {
            Some(true) => Self::Enabled,
            Some(false) => Self::Disabled,
            None => Self::Unspecified,
        }
    }
}

impl ConversionMode {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "into-default-style-data-style" => Ok(Self::IntoDefaultStyleDataStyle),
            "into-english-number" => Ok(Self::IntoEnglishNumber),
            "keep-text" => Ok(Self::KeepText),
            _ => Err(invalid(format!("invalid office:conversion-mode '{value}'"))),
        }
    }
}

/// A non-executing `office:dde-source` declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Source {
    application: String,
    topic: String,
    item: String,
    name: Option<String>,
    conversion_mode: ConversionMode,
    automatic_update: AutomaticUpdate,
}

impl Source {
    /// Creates a detached, inert DDE source descriptor.
    ///
    /// # Errors
    ///
    /// Returns an error when a required identifier is empty, too large, or
    /// contains invalid XML character data.
    pub fn new(
        application: impl Into<String>,
        topic: impl Into<String>,
        item: impl Into<String>,
    ) -> Result<Self> {
        let source = Self {
            application: application.into(),
            topic: topic.into(),
            item: item.into(),
            name: None,
            conversion_mode: ConversionMode::Unspecified,
            automatic_update: AutomaticUpdate::Unspecified,
        };
        validate_required_source_value("office:dde-application", &source.application)?;
        validate_required_source_value("office:dde-topic", &source.topic)?;
        validate_required_source_value("office:dde-item", &source.item)?;
        Ok(source)
    }

    /// Returns a copy with a checked optional source name.
    ///
    /// # Errors
    ///
    /// Returns an error when `name` is empty, too large, or contains invalid
    /// XML character data.
    pub fn named(mut self, source_name: impl Into<String>) -> Result<Self> {
        let name = source_name.into();
        validate_required_source_value("office:name", &name)?;
        self.name = Some(name);
        Ok(self)
    }

    /// Returns a copy with the requested conversion policy.
    #[must_use]
    pub const fn with_conversion_mode(mut self, mode: ConversionMode) -> Self {
        self.conversion_mode = mode;
        self
    }

    /// Returns a copy with the requested automatic-update policy.
    #[must_use]
    pub const fn with_automatic_update(mut self, policy: AutomaticUpdate) -> Self {
        self.automatic_update = policy;
        self
    }

    #[must_use]
    pub fn application(&self) -> &str {
        &self.application
    }

    #[must_use]
    pub fn topic(&self) -> &str {
        &self.topic
    }

    #[must_use]
    pub fn item(&self) -> &str {
        &self.item
    }

    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    #[must_use]
    pub const fn conversion_mode(&self) -> ConversionMode {
        self.conversion_mode
    }

    #[must_use]
    pub const fn automatic_update(&self) -> AutomaticUpdate {
        self.automatic_update
    }
}

/// One formula DDE link and its exact cached `table:table` subtree.
#[derive(Clone, Debug)]
pub struct Link {
    source: Source,
    content: Arc<str>,
    cached_table: Range<usize>,
}

impl Link {
    #[must_use]
    pub fn source(&self) -> &Source {
        &self.source
    }

    /// Return cached table XML without copying or interpreting its values.
    #[must_use]
    pub fn cached_table_xml(&self) -> &str {
        &self.content[self.cached_table.clone()]
    }
}

/// A sheet-local DDE source declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetSource {
    sheet: String,
    source: Source,
}

impl SheetSource {
    #[must_use]
    pub fn sheet(&self) -> &str {
        &self.sheet
    }

    #[must_use]
    pub fn source(&self) -> &Source {
        &self.source
    }
}

/// Immutable source-bound inventory of all spreadsheet DDE declarations.
#[derive(Clone, Debug)]
pub struct Snapshot {
    content: Arc<str>,
    inventory: Arc<Inventory>,
    limits: Limits,
    context: ExecutionContext,
    enforce_context_lineage: bool,
    // These reservations pin retained source/candidate bytes for the
    // lifetime of a snapshot, including when it is held by a detached patch.
    _input_reservation: Arc<Reservation>,
    _memory_reservation: Arc<Reservation>,
    _output_reservation: Option<Arc<Reservation>>,
}

#[derive(Clone, Debug)]
struct Inventory {
    links: Vec<Link>,
    sheet_sources: Vec<SheetSource>,
    /// Physical worksheet index for every entry in `sheet_sources`.
    ///
    /// Sheet names are not required to be unique in a malformed or partially
    /// authored source.  Transactions therefore resolve an existing source
    /// by this occurrence index instead of searching for the first matching
    /// name.
    sheet_table_indices: Vec<usize>,
    table_names: Vec<Option<String>>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum DocumentRoot {
    Content,
    Flat,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RawEventKind {
    Start,
    Empty,
    End,
    Other,
}

#[derive(Clone, Copy, Debug)]
struct RawEvent {
    kind: RawEventKind,
    end: usize,
    namespace_bytes: usize,
    namespace_bindings: usize,
    element_name_bytes: usize,
}

#[derive(Clone, Copy, Debug)]
struct NamespaceScope {
    namespace_bytes: usize,
    namespace_bindings: usize,
    element_name_bytes: usize,
    is_start: bool,
}

/// Charges storage that `NsReader` may allocate while resolving the next
/// event. Slice events borrow the source, but namespace declarations and
/// checked end-tag names are copied into resolver-owned buffers before the
/// event is returned. The raw event envelope lets us reserve that bounded
/// storage before invoking the reader.
struct ParserScratch {
    reservation: Reservation,
    reserved_bound: usize,
    depth_limit: usize,
    scopes: Vec<NamespaceScope>,
    depth_reservations: Vec<Reservation>,
    pending_empty_depth: Option<Reservation>,
    active_namespace_bytes: usize,
    active_namespace_bindings: usize,
    active_element_name_bytes: usize,
    active_start_elements: usize,
}

impl ParserScratch {
    fn new(context: &ExecutionContext, depth: usize) -> Result<Self> {
        let initial_capacity = depth.min(INITIAL_SCOPE_CAPACITY);
        let initial_bound = vector_capacity_bound(
            initial_capacity,
            initial_capacity,
            NAMESPACE_RESOLVER_BASE_BYTES,
        )?;
        let reservation = context
            .reserve(
                Resource::Memory,
                u64::try_from(initial_bound)
                    .map_err(|_| invalid("DDE parser scratch size exceeds u64"))?,
            )
            .map_err(map_execution)?;
        let mut scopes = Vec::new();
        scopes
            .try_reserve_exact(initial_capacity)
            .map_err(|_| invalid("DDE parser scope allocation failed"))?;
        let mut depth_reservations = Vec::new();
        depth_reservations
            .try_reserve_exact(initial_capacity)
            .map_err(|_| invalid("DDE parser depth stack allocation failed"))?;
        Ok(Self {
            reservation,
            reserved_bound: initial_bound,
            depth_limit: depth,
            scopes,
            depth_reservations,
            pending_empty_depth: None,
            active_namespace_bytes: 0,
            active_namespace_bindings: 0,
            active_element_name_bytes: 0,
            active_start_elements: 0,
        })
    }

    fn prepare(&mut self, event: RawEvent, context: &ExecutionContext) -> Result<()> {
        match event.kind {
            RawEventKind::Start | RawEventKind::Empty => {
                let event_depth = self
                    .active_start_elements
                    .checked_add(1)
                    .ok_or_else(|| invalid("DDE XML depth overflow"))?;
                if event_depth > self.depth_limit {
                    return Err(invalid("DDE XML exceeds the nesting limit"));
                }
                let namespace_bytes = self
                    .active_namespace_bytes
                    .checked_add(event.namespace_bytes)
                    .ok_or_else(|| invalid("DDE namespace storage size overflows"))?;
                let namespace_bindings = self
                    .active_namespace_bindings
                    .checked_add(event.namespace_bindings)
                    .ok_or_else(|| invalid("DDE namespace binding count overflows"))?;
                let (element_name_bytes, active_start_elements) =
                    if event.kind == RawEventKind::Start {
                        (
                            self.active_element_name_bytes
                                .checked_add(event.element_name_bytes)
                                .ok_or_else(|| invalid("DDE element-name storage overflows"))?,
                            self.active_start_elements
                                .checked_add(1)
                                .ok_or_else(|| invalid("DDE open-element count overflows"))?,
                        )
                    } else {
                        (
                            self.active_element_name_bytes
                                .checked_add(event.element_name_bytes)
                                .ok_or_else(|| invalid("DDE event-name storage overflows"))?,
                            self.active_start_elements,
                        )
                    };
                self.ensure_scope_capacity(context)?;
                let depth_reservation =
                    context.reserve(Resource::Depth, 1).map_err(map_execution)?;
                if event.kind == RawEventKind::Start {
                    self.ensure_depth_reservation_capacity(context)?;
                }
                self.reserve_bound(
                    self.storage_bound(
                        namespace_bytes,
                        namespace_bindings,
                        element_name_bytes,
                        active_start_elements,
                    )?,
                    context,
                )?;
                self.scopes.push(NamespaceScope {
                    namespace_bytes: event.namespace_bytes,
                    namespace_bindings: event.namespace_bindings,
                    element_name_bytes: event.element_name_bytes,
                    is_start: event.kind == RawEventKind::Start,
                });
                if event.kind == RawEventKind::Start {
                    self.depth_reservations.push(depth_reservation);
                } else {
                    self.pending_empty_depth = Some(depth_reservation);
                }
                self.active_namespace_bytes = namespace_bytes;
                self.active_namespace_bindings = namespace_bindings;
                self.active_element_name_bytes = element_name_bytes;
                self.active_start_elements = active_start_elements;
            },
            RawEventKind::End => {
                if self.scopes.is_empty() || self.depth_reservations.is_empty() {
                    return Err(invalid("DDE namespace scope stack underflow"));
                }
                self.reserve_bound(
                    self.storage_bound(
                        self.active_namespace_bytes,
                        self.active_namespace_bindings,
                        self.active_element_name_bytes
                            .checked_add(event.element_name_bytes)
                            .ok_or_else(|| invalid("DDE event-name storage overflows"))?,
                        self.active_start_elements,
                    )?,
                    context,
                )?;
            },
            RawEventKind::Other => {},
        }
        Ok(())
    }

    fn finish(&mut self, expected: RawEventKind, actual: RawEventKind) -> Result<()> {
        if expected != actual {
            return Err(invalid(
                "DDE raw XML event classification disagrees with reader",
            ));
        }
        match actual {
            RawEventKind::Empty => {
                self.pop_scope()?;
                drop(self.pending_empty_depth.take());
            },
            RawEventKind::End => {
                self.pop_scope()?;
                drop(self.depth_reservations.pop());
            },
            RawEventKind::Start | RawEventKind::Other => {},
        }
        Ok(())
    }

    fn pop_scope(&mut self) -> Result<()> {
        let scope = self
            .scopes
            .pop()
            .ok_or_else(|| invalid("DDE namespace scope stack underflow"))?;
        self.active_namespace_bytes = self
            .active_namespace_bytes
            .checked_sub(scope.namespace_bytes)
            .ok_or_else(|| invalid("DDE namespace scope accounting underflow"))?;
        self.active_namespace_bindings = self
            .active_namespace_bindings
            .checked_sub(scope.namespace_bindings)
            .ok_or_else(|| invalid("DDE namespace binding accounting underflow"))?;
        if scope.is_start {
            self.active_element_name_bytes = self
                .active_element_name_bytes
                .checked_sub(scope.element_name_bytes)
                .ok_or_else(|| invalid("DDE element-name accounting underflow"))?;
            self.active_start_elements = self
                .active_start_elements
                .checked_sub(1)
                .ok_or_else(|| invalid("DDE open-element accounting underflow"))?;
        }
        Ok(())
    }

    fn storage_bound(
        &self,
        namespace_bytes: usize,
        namespace_bindings: usize,
        element_name_bytes: usize,
        active_start_elements: usize,
    ) -> Result<usize> {
        let logical = storage_bound(
            namespace_bytes,
            namespace_bindings,
            element_name_bytes,
            active_start_elements,
        )?;
        let vectors = vector_capacity_bound(
            self.scopes.capacity(),
            self.depth_reservations.capacity(),
            NAMESPACE_RESOLVER_BASE_BYTES,
        )?;
        Ok(logical.max(vectors))
    }

    fn ensure_scope_capacity(&mut self, context: &ExecutionContext) -> Result<()> {
        if self.scopes.len() < self.scopes.capacity() {
            return Ok(());
        }
        let current = self.scopes.capacity();
        let desired = if current == 0 {
            1
        } else {
            current
                .checked_mul(2)
                .ok_or_else(|| invalid("DDE parser scope capacity overflows"))?
                .min(self.depth_limit)
        };
        if desired <= current {
            return Err(invalid("DDE parser scope stack exceeds its depth limit"));
        }
        let bound = vector_capacity_bound(
            desired,
            self.depth_reservations.capacity(),
            NAMESPACE_RESOLVER_BASE_BYTES,
        )?;
        self.reserve_bound(bound, context)?;
        self.scopes
            .try_reserve_exact(desired - current)
            .map_err(|_| invalid("DDE parser scope allocation failed"))?;
        Ok(())
    }

    fn ensure_depth_reservation_capacity(&mut self, context: &ExecutionContext) -> Result<()> {
        if self.depth_reservations.len() < self.depth_reservations.capacity() {
            return Ok(());
        }
        let current = self.depth_reservations.capacity();
        let desired = if current == 0 {
            1
        } else {
            current
                .checked_mul(2)
                .ok_or_else(|| invalid("DDE parser depth capacity overflows"))?
                .min(self.depth_limit)
        };
        if desired <= current {
            return Err(invalid("DDE parser depth stack exceeds its depth limit"));
        }
        let bound = vector_capacity_bound(
            self.scopes.capacity(),
            desired,
            NAMESPACE_RESOLVER_BASE_BYTES,
        )?;
        self.reserve_bound(bound, context)?;
        self.depth_reservations
            .try_reserve_exact(desired - current)
            .map_err(|_| invalid("DDE parser depth stack allocation failed"))?;
        Ok(())
    }

    fn reserve_bound(&mut self, desired: usize, context: &ExecutionContext) -> Result<()> {
        if desired <= self.reserved_bound {
            return Ok(());
        }
        let amount = desired
            .checked_sub(self.reserved_bound)
            .ok_or_else(|| invalid("DDE parser scratch reservation underflow"))?;
        let reservation = context
            .reserve(
                Resource::Memory,
                u64::try_from(amount)
                    .map_err(|_| invalid("DDE parser scratch size exceeds u64"))?,
            )
            .map_err(map_execution)?;
        if let Err(_reservation) = self.reservation.try_merge(reservation) {
            return Err(invalid(
                "DDE parser memory reservations have different owners",
            ));
        }
        self.reserved_bound = desired;
        Ok(())
    }
}

/// A single budget reservation that grows as retained parser projections are
/// appended.  Keeping one token on `Snapshot` makes clones and detached
/// patches retain the same bounded source inventory without cloning tokens.
struct MemoryLedger {
    reservation: Reservation,
}

impl MemoryLedger {
    fn new(reservation: Reservation) -> Self {
        Self { reservation }
    }

    fn retain(&mut self, context: &ExecutionContext, amount: usize) -> Result<()> {
        if amount == 0 {
            return Ok(());
        }
        let amount =
            u64::try_from(amount).map_err(|_| invalid("DDE retained memory size exceeds u64"))?;
        let reservation = context
            .reserve(Resource::Memory, amount)
            .map_err(map_execution)?;
        if let Err(_reservation) = self.reservation.try_merge(reservation) {
            return Err(invalid("DDE memory reservations have different owners"));
        }
        Ok(())
    }

    fn into_reservation(self) -> Reservation {
        self.reservation
    }
}

fn admit_content(
    content_xml: &str,
    requested_limits: Limits,
    context: &ExecutionContext,
) -> Result<(Limits, Reservation, MemoryLedger)> {
    let limits = requested_limits.validate()?;
    if content_xml.len() > limits.input_bytes {
        return Err(Error::ResourceLimit {
            resource: "input bytes",
            actual: content_xml.len(),
            maximum: limits.input_bytes,
        });
    }
    context.check().map_err(map_execution)?;
    let input_amount =
        u64::try_from(content_xml.len()).map_err(|_| invalid("DDE content length exceeds u64"))?;
    let input_reservation = context
        .reserve(Resource::InputBytes, input_amount)
        .map_err(map_execution)?;
    let retained_memory = context
        .reserve(Resource::Memory, input_amount)
        .map_err(map_execution)?;
    Ok((
        limits,
        input_reservation,
        MemoryLedger::new(retained_memory),
    ))
}

impl Snapshot {
    /// Parse the default-bounded inert DDE inventory.
    ///
    /// # Errors
    ///
    /// Returns an error when XML is malformed, violates the ODF DDE grammar,
    /// or exceeds a default resource limit.
    pub fn parse(content_xml: &str) -> Result<Self> {
        Self::parse_with(content_xml, Limits::default())
    }

    /// Parse the inert DDE inventory under caller-provided resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error when XML is malformed, violates the ODF DDE grammar,
    /// or exceeds `limits`.
    pub fn parse_with(content_xml: &str, requested_limits: Limits) -> Result<Self> {
        let context = default_context();
        Self::parse_with_context_inner(content_xml, requested_limits, &context, false)
    }

    /// Parse the inert DDE inventory while retaining a caller-owned execution
    /// context for subsequent editing and publication.
    ///
    /// The context is checked before every parser step and its input, memory,
    /// and work dimensions are charged.  The same budget lineage is required
    /// again by [`Edit::commit`], so a numerically equivalent unrelated budget
    /// cannot be substituted for the retained operation context.
    pub fn parse_with_context(
        content_xml: &str,
        requested_limits: Limits,
        context: &ExecutionContext,
    ) -> Result<Self> {
        Self::parse_with_context_inner(content_xml, requested_limits, context, true)
    }

    fn parse_with_context_inner(
        content_xml: &str,
        requested_limits: Limits,
        context: &ExecutionContext,
        enforce_context_lineage: bool,
    ) -> Result<Self> {
        // Admit the source before allocating the retained Arc.  This keeps
        // the ordinary string entry point within the caller's memory policy.
        let (limits, input_reservation, retained_memory) =
            admit_content(content_xml, requested_limits, context)?;
        let content = Arc::<str>::from(content_xml);
        Self::parse_admitted_content(
            content,
            limits,
            context,
            enforce_context_lineage,
            input_reservation,
            retained_memory,
        )
    }

    /// Parse an already retained source string without making a second full
    /// `content.xml` allocation.  The source-backed adapter uses this hook
    /// after its package owner has admitted and retained the content bytes.
    pub(crate) fn parse_shared_with_context(
        content: Arc<str>,
        requested_limits: Limits,
        context: &ExecutionContext,
        enforce_context_lineage: bool,
    ) -> Result<Self> {
        let (limits, input_reservation, retained_memory) =
            admit_content(content.as_ref(), requested_limits, context)?;
        Self::parse_admitted_content(
            content,
            limits,
            context,
            enforce_context_lineage,
            input_reservation,
            retained_memory,
        )
    }

    fn parse_admitted_content(
        content: Arc<str>,
        limits: Limits,
        context: &ExecutionContext,
        enforce_context_lineage: bool,
        input_reservation: Reservation,
        mut retained_memory: MemoryLedger,
    ) -> Result<Self> {
        let content_xml = content.as_ref();

        let mut reader = NsReader::from_str(content_xml);
        reader.config_mut().enable_all_checks(true);
        reader.config_mut().trim_text(false);
        reader
            .resolver_mut()
            .set_max_declarations_per_element(MAX_ATTRIBUTES);
        let mut state = ParserState::new();
        let source_offset = if content_xml.as_bytes().starts_with(UTF8_BOM) {
            UTF8_BOM.len()
        } else {
            0
        };
        let mut parser_scratch = ParserScratch::new(context, limits.depth())?;

        loop {
            context.check().map_err(map_execution)?;
            context.consume(Resource::Work, 1).map_err(map_execution)?;
            let event_start = xml_position(&reader)?
                .checked_add(source_offset)
                .ok_or(Error::PositionOverflow)?;
            let raw_event = raw_event_for_admission(content_xml, event_start);
            parser_scratch.prepare(raw_event, context)?;
            let raw_event_bytes = raw_event
                .end
                .checked_sub(event_start)
                .ok_or_else(|| invalid("DDE raw XML event position moved backwards"))?;
            if raw_event_bytes != 0 {
                context
                    .consume(
                        Resource::Work,
                        u64::try_from(raw_event_bytes)
                            .map_err(|_| invalid("DDE XML event size exceeds u64"))?,
                    )
                    .map_err(map_execution)?;
            }
            let (resolved_namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::InvalidXml(error.to_string()))?;
            let namespace = namespace_kind_checked(&resolved_namespace)?;
            let event_end = xml_position(&reader)?
                .checked_add(source_offset)
                .ok_or(Error::PositionOverflow)?;
            let event_bytes = event_end
                .checked_sub(event_start)
                .ok_or_else(|| invalid("DDE XML reader position moved backwards"))?;
            if event_bytes > raw_event_bytes {
                let work = u64::try_from(event_bytes - raw_event_bytes)
                    .map_err(|_| invalid("DDE XML event size exceeds u64"))?;
                context
                    .consume(Resource::Work, work)
                    .map_err(map_execution)?;
            }
            let eof = matches!(&event, Event::Eof);
            let actual_kind = event_kind(&event);
            state.process(
                event,
                namespace,
                event_start,
                event_end,
                &reader,
                &limits,
                context,
                &mut retained_memory,
                &content,
            )?;
            parser_scratch.finish(raw_event.kind, actual_kind)?;
            if eof {
                break;
            }
        }
        if state.depth != 0
            || state.root.is_none()
            || !state.root_closed
            || !state.body_seen
            || !state.spreadsheet_seen
            || state.body_depth.is_some()
            || state.link.is_some()
            || state.cached_depth.is_some()
            || state.source_depth.is_some()
            || !parser_scratch.scopes.is_empty()
            || !parser_scratch.depth_reservations.is_empty()
            || parser_scratch.pending_empty_depth.is_some()
        {
            return Err(invalid(
                "DDE XML lacks a complete document-content envelope",
            ));
        }
        Ok(Self {
            content,
            inventory: Arc::new(Inventory {
                links: state.links,
                sheet_sources: state.sheet_sources,
                sheet_table_indices: state.sheet_table_indices,
                table_names: state.table_names,
            }),
            limits,
            context: context.clone(),
            enforce_context_lineage,
            _input_reservation: Arc::new(input_reservation),
            _memory_reservation: Arc::new(retained_memory.into_reservation()),
            _output_reservation: None,
        })
    }

    #[must_use]
    pub fn source_xml(&self) -> &str {
        &self.content
    }

    #[must_use]
    pub fn links(&self) -> &[Link] {
        &self.inventory.links
    }

    #[must_use]
    pub fn sheet_sources(&self) -> &[SheetSource] {
        &self.inventory.sheet_sources
    }

    /// Returns spreadsheet table names in source order.  `None` represents an
    /// ordinary unnamed table; it can only be selected by position for a
    /// source declaration edit.
    #[must_use]
    pub fn table_names(&self) -> &[Option<String>] {
        &self.inventory.table_names
    }

    /// Begin an immutable, source-checked DDE edit.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit::new(self.clone())
    }

    pub(crate) fn limits(&self) -> Limits {
        self.limits
    }

    pub(crate) fn context(&self) -> &ExecutionContext {
        &self.context
    }

    pub(crate) fn enforce_context_lineage(&self) -> bool {
        self.enforce_context_lineage
    }

    fn inventory(&self) -> &Inventory {
        &self.inventory
    }
}

struct ParserState {
    depth: usize,
    root: Option<DocumentRoot>,
    root_closed: bool,
    declaration_seen: bool,
    prolog_markup_seen: bool,
    body_seen: bool,
    body_depth: Option<usize>,
    spreadsheet_seen: bool,
    spreadsheet_depth: Option<usize>,
    links_depth: Option<usize>,
    links_seen: bool,
    link: Option<LinkBuilder>,
    source_depth: Option<usize>,
    cached_depth: Option<usize>,
    sheet: Option<SheetBuilder>,
    links: Vec<Link>,
    sheet_sources: Vec<SheetSource>,
    sheet_table_indices: Vec<usize>,
    table_names: Vec<Option<String>>,
}

impl ParserState {
    fn new() -> Self {
        Self {
            depth: 0,
            root: None,
            root_closed: false,
            declaration_seen: false,
            prolog_markup_seen: false,
            body_seen: false,
            body_depth: None,
            spreadsheet_seen: false,
            spreadsheet_depth: None,
            links_depth: None,
            links_seen: false,
            link: None,
            source_depth: None,
            cached_depth: None,
            sheet: None,
            links: Vec::new(),
            sheet_sources: Vec::new(),
            sheet_table_indices: Vec::new(),
            table_names: Vec::new(),
        }
    }

    fn process(
        &mut self,
        event: Event<'_>,
        namespace: NamespaceKind,
        event_start: usize,
        event_end: usize,
        reader: &NsReader<&[u8]>,
        limits: &Limits,
        context: &ExecutionContext,
        retained_memory: &mut MemoryLedger,
        content: &Arc<str>,
    ) -> Result<()> {
        match event {
            Event::Start(element) => {
                validate_attributes(&element, reader, limits.text_bytes, context)?;
                self.depth = self
                    .depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML depth overflow"))?;
                if self.depth > limits.depth {
                    return Err(invalid("DDE XML exceeds the nesting limit"));
                }
                let local = element.local_name();
                let local = local.as_ref();
                let is_spreadsheet = is(namespace, local, NamespaceKind::Office, b"spreadsheet");
                if self.depth == 1 {
                    if self.root.is_some() || self.root_closed {
                        return Err(invalid("DDE XML has multiple document roots"));
                    }
                    let root = if is(namespace, local, NamespaceKind::Office, b"document-content") {
                        DocumentRoot::Content
                    } else if is(namespace, local, NamespaceKind::Office, b"document") {
                        let mimetype = required_attr_with_context(
                            &element,
                            reader,
                            OFFICE,
                            b"mimetype",
                            limits.text_bytes,
                            context,
                        )?;
                        if mimetype.as_bytes() != ODF_SPREADSHEET_MIMETYPE {
                            return Err(invalid("flat ODS document has the wrong office:mimetype"));
                        }
                        DocumentRoot::Flat
                    } else {
                        return Err(invalid(
                            "DDE XML root must be office:document-content or flat office:document",
                        ));
                    };
                    self.root = Some(root);
                } else {
                    if self.root.is_none() || self.root_closed {
                        return Err(invalid("DDE XML contains content outside its root"));
                    }
                    if is(namespace, local, NamespaceKind::Office, b"document-content")
                        || is(namespace, local, NamespaceKind::Office, b"document")
                    {
                        return Err(invalid("DDE XML has a nested office document root"));
                    }
                    if is(namespace, local, NamespaceKind::Office, b"body") {
                        if self.depth != 2 || self.body_seen {
                            return Err(invalid(
                                "office:body must be the unique direct root child",
                            ));
                        }
                        self.body_seen = true;
                        self.body_depth = Some(self.depth);
                    } else if is_spreadsheet {
                        if self.depth != 3 || self.body_depth != Some(2) || self.spreadsheet_seen {
                            return Err(invalid(
                                "office:spreadsheet must be the unique direct office:body child",
                            ));
                        }
                        self.spreadsheet_seen = true;
                        self.spreadsheet_depth = Some(self.depth);
                    } else if self.body_depth == Some(2)
                        && self.depth == 3
                        && namespace == NamespaceKind::Office
                    {
                        return Err(invalid(
                            "office:body must contain a direct office:spreadsheet",
                        ));
                    }
                }

                if is_spreadsheet {
                    return Ok(());
                }
                if self
                    .spreadsheet_depth
                    .is_some_and(|value| self.depth == value + 1)
                    && is(namespace, local, NamespaceKind::Table, b"dde-links")
                {
                    if self.links_seen || self.links_depth.replace(self.depth).is_some() {
                        return Err(invalid("duplicate table:dde-links owner"));
                    }
                    self.links_seen = true;
                } else if self
                    .links_depth
                    .is_some_and(|value| self.depth == value + 1)
                    && is(namespace, local, NamespaceKind::Table, b"dde-link")
                {
                    if self.link.replace(LinkBuilder::new(self.depth)).is_some() {
                        return Err(invalid("nested table:dde-link"));
                    }
                } else if self
                    .link
                    .as_ref()
                    .is_some_and(|value| self.depth == value.depth + 1)
                    && is(namespace, local, NamespaceKind::Office, b"dde-source")
                {
                    let value =
                        parse_source_with_context(&element, reader, limits.text_bytes, context)?;
                    let source_memory = source_memory_bytes(&value)?;
                    retained_memory.retain(context, source_memory)?;
                    let Some(builder) = self.link.as_mut() else {
                        return Err(invalid("DDE link parser state is missing"));
                    };
                    if builder.source.replace(value).is_some() || builder.cached.is_some() {
                        return Err(invalid("office:dde-source must be the first link child"));
                    }
                    self.source_depth = Some(self.depth);
                } else if self
                    .link
                    .as_ref()
                    .is_some_and(|value| self.depth == value.depth + 1)
                    && is(namespace, local, NamespaceKind::Table, b"table")
                {
                    let Some(builder) = self.link.as_mut() else {
                        return Err(invalid("DDE link parser state is missing"));
                    };
                    if builder.source.is_none() || builder.cached.is_some() {
                        return Err(invalid("cached table must follow exactly one DDE source"));
                    }
                    builder.cached = Some(event_start..event_start);
                    self.cached_depth = Some(self.depth);
                } else if self.source_depth.is_some() {
                    return Err(invalid("office:dde-source must not contain child elements"));
                } else if self.link.is_some()
                    && self.cached_depth.is_none()
                    && self.source_depth.is_none()
                {
                    return Err(invalid("unsupported child in table:dde-link"));
                } else if self
                    .spreadsheet_depth
                    .is_some_and(|value| self.depth == value + 1)
                    && is(namespace, local, NamespaceKind::Table, b"table")
                {
                    let table_index = push_table_name(
                        &mut self.table_names,
                        optional_attr_with_context(
                            &element,
                            reader,
                            TABLE,
                            b"name",
                            limits.text_bytes,
                            context,
                        )?,
                        retained_memory,
                        context,
                    )?;
                    self.sheet = Some(SheetBuilder {
                        depth: self.depth,
                        table_index,
                        source_seen: false,
                    });
                } else if self
                    .sheet
                    .as_ref()
                    .is_some_and(|value| self.depth == value.depth + 1)
                    && is(namespace, local, NamespaceKind::Office, b"dde-source")
                {
                    let value =
                        parse_source_with_context(&element, reader, limits.text_bytes, context)?;
                    let table_index = self
                        .sheet
                        .as_ref()
                        .ok_or_else(|| invalid("DDE sheet parser state is missing"))?
                        .table_index;
                    let source_seen = self
                        .sheet
                        .as_ref()
                        .ok_or_else(|| invalid("DDE sheet parser state is missing"))?
                        .source_seen;
                    if source_seen {
                        return Err(invalid("duplicate sheet office:dde-source"));
                    }
                    push_sheet_source(
                        &mut self.sheet_sources,
                        &mut self.sheet_table_indices,
                        table_index,
                        &self.table_names,
                        value,
                        limits.sheet_sources,
                        retained_memory,
                        context,
                    )?;
                    if let Some(current) = self.sheet.as_mut() {
                        current.source_seen = true;
                    }
                    self.source_depth = Some(self.depth);
                }
            },
            Event::Empty(element) => {
                validate_attributes(&element, reader, limits.text_bytes, context)?;
                let event_depth = self
                    .depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML depth overflow"))?;
                let local = element.local_name();
                let local = local.as_ref();
                let is_spreadsheet = is(namespace, local, NamespaceKind::Office, b"spreadsheet");
                if event_depth == 1 {
                    return Err(invalid(
                        "DDE document root must contain office:body and office:spreadsheet",
                    ));
                }
                if self.root.is_none() || self.root_closed {
                    return Err(invalid("DDE XML contains content outside its root"));
                }
                if is(namespace, local, NamespaceKind::Office, b"document-content")
                    || is(namespace, local, NamespaceKind::Office, b"document")
                {
                    return Err(invalid("DDE XML has a nested office document root"));
                }
                if is(namespace, local, NamespaceKind::Office, b"body") {
                    return Err(invalid(
                        "office:body must contain a direct office:spreadsheet",
                    ));
                }
                if is_spreadsheet {
                    if event_depth != 3 || self.body_depth != Some(2) || self.spreadsheet_seen {
                        return Err(invalid(
                            "office:spreadsheet must be the unique direct office:body child",
                        ));
                    }
                    self.spreadsheet_seen = true;
                } else if self.body_depth == Some(2)
                    && event_depth == 3
                    && namespace == NamespaceKind::Office
                {
                    return Err(invalid(
                        "office:body must contain a direct office:spreadsheet",
                    ));
                }
                if is(namespace, local, NamespaceKind::Table, b"dde-links")
                    && self
                        .spreadsheet_depth
                        .is_some_and(|value| event_depth == value + 1)
                {
                    return Err(invalid("table:dde-links must contain a link"));
                } else if is(namespace, local, NamespaceKind::Table, b"dde-link")
                    && self
                        .links_depth
                        .is_some_and(|value| event_depth == value + 1)
                {
                    return Err(invalid("table:dde-link requires a source and cached table"));
                } else if self.source_depth.is_some() {
                    return Err(invalid("office:dde-source must not contain child elements"));
                } else if self
                    .link
                    .as_ref()
                    .is_some_and(|value| event_depth == value.depth + 1)
                    && is(namespace, local, NamespaceKind::Office, b"dde-source")
                {
                    let value =
                        parse_source_with_context(&element, reader, limits.text_bytes, context)?;
                    let source_memory = source_memory_bytes(&value)?;
                    retained_memory.retain(context, source_memory)?;
                    let Some(builder) = self.link.as_mut() else {
                        return Err(invalid("DDE link parser state is missing"));
                    };
                    if builder.source.replace(value).is_some() || builder.cached.is_some() {
                        return Err(invalid("office:dde-source must be the first link child"));
                    }
                } else if self
                    .link
                    .as_ref()
                    .is_some_and(|value| event_depth == value.depth + 1)
                    && is(namespace, local, NamespaceKind::Table, b"table")
                {
                    let Some(builder) = self.link.as_mut() else {
                        return Err(invalid("DDE link parser state is missing"));
                    };
                    if builder.source.is_none() || builder.cached.is_some() {
                        return Err(invalid("cached table must follow exactly one DDE source"));
                    }
                    if event_end
                        .checked_sub(event_start)
                        .is_none_or(|value| value > limits.cached_table_bytes)
                    {
                        return Err(invalid("DDE cached table exceeds its byte limit"));
                    }
                    builder.cached = Some(event_start..event_end);
                } else if self
                    .sheet
                    .as_ref()
                    .is_some_and(|value| event_depth == value.depth + 1)
                    && is(namespace, local, NamespaceKind::Office, b"dde-source")
                {
                    let value =
                        parse_source_with_context(&element, reader, limits.text_bytes, context)?;
                    let table_index = self
                        .sheet
                        .as_ref()
                        .ok_or_else(|| invalid("DDE sheet parser state is missing"))?
                        .table_index;
                    let source_seen = self
                        .sheet
                        .as_ref()
                        .ok_or_else(|| invalid("DDE sheet parser state is missing"))?
                        .source_seen;
                    if source_seen {
                        return Err(invalid("duplicate sheet office:dde-source"));
                    }
                    push_sheet_source(
                        &mut self.sheet_sources,
                        &mut self.sheet_table_indices,
                        table_index,
                        &self.table_names,
                        value,
                        limits.sheet_sources,
                        retained_memory,
                        context,
                    )?;
                    if let Some(current) = self.sheet.as_mut() {
                        current.source_seen = true;
                    }
                } else if self
                    .spreadsheet_depth
                    .is_some_and(|value| event_depth == value + 1)
                    && is(namespace, local, NamespaceKind::Table, b"table")
                {
                    push_table_name(
                        &mut self.table_names,
                        optional_attr_with_context(
                            &element,
                            reader,
                            TABLE,
                            b"name",
                            limits.text_bytes,
                            context,
                        )?,
                        retained_memory,
                        context,
                    )?;
                }
            },
            Event::End(element) => {
                if self.depth == 0 {
                    return Err(invalid("DDE XML element stack underflow"));
                }
                let local = element.local_name();
                let local = local.as_ref();
                if self.cached_depth == Some(self.depth)
                    && is(namespace, local, NamespaceKind::Table, b"table")
                {
                    let Some(builder) = self.link.as_mut() else {
                        return Err(invalid("DDE cached-table parser state is missing"));
                    };
                    let Some(cached) = builder.cached.as_ref() else {
                        return Err(invalid("DDE cached-table start is missing"));
                    };
                    let cached_start = cached.start;
                    let cached_size = event_end
                        .checked_sub(cached_start)
                        .ok_or_else(|| invalid("DDE cached-table position is invalid"))?;
                    if cached_size > limits.cached_table_bytes {
                        return Err(invalid("DDE cached table exceeds its byte limit"));
                    }
                    builder.cached = Some(cached_start..event_end);
                    self.cached_depth = None;
                } else if self.source_depth == Some(self.depth)
                    && is(namespace, local, NamespaceKind::Office, b"dde-source")
                {
                    self.source_depth = None;
                } else if self
                    .link
                    .as_ref()
                    .is_some_and(|value| self.depth == value.depth)
                    && is(namespace, local, NamespaceKind::Table, b"dde-link")
                {
                    let Some(builder) = self.link.take() else {
                        return Err(invalid("DDE link parser state is missing"));
                    };
                    if self.links.len() >= limits.links {
                        return Err(invalid("DDE link count exceeds its limit"));
                    }
                    context
                        .consume(Resource::Objects, 1)
                        .map_err(map_execution)?;
                    retained_memory.retain(context, size_of::<Link>())?;
                    self.links
                        .try_reserve(1)
                        .map_err(|_| invalid("DDE link catalog allocation failed"))?;
                    self.links.push(builder.finish(Arc::clone(content))?);
                } else if self.links_depth == Some(self.depth)
                    && is(namespace, local, NamespaceKind::Table, b"dde-links")
                {
                    if self.links.is_empty() {
                        return Err(invalid("table:dde-links must contain a link"));
                    }
                    self.links_depth = None;
                } else if self
                    .sheet
                    .as_ref()
                    .is_some_and(|value| self.depth == value.depth)
                    && is(namespace, local, NamespaceKind::Table, b"table")
                {
                    self.sheet = None;
                } else if self.spreadsheet_depth == Some(self.depth)
                    && is_spreadsheet(namespace, local)
                {
                    self.spreadsheet_depth = None;
                }
                if self.depth == 2 && is(namespace, local, NamespaceKind::Office, b"body") {
                    if !self.body_seen {
                        return Err(invalid("office:body is missing"));
                    }
                    self.body_depth = None;
                }
                if self.depth == 1 {
                    if !self.root_matches(namespace, local) {
                        return Err(invalid("DDE XML root end tag is invalid"));
                    }
                    self.root_closed = true;
                }
                self.depth = self
                    .depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("DDE XML element stack underflow"))?;
            },
            Event::Text(text) if self.source_depth.is_some() => {
                if decode_text_is_non_whitespace(&text, context)? {
                    return Err(invalid("office:dde-source must be empty"));
                }
            },
            Event::CData(text) if self.source_depth.is_some() => {
                if decode_cdata_is_non_whitespace(&text, context)? {
                    return Err(invalid("office:dde-source must be empty"));
                }
            },
            Event::Text(text)
                if self.cached_depth.is_none()
                    && (self.link.is_some() || self.links_depth.is_some()) =>
            {
                if decode_text_is_non_whitespace(&text, context)? {
                    return Err(invalid("DDE containers must not contain text"));
                }
            },
            Event::CData(text)
                if self.cached_depth.is_none()
                    && (self.link.is_some() || self.links_depth.is_some()) =>
            {
                if decode_cdata_is_non_whitespace(&text, context)? {
                    return Err(invalid("DDE containers must not contain CDATA"));
                }
            },
            Event::GeneralRef(reference) if self.link.is_some() || self.links_depth.is_some() => {
                if self.cached_depth.is_none() || !valid_general_ref(reference.as_ref()) {
                    return Err(invalid(
                        "DDE containers contain an unsupported entity reference",
                    ));
                }
            },
            Event::GeneralRef(reference) => {
                if self.root.is_none() || self.root_closed || !valid_general_ref(reference.as_ref())
                {
                    return Err(invalid("DDE XML contains an unsupported entity reference"));
                }
            },
            Event::DocType(_) => return Err(invalid("DTD content is not accepted")),
            Event::Eof => {},
            Event::Text(text) => {
                let bytes: &[u8] = text.as_ref();
                if (self.root.is_none() || self.root_closed)
                    && bytes.iter().any(|byte| !byte.is_ascii_whitespace())
                {
                    return Err(invalid("DDE XML has text outside its root"));
                }
            },
            Event::CData(_) if self.root.is_none() || self.root_closed => {
                return Err(invalid("DDE XML has CDATA outside its root"));
            },
            Event::CData(_) => {},
            Event::Decl(_) => {
                if self.root.is_some()
                    || self.root_closed
                    || self.declaration_seen
                    || self.prolog_markup_seen
                {
                    return Err(invalid(
                        "DDE XML declaration must be the first document markup",
                    ));
                }
                self.declaration_seen = true;
            },
            Event::Comment(_) | Event::PI(_) if self.root.is_none() => {
                if !self.declaration_seen {
                    self.prolog_markup_seen = true;
                }
            },
            Event::Comment(_) | Event::PI(_) => {},
        }
        Ok(())
    }

    fn root_matches(&self, namespace: NamespaceKind, local: &[u8]) -> bool {
        match self.root {
            Some(DocumentRoot::Content) => {
                is(namespace, local, NamespaceKind::Office, b"document-content")
            },
            Some(DocumentRoot::Flat) => is(namespace, local, NamespaceKind::Office, b"document"),
            None => false,
        }
    }
}

struct LinkBuilder {
    depth: usize,
    source: Option<Source>,
    cached: Option<Range<usize>>,
}

impl LinkBuilder {
    fn new(depth: usize) -> Self {
        Self {
            depth,
            source: None,
            cached: None,
        }
    }

    fn finish(self, content: Arc<str>) -> Result<Link> {
        Ok(Link {
            source: self
                .source
                .ok_or_else(|| invalid("DDE link has no source"))?,
            content,
            cached_table: self
                .cached
                .ok_or_else(|| invalid("DDE link has no cached table"))?,
        })
    }
}

struct SheetBuilder {
    depth: usize,
    table_index: usize,
    source_seen: bool,
}

fn storage_bound(
    namespace_bytes: usize,
    namespace_bindings: usize,
    element_name_bytes: usize,
    active_start_elements: usize,
) -> Result<usize> {
    let namespace_bytes = namespace_bytes
        .checked_mul(PARSER_CAPACITY_FACTOR)
        .ok_or_else(|| invalid("DDE namespace scratch size overflows"))?;
    let namespace_binding_bytes = namespace_bindings
        .checked_mul(NAMESPACE_BINDING_BYTES)
        .and_then(|value| value.checked_mul(PARSER_CAPACITY_FACTOR))
        .ok_or_else(|| invalid("DDE namespace binding scratch size overflows"))?;
    let element_name_bytes = element_name_bytes
        .checked_mul(PARSER_CAPACITY_FACTOR)
        .ok_or_else(|| invalid("DDE element-name scratch size overflows"))?;
    let open_element_bytes = active_start_elements
        .checked_mul(size_of::<usize>())
        .and_then(|value| value.checked_mul(PARSER_CAPACITY_FACTOR))
        .ok_or_else(|| invalid("DDE open-element scratch size overflows"))?;
    NAMESPACE_RESOLVER_BASE_BYTES
        .checked_add(namespace_bytes)
        .and_then(|value| value.checked_add(namespace_binding_bytes))
        .and_then(|value| value.checked_add(element_name_bytes))
        .and_then(|value| value.checked_add(open_element_bytes))
        .ok_or_else(|| invalid("DDE parser scratch size overflows"))
}

fn vector_capacity_bound(
    scope_capacity: usize,
    depth_capacity: usize,
    base: usize,
) -> Result<usize> {
    let scope_bytes = scope_capacity
        .checked_mul(size_of::<NamespaceScope>())
        .and_then(|value| value.checked_mul(PARSER_CAPACITY_FACTOR))
        .ok_or_else(|| invalid("DDE parser scope memory size overflows"))?;
    let depth_bytes = depth_capacity
        .checked_mul(size_of::<Reservation>())
        .and_then(|value| value.checked_mul(PARSER_CAPACITY_FACTOR))
        .ok_or_else(|| invalid("DDE parser depth stack memory size overflows"))?;
    base.checked_add(scope_bytes)
        .and_then(|value| value.checked_add(depth_bytes))
        .ok_or_else(|| invalid("DDE parser scratch memory size overflows"))
}

fn raw_event_for_admission(content: &str, start: usize) -> RawEvent {
    let bytes = content.as_bytes();
    let mut cursor = start;
    if cursor == 0 && bytes.starts_with(UTF8_BOM) {
        cursor = cursor.saturating_add(UTF8_BOM.len());
    }
    let Some(&first) = bytes.get(cursor) else {
        return RawEvent {
            kind: RawEventKind::Other,
            end: cursor.min(bytes.len()),
            namespace_bytes: 0,
            namespace_bindings: 0,
            element_name_bytes: 0,
        };
    };
    if first != b'<' {
        return RawEvent {
            kind: RawEventKind::Other,
            end: raw_text_end(bytes, cursor),
            namespace_bytes: 0,
            namespace_bindings: 0,
            element_name_bytes: 0,
        };
    }
    let Some(&next) = bytes.get(cursor + 1) else {
        return RawEvent {
            kind: RawEventKind::Other,
            end: bytes.len(),
            namespace_bytes: 0,
            namespace_bindings: 0,
            element_name_bytes: 0,
        };
    };
    if next == b'/' {
        let end = find_unquoted_gt(bytes, cursor);
        return RawEvent {
            kind: RawEventKind::End,
            end: end.map_or(bytes.len(), |value| value.saturating_add(1)),
            namespace_bytes: 0,
            namespace_bindings: 0,
            element_name_bytes: end
                .and_then(|value| tag_name_length(&bytes[cursor..=value]))
                .unwrap_or(bytes.len().saturating_sub(cursor)),
        };
    }
    if next == b'!' || next == b'?' {
        return RawEvent {
            kind: RawEventKind::Other,
            end: find_unquoted_gt(bytes, cursor)
                .map_or(bytes.len(), |value| value.saturating_add(1)),
            namespace_bytes: 0,
            namespace_bindings: 0,
            element_name_bytes: 0,
        };
    }
    let event_end =
        find_unquoted_gt(bytes, cursor).map_or(bytes.len(), |value| value.saturating_add(1));
    let tag = &bytes[cursor..event_end];
    let kind = if tag.get(..tag.len().saturating_sub(1)).is_some_and(|value| {
        value.iter().rev().find(|byte| !byte.is_ascii_whitespace()) == Some(&b'/')
    }) {
        RawEventKind::Empty
    } else {
        RawEventKind::Start
    };
    let (namespace_bytes, namespace_bindings, element_name_bytes) = namespace_declarations(tag)
        .unwrap_or_else(|| {
            (
                tag.len(),
                MAX_ATTRIBUTES,
                tag_name_length(tag).unwrap_or(tag.len()),
            )
        });
    RawEvent {
        kind,
        end: event_end,
        namespace_bytes,
        namespace_bindings,
        element_name_bytes,
    }
}

fn raw_text_end(bytes: &[u8], start: usize) -> usize {
    bytes
        .get(start..)
        .and_then(|value| value.iter().position(|byte| matches!(*byte, b'<' | b'&')))
        .and_then(|offset| start.checked_add(offset))
        .unwrap_or(bytes.len())
}

fn find_unquoted_gt(bytes: &[u8], start: usize) -> Option<usize> {
    let mut quote = None;
    for (offset, byte) in bytes.get(start..)?.iter().copied().enumerate() {
        match quote {
            Some(value) if byte == value => quote = None,
            Some(_) => {},
            None if byte == b'\'' || byte == b'"' => quote = Some(byte),
            None if byte == b'>' => return start.checked_add(offset),
            None => {},
        }
    }
    None
}

fn namespace_declarations(tag: &[u8]) -> Option<(usize, usize, usize)> {
    if tag.first() != Some(&b'<') || tag.last() != Some(&b'>') {
        return None;
    }
    let mut cursor = 1usize;
    skip_ascii_whitespace(tag, &mut cursor);
    let name_start = cursor;
    while let Some(&byte) = tag.get(cursor) {
        if byte.is_ascii_whitespace() || byte == b'/' || byte == b'>' {
            break;
        }
        cursor += 1;
    }
    if cursor == name_start {
        return None;
    }
    let mut element_name_bytes = cursor - name_start;
    let mut namespace_bytes = 0usize;
    let mut namespace_bindings = 0usize;
    loop {
        skip_ascii_whitespace(tag, &mut cursor);
        if matches!(tag.get(cursor), Some(b'>') | Some(b'/') | None) {
            return Some((namespace_bytes, namespace_bindings, element_name_bytes));
        }
        let attribute_start = cursor;
        while let Some(&byte) = tag.get(cursor) {
            if byte.is_ascii_whitespace() || matches!(byte, b'=' | b'/' | b'>') {
                break;
            }
            cursor += 1;
        }
        if cursor == attribute_start {
            return None;
        }
        let attribute_name = &tag[attribute_start..cursor];
        element_name_bytes = element_name_bytes.checked_add(attribute_name.len())?;
        skip_ascii_whitespace(tag, &mut cursor);
        if tag.get(cursor) != Some(&b'=') {
            return None;
        }
        cursor += 1;
        skip_ascii_whitespace(tag, &mut cursor);
        let quote = *tag.get(cursor)?;
        if quote != b'\'' && quote != b'"' {
            return None;
        }
        cursor += 1;
        let value_start = cursor;
        while let Some(&byte) = tag.get(cursor) {
            if byte == quote {
                break;
            }
            cursor += 1;
        }
        if tag.get(cursor) != Some(&quote) {
            return None;
        }
        let value_len = cursor - value_start;
        if attribute_name == b"xmlns" {
            namespace_bytes = namespace_bytes.checked_add(value_len)?;
            namespace_bindings = namespace_bindings.checked_add(1)?;
        } else if let Some(prefix) = attribute_name.strip_prefix(b"xmlns:") {
            namespace_bytes = namespace_bytes
                .checked_add(prefix.len())?
                .checked_add(value_len)?;
            namespace_bindings = namespace_bindings.checked_add(1)?;
        }
        cursor += 1;
    }
}

fn tag_name_length(tag: &[u8]) -> Option<usize> {
    if tag.first() != Some(&b'<') {
        return None;
    }
    let mut cursor = 1usize;
    if tag.get(cursor) == Some(&b'/') {
        cursor += 1;
    }
    while let Some(&byte) = tag.get(cursor) {
        if byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>') {
            break;
        }
        cursor += 1;
    }
    (cursor > 1).then_some(cursor - 1)
}

fn skip_ascii_whitespace(bytes: &[u8], cursor: &mut usize) {
    while bytes.get(*cursor).is_some_and(u8::is_ascii_whitespace) {
        *cursor += 1;
    }
}

fn event_kind(event: &Event<'_>) -> RawEventKind {
    match event {
        Event::Start(_) => RawEventKind::Start,
        Event::Empty(_) => RawEventKind::Empty,
        Event::End(_) => RawEventKind::End,
        _ => RawEventKind::Other,
    }
}

fn parse_source_with_context(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    text_bytes: usize,
    context: &ExecutionContext,
) -> Result<Source> {
    let application = required_attr_with_context(
        element,
        reader,
        OFFICE,
        b"dde-application",
        text_bytes,
        context,
    )?;
    let topic =
        required_attr_with_context(element, reader, OFFICE, b"dde-topic", text_bytes, context)?;
    let item =
        required_attr_with_context(element, reader, OFFICE, b"dde-item", text_bytes, context)?;
    let source_name =
        optional_attr_with_context(element, reader, OFFICE, b"name", text_bytes, context)?;
    let conversion_mode = optional_attr_with_context(
        element,
        reader,
        OFFICE,
        b"conversion-mode",
        text_bytes,
        context,
    )?
    .as_deref()
    .map(ConversionMode::parse)
    .transpose()?
    .unwrap_or(ConversionMode::Unspecified);
    let automatic_update = optional_attr_with_context(
        element,
        reader,
        OFFICE,
        b"automatic-update",
        text_bytes,
        context,
    )?
    .as_deref()
    .map(parse_bool)
    .transpose()?
    .into();
    let mut source = Source::new(application, topic, item)?;
    if let Some(name) = source_name {
        source = source.named(name)?;
    }
    Ok(source
        .with_conversion_mode(conversion_mode)
        .with_automatic_update(automatic_update))
}

fn required_attr_with_context(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    namespace: &[u8],
    local: &[u8],
    limit: usize,
    context: &ExecutionContext,
) -> Result<String> {
    optional_attr_with_context(element, reader, namespace, local, limit, context)?.ok_or_else(
        || {
            invalid(format!(
                "missing required attribute {}",
                String::from_utf8_lossy(local)
            ))
        },
    )
}

fn optional_attr_with_context(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    namespace: &[u8],
    local: &[u8],
    limit: usize,
    context: &ExecutionContext,
) -> Result<Option<String>> {
    let mut value = None;
    for raw_attribute in element.attributes().with_checks(true) {
        let attribute =
            raw_attribute.map_err(|error| invalid(format!("invalid DDE attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == namespace)
            && name.as_ref() == local
        {
            if value.is_some() {
                return Err(invalid("duplicate DDE attribute"));
            }
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| invalid(format!("invalid DDE attribute value: {error}")))?
                .into_owned();
            let decoded_memory = context
                .reserve(
                    Resource::Memory,
                    u64::try_from(decoded.len())
                        .map_err(|_| invalid("DDE attribute size exceeds u64"))?,
                )
                .map_err(map_execution)?;
            if decoded.len() > limit || !xml_text_is_valid(&decoded) {
                return Err(invalid("invalid or oversized DDE attribute value"));
            }
            drop(decoded_memory);
            value = Some(decoded);
        }
    }
    Ok(value)
}

fn validate_attributes(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    limit: usize,
    context: &ExecutionContext,
) -> Result<()> {
    let mut count = 0usize;
    for raw_attribute in element.attributes().with_checks(true) {
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("DDE attribute count overflow"))?;
        if count > MAX_ATTRIBUTES {
            return Err(invalid("DDE element exceeds the attribute count limit"));
        }
        let attribute =
            raw_attribute.map_err(|error| invalid(format!("invalid DDE attribute: {error}")))?;
        let (resolved, _name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(resolved, ResolveResult::Unknown(_)) {
            return Err(invalid("DDE XML contains an undeclared attribute prefix"));
        }
        let decoded = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid DDE attribute value: {error}")))?;
        let decoded_len = decoded.len();
        let _decoded_memory = context
            .reserve(
                Resource::Memory,
                u64::try_from(decoded_len)
                    .map_err(|_| invalid("DDE attribute size exceeds u64"))?,
            )
            .map_err(map_execution)?;
        if decoded_len > limit || !xml_text_is_valid(&decoded) {
            return Err(invalid("invalid or oversized DDE attribute value"));
        }
    }
    Ok(())
}

fn push_table_name(
    table_names: &mut Vec<Option<String>>,
    name: Option<String>,
    retained_memory: &mut MemoryLedger,
    context: &ExecutionContext,
) -> Result<usize> {
    let index = table_names.len();
    context
        .consume(Resource::Objects, 1)
        .map_err(map_execution)?;
    let amount = size_of::<Option<String>>()
        .checked_add(name.as_ref().map_or(0, String::len))
        .ok_or_else(|| invalid("DDE table-name memory size overflows"))?;
    retained_memory.retain(context, amount)?;
    table_names
        .try_reserve(1)
        .map_err(|_| invalid("DDE table catalog allocation failed"))?;
    table_names.push(name);
    Ok(index)
}

fn push_sheet_source(
    sheet_sources: &mut Vec<SheetSource>,
    sheet_table_indices: &mut Vec<usize>,
    table_index: usize,
    table_names: &[Option<String>],
    source: Source,
    limit: usize,
    retained_memory: &mut MemoryLedger,
    context: &ExecutionContext,
) -> Result<()> {
    if sheet_sources.len() >= limit {
        return Err(invalid("sheet DDE source count exceeds its limit"));
    }
    let sheet_name = table_names
        .get(table_index)
        .and_then(Option::as_ref)
        .cloned()
        .ok_or_else(|| invalid("a sheet-local office:dde-source requires table:name"))?;
    let source_memory = source_memory_bytes(&source)?;
    let amount = size_of::<SheetSource>()
        .checked_add(sheet_name.len())
        .and_then(|value| value.checked_add(source_memory))
        .and_then(|value| value.checked_add(size_of::<usize>()))
        .ok_or_else(|| invalid("DDE sheet-source memory size overflows"))?;
    context
        .consume(Resource::Objects, 1)
        .map_err(map_execution)?;
    retained_memory.retain(context, amount)?;
    sheet_sources
        .try_reserve(1)
        .map_err(|_| invalid("DDE sheet-source allocation failed"))?;
    sheet_table_indices
        .try_reserve(1)
        .map_err(|_| invalid("DDE sheet-source index allocation failed"))?;
    sheet_sources.push(SheetSource {
        sheet: sheet_name,
        source,
    });
    sheet_table_indices.push(table_index);
    Ok(())
}

fn source_memory_bytes(source: &Source) -> Result<usize> {
    let mut amount = size_of::<Source>();
    for value in [source.application(), source.topic(), source.item()] {
        amount = amount
            .checked_add(value.len())
            .ok_or_else(|| invalid("DDE source memory size overflows"))?;
    }
    if let Some(name) = source.name() {
        amount = amount
            .checked_add(name.len())
            .ok_or_else(|| invalid("DDE source name memory size overflows"))?;
    }
    Ok(amount)
}

fn decode_text_is_non_whitespace(
    text: &quick_xml::events::BytesText<'_>,
    context: &ExecutionContext,
) -> Result<bool> {
    let value = text
        .xml_content(XmlVersion::Explicit1_0)
        .map_err(|error| invalid(format!("invalid DDE XML text: {error}")))?;
    let _decoded_memory = context
        .reserve(
            Resource::Memory,
            u64::try_from(value.len()).map_err(|_| invalid("DDE text size exceeds u64"))?,
        )
        .map_err(map_execution)?;
    Ok(!value.trim().is_empty())
}

fn decode_cdata_is_non_whitespace(
    text: &quick_xml::events::BytesCData<'_>,
    context: &ExecutionContext,
) -> Result<bool> {
    let value = text
        .xml_content(XmlVersion::Explicit1_0)
        .map_err(|error| invalid(format!("invalid DDE XML CDATA: {error}")))?;
    let _decoded_memory = context
        .reserve(
            Resource::Memory,
            u64::try_from(value.len()).map_err(|_| invalid("DDE CDATA size exceeds u64"))?,
        )
        .map_err(map_execution)?;
    Ok(!value.trim().is_empty())
}

fn valid_general_ref(reference: &[u8]) -> bool {
    match reference {
        b"amp" | b"lt" | b"gt" | b"apos" | b"quot" => true,
        value if value.starts_with(b"#x") => std::str::from_utf8(&value[2..])
            .ok()
            .and_then(|value| u32::from_str_radix(value, 16).ok())
            .is_some_and(valid_xml_scalar),
        value if value.starts_with(b"#") => {
            value[1..].iter().all(|byte| byte.is_ascii_digit())
                && std::str::from_utf8(&value[1..])
                    .ok()
                    .and_then(|value| value.parse::<u32>().ok())
                    .is_some_and(valid_xml_scalar)
        },
        _ => false,
    }
}

fn valid_xml_scalar(value: u32) -> bool {
    matches!(value, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0x10FFFF)
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("invalid XML boolean '{value}'"))),
    }
}

fn xml_text_is_valid(value: &str) -> bool {
    !value.chars().any(|character| {
        matches!(
            character,
            '\u{0000}'..='\u{0008}' | '\u{000B}'..='\u{000C}' | '\u{000E}'..='\u{001F}'
        )
    })
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == TABLE => NamespaceKind::Table,
        ResolveResult::Bound(Namespace(uri)) if *uri == MCE => NamespaceKind::MarkupCompatibility,
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn namespace_kind_checked(namespace: &ResolveResult<'_>) -> Result<NamespaceKind> {
    if matches!(namespace, ResolveResult::Unknown(_)) {
        return Err(invalid("DDE XML contains an undeclared element prefix"));
    }
    Ok(namespace_kind(namespace))
}

fn is_spreadsheet(namespace: NamespaceKind, local: &[u8]) -> bool {
    is(namespace, local, NamespaceKind::Office, b"spreadsheet")
}

fn is(namespace: NamespaceKind, local: &[u8], expected_ns: NamespaceKind, expected: &[u8]) -> bool {
    namespace == expected_ns && local == expected
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidStructure(message.into())
}

/// Construct the finite default context used by the convenience DDE APIs.
pub(crate) fn default_context() -> ExecutionContext {
    let (_cancellation_source, token) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("non-zero worker count"),
        NonZeroUsize::new(1).expect("non-zero task count"),
        NonZeroU64::new(1024 * 1024).expect("non-zero in-flight byte limit"),
        0,
    )
    .expect("fixed DDE execution policy is valid");
    ExecutionContext::new(
        Budget::root(
            "ods-dde",
            litchi_core::Limits::for_profile(Profile::TrustedBatch),
        ),
        token,
        execution_limits,
    )
}

pub(crate) fn map_execution(error: ExecutionError) -> Error {
    match error {
        ExecutionError::Cancelled => Error::Cancelled,
        ExecutionError::ResourceLimit(limit) => Error::ExecutionLimit(limit),
        other => invalid(format!("DDE execution policy rejected operation: {other}")),
    }
}

fn xml_position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_position_error| Error::PositionOverflow)
}

fn validate_required_source_value(name: &'static str, value: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("DDE attribute {name} must not be empty")));
    }
    if value.len() > MAX_TEXT_BYTES {
        return Err(Error::ResourceLimit {
            resource: name,
            actual: value.len(),
            maximum: MAX_TEXT_BYTES,
        });
    }
    if !xml_text_is_valid(value) {
        return Err(invalid(format!(
            "DDE attribute {name} contains invalid XML character data"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    const CONTENT_PREFIX: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">"#;
    const CONTENT_SUFFIX: &str = r#"</office:spreadsheet></office:body></office:document-content>"#;

    fn content(body: &str) -> String {
        format!(r#"{CONTENT_PREFIX}<office:body><office:spreadsheet>{body}{CONTENT_SUFFIX}"#)
    }

    #[test]
    fn accepts_flat_spreadsheet_document_root() {
        let flat = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet"><office:body><office:spreadsheet/></office:body></office:document>"#;
        assert!(Snapshot::parse(flat).is_ok());
    }

    #[test]
    fn rejects_spreadsheet_hidden_under_foreign_or_mce_owner() {
        let foreign = content(
            r#"<vendor:wrapper xmlns:vendor="urn:example:vendor"><office:spreadsheet/></vendor:wrapper>"#,
        );
        assert!(Snapshot::parse(&foreign).is_err());

        let mce = content(
            r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback><office:spreadsheet/></mc:Fallback></mc:AlternateContent>"#,
        );
        assert!(Snapshot::parse(&mce).is_err());
    }

    #[test]
    fn rejects_noncanonical_document_body_spreadsheet_ancestry() {
        let wrong_root = r#"<office:spreadsheet xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#;
        assert!(Snapshot::parse(wrong_root).is_err());

        let duplicate_body =
            format!(r#"{CONTENT_PREFIX}<office:body/><office:body/></office:document-content>"#);
        assert!(Snapshot::parse(&duplicate_body).is_err());

        let duplicate_spreadsheet = format!(
            r#"{CONTENT_PREFIX}<office:body><office:spreadsheet/><office:spreadsheet/></office:body></office:document-content>"#
        );
        assert!(Snapshot::parse(&duplicate_spreadsheet).is_err());

        let office_sibling = format!(
            r#"{CONTENT_PREFIX}<office:body><office:automatic-styles/><office:spreadsheet/></office:body></office:document-content>"#
        );
        assert!(Snapshot::parse(&office_sibling).is_err());

        let wrong_flat_mimetype = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.text"><office:body><office:spreadsheet/></office:body></office:document>"#;
        assert!(Snapshot::parse(wrong_flat_mimetype).is_err());
    }

    #[test]
    fn rejects_unresolved_namespace_prefixes() {
        let unresolved_element = content(r#"<unknown:extension/>"#);
        assert!(Snapshot::parse(&unresolved_element).is_err());

        let unresolved_attribute = content(r#"<office:spreadsheet unknown:flag="true"/>"#);
        assert!(Snapshot::parse(&unresolved_attribute).is_err());
    }

    #[test]
    fn rejects_non_initial_or_repeated_xml_declarations() {
        let repeated = format!(
            r#"<?xml version="1.0"?><?xml version="1.0"?>{CONTENT_PREFIX}<office:body><office:spreadsheet/></office:body></office:document-content>"#
        );
        assert!(Snapshot::parse(&repeated).is_err());

        let late = format!(
            r#"<!--before declaration--><?xml version="1.0"?>{CONTENT_PREFIX}<office:body><office:spreadsheet/></office:body></office:document-content>"#
        );
        assert!(Snapshot::parse(&late).is_err());
    }

    fn parser_context(scope: &'static str, memory: u64, work: u64) -> ExecutionContext {
        parser_context_with_depth(scope, memory, work, 1_024)
    }

    fn parser_context_with_depth(
        scope: &'static str,
        memory: u64,
        work: u64,
        depth: u64,
    ) -> ExecutionContext {
        let (_cancellation_source, token) = CancellationSource::pair();
        let budget = Budget::root(
            scope,
            litchi_core::Limits::new(memory, 1_000_000, 1_000_000, 10_000, depth, work),
        );
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("worker"),
            NonZeroUsize::new(1).expect("task"),
            NonZeroU64::new(1_000_000).expect("in-flight bytes"),
            0,
        )
        .expect("execution limits");
        ExecutionContext::new(budget, token, execution_limits)
    }

    #[test]
    fn parser_reports_finite_work_and_memory_budget_refusals() {
        let work_context = parser_context("dde-parser-work", 1_000_000, 1);
        let work_error =
            Snapshot::parse_with_context(&content(""), Limits::default(), &work_context)
                .expect_err("event work must be bounded");
        let Error::ExecutionLimit(work_limit) = work_error else {
            panic!("expected structured work limit");
        };
        assert_eq!(work_limit.resource, Resource::Work);
        assert_eq!(work_limit.scope.as_ref(), "dde-parser-work");

        let memory_context = parser_context("dde-parser-memory", 1, 1_000_000);
        let memory_error =
            Snapshot::parse_with_context(&content(""), Limits::default(), &memory_context)
                .expect_err("retained source memory must be bounded");
        let Error::ExecutionLimit(memory_limit) = memory_error else {
            panic!("expected structured memory limit");
        };
        assert_eq!(memory_limit.resource, Resource::Memory);
        assert_eq!(memory_limit.scope.as_ref(), "dde-parser-memory");
    }

    #[test]
    fn parser_checks_context_depth_and_releases_sibling_reservations() {
        let low_depth_context =
            parser_context_with_depth("dde-parser-depth", 1_000_000, 1_000_000, 2);
        let depth_error =
            Snapshot::parse_with_context(&content(""), Limits::default(), &low_depth_context)
                .expect_err("root/body/spreadsheet nesting must consume context depth");
        let Error::ExecutionLimit(depth_limit) = depth_error else {
            panic!("expected structured depth limit");
        };
        assert_eq!(depth_limit.resource, Resource::Depth);
        assert_eq!(depth_limit.observed, 3);
        assert_eq!(depth_limit.limit, 2);
        assert_eq!(depth_limit.scope.as_ref(), "dde-parser-depth");

        let sibling_body = (0..16).map(|_| "<table:table/>").collect::<String>();
        let sibling_context =
            parser_context_with_depth("dde-parser-depth-siblings", 1_000_000, 1_000_000, 4);
        Snapshot::parse_with_context(&content(&sibling_body), Limits::default(), &sibling_context)
            .expect("sibling empty tables must release active depth reservations");
    }

    #[test]
    fn namespace_resolution_is_admitted_before_reader_allocation() {
        let giant_uri = format!("urn:giant:{}", "x".repeat(8_192));
        let source = format!(
            r#"<office:document-content xmlns:office="{office}" xmlns:table="{table}" xmlns:giant="{giant}"><office:body><office:spreadsheet/></office:body></office:document-content>"#,
            office = String::from_utf8_lossy(OFFICE),
            table = String::from_utf8_lossy(TABLE),
            giant = giant_uri,
        );
        let parser_capacity = vector_capacity_bound(
            INITIAL_SCOPE_CAPACITY,
            INITIAL_SCOPE_CAPACITY,
            NAMESPACE_RESOLVER_BASE_BYTES,
        )
        .expect("test parser capacity bound");
        let context = parser_context(
            "dde-parser-namespace",
            u64::try_from(source.len() + parser_capacity + 1).expect("test memory size"),
            1_000_000,
        );
        let error = Snapshot::parse_with_context(&source, Limits::default(), &context)
            .expect_err("giant namespace must be refused before resolver allocation");
        let Error::ExecutionLimit(limit) = error else {
            panic!("expected structured namespace memory limit");
        };
        assert_eq!(limit.resource, Resource::Memory);
        assert_eq!(limit.scope.as_ref(), "dde-parser-namespace");
    }

    #[test]
    fn unresolved_name_storage_is_admitted_before_reader_resolution() {
        let giant_prefix = format!("unknown{}", "x".repeat(8_192));
        let source = format!(
            r#"<office:document-content xmlns:office="{office}" {prefix}:flag="value"><office:body><office:spreadsheet/></office:body></office:document-content>"#,
            office = String::from_utf8_lossy(OFFICE),
            prefix = giant_prefix,
        );
        let parser_capacity = vector_capacity_bound(
            INITIAL_SCOPE_CAPACITY,
            INITIAL_SCOPE_CAPACITY,
            NAMESPACE_RESOLVER_BASE_BYTES,
        )
        .expect("test parser capacity bound");
        let context = parser_context(
            "dde-parser-unresolved-name",
            u64::try_from(source.len() + parser_capacity + 1).expect("test memory size"),
            1_000_000,
        );
        let error = Snapshot::parse_with_context(&source, Limits::default(), &context)
            .expect_err("unresolved name must be refused before resolver allocation");
        let Error::ExecutionLimit(limit) = error else {
            panic!("expected structured unresolved-name memory limit");
        };
        assert_eq!(limit.resource, Resource::Memory);
        assert_eq!(limit.scope.as_ref(), "dde-parser-unresolved-name");
    }

    #[test]
    fn preserves_predefined_and_numeric_cache_references_but_rejects_custom_entities() {
        let cache = r#"<table:dde-links><table:dde-link><office:dde-source office:dde-application="app" office:dde-topic="topic" office:dde-item="item"/><table:table><table:table-row><table:table-cell office:string-value="a &amp; b">a &amp; b &#x20;</table:table-cell></table:table-row></table:table></table:dde-link></table:dde-links>"#;
        let snapshot = Snapshot::parse(&content(cache)).expect("predefined references are inert");
        assert!(snapshot.links()[0].cached_table_xml().contains("&amp;"));
        assert!(snapshot.links()[0].cached_table_xml().contains("&#x20;"));

        let custom = cache.replace("&#x20;", "&custom-entity;");
        assert!(Snapshot::parse(&content(&custom)).is_err());
    }

    #[test]
    fn leading_bom_keeps_cached_table_range_in_source_coordinates() {
        let cache = r#"<table:dde-links><table:dde-link><office:dde-source office:dde-application="app" office:dde-topic="topic" office:dde-item="item"/><table:table/></table:dde-link></table:dde-links>"#;
        let mut source = String::from_utf8_lossy(UTF8_BOM).into_owned();
        source.push_str(&content(cache));
        let snapshot = Snapshot::parse(&source).expect("UTF-8 BOM is accepted");
        assert_eq!(snapshot.links()[0].cached_table_xml(), "<table:table/>");
    }

    #[test]
    fn retains_physical_table_index_for_duplicate_sheet_names() {
        let body = r#"<table:table table:name="Same"><office:dde-source office:dde-application="app" office:dde-topic="topic" office:dde-item="one"/></table:table><table:table table:name="Same"><office:dde-source office:dde-application="app" office:dde-topic="topic" office:dde-item="two"/></table:table>"#;
        let snapshot = Snapshot::parse(&content(body)).expect("duplicate names are indexed");
        assert_eq!(snapshot.sheet_sources().len(), 2);
        assert_eq!(snapshot.inventory().sheet_table_indices, [0, 1]);
    }
}
