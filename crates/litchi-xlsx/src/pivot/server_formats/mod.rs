//! Source-bound support for the MS-XLSX `pivotTableServerFormats` extension.
//!
//! The extension is deliberately kept in a separate owner.  It is a payload
//! of a PivotTable definition, but it is only meaningful when the workbook
//! contains the `pivotTableReferences` relationship closure for a
//! non-worksheet PivotTable and the associated cache is an external cache
//! with a `pivotCacheIdVersion` extension.  The owner captures the original
//! PivotTable XML in an `Arc` and retains byte ranges for the required ordered
//! `serverFormat` leaves and their optional scalar attributes.  Unknown
//! extension bytes and all other package members remain source material.
//!
//! The ordinary workbook surface is `Workbook::pivot_table_server_formats`
//! (also available as `Workbook::pivot_server_formats`) for typed reads and
//! `Workbook::edit_pivot_table` for a source-checked edit.  Selectors are a
//! semantic PivotTable name or workbook position; callers do not provide OPC
//! relationship IDs or Part names.  An edit stages `ServerFormat` values or
//! explicit `AttributeEdit::{Keep,Set,Clear}` operations for `culture` and
//! `format`, then publishes through `WorkbookCommit::workbook` and its exact
//! reversible `WorkbookPatch`.
//!
//! This owner covers the C510 `pivotTableServerFormats` payload under a
//! non-worksheet PivotTable and its required `pivotTableReferences`/external
//! cache closure, including the recognized ABF5 cache-version extension.  It
//! reads and writes the required ordered `x15:serverFormat` leaves and their
//! two optional scalar attributes.  Structural edits preserve known
//! `pivotValueCellExtra@in` associations or fail closed when the source cannot
//! prove them.  It does not infer newer extension URIs, invent index mappings,
//! or claim native Excel PivotTable authoring or refresh semantics.

use std::collections::{HashMap, HashSet};
use std::mem::size_of;
use std::ops::Range;
use std::sync::Arc;

use litchi_core::xml::ReaderOrigin;
use litchi_core::{Resource, ResourceLimit};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, OwnedRelationships, PackURI, Part, ReadLimits};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{LocalName, Prefix, PrefixDeclaration, QName};
use quick_xml::reader::Reader;

use litchi_ooxml_common::xml_name::is_qualified_name;

use crate::Workbook;
use crate::error::{Error, Result, invalid};
use crate::raw;
use crate::source_attributes::{
    append_escaped_xstring, escaped_xstring_len, try_escaped_xstring, validate_xml_characters,
    value_span,
};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

mod relationships;
use relationships::RelationshipIndex;

pub mod cached_unique_names;
pub mod table_data;

const CORE_NS: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_CORE_NS: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const EXT_NS: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const X14_NS: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const MCE_NS: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const REL_NS: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";

const PIVOT_TABLE_SERVER_FORMATS_URI: &str = "{C510F80B-63DE-4267-81D5-13C33094786E}";
const PIVOT_TABLE_DATA_URI: &str = "{44433962-1CF7-4059-B4EE-95C3D5FFCF73}";
const PIVOT_TABLE_REFERENCES_URI: &str = "{983426D0-5260-488c-9760-48F4B6AC55F4}";
const PIVOT_CACHE_DEFINITION_URI: &str = "{725AE2AE-9491-48BE-B2B4-4EB974FC3084}";
const PIVOT_CACHE_ID_VERSION_URI: &str = "{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}";
const CACHE_SOURCE_URI: &str = "{F057638F-6D5F-4E77-A914-E7F072B9BCA8}";
const CONNECTION_MODEL_URI: &str = "{DE250136-89BD-433C-8126-D09CA5730AF9}";
const CONNECTIONS_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
const STRICT_CONNECTIONS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/connections";
const CONNECTIONS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";

const MAX_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_NODES: usize = 100_000;
const MAX_NAMESPACE_DECLARATIONS: usize = 256;
const MAX_NAMESPACE_BYTES: usize = 16 * 1024;
const MAX_NAME_BYTES: usize = 16 * 1024;
const MAX_FRAGMENT_BYTES: usize = 1024 * 1024;
const MAX_SERVER_FORMATS: usize = (1usize << 31) - 1;
const MAX_REFERENCE_COUNT: usize = (1usize << 31) - 1;
const MAX_ATTRIBUTE_TEXT_BYTES: usize = 1024 * 1024;

#[derive(Clone, Copy, Debug)]
struct XmlScanLimits {
    part_bytes: usize,
    events: usize,
    depth: usize,
    attribute_bytes: usize,
    namespace_bytes: usize,
    name_bytes: usize,
    namespace_declarations: usize,
}

impl XmlScanLimits {
    fn from_read_limits(limits: ReadLimits) -> Self {
        let part_bytes = usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX);
        Self {
            part_bytes: MAX_PART_BYTES.min(part_bytes),
            events: MAX_NODES.min(limits.max_xml_events()),
            depth: MAX_DEPTH.min(limits.max_xml_depth()),
            attribute_bytes: MAX_ATTRIBUTE_TEXT_BYTES.min(limits.max_xml_attribute_bytes()),
            namespace_bytes: MAX_NAMESPACE_BYTES.min(limits.max_xml_attribute_bytes()),
            name_bytes: MAX_NAME_BYTES.min(limits.max_xml_attribute_bytes()),
            namespace_declarations: MAX_NAMESPACE_DECLARATIONS,
        }
    }
}

/// Build the explicit MCE profile used for the workbook catalog.  The raw
/// catalog parser is also used by callers outside this owner and therefore
/// retains its compatibility wrapper, but this graph must never silently use
/// an unbounded/default policy while admitting a package.
fn catalog_mce_limits(limits: ReadLimits) -> litchi_ooxml_common::mce::Limits {
    let part_bytes = usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX);
    let input = MAX_PART_BYTES.min(part_bytes);
    let aggregate_bytes = usize::try_from(limits.max_total_part_bytes()).unwrap_or(usize::MAX);
    let caller_tokens = limits
        .max_xml_events()
        .min(limits.max_xml_attribute_bytes());
    let directive_tokens = MAX_NODES.min(caller_tokens);
    litchi_ooxml_common::mce::Limits {
        max_input_bytes: input,
        // MCE may expand a small source substantially while selecting a
        // branch. Keep both the intermediate output and directive state
        // within the caller's XML/Part policy before the raw catalog sees it.
        max_output_bytes: MAX_PART_BYTES.min(aggregate_bytes),
        max_depth: MAX_DEPTH.min(limits.max_xml_depth()),
        max_namespace_bindings: MAX_NAMESPACE_DECLARATIONS.min(caller_tokens),
        max_directive_tokens: directive_tokens,
        max_choices_per_alternate: MAX_REFERENCE_COUNT.min(1024).min(directive_tokens),
        max_attributes_per_element: litchi_ooxml_common::mce::DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT,
    }
}

/// Semantic selector for a workbook PivotTable.
///
/// `Name` and `Position` are workbook-facing identities.  OPC relationship
/// IDs and Part URIs are intentionally not accepted here.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotTableSelector<'a> {
    /// Select the unique PivotTable name in the workbook.
    Name(&'a str),
    /// Select the zero-based position in the workbook's non-worksheet
    /// `pivotTableReferences` collection.
    Position(usize),
}

impl<'a> From<&'a str> for PivotTableSelector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

impl<'a> From<&'a String> for PivotTableSelector<'a> {
    fn from(value: &'a String) -> Self {
        Self::Name(value.as_str())
    }
}

impl From<usize> for PivotTableSelector<'static> {
    fn from(value: usize) -> Self {
        Self::Position(value)
    }
}

/// A workbook-contextual, typed view of one semantic PivotTable.
///
/// This is the ordinary Workbook-facing read surface.  It intentionally
/// contains only the resolved semantic collection; source XML, relationship
/// IDs, and Part names remain private to the source-bound owner.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PivotTableView {
    value: PivotTableServerFormats,
}

impl PivotTableView {
    /// The resolved server-format collection.
    #[must_use]
    pub fn server_formats(&self) -> &PivotTableServerFormats {
        &self.value
    }

    /// Alias for callers using the extension element's collection name.
    #[must_use]
    pub fn formats(&self) -> &[ServerFormat] {
        self.value.formats()
    }

    /// Semantic PivotTable name.
    #[must_use]
    pub fn table_name(&self) -> &str {
        self.value.table_name()
    }

    /// Logical workbook PivotCache ID.
    #[must_use]
    pub const fn cache_id(&self) -> u32 {
        self.value.cache_id()
    }
}

/// The optional `culture` or `format` value on one server format.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ServerFormat {
    /// Opaque SpreadsheetML `ST_Xstring` culture value, when present.
    pub culture: Option<String>,
    /// Opaque SpreadsheetML `ST_Xstring` number-format value, when present.
    pub format: Option<String>,
}

impl ServerFormat {
    /// Build a server-format value from optional decoded attributes.
    #[must_use]
    pub const fn new(culture: Option<String>, format: Option<String>) -> Self {
        Self { culture, format }
    }
}

/// One explicit scalar edit.  `Keep`, `Set`, and `Clear` are intentionally
/// distinct so an omitted field cannot accidentally clear or retain a value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AttributeEdit {
    /// Leave this attribute at its source value.
    Keep,
    /// Set or replace this attribute with decoded SpreadsheetML text.
    Set(String),
    /// Remove this attribute from the source start tag.
    Clear,
}

impl AttributeEdit {
    /// Construct a `Set` operation.
    #[must_use]
    pub fn set(value: impl Into<String>) -> Self {
        Self::Set(value.into())
    }

    /// Construct a `Clear` operation.
    #[must_use]
    pub const fn clear() -> Self {
        Self::Clear
    }

    /// Construct a `Keep` operation.
    #[must_use]
    pub const fn keep() -> Self {
        Self::Keep
    }
}

/// Explicit per-attribute update for one server format.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ServerFormatEdit {
    /// `culture` update.
    pub culture: AttributeEdit,
    /// `format` update.
    pub format: AttributeEdit,
}

impl ServerFormatEdit {
    /// Keep both source attributes.
    #[must_use]
    pub const fn keep() -> Self {
        Self {
            culture: AttributeEdit::Keep,
            format: AttributeEdit::Keep,
        }
    }

    /// Set both optional attributes explicitly, clearing an omitted value.
    #[must_use]
    pub fn from_value(value: ServerFormat) -> Self {
        Self {
            culture: value
                .culture
                .map_or(AttributeEdit::Clear, AttributeEdit::Set),
            format: value
                .format
                .map_or(AttributeEdit::Clear, AttributeEdit::Set),
        }
    }
}

/// A typed, source-bound read of one `pivotTableServerFormats` payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotTableServerFormats {
    table_name: String,
    cache_id: u32,
    formats: Box<[ServerFormat]>,
    diagnostic_index_boundary: bool,
    mce_ambiguous: bool,
}

impl PivotTableServerFormats {
    /// The semantic PivotTable name.
    #[must_use]
    pub fn table_name(&self) -> &str {
        &self.table_name
    }

    /// The semantic workbook PivotCache ID closed by the table and cache
    /// relationship graph.
    #[must_use]
    pub const fn cache_id(&self) -> u32 {
        self.cache_id
    }

    /// The bounded ordered server-format list.
    #[must_use]
    pub fn formats(&self) -> &[ServerFormat] {
        &self.formats
    }

    /// Alias for callers using the XML collection name.
    #[must_use]
    pub fn server_formats(&self) -> &[ServerFormat] {
        self.formats()
    }

    /// Whether one `pivotValueCellExtra@in` used the unresolved inclusive
    /// endpoint.  Such a source remains readable but cannot be changed.
    #[must_use]
    pub const fn has_diagnostic_index_boundary(&self) -> bool {
        self.diagnostic_index_boundary
    }

    /// Whether the recognized payload is selected through an MCE branch.
    /// Such a source is readable but scalar publication is conservatively
    /// refused because branch selection cannot be proven by this owner.
    #[must_use]
    pub const fn has_ambiguous_mce_owner(&self) -> bool {
        self.mce_ambiguous
    }
}

/// An exact source-bound snapshot of the selected PivotTable extension.
#[derive(Clone, Debug)]
pub struct Snapshot {
    value: PivotTableServerFormats,
    table: SourcePart,
    cache: SourcePart,
    connections: Option<SourcePart>,
    workbook: SourcePart,
    workbook_owner: Arc<Vec<u8>>,
    workbook_context: Arc<Vec<u8>>,
    table_owner: Range<usize>,
    extension_owner: Range<usize>,
    count: AttributeSource,
    entries: Box<[EntrySource]>,
    index_refs: Box<[IndexReference]>,
    opaque_index_refs: bool,
    selection: usize,
}

impl Snapshot {
    /// Resolve and read a semantic PivotTable selector from an owned OPC
    /// package.
    pub fn load<'a>(
        package: &OpcPackage,
        selector: impl Into<PivotTableSelector<'a>>,
    ) -> Result<Self> {
        let selector = selector.into();
        let graph = Graph::load(package)?;
        graph.snapshot(selector)
    }

    /// Alias emphasizing that this state is tied to exact source bytes.
    pub fn read<'a>(
        package: &OpcPackage,
        selector: impl Into<PivotTableSelector<'a>>,
    ) -> Result<Self> {
        Self::load(package, selector)
    }

    /// The typed server-format collection.
    #[must_use]
    pub fn server_formats(&self) -> &PivotTableServerFormats {
        &self.value
    }

    /// Alias for callers that use the extension element's name.
    #[must_use]
    pub fn formats(&self) -> &[ServerFormat] {
        self.value.formats()
    }

    /// Semantic PivotTable name.
    #[must_use]
    pub fn table_name(&self) -> &str {
        self.value.table_name()
    }

    /// Logical PivotCache ID.
    #[must_use]
    pub const fn cache_id(&self) -> u32 {
        self.value.cache_id()
    }

    /// Whether the source carries the unresolved `in == count` diagnostic.
    #[must_use]
    pub const fn has_diagnostic_index_boundary(&self) -> bool {
        self.value.has_diagnostic_index_boundary()
    }

    /// Whether this source requires an MCE branch-preserving edit path.
    #[must_use]
    pub const fn has_ambiguous_mce_owner(&self) -> bool {
        self.value.has_ambiguous_mce_owner()
    }

    /// Whether an unproven `pivotValueCellExtra@in` source prevents a
    /// structural list edit from preserving index associations.
    #[must_use]
    pub const fn has_opaque_index_references(&self) -> bool {
        self.opaque_index_refs
    }

    /// Exact source XML for the PivotTable definition Part.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.table.bytes.as_slice()
    }

    /// Share the exact source XML allocation.
    #[must_use]
    pub fn source_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.table.bytes)
    }

    /// Internal physical identity retained for diagnostics and patch
    /// verification.  Ordinary selectors never require callers to provide it.
    #[must_use]
    pub fn table_part(&self) -> &PackURI {
        &self.table.name
    }

    /// The exact owner range containing the PivotTable definition root.
    #[must_use]
    pub fn source_owner_range(&self) -> Range<usize> {
        self.table_owner.clone()
    }

    fn same_source(&self, other: &Self) -> bool {
        self.selection == other.selection
            && self.table.same_source(&other.table)
            && self.cache.same_source(&other.cache)
            && same_optional_source(&self.connections, &other.connections)
            && self.workbook.same_source(&other.workbook)
    }

    fn same_readset(&self, other: &Self) -> bool {
        self.selection == other.selection
            && self.table.same_source(&other.table)
            && self.cache.same_source(&other.cache)
            && same_optional_source(&self.connections, &other.connections)
            && self.workbook.same_closure(&other.workbook)
            && self.workbook_owner.as_slice() == other.workbook_owner.as_slice()
            && self.workbook_context.as_slice() == other.workbook_context.as_slice()
    }

    fn same_closure(&self, other: &Self) -> bool {
        self.table.same_closure(&other.table)
            && self.cache.same_closure(&other.cache)
            && same_optional_closure(&self.connections, &other.connections)
            && self.workbook.same_closure(&other.workbook)
    }
}

#[derive(Clone, Debug)]
struct EntrySource {
    culture: Option<AttributeSource>,
    format: Option<AttributeSource>,
    start_tag: Range<usize>,
    element: Range<usize>,
    qname: Vec<u8>,
    namespace_decl: Option<Vec<u8>>,
}

#[derive(Clone, Debug)]
struct AttributeSource {
    value: Range<usize>,
    whole: Range<usize>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SourcePart {
    name: PackURI,
    content_type: String,
    bytes: Arc<Vec<u8>>,
    relationships: Option<OwnedRelationships>,
    incoming: Vec<RelationshipState>,
}

impl SourcePart {
    fn capture(index: &RelationshipIndex<'_>, part: &dyn Part) -> Result<Self> {
        let name = part.partname().clone();
        let relationships = if name.as_str() == "/" {
            Some(index.root_source().clone())
        } else {
            index.source(&name)?
        };
        let source_ref = index.source_ref(&name)?;
        if let (Some(relationships), Some(source_ref)) = (relationships.as_ref(), source_ref)
            && relationships.owner() != source_ref.owner()
        {
            return Err(invalid("PivotTable relationship source owner mismatch"));
        }
        let incoming = index.capture_incoming(&name)?;
        Ok(Self {
            name,
            content_type: part.content_type().to_owned(),
            bytes: part.blob_arc(),
            relationships,
            incoming,
        })
    }

    fn same_source(&self, other: &Self) -> bool {
        self.name == other.name
            && self.content_type == other.content_type
            && self.bytes.as_slice() == other.bytes.as_slice()
            && self.same_closure(other)
    }

    fn same_closure(&self, other: &Self) -> bool {
        self.name == other.name
            && self.content_type == other.content_type
            && self.relationships == other.relationships
            && self.incoming == other.incoming
    }

    /// Bytes retained by the source-bound relationship closure in addition to
    /// the owning Part blob.  The relationship XML is shared by an Arc, but
    /// it remains live for exact read-set/inverse checks; incoming edges own
    /// their copied strings and therefore need their own aggregate admission.
    fn retained_closure_bytes(&self) -> Result<usize> {
        let mut retained = self
            .name
            .as_str()
            .len()
            .checked_add(self.content_type.len())
            .ok_or_else(|| invalid("PivotTable source closure bytes overflow"))?;
        if let Some(relationships) = &self.relationships {
            retained = retained
                .checked_add(relationships.bytes().len())
                .and_then(|bytes| bytes.checked_add(relationships.owner().as_str().len()))
                .and_then(|bytes| bytes.checked_add(size_of::<OwnedRelationships>()))
                .ok_or_else(|| invalid("PivotTable relationship source bytes overflow"))?;
        }
        retained = retained
            .checked_add(size_of::<Vec<RelationshipState>>())
            .and_then(|bytes| {
                bytes.checked_add(
                    self.incoming
                        .len()
                        .checked_mul(size_of::<RelationshipState>())?,
                )
            })
            .ok_or_else(|| invalid("PivotTable incoming relationship bytes overflow"))?;
        for relationship in &self.incoming {
            for length in [
                relationship.source.len(),
                relationship.id.len(),
                relationship.reltype.len(),
                relationship.target.len(),
            ] {
                retained = retained
                    .checked_add(length)
                    .ok_or_else(|| invalid("PivotTable incoming relationship bytes overflow"))?;
            }
        }
        Ok(retained)
    }
}

fn same_optional_source(left: &Option<SourcePart>, right: &Option<SourcePart>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => left.same_source(right),
        _ => false,
    }
}

fn same_optional_closure(left: &Option<SourcePart>, right: &Option<SourcePart>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => left.same_closure(right),
        _ => false,
    }
}

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
struct RelationshipState {
    source: String,
    id: String,
    reltype: String,
    target: String,
    external: bool,
}

/// Failure-atomic edits over one PivotTable's ordered server-format list and
/// scalar metadata.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    staged: Vec<ServerFormat>,
    staged_sources: Vec<Option<usize>>,
    selection: usize,
}

impl<'a> Transaction<'a> {
    /// Resolve a semantic PivotTable selector and begin an isolated edit.
    pub fn new<'selector>(
        target: &'a mut OpcPackage,
        selector: impl Into<PivotTableSelector<'selector>>,
    ) -> Result<Self> {
        let before = Snapshot::load(target, selector)?;
        ensure_editable(&before)?;
        let mut staged = Vec::new();
        staged
            .try_reserve_exact(before.formats().len())
            .map_err(|source| Error::Allocation {
                resource: "PivotTable staged server-format values",
                source,
            })?;
        staged.extend_from_slice(before.formats());
        let mut staged_sources = Vec::new();
        staged_sources
            .try_reserve_exact(before.formats().len())
            .map_err(|source| Error::Allocation {
                resource: "PivotTable staged server-format source order",
                source,
            })?;
        staged_sources.extend((0..before.formats().len()).map(Some));
        Ok(Self {
            selection: before.selection,
            target,
            before,
            staged,
            staged_sources,
        })
    }

    /// Exact source state captured at transaction start.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged ordered server-format values.
    #[must_use]
    pub fn server_formats(&self) -> &[ServerFormat] {
        &self.staged
    }

    /// Set, replace, or clear both optional attributes explicitly.  `None`
    /// means clear for this method; use [`Self::update_server_format`] when a
    /// field must be retained.
    pub fn set_server_format(&mut self, index: usize, value: ServerFormat) -> Result<bool> {
        let slot = self
            .staged
            .get_mut(index)
            .ok_or_else(|| invalid("server-format index is out of range"))?;
        if *slot == value {
            return Ok(false);
        }
        let text_limit = caller_attribute_limit(self.target.read_limits());
        validate_text(&value.culture, text_limit)?;
        validate_text(&value.format, text_limit)?;
        *slot = value;
        Ok(true)
    }

    /// Insert one ordered `serverFormat` leaf before `index`.
    ///
    /// The collection is schema-required and non-empty, so insertion accepts
    /// `0..=len` and publication updates the required `count` attribute.
    pub fn insert_server_format(&mut self, index: usize, value: ServerFormat) -> Result<()> {
        if index > self.staged.len() {
            return Err(invalid("server-format insertion index is out of range"));
        }
        if self.staged.len() >= MAX_SERVER_FORMATS {
            return Err(invalid("server-format collection exceeds its count limit"));
        }
        let text_limit = caller_attribute_limit(self.target.read_limits());
        validate_text(&value.culture, text_limit)?;
        validate_text(&value.format, text_limit)?;
        self.staged
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "PivotTable staged server-format values",
                source,
            })?;
        self.staged_sources
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "PivotTable staged server-format source order",
                source,
            })?;
        self.staged.insert(index, value);
        self.staged_sources.insert(index, None);
        Ok(())
    }

    /// Append one ordered `serverFormat` leaf.
    pub fn push_server_format(&mut self, value: ServerFormat) -> Result<()> {
        self.insert_server_format(self.staged.len(), value)
    }

    /// Remove one ordered `serverFormat` leaf.
    ///
    /// Removing the last leaf is refused because the schema requires at least
    /// one `serverFormat` child.  Removing the whole payload/container is a
    /// separate lifecycle operation and is not inferred here.
    pub fn remove_server_format(&mut self, index: usize) -> Result<ServerFormat> {
        if self.staged.len() == 1 {
            return Err(invalid(
                "pivotTableServerFormats must retain one server-format child",
            ));
        }
        if index >= self.staged.len() {
            return Err(invalid("server-format removal index is out of range"));
        }
        self.staged_sources.remove(index);
        Ok(self.staged.remove(index))
    }

    /// Move one leaf, interpreting `to` in the final ordered sequence.
    pub fn move_server_format(&mut self, from: usize, to: usize) -> Result<()> {
        if from >= self.staged.len() || to >= self.staged.len() {
            return Err(invalid("server-format move index is out of range"));
        }
        if from != to {
            let value = self.staged.remove(from);
            let source = self.staged_sources.remove(from);
            self.staged.insert(to, value);
            self.staged_sources.insert(to, source);
        }
        Ok(())
    }

    /// Reorder the leaves using a final-position-to-old-position permutation.
    pub fn reorder_server_formats(&mut self, order: &[usize]) -> Result<()> {
        let len = self.staged.len();
        validate_permutation(order, len)?;
        let mut values = Vec::new();
        values
            .try_reserve_exact(len)
            .map_err(|source| Error::Allocation {
                resource: "PivotTable reordered server-format values",
                source,
            })?;
        for &index in order {
            values.push(clone_server_format(&self.staged[index])?);
        }
        let mut sources = Vec::new();
        sources
            .try_reserve_exact(len)
            .map_err(|source| Error::Allocation {
                resource: "PivotTable reordered server-format source order",
                source,
            })?;
        for &index in order {
            sources.push(self.staged_sources[index]);
        }
        self.staged = values;
        self.staged_sources = sources;
        Ok(())
    }

    /// Apply explicit `Keep`/`Set`/`Clear` operations to one server format.
    pub fn update_server_format(&mut self, index: usize, edit: ServerFormatEdit) -> Result<bool> {
        let current = self
            .staged
            .get(index)
            .cloned()
            .ok_or_else(|| invalid("server-format index is out of range"))?;
        let mut next = current;
        let text_limit = caller_attribute_limit(self.target.read_limits());
        apply_attribute(&mut next.culture, edit.culture, text_limit)?;
        apply_attribute(&mut next.format, edit.format, text_limit)?;
        self.set_server_format(index, next)
    }

    /// Set or clear only the `culture` attribute.
    pub fn set_culture(&mut self, index: usize, value: Option<String>) -> Result<bool> {
        self.update_server_format(
            index,
            ServerFormatEdit {
                culture: value.map_or(AttributeEdit::Clear, AttributeEdit::Set),
                format: AttributeEdit::Keep,
            },
        )
    }

    /// Set or clear only the `format` attribute.
    pub fn set_format(&mut self, index: usize, value: Option<String>) -> Result<bool> {
        self.update_server_format(
            index,
            ServerFormatEdit {
                culture: AttributeEdit::Keep,
                format: value.map_or(AttributeEdit::Clear, AttributeEdit::Set),
            },
        )
    }

    /// Clear the `culture` attribute.
    pub fn clear_culture(&mut self, index: usize) -> Result<bool> {
        self.update_server_format(
            index,
            ServerFormatEdit {
                culture: AttributeEdit::Clear,
                format: AttributeEdit::Keep,
            },
        )
    }

    /// Clear the `format` attribute.
    pub fn clear_format(&mut self, index: usize) -> Result<bool> {
        self.update_server_format(
            index,
            ServerFormatEdit {
                culture: AttributeEdit::Keep,
                format: AttributeEdit::Clear,
            },
        )
    }

    /// Whether the staged semantic list differs from the source list.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.formats() != self.staged.as_slice()
            || self
                .staged_sources
                .iter()
                .enumerate()
                .any(|(index, source)| *source != Some(index))
    }

    /// Validate, rewrite, reopen, and atomically publish the staged edit.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        if self.target.is_signed() || self.target.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let current = Snapshot::load(self.target, PivotTableSelector::Position(self.selection))?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: self.before.table_part().to_string(),
            });
        }
        let output = rewrite_table(
            &self.before,
            &self.staged,
            &self.staged_sources,
            caller_part_limit(self.target.read_limits()),
            self.target.read_limits().max_total_part_bytes(),
            caller_attribute_limit(self.target.read_limits()),
            package_total_bytes(self.target)?,
        )?;
        let mut candidate = self.target.clone();
        candidate
            .get_part_mut(self.before.table_part())?
            .set_blob_shared(Arc::new(output));
        validate_candidate_limits(&candidate)?;
        let after = Snapshot::load(&candidate, PivotTableSelector::Position(self.selection))?;
        if after.formats() != self.staged.as_slice()
            || !after.same_closure(&self.before)
            || after.table_owner.start != self.before.table_owner.start
        {
            return Err(invalid(
                "PivotTable server-format publication changed semantic state or source closure",
            ));
        }
        let patch = Patch::new(self.before, after.clone());
        *self.target = candidate;
        Ok(Commit::new(after, patch, true))
    }
}

fn ensure_editable(before: &Snapshot) -> Result<()> {
    if before.has_diagnostic_index_boundary() {
        return Err(invalid(
            "pivotTableServerFormats has an ambiguous pivotValueCellExtra@in == count boundary",
        ));
    }
    if before.has_ambiguous_mce_owner() {
        return Err(invalid(
            "pivotTableServerFormats selected through an MCE branch is read-only",
        ));
    }
    Ok(())
}

/// Load one source-bound PivotTable server-format snapshot.
pub fn load<'a>(
    package: &OpcPackage,
    selector: impl Into<PivotTableSelector<'a>>,
) -> Result<Snapshot> {
    Snapshot::load(package, selector)
}

/// Start a source-bound PivotTable server-format transaction.
pub fn edit<'package, 'selector>(
    package: &'package mut OpcPackage,
    selector: impl Into<PivotTableSelector<'selector>>,
) -> Result<Transaction<'package>> {
    Transaction::new(package, selector)
}

/// Apply an exact source-checked patch atomically.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<()> {
    patch.apply(package)
}

/// An exact source-checked reversible edit.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Required source state.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact state produced by application.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact source-bound inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply the patch after validating source closure and semantic readback.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        self.apply_checked(package, false)
    }

    fn apply_readset(&self, package: &mut OpcPackage) -> Result<()> {
        self.apply_checked(package, true)
    }

    fn apply_checked(
        &self,
        package: &mut OpcPackage,
        allow_unrelated_workbook_changes: bool,
    ) -> Result<()> {
        let current = Snapshot::load(package, PivotTableSelector::Position(self.before.selection))?;
        let source_matches = if allow_unrelated_workbook_changes {
            current.same_readset(&self.before)
        } else {
            current.same_source(&self.before)
        };
        if !source_matches {
            return Err(Error::PatchConflict {
                part: self.before.table_part().to_string(),
            });
        }
        if self.is_empty() {
            return Ok(());
        }
        if package.is_signed() || package.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let mut candidate = package.clone();
        candidate
            .get_part_mut(self.after.table_part())?
            .set_blob_shared(Arc::clone(&self.after.table.bytes));
        validate_candidate_limits(&candidate)?;
        let resulting = Snapshot::load(
            &candidate,
            PivotTableSelector::Position(self.after.selection),
        )?;
        let target_matches = if allow_unrelated_workbook_changes {
            resulting.same_readset(&self.after)
        } else {
            resulting.same_source(&self.after)
        };
        if resulting.formats() != self.after.formats() || !target_matches {
            return Err(invalid(
                "PivotTable server-format patch verification failed",
            ));
        }
        *package = candidate;
        Ok(())
    }
}

/// A committed transaction and reversible patch.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn new(snapshot: Snapshot, patch: Patch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            changed,
        }
    }

    /// Resulting immutable source-bound state.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact reversible patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Whether authored values changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Consume the commit into its snapshot and patch.
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Start an ordinary semantic edit against an immutable [`Workbook`] snapshot.
///
/// The package owner and its source-bound ranges stay private.  Callers stage
/// only typed server-format values and receive a new Workbook through
/// [`WorkbookCommit::workbook`] after commit.
pub struct WorkbookTransaction {
    source: Workbook,
    before: Snapshot,
    staged: Vec<ServerFormat>,
    staged_sources: Vec<Option<usize>>,
    selection: usize,
}

impl WorkbookTransaction {
    pub(crate) fn new<'selector>(
        source: &Workbook,
        selector: impl Into<PivotTableSelector<'selector>>,
    ) -> Result<Self> {
        let before = Snapshot::load(source.pivot_package(), selector)?;
        ensure_editable(&before)?;
        let mut staged = Vec::new();
        staged
            .try_reserve_exact(before.formats().len())
            .map_err(|source| Error::Allocation {
                resource: "Workbook staged pivot server-format values",
                source,
            })?;
        staged.extend_from_slice(before.formats());
        let mut staged_sources = Vec::new();
        staged_sources
            .try_reserve_exact(before.formats().len())
            .map_err(|source| Error::Allocation {
                resource: "Workbook staged pivot server-format source order",
                source,
            })?;
        staged_sources.extend((0..before.formats().len()).map(Some));
        Ok(Self {
            source: source.clone(),
            selection: before.selection,
            before,
            staged,
            staged_sources,
        })
    }

    /// Exact typed state captured at transaction start.
    #[must_use]
    pub fn before(&self) -> &PivotTableServerFormats {
        self.before.server_formats()
    }

    /// Currently staged ordered server-format values.
    #[must_use]
    pub fn server_formats(&self) -> &[ServerFormat] {
        &self.staged
    }

    /// Alias for callers using the extension element's name.
    #[must_use]
    pub fn formats(&self) -> &[ServerFormat] {
        self.server_formats()
    }

    /// Set, replace, or clear both optional attributes explicitly.
    pub fn set_server_format(&mut self, index: usize, value: ServerFormat) -> Result<bool> {
        let slot = self
            .staged
            .get_mut(index)
            .ok_or_else(|| invalid("server-format index is out of range"))?;
        if *slot == value {
            return Ok(false);
        }
        let text_limit = caller_attribute_limit(self.source.pivot_package().read_limits());
        validate_text(&value.culture, text_limit)?;
        validate_text(&value.format, text_limit)?;
        *slot = value;
        Ok(true)
    }

    /// Insert one ordered `serverFormat` leaf before `index`.
    pub fn insert_server_format(&mut self, index: usize, value: ServerFormat) -> Result<()> {
        if index > self.staged.len() {
            return Err(invalid("server-format insertion index is out of range"));
        }
        if self.staged.len() >= MAX_SERVER_FORMATS {
            return Err(invalid("server-format collection exceeds its count limit"));
        }
        let text_limit = caller_attribute_limit(self.source.pivot_package().read_limits());
        validate_text(&value.culture, text_limit)?;
        validate_text(&value.format, text_limit)?;
        self.staged
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "Workbook staged pivot server-format values",
                source,
            })?;
        self.staged_sources
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "Workbook staged pivot server-format source order",
                source,
            })?;
        self.staged.insert(index, value);
        self.staged_sources.insert(index, None);
        Ok(())
    }

    /// Append one ordered `serverFormat` leaf.
    pub fn push_server_format(&mut self, value: ServerFormat) -> Result<()> {
        self.insert_server_format(self.staged.len(), value)
    }

    /// Remove one ordered `serverFormat` leaf while retaining the required
    /// non-empty collection.
    pub fn remove_server_format(&mut self, index: usize) -> Result<ServerFormat> {
        if self.staged.len() == 1 {
            return Err(invalid(
                "pivotTableServerFormats must retain one server-format child",
            ));
        }
        if index >= self.staged.len() {
            return Err(invalid("server-format removal index is out of range"));
        }
        self.staged_sources.remove(index);
        Ok(self.staged.remove(index))
    }

    /// Move one leaf, interpreting `to` in the final ordered sequence.
    pub fn move_server_format(&mut self, from: usize, to: usize) -> Result<()> {
        if from >= self.staged.len() || to >= self.staged.len() {
            return Err(invalid("server-format move index is out of range"));
        }
        if from != to {
            let value = self.staged.remove(from);
            let source = self.staged_sources.remove(from);
            self.staged.insert(to, value);
            self.staged_sources.insert(to, source);
        }
        Ok(())
    }

    /// Reorder the leaves using a final-position-to-old-position permutation.
    pub fn reorder_server_formats(&mut self, order: &[usize]) -> Result<()> {
        let len = self.staged.len();
        validate_permutation(order, len)?;
        let mut values = Vec::new();
        values
            .try_reserve_exact(len)
            .map_err(|source| Error::Allocation {
                resource: "Workbook reordered server-format values",
                source,
            })?;
        for &index in order {
            values.push(clone_server_format(&self.staged[index])?);
        }
        let mut sources = Vec::new();
        sources
            .try_reserve_exact(len)
            .map_err(|source| Error::Allocation {
                resource: "Workbook reordered server-format source order",
                source,
            })?;
        for &index in order {
            sources.push(self.staged_sources[index]);
        }
        self.staged = values;
        self.staged_sources = sources;
        Ok(())
    }

    /// Apply explicit `Keep`/`Set`/`Clear` operations to one server format.
    pub fn update_server_format(&mut self, index: usize, edit: ServerFormatEdit) -> Result<bool> {
        let current = self
            .staged
            .get(index)
            .cloned()
            .ok_or_else(|| invalid("server-format index is out of range"))?;
        let mut next = current;
        let text_limit = caller_attribute_limit(self.source.pivot_package().read_limits());
        apply_attribute(&mut next.culture, edit.culture, text_limit)?;
        apply_attribute(&mut next.format, edit.format, text_limit)?;
        self.set_server_format(index, next)
    }

    /// Set or clear only the `culture` attribute.
    pub fn set_culture(&mut self, index: usize, value: Option<String>) -> Result<bool> {
        self.update_server_format(
            index,
            ServerFormatEdit {
                culture: value.map_or(AttributeEdit::Clear, AttributeEdit::Set),
                format: AttributeEdit::Keep,
            },
        )
    }

    /// Set or clear only the `format` attribute.
    pub fn set_format(&mut self, index: usize, value: Option<String>) -> Result<bool> {
        self.update_server_format(
            index,
            ServerFormatEdit {
                culture: AttributeEdit::Keep,
                format: value.map_or(AttributeEdit::Clear, AttributeEdit::Set),
            },
        )
    }

    /// Clear the `culture` attribute.
    pub fn clear_culture(&mut self, index: usize) -> Result<bool> {
        self.update_server_format(
            index,
            ServerFormatEdit {
                culture: AttributeEdit::Clear,
                format: AttributeEdit::Keep,
            },
        )
    }

    /// Clear the `format` attribute.
    pub fn clear_format(&mut self, index: usize) -> Result<bool> {
        self.update_server_format(
            index,
            ServerFormatEdit {
                culture: AttributeEdit::Keep,
                format: AttributeEdit::Clear,
            },
        )
    }

    /// Whether the staged semantic list differs from the source list.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.formats() != self.staged.as_slice()
            || self
                .staged_sources
                .iter()
                .enumerate()
                .any(|(index, source)| *source != Some(index))
    }

    /// Validate, reopen, and publish a new immutable Workbook snapshot.
    pub fn commit(self) -> Result<WorkbookCommit> {
        let source = self.source;
        let staged = self.staged;
        let staged_sources = self.staged_sources;
        let selection = self.selection;
        let mut candidate = source.pivot_package().clone();
        let mut transaction =
            Transaction::new(&mut candidate, PivotTableSelector::Position(selection))?;
        transaction.staged = staged;
        transaction.staged_sources = staged_sources;
        let committed = transaction.commit()?;
        let changed = committed.changed();
        let low_patch = committed.patch().clone();
        let workbook = source.adopt_published_package(candidate)?;
        let patch = WorkbookPatch::new(source, workbook.clone(), low_patch);
        Ok(WorkbookCommit::new(workbook, patch, changed))
    }
}

/// An exact reversible patch over ordinary Workbook snapshots.
#[derive(Clone, Debug)]
pub struct WorkbookPatch {
    before: Workbook,
    after: Workbook,
    source_patch: Patch,
}

impl WorkbookPatch {
    fn new(before: Workbook, after: Workbook, source_patch: Patch) -> Self {
        Self {
            before,
            after,
            source_patch,
        }
    }

    /// Typed semantic state required before application.
    #[must_use]
    pub fn before(&self) -> &PivotTableServerFormats {
        self.source_patch.before().server_formats()
    }

    /// Typed semantic state produced by application.
    #[must_use]
    pub fn after(&self) -> &PivotTableServerFormats {
        self.source_patch.after().server_formats()
    }

    /// Whether this patch preserves the complete source owner byte-for-byte.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.source_patch.is_empty()
    }

    /// Return the exact source-bound inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            source_patch: self.source_patch.inverse(),
        }
    }

    /// Apply after checking the complete source-bound owner closure.
    pub fn apply(&self, source: &Workbook) -> Result<WorkbookCommit> {
        source.ensure_mutation_allowed("apply_pivot_table_server_formats_patch")?;
        let mut candidate = source.pivot_package().clone();
        self.source_patch.apply_readset(&mut candidate)?;
        let workbook = source.adopt_published_package(candidate)?;
        Ok(WorkbookCommit::new(
            workbook.clone(),
            Self::new(source.clone(), workbook, self.source_patch.clone()),
            !self.is_empty(),
        ))
    }
}

/// A committed ordinary Workbook server-format edit.
#[derive(Clone, Debug)]
pub struct WorkbookCommit {
    workbook: Workbook,
    patch: WorkbookPatch,
    changed: bool,
}

impl WorkbookCommit {
    fn new(workbook: Workbook, patch: WorkbookPatch, changed: bool) -> Self {
        Self {
            workbook,
            patch,
            changed,
        }
    }

    /// Resulting immutable Workbook snapshot.
    #[must_use]
    pub fn workbook(&self) -> &Workbook {
        &self.workbook
    }

    /// Alias for callers using snapshot terminology.
    #[must_use]
    pub fn snapshot(&self) -> &PivotTableServerFormats {
        self.patch.after()
    }

    /// Exact reversible Workbook patch.
    #[must_use]
    pub fn patch(&self) -> &WorkbookPatch {
        &self.patch
    }

    /// Whether authored values changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Consume the commit into its Workbook and patch.
    pub fn into_parts(self) -> (Workbook, WorkbookPatch) {
        (self.workbook, self.patch)
    }
}

/// Resolve a semantic PivotTable through the ordinary Workbook facade.
pub(crate) fn view_workbook<'selector>(
    workbook: &Workbook,
    selector: impl Into<PivotTableSelector<'selector>>,
) -> Result<PivotTableView> {
    let snapshot = Snapshot::load(workbook.pivot_package(), selector)?;
    Ok(PivotTableView {
        value: snapshot.server_formats().clone(),
    })
}

/// Start a semantic Workbook edit without exposing its OPC owner.
pub(crate) fn edit_workbook<'selector>(
    workbook: &Workbook,
    selector: impl Into<PivotTableSelector<'selector>>,
) -> Result<WorkbookTransaction> {
    WorkbookTransaction::new(workbook, selector)
}

struct Graph {
    workbook: SourcePart,
    connections: Option<SourcePart>,
    workbook_owner: Arc<Vec<u8>>,
    workbook_context: Arc<Vec<u8>>,
    workbook_mce_ambiguous: bool,
    workbook_closure_diagnostic: bool,
    /// Ordinary worksheet table names that collide with a PivotTable name.
    /// They make only a textual selector ambiguous; position selectors remain
    /// valid and the collision is not a PivotTable schema failure.
    ordinary_name_collisions: HashSet<String>,
    refs: Vec<Reference>,
    caches: HashMap<u32, CacheInfo>,
}

struct Reference {
    table_uri: PackURI,
    name: String,
    cache_id: u32,
    table: TableInfo,
}

struct WorkbookReferences {
    refs: Vec<Reference>,
    owner: Arc<Vec<u8>>,
    context: Arc<Vec<u8>>,
    mce_ambiguous: bool,
    diagnostic: bool,
}

struct TableInfo {
    name: String,
    cache_id: u32,
    cache_uri: PackURI,
    source: SourcePart,
    owner: Range<usize>,
    payload: Option<PayloadInfo>,
}

struct CacheInfo {
    uri: PackURI,
    source: SourcePart,
    mce_ambiguous: bool,
    closure_diagnostic: bool,
}

#[derive(Clone, Copy, Debug, Default)]
struct CacheClosureStatus {
    mce_ambiguous: bool,
    diagnostic: bool,
}

#[derive(Debug)]
struct ConnectionCatalog {
    by_id: HashMap<u32, String>,
    by_name: HashMap<String, u32>,
    model_ids: HashSet<u32>,
}

struct PayloadInfo {
    entries: Vec<ParsedEntry>,
    owner: Range<usize>,
    count: AttributeSource,
    index_refs: Vec<IndexReference>,
    opaque_index_refs: bool,
    diagnostic_index_boundary: bool,
    mce_ambiguous: bool,
}

struct ParsedEntry {
    value: ServerFormat,
    source: EntrySource,
}

#[derive(Clone, Debug)]
struct IndexReference {
    value: Range<usize>,
    index: usize,
}

impl Graph {
    fn load(package: &OpcPackage) -> Result<Self> {
        Self::load_with_owner(package, true, false, false)
    }

    /// Load the common workbook/PivotTable/cache relationship graph for the
    /// `pivotTableData` owner.  The data owner validates its own payload after
    /// this shared closure has been established, so the C510 server-format
    /// list is deliberately not required here.
    pub(super) fn load_for_table_data(package: &OpcPackage) -> Result<Self> {
        Self::load_with_owner(package, false, true, true)
    }

    fn load_with_owner(
        package: &OpcPackage,
        parse_server_formats: bool,
        require_cache_definition_id: bool,
        allow_empty_connection_names: bool,
    ) -> Result<Self> {
        validate_candidate_limits(package)?;
        let relationship_index = RelationshipIndex::build(package)?;
        let workbook_part = package.main_document_part()?;
        if workbook_part.blob().len() > MAX_PART_BYTES
            || workbook_part.blob().len() > caller_part_limit(package.read_limits())
        {
            return Err(invalid("workbook Part exceeds the caller's Part limit"));
        }
        let workbook = SourcePart::capture(&relationship_index, workbook_part)?;
        // Run the owner scanner before the general catalog parser so caller
        // XML event/depth/attribute ceilings admit the workbook bytes before
        // any catalog projection can retain expanded values.
        let workbook_references = parse_workbook_references(
            package,
            workbook_part,
            &relationship_index,
            parse_server_formats,
            !parse_server_formats,
        )?;
        let catalog = raw::parse_catalog_with_mce(
            workbook_part.blob(),
            &litchi_ooxml_common::mce::Capabilities::ooxml_baseline(),
            &catalog_mce_limits(package.read_limits()),
        )?;
        let mut connections = None;
        let mut connection_catalog = None;
        let refs = workbook_references.refs;
        if refs.is_empty() || refs.len() > MAX_REFERENCE_COUNT {
            return Err(invalid(
                "pivotTableReferences must contain between one and fewer than 2^31 references",
            ));
        }

        validate_table_incoming_closure(package, workbook_part, &refs, &relationship_index)?;

        let mut names = HashMap::<String, PackURI>::new();
        names
            .try_reserve(refs.len())
            .map_err(|source| Error::Allocation {
                resource: "PivotTable name index",
                source,
            })?;
        for reference in &refs {
            let table = &reference.table;
            if table.name != reference.name || table.cache_id != reference.cache_id {
                return Err(invalid(
                    "workbook pivotTableReference does not match its PivotTable definition",
                ));
            }
            if names
                .insert(table.name.clone(), reference.table_uri.clone())
                .is_some()
            {
                return Err(invalid("PivotTable name is not unique in the workbook"));
            }
        }

        // The base workbook catalog supplies the semantic cache IDs.  Every
        // referenced cache is then closed through the table relationship and
        // the cache-definition owner.
        let mut caches = HashMap::new();
        caches
            .try_reserve(catalog.pivot_caches.len())
            .map_err(|source| Error::Allocation {
                resource: "PivotCache graph",
                source,
            })?;
        for cache in &catalog.pivot_caches {
            let relation = workbook_part
                .rels()
                .get(&cache.relationship_id)
                .ok_or_else(|| invalid("workbook PivotCache relationship is missing"))?;
            if relation.is_external()
                || relation.target_query().is_some()
                || relation.target_fragment().is_some()
                || !matches!(
                    relation.reltype(),
                    rt::PIVOT_CACHE_DEFINITION | rt::STRICT_PIVOT_CACHE_DEFINITION
                )
            {
                return Err(invalid("workbook PivotCache relationship is invalid"));
            }
            let uri = relation.target_partname()?;
            let canonical_uri = relationship_index
                .canonical_part(&uri)?
                .ok_or_else(|| invalid("workbook PivotCache target is not a package Part"))?;
            let part = package.get_part(canonical_uri)?;
            if part.content_type() != ct::SML_PIVOT_CACHE_DEFINITION {
                return Err(invalid(
                    "workbook PivotCache relationship targets the wrong content type",
                ));
            }
            if caches
                .insert(
                    cache.cache_id,
                    CacheInfo {
                        source: SourcePart::capture(&relationship_index, part)?,
                        uri: canonical_uri.clone(),
                        mce_ambiguous: false,
                        closure_diagnostic: false,
                    },
                )
                .is_some()
            {
                return Err(invalid("workbook contains duplicate PivotCache IDs"));
            }
        }
        let mut validated_cache_uris = HashMap::new();
        validated_cache_uris
            .try_reserve(refs.len())
            .map_err(|source| Error::Allocation {
                resource: "PivotCache validation index",
                source,
            })?;
        for reference in &refs {
            let table = &reference.table;
            let cache = caches.get(&table.cache_id).ok_or_else(|| {
                invalid("PivotTable cache ID is absent from workbook pivotCaches")
            })?;
            if cache.uri != table.cache_uri {
                return Err(invalid(
                    "PivotTable cache relationship does not match workbook cache ID",
                ));
            }
            let cache_part = package.get_part(&cache.uri)?;
            if cache_part.blob().len() > MAX_PART_BYTES
                || cache_part.blob().len() > caller_part_limit(package.read_limits())
            {
                return Err(invalid("PivotCache Part exceeds the caller's Part limit"));
            }
            if let Some(previous_id) = validated_cache_uris.get(&cache.uri) {
                if *previous_id != table.cache_id {
                    return Err(invalid(
                        "one PivotCache Part is bound to multiple semantic cache IDs",
                    ));
                }
            } else {
                validated_cache_uris.insert(cache.uri.clone(), table.cache_id);
                let cache_scan = if require_cache_definition_id {
                    scan_xml_with_mce(
                        cache_part.blob(),
                        "pivotCacheDefinition",
                        package.read_limits(),
                    )?
                } else {
                    scan_xml(
                        cache_part.blob(),
                        "pivotCacheDefinition",
                        package.read_limits(),
                    )?
                };
                let cache_root = cache_scan
                    .elements
                    .iter()
                    .find(|element| element.parent_index.is_none())
                    .ok_or_else(|| invalid("PivotCache Part has no root"))?;
                let cache_source = cache_scan
                    .elements
                    .iter()
                    .find(|element| {
                        element.parent_index == Some(cache_root.index)
                            && element.ns == cache_root.ns
                            && element.local == b"cacheSource"
                    })
                    .ok_or_else(|| invalid("PivotCache has no cacheSource"))?;
                if cache_requires_connection_route(&cache_scan, cache_root, cache_source)?
                    && connection_catalog.is_none()
                {
                    if let Some((source, catalog)) = load_connections(
                        package,
                        workbook_part,
                        &relationship_index,
                        false,
                        allow_empty_connection_names,
                    )? {
                        connections = Some(source);
                        connection_catalog = Some(catalog);
                    }
                }
                let cache_status = validate_cache_closure(
                    &cache_scan,
                    table.cache_id,
                    connection_catalog.as_ref(),
                    require_cache_definition_id,
                    allow_empty_connection_names,
                )?;
                if let Some(cache_info) = caches.get_mut(&table.cache_id) {
                    cache_info.mce_ambiguous = cache_status.mce_ambiguous;
                    cache_info.closure_diagnostic = cache_status.diagnostic;
                }
            }
        }

        // All PivotTable names in worksheet-owned tables also participate in
        // the workbook uniqueness rule.  Ordinary worksheet table names are
        // retained only as a conservative textual-selector ambiguity policy.
        let worksheet_names = worksheet_name_index(
            package,
            &catalog,
            &relationship_index,
            !parse_server_formats,
        )?;
        for name in &worksheet_names.pivot_names {
            if names.contains_key(name) {
                return Err(invalid("PivotTable name is not unique in the workbook"));
            }
        }
        let mut ordinary_name_collisions = HashSet::new();
        ordinary_name_collisions
            .try_reserve(worksheet_names.ordinary_names.len())
            .map_err(|source| Error::Allocation {
                resource: "ordinary worksheet table selector index",
                source,
            })?;
        for name in worksheet_names.ordinary_names {
            if names.contains_key(&name) {
                ordinary_name_collisions.insert(name);
            }
        }

        Ok(Self {
            workbook,
            connections,
            workbook_owner: workbook_references.owner,
            workbook_context: workbook_references.context,
            workbook_mce_ambiguous: workbook_references.mce_ambiguous,
            workbook_closure_diagnostic: workbook_references.diagnostic,
            ordinary_name_collisions,
            refs,
            caches,
        })
    }

    fn snapshot<'s>(&'s self, selector: PivotTableSelector<'s>) -> Result<Snapshot> {
        let index = match selector {
            PivotTableSelector::Position(index) => {
                if index >= self.refs.len() {
                    return Err(invalid("PivotTable selector position is out of range"));
                }
                index
            },
            PivotTableSelector::Name(name) => {
                let mut found = None;
                for (index, reference) in self.refs.iter().enumerate() {
                    if reference.name == name {
                        if found.is_some() {
                            return Err(invalid("PivotTable selector name is ambiguous"));
                        }
                        found = Some(index);
                    }
                }
                found.ok_or_else(|| invalid("PivotTable selector name did not resolve"))?
            },
        };
        let reference = self
            .refs
            .get(index)
            .ok_or_else(|| invalid("PivotTable selector did not resolve"))?;
        let table = &reference.table;
        let payload = table
            .payload
            .as_ref()
            .ok_or_else(|| invalid("PivotTable server-format payload is not loaded"))?;
        let cache_uri = self
            .caches
            .get(&table.cache_id)
            .ok_or_else(|| invalid("PivotTable cache graph is incomplete"))?;
        let table_source = table.source.clone();
        let cache_source = cache_uri.source.clone();
        let mut boxed = Vec::new();
        boxed
            .try_reserve_exact(payload.entries.len())
            .map_err(|source| Error::Allocation {
                resource: "PivotTable server-format semantic values",
                source,
            })?;
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(payload.entries.len())
            .map_err(|source| Error::Allocation {
                resource: "PivotTable server-format source ranges",
                source,
            })?;
        for entry in &payload.entries {
            boxed.push(entry.value.clone());
            entries.push(entry.source.clone());
        }
        let mut index_refs = Vec::new();
        index_refs
            .try_reserve_exact(payload.index_refs.len())
            .map_err(|source| Error::Allocation {
                resource: "PivotTable server-format index references",
                source,
            })?;
        index_refs.extend(payload.index_refs.iter().cloned());
        Ok(Snapshot {
            value: PivotTableServerFormats {
                table_name: table.name.clone(),
                cache_id: table.cache_id,
                formats: boxed.into_boxed_slice(),
                diagnostic_index_boundary: payload.diagnostic_index_boundary,
                mce_ambiguous: payload.mce_ambiguous,
            },
            table: table_source,
            cache: cache_source,
            connections: self.connections.clone(),
            workbook: self.workbook.clone(),
            workbook_owner: Arc::clone(&self.workbook_owner),
            workbook_context: Arc::clone(&self.workbook_context),
            table_owner: table.owner.clone(),
            extension_owner: payload.owner.clone(),
            count: payload.count.clone(),
            entries: entries.into_boxed_slice(),
            index_refs: index_refs.into_boxed_slice(),
            opaque_index_refs: payload.opaque_index_refs,
            selection: index,
        })
    }

    pub(super) fn ordinary_name_collision(&self, name: &str) -> bool {
        self.ordinary_name_collisions.contains(name)
    }
}

fn apply_attribute(
    slot: &mut Option<String>,
    edit: AttributeEdit,
    maximum_text_bytes: usize,
) -> Result<()> {
    match edit {
        AttributeEdit::Keep => {},
        AttributeEdit::Set(value) => {
            validate_text_value(&value, maximum_text_bytes)?;
            *slot = Some(value);
        },
        AttributeEdit::Clear => *slot = None,
    }
    Ok(())
}

fn validate_permutation(order: &[usize], length: usize) -> Result<()> {
    if order.len() != length {
        return Err(invalid("server-format reorder must include every leaf"));
    }
    let mut seen = Vec::new();
    seen.try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "PivotTable server-format reorder check",
            source,
        })?;
    seen.resize(length, false);
    for &index in order {
        let slot = seen
            .get_mut(index)
            .ok_or_else(|| invalid("server-format reorder index is out of range"))?;
        if *slot {
            return Err(invalid("server-format reorder contains a duplicate index"));
        }
        *slot = true;
    }
    Ok(())
}

fn clone_server_format(value: &ServerFormat) -> Result<ServerFormat> {
    Ok(ServerFormat {
        culture: try_clone_string(value.culture.as_deref(), "server-format culture")?,
        format: try_clone_string(value.format.as_deref(), "server-format format")?,
    })
}

fn try_clone_string(value: Option<&str>, resource: &'static str) -> Result<Option<String>> {
    let Some(value) = value else {
        return Ok(None);
    };
    let mut cloned = String::new();
    cloned
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    cloned.push_str(value);
    Ok(Some(cloned))
}

fn validate_text(value: &Option<String>, maximum_text_bytes: usize) -> Result<()> {
    if let Some(value) = value {
        validate_text_value(value, maximum_text_bytes)?;
    }
    Ok(())
}

fn validate_text_value(value: &str, maximum_text_bytes: usize) -> Result<()> {
    if value.len() > MAX_ATTRIBUTE_TEXT_BYTES || value.len() > maximum_text_bytes {
        return Err(invalid("server-format attribute exceeds its text limit"));
    }
    Ok(())
}

fn validate_candidate_limits(package: &OpcPackage) -> Result<()> {
    let limits = package.read_limits();
    if package.part_count() > limits.max_parts() {
        return Err(invalid(
            "PivotTable candidate exceeds the caller's Part count limit",
        ));
    }
    let mut total = 0u64;
    let maximum_part_bytes = caller_part_limit(limits);
    // The per-Part and aggregate ceilings are charged against inflated
    // bytes, so this census decodes every payload (ADR 0030).
    for part in package.try_iter_parts() {
        let bytes = part?.blob().len();
        if bytes > MAX_PART_BYTES || bytes > maximum_part_bytes {
            return Err(invalid(
                "PivotTable candidate Part exceeds the caller's Part limit",
            ));
        }
        total = total
            .checked_add(bytes as u64)
            .ok_or_else(|| invalid("PivotTable candidate Part bytes overflow"))?;
    }
    if total > limits.max_total_part_bytes() {
        return Err(invalid(
            "PivotTable candidate exceeds the caller's aggregate Part limit",
        ));
    }
    Ok(())
}

fn caller_part_limit(limits: ReadLimits) -> usize {
    usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX)
}

fn caller_attribute_limit(limits: ReadLimits) -> usize {
    MAX_ATTRIBUTE_TEXT_BYTES.min(limits.max_xml_attribute_bytes())
}

fn package_total_bytes(package: &OpcPackage) -> Result<u64> {
    // The aggregate Part ceiling is charged against inflated bytes, so this
    // census decodes every payload (ADR 0030).
    package.try_iter_parts().try_fold(0u64, |total, part| {
        total
            .checked_add(u64::try_from(part?.blob().len()).unwrap_or(u64::MAX))
            .ok_or_else(|| invalid("PivotTable candidate Part bytes overflow"))
    })
}

fn rewrite_table(
    before: &Snapshot,
    staged: &[ServerFormat],
    staged_sources: &[Option<usize>],
    maximum_output_bytes: usize,
    maximum_total_bytes: u64,
    maximum_attribute_bytes: usize,
    current_total_bytes: u64,
) -> Result<Vec<u8>> {
    let scalar_shape = staged.len() == before.entries.len()
        && staged_sources.len() == before.entries.len()
        && staged_sources
            .iter()
            .enumerate()
            .all(|(index, source)| *source == Some(index));
    if scalar_shape {
        return rewrite_scalar_table(
            before,
            staged,
            maximum_output_bytes,
            maximum_total_bytes,
            maximum_attribute_bytes,
            current_total_bytes,
        );
    }
    rewrite_structural_table(
        before,
        staged,
        staged_sources,
        maximum_output_bytes,
        maximum_total_bytes,
        maximum_attribute_bytes,
        current_total_bytes,
    )
}

fn rewrite_scalar_table(
    before: &Snapshot,
    staged: &[ServerFormat],
    maximum_output_bytes: usize,
    maximum_total_bytes: u64,
    maximum_attribute_bytes: usize,
    current_total_bytes: u64,
) -> Result<Vec<u8>> {
    let source = before.table.bytes.as_slice();
    let owner = source
        .get(before.extension_owner.clone())
        .ok_or_else(|| invalid("PivotTable server-format owner range is invalid"))?;
    if owner.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid(
            "pivotTableServerFormats owner fragment exceeds limit",
        ));
    }
    let mut output_len = source.len();
    let mut owner_len = owner.len();
    let mut has_edits = false;
    for ((entry, previous), next) in before.entries.iter().zip(before.formats()).zip(staged) {
        has_edits |= apply_planned_attribute_delta(
            source,
            &entry.start_tag,
            b"culture",
            entry.culture.as_ref(),
            previous.culture.as_ref(),
            next.culture.as_deref(),
            &mut output_len,
            &mut owner_len,
            maximum_attribute_bytes,
        )?;
        has_edits |= apply_planned_attribute_delta(
            source,
            &entry.start_tag,
            b"format",
            entry.format.as_ref(),
            previous.format.as_ref(),
            next.format.as_deref(),
            &mut output_len,
            &mut owner_len,
            maximum_attribute_bytes,
        )?;
    }
    if !has_edits {
        return Ok(source.to_vec());
    }
    if owner_len > MAX_FRAGMENT_BYTES {
        return Err(invalid(
            "rewritten pivotTableServerFormats owner fragment exceeds limit",
        ));
    }
    if output_len > MAX_PART_BYTES || output_len > maximum_output_bytes {
        return Err(invalid("PivotTable output exceeds the caller's Part limit"));
    }
    let output_len_u64 = u64::try_from(output_len).unwrap_or(u64::MAX);
    let source_len = u64::try_from(source.len()).unwrap_or(u64::MAX);
    let prospective_total = current_total_bytes
        .checked_sub(source_len)
        .and_then(|total| total.checked_add(output_len_u64))
        .ok_or_else(|| invalid("PivotTable candidate aggregate bytes overflow"))?;
    if prospective_total > maximum_total_bytes {
        return Err(invalid(
            "PivotTable output exceeds the caller's aggregate Part limit",
        ));
    }
    let mut edits = Vec::<(Range<usize>, Vec<u8>)>::new();
    for ((entry, previous), next) in before.entries.iter().zip(before.formats()).zip(staged) {
        push_named_attribute_edit(
            source,
            &entry.start_tag,
            b"culture",
            entry.culture.as_ref(),
            previous.culture.as_ref(),
            next.culture.as_deref(),
            &mut edits,
        )?;
        push_named_attribute_edit(
            source,
            &entry.start_tag,
            b"format",
            entry.format.as_ref(),
            previous.format.as_ref(),
            next.format.as_deref(),
            &mut edits,
        )?;
    }
    edits.sort_by_key(|edit| std::cmp::Reverse(edit.0.start));
    for pair in edits.windows(2) {
        if pair[0].0.start < pair[1].0.end {
            return Err(invalid("server-format source edit ranges overlap"));
        }
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "PivotTable server-format output",
            source,
        })?;
    output.extend_from_slice(source);
    for (range, value) in edits {
        output.splice(range, value);
    }
    Ok(output)
}

#[derive(Clone, Copy)]
struct LeafAttributePlan<'a> {
    start: usize,
    end: usize,
    name: &'static [u8],
    value: Option<&'a str>,
}

fn rewrite_structural_table(
    before: &Snapshot,
    staged: &[ServerFormat],
    staged_sources: &[Option<usize>],
    maximum_output_bytes: usize,
    maximum_total_bytes: u64,
    maximum_attribute_bytes: usize,
    current_total_bytes: u64,
) -> Result<Vec<u8>> {
    let source = before.table.bytes.as_slice();
    let owner = source
        .get(before.extension_owner.clone())
        .ok_or_else(|| invalid("PivotTable server-format owner range is invalid"))?;
    if owner.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid(
            "pivotTableServerFormats owner fragment exceeds limit",
        ));
    }
    if staged.is_empty() || staged.len() > MAX_SERVER_FORMATS {
        return Err(invalid(
            "pivotTableServerFormats must retain between one and fewer than 2^31 leaves",
        ));
    }
    if staged_sources.len() != staged.len() {
        return Err(invalid(
            "server-format source order does not match staged values",
        ));
    }
    if before.value.has_diagnostic_index_boundary() {
        return Err(invalid(
            "pivotTableServerFormats has an ambiguous pivotValueCellExtra@in == count boundary",
        ));
    }
    if before.opaque_index_refs {
        return Err(invalid(
            "pivotTableServerFormats has an unproven pivotValueCellExtra@in reference",
        ));
    }
    let text_limit = maximum_attribute_bytes;
    for value in staged {
        validate_text(&value.culture, text_limit)?;
        validate_text(&value.format, text_limit)?;
    }

    let old_len = before.entries.len();
    let mut old_to_new = Vec::new();
    old_to_new
        .try_reserve_exact(old_len)
        .map_err(|source| Error::Allocation {
            resource: "PivotTable server-format source permutation",
            source,
        })?;
    old_to_new.resize(old_len, usize::MAX);
    for (new_index, source_index) in staged_sources.iter().enumerate() {
        let Some(source_index) = source_index else {
            continue;
        };
        let slot = old_to_new
            .get_mut(*source_index)
            .ok_or_else(|| invalid("server-format source order is out of range"))?;
        if *slot != usize::MAX {
            return Err(invalid("server-format source order contains a duplicate"));
        }
        *slot = new_index;
    }

    let count_changed = staged.len() != old_len;
    let count_bytes = decimal_u32(
        u32::try_from(staged.len())
            .map_err(|_| invalid("pivotTableServerFormats count does not fit unsignedInt"))?,
    );
    if count_changed
        && (before.count.value.start < before.extension_owner.start
            || before.count.value.end > before.extension_owner.end)
    {
        return Err(invalid(
            "pivotTableServerFormats count source range is outside its owner",
        ));
    }
    let mut owner_len = owner.len();
    if count_changed {
        owner_len = owner_len
            .checked_sub(
                before
                    .count
                    .value
                    .end
                    .saturating_sub(before.count.value.start),
            )
            .and_then(|length| length.checked_add(count_bytes.len))
            .ok_or_else(|| invalid("PivotTable server-format owner length overflows"))?;
    }
    for old_index in 0..old_len {
        let old_entry = &before.entries[old_index];
        let replacement_len = if old_index < staged.len() {
            staged_child_len(
                before,
                staged,
                staged_sources,
                old_index,
                maximum_attribute_bytes,
            )?
        } else {
            0
        };
        owner_len = owner_len
            .checked_sub(
                old_entry
                    .element
                    .end
                    .saturating_sub(old_entry.element.start),
            )
            .and_then(|length| length.checked_add(replacement_len))
            .ok_or_else(|| invalid("PivotTable server-format owner length overflows"))?;
    }
    for new_index in old_len..staged.len() {
        owner_len = owner_len
            .checked_add(staged_child_len(
                before,
                staged,
                staged_sources,
                new_index,
                maximum_attribute_bytes,
            )?)
            .ok_or_else(|| invalid("PivotTable server-format owner length overflows"))?;
    }
    if owner_len > MAX_FRAGMENT_BYTES {
        return Err(invalid(
            "rewritten pivotTableServerFormats owner fragment exceeds limit",
        ));
    }

    let mut output_len = source.len();
    output_len = output_len
        .checked_sub(owner.len())
        .and_then(|length| length.checked_add(owner_len))
        .ok_or_else(|| invalid("PivotTable server-format output length overflows"))?;
    let mut previous_reference_end = 0usize;
    for reference in &before.index_refs {
        if reference.value.start < previous_reference_end
            || reference.value.end > source.len()
            || reference.value.start > reference.value.end
        {
            return Err(invalid(
                "pivotValueCellExtra@in source ranges are not ordered",
            ));
        }
        previous_reference_end = reference.value.end;
        let new_index = *old_to_new
            .get(reference.index)
            .ok_or_else(|| invalid("pivotValueCellExtra@in source index is out of range"))?;
        if new_index == usize::MAX {
            return Err(invalid(
                "removing a server-format leaf would orphan pivotValueCellExtra@in",
            ));
        }
        if new_index != reference.index {
            let old_digits = reference.value.end.saturating_sub(reference.value.start);
            let new_digits = decimal_u32(
                u32::try_from(new_index)
                    .map_err(|_| invalid("pivotValueCellExtra@in does not fit unsignedInt"))?,
            );
            output_len = output_len
                .checked_sub(old_digits)
                .and_then(|length| length.checked_add(new_digits.len))
                .ok_or_else(|| invalid("PivotTable server-format output length overflows"))?;
        }
    }
    if owner_len > MAX_PART_BYTES
        || owner_len > maximum_output_bytes
        || output_len > MAX_PART_BYTES
        || output_len > maximum_output_bytes
    {
        return Err(invalid("PivotTable output exceeds the caller's Part limit"));
    }
    let output_len_u64 = u64::try_from(output_len).unwrap_or(u64::MAX);
    let source_len = u64::try_from(source.len()).unwrap_or(u64::MAX);
    let prospective_total = current_total_bytes
        .checked_sub(source_len)
        .and_then(|total| total.checked_add(output_len_u64))
        .ok_or_else(|| invalid("PivotTable candidate aggregate bytes overflow"))?;
    if prospective_total > maximum_total_bytes {
        return Err(invalid(
            "PivotTable output exceeds the caller's aggregate Part limit",
        ));
    }

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "PivotTable structural server-format output",
            source,
        })?;
    append_structural_table(
        &mut output,
        source,
        before,
        staged,
        staged_sources,
        &old_to_new,
        count_changed,
        &count_bytes,
        maximum_attribute_bytes,
    )?;
    if output.len() != output_len {
        return Err(invalid(
            "PivotTable structural server-format output length mismatch",
        ));
    }
    Ok(output)
}

fn staged_child_len(
    before: &Snapshot,
    staged: &[ServerFormat],
    staged_sources: &[Option<usize>],
    new_index: usize,
    maximum_attribute_bytes: usize,
) -> Result<usize> {
    let Some(source_index) = staged_sources[new_index] else {
        let include_open =
            target_leaf_includes_open(before.table.bytes.as_slice(), before, new_index);
        return new_leaf_len(
            &before.entries[0].qname,
            before.entries[0].namespace_decl.as_deref(),
            &staged[new_index],
            maximum_attribute_bytes,
            include_open,
        );
    };
    let include_open = target_leaf_includes_open(before.table.bytes.as_slice(), before, new_index);
    existing_leaf_len(
        before.table.bytes.as_slice(),
        &before.entries[source_index],
        &before.formats()[source_index],
        &staged[new_index],
        maximum_attribute_bytes,
        include_open,
    )
}

fn target_leaf_includes_open(source: &[u8], before: &Snapshot, new_index: usize) -> bool {
    before.entries.get(new_index).is_none_or(|entry| {
        entry
            .element
            .start
            .checked_sub(1)
            .and_then(|index| source.get(index))
            != Some(&b'<')
    })
}

fn existing_leaf_len(
    source: &[u8],
    entry: &EntrySource,
    previous: &ServerFormat,
    next: &ServerFormat,
    maximum_attribute_bytes: usize,
    include_open: bool,
) -> Result<usize> {
    let plans = leaf_attribute_plans(source, entry, previous, next, maximum_attribute_bytes)?;
    let source_has_open = source.get(entry.element.start) == Some(&b'<');
    let mut length = entry.element.end.saturating_sub(entry.element.start);
    if include_open != source_has_open {
        length = if include_open {
            length
                .checked_add(1)
                .ok_or_else(|| invalid("server-format leaf length overflows"))?
        } else {
            length
                .checked_sub(1)
                .ok_or_else(|| invalid("server-format leaf length underflows"))?
        };
    }
    for plan in plans.into_iter().flatten() {
        let added = match plan.value {
            Some(value) => escaped_attribute_len(value)?,
            None => 0,
        };
        if plan.start == plan.end {
            let attribute_len = 4usize
                .checked_add(plan.name.len())
                .and_then(|length| length.checked_add(added))
                .ok_or_else(|| invalid("server-format leaf length overflows"))?;
            length = length
                .checked_add(attribute_len)
                .ok_or_else(|| invalid("server-format leaf length overflows"))?;
        } else {
            length = length
                .checked_sub(plan.end.saturating_sub(plan.start))
                .and_then(|length| length.checked_add(added))
                .ok_or_else(|| invalid("server-format leaf length overflows"))?;
        }
    }
    Ok(length)
}

fn new_leaf_len(
    qname: &[u8],
    namespace_decl: Option<&[u8]>,
    value: &ServerFormat,
    maximum_attribute_bytes: usize,
    include_open: bool,
) -> Result<usize> {
    let mut length = usize::from(include_open)
        .checked_add(qname.len())
        .and_then(|length| length.checked_add(2))
        .ok_or_else(|| invalid("server-format leaf length overflows"))?;
    if let Some(namespace_decl) = namespace_decl {
        length = length
            .checked_add(namespace_decl.len())
            .ok_or_else(|| invalid("server-format namespace length overflows"))?;
    }
    for (name, value) in [
        (b"culture".as_slice(), value.culture.as_deref()),
        (b"format".as_slice(), value.format.as_deref()),
    ] {
        if let Some(value) = value {
            let encoded = checked_attribute_value_len(name, value, maximum_attribute_bytes)?;
            let attribute_len = 4usize
                .checked_add(name.len())
                .and_then(|length| length.checked_add(encoded))
                .ok_or_else(|| invalid("server-format leaf length overflows"))?;
            length = length
                .checked_add(attribute_len)
                .ok_or_else(|| invalid("server-format leaf length overflows"))?;
        }
    }
    Ok(length)
}

fn checked_attribute_value_len(
    name: &[u8],
    value: &str,
    maximum_attribute_bytes: usize,
) -> Result<usize> {
    let encoded = escaped_attribute_len(value)?;
    if name
        .len()
        .checked_add(encoded)
        .is_none_or(|length| length > maximum_attribute_bytes)
    {
        return Err(invalid("server-format attribute exceeds its text limit"));
    }
    Ok(encoded)
}

fn leaf_attribute_plans<'a>(
    source: &[u8],
    entry: &EntrySource,
    previous: &ServerFormat,
    next: &'a ServerFormat,
    maximum_attribute_bytes: usize,
) -> Result<[Option<LeafAttributePlan<'a>>; 2]> {
    let mut plans = [
        attribute_plan(
            source,
            entry.start_tag.clone(),
            b"culture",
            entry.culture.as_ref(),
            previous.culture.as_ref(),
            next.culture.as_deref(),
            maximum_attribute_bytes,
        )?,
        attribute_plan(
            source,
            entry.start_tag.clone(),
            b"format",
            entry.format.as_ref(),
            previous.format.as_ref(),
            next.format.as_deref(),
            maximum_attribute_bytes,
        )?,
    ];
    if let (Some(left), Some(right)) = (plans[0], plans[1])
        && right.start < left.start
    {
        plans.swap(0, 1);
    }
    Ok(plans)
}

fn attribute_plan<'a>(
    source: &[u8],
    start_tag: Range<usize>,
    name: &'static [u8],
    current: Option<&AttributeSource>,
    previous: Option<&String>,
    next: Option<&'a str>,
    maximum_attribute_bytes: usize,
) -> Result<Option<LeafAttributePlan<'a>>> {
    if previous.map(String::as_str) == next {
        return Ok(None);
    }
    if let Some(value) = next {
        checked_attribute_value_len(name, value, maximum_attribute_bytes)?;
    }
    let plan = match (current, next) {
        (Some(current), Some(value)) => LeafAttributePlan {
            start: current.value.start,
            end: current.value.end,
            name,
            value: Some(value),
        },
        (Some(current), None) => LeafAttributePlan {
            start: current.whole.start,
            end: current.whole.end,
            name,
            value: None,
        },
        (None, Some(value)) => {
            let insertion = if start_tag.end >= 2
                && source.get(start_tag.end - 2..start_tag.end) == Some(b"/>")
            {
                start_tag.end - 2
            } else {
                start_tag
                    .end
                    .checked_sub(1)
                    .ok_or_else(|| invalid("server-format start tag range underflow"))?
            };
            if source.get(insertion) != Some(&b'>')
                && source.get(insertion..insertion.saturating_add(2)) != Some(b"/>")
            {
                return Err(invalid(
                    "server-format attribute insertion point is ambiguous",
                ));
            }
            LeafAttributePlan {
                start: insertion,
                end: insertion,
                name,
                value: Some(value),
            }
        },
        (None, None) => return Ok(None),
    };
    Ok(Some(plan))
}

#[derive(Clone, Copy)]
struct DecimalBytes {
    bytes: [u8; 10],
    len: usize,
}

fn decimal_u32(value: u32) -> DecimalBytes {
    let mut bytes = [0; 10];
    let mut value = value;
    let mut len = 0usize;
    loop {
        bytes[9 - len] = b'0' + (value % 10) as u8;
        len += 1;
        value /= 10;
        if value == 0 {
            break;
        }
    }
    bytes.copy_within(10 - len..10, 0);
    DecimalBytes { bytes, len }
}

fn append_structural_table(
    output: &mut Vec<u8>,
    source: &[u8],
    before: &Snapshot,
    staged: &[ServerFormat],
    staged_sources: &[Option<usize>],
    old_to_new: &[usize],
    count_changed: bool,
    count: &DecimalBytes,
    maximum_attribute_bytes: usize,
) -> Result<()> {
    let owner = &before.extension_owner;
    let mut cursor = 0usize;
    let mut owner_written = false;
    for reference in &before.index_refs {
        if !owner_written && owner.start <= reference.value.start {
            let mut ignored_count = false;
            append_source_segment(
                output,
                source,
                cursor..owner.start,
                None,
                &mut ignored_count,
            )?;
            append_structural_owner(
                output,
                source,
                before,
                staged,
                staged_sources,
                count_changed,
                count,
                maximum_attribute_bytes,
            )?;
            cursor = owner.end;
            owner_written = true;
        }
        if reference.value.start < cursor
            || reference.value.end < reference.value.start
            || (reference.value.start < owner.end && reference.value.end > owner.start)
        {
            return Err(invalid(
                "pivotValueCellExtra@in overlaps the server-format owner",
            ));
        }
        let mut ignored_count = false;
        append_source_segment(
            output,
            source,
            cursor..reference.value.start,
            None,
            &mut ignored_count,
        )?;
        let new_index = *old_to_new
            .get(reference.index)
            .ok_or_else(|| invalid("pivotValueCellExtra@in source index is out of range"))?;
        if new_index == usize::MAX {
            return Err(invalid(
                "removing a server-format leaf would orphan pivotValueCellExtra@in",
            ));
        }
        if new_index == reference.index {
            output.extend_from_slice(&source[reference.value.clone()]);
        } else {
            let digits = decimal_u32(
                u32::try_from(new_index)
                    .map_err(|_| invalid("pivotValueCellExtra@in does not fit unsignedInt"))?,
            );
            output.extend_from_slice(&digits.bytes[..digits.len]);
        }
        cursor = reference.value.end;
    }
    if !owner_written {
        let mut ignored_count = false;
        append_source_segment(
            output,
            source,
            cursor..owner.start,
            None,
            &mut ignored_count,
        )?;
        append_structural_owner(
            output,
            source,
            before,
            staged,
            staged_sources,
            count_changed,
            count,
            maximum_attribute_bytes,
        )?;
        cursor = owner.end;
    }
    let mut ignored_count = false;
    append_source_segment(
        output,
        source,
        cursor..source.len(),
        None,
        &mut ignored_count,
    )
}

fn append_structural_owner(
    output: &mut Vec<u8>,
    source: &[u8],
    before: &Snapshot,
    staged: &[ServerFormat],
    staged_sources: &[Option<usize>],
    count_changed: bool,
    count: &DecimalBytes,
    maximum_attribute_bytes: usize,
) -> Result<()> {
    let mut cursor = before.extension_owner.start;
    let mut count_written = false;
    for old_index in 0..before.entries.len() {
        let entry = &before.entries[old_index];
        append_source_segment(
            output,
            source,
            cursor..entry.element.start,
            count_changed.then_some((before.count.value.clone(), count)),
            &mut count_written,
        )?;
        if old_index < staged.len() {
            append_staged_child(
                output,
                source,
                before,
                staged,
                staged_sources,
                old_index,
                maximum_attribute_bytes,
            )?;
        }
        cursor = entry.element.end;
    }
    for new_index in before.entries.len()..staged.len() {
        append_staged_child(
            output,
            source,
            before,
            staged,
            staged_sources,
            new_index,
            maximum_attribute_bytes,
        )?;
    }
    append_source_segment(
        output,
        source,
        cursor..before.extension_owner.end,
        count_changed.then_some((before.count.value.clone(), count)),
        &mut count_written,
    )?;
    if count_changed && !count_written {
        return Err(invalid(
            "pivotTableServerFormats count source range is not in its owner",
        ));
    }
    Ok(())
}

fn append_staged_child(
    output: &mut Vec<u8>,
    source: &[u8],
    before: &Snapshot,
    staged: &[ServerFormat],
    staged_sources: &[Option<usize>],
    new_index: usize,
    maximum_attribute_bytes: usize,
) -> Result<()> {
    let Some(source_index) = staged_sources[new_index] else {
        let template = &before.entries[0];
        let include_open = target_leaf_includes_open(source, before, new_index);
        return append_new_leaf(
            output,
            &template.qname,
            template.namespace_decl.as_deref(),
            &staged[new_index],
            maximum_attribute_bytes,
            include_open,
        );
    };
    let include_open = target_leaf_includes_open(source, before, new_index);
    append_existing_leaf(
        output,
        source,
        &before.entries[source_index],
        &before.formats()[source_index],
        &staged[new_index],
        maximum_attribute_bytes,
        include_open,
    )
}

fn append_existing_leaf(
    output: &mut Vec<u8>,
    source: &[u8],
    entry: &EntrySource,
    previous: &ServerFormat,
    next: &ServerFormat,
    maximum_attribute_bytes: usize,
    include_open: bool,
) -> Result<()> {
    let plans = leaf_attribute_plans(source, entry, previous, next, maximum_attribute_bytes)?;
    let source_has_open = source.get(entry.element.start) == Some(&b'<');
    let source_start = if source_has_open && !include_open {
        entry
            .element
            .start
            .checked_add(1)
            .ok_or_else(|| invalid("server-format leaf source range overflows"))?
    } else {
        entry.element.start
    };
    if include_open && !source_has_open {
        output.push(b'<');
    }
    let mut cursor = source_start;
    for plan in plans.into_iter().flatten() {
        if plan.start < cursor || plan.end < plan.start || plan.end > entry.element.end {
            return Err(invalid("server-format attribute source range is invalid"));
        }
        output.extend_from_slice(&source[cursor..plan.start]);
        if let Some(value) = plan.value {
            if plan.start == plan.end {
                output.push(b' ');
                output.extend_from_slice(plan.name);
                output.extend_from_slice(b"=\"");
                append_encoded_xstring(output, value)?;
                output.push(b'"');
            } else {
                append_encoded_xstring(output, value)?;
            }
        }
        cursor = if plan.start == plan.end {
            plan.start
        } else {
            plan.end
        };
    }
    output.extend_from_slice(&source[cursor..entry.element.end]);
    Ok(())
}

fn append_new_leaf(
    output: &mut Vec<u8>,
    qname: &[u8],
    namespace_decl: Option<&[u8]>,
    value: &ServerFormat,
    maximum_attribute_bytes: usize,
    include_open: bool,
) -> Result<()> {
    if include_open {
        output.push(b'<');
    }
    output.extend_from_slice(qname);
    if let Some(namespace_decl) = namespace_decl {
        output.extend_from_slice(namespace_decl);
    }
    for (name, value) in [
        (b"culture".as_slice(), value.culture.as_deref()),
        (b"format".as_slice(), value.format.as_deref()),
    ] {
        if let Some(value) = value {
            checked_attribute_value_len(name, value, maximum_attribute_bytes)?;
            output.push(b' ');
            output.extend_from_slice(name);
            output.extend_from_slice(b"=\"");
            append_encoded_xstring(output, value)?;
            output.push(b'"');
        }
    }
    output.extend_from_slice(b"/>");
    Ok(())
}

fn append_encoded_xstring(output: &mut Vec<u8>, value: &str) -> Result<()> {
    let encoded_len = escaped_xstring_len(value)?;
    let required = output
        .len()
        .checked_add(encoded_len)
        .ok_or_else(|| invalid("escaped server-format attribute length overflows"))?;
    if required > output.capacity() {
        return Err(invalid(
            "encoded server-format attribute exceeds its preflight output capacity",
        ));
    }
    let previous_len = output.len();
    append_escaped_xstring(output, value);
    if output.len() != required || output.len() - previous_len != encoded_len {
        return Err(invalid(
            "encoded server-format attribute length disagrees with its preflight",
        ));
    }
    Ok(())
}

fn append_source_segment(
    output: &mut Vec<u8>,
    source: &[u8],
    range: Range<usize>,
    count: Option<(Range<usize>, &DecimalBytes)>,
    count_written: &mut bool,
) -> Result<()> {
    if range.start > range.end || range.end > source.len() {
        return Err(invalid("PivotTable source segment is invalid"));
    }
    if let Some((count_range, count_value)) = count {
        if count_range.start < range.end && count_range.end > range.start {
            if *count_written || count_range.start < range.start || count_range.end > range.end {
                return Err(invalid("PivotTable count source range overlaps its owner"));
            }
            output.extend_from_slice(&source[range.start..count_range.start]);
            output.extend_from_slice(&count_value.bytes[..count_value.len]);
            output.extend_from_slice(&source[count_range.end..range.end]);
            *count_written = true;
            return Ok(());
        }
    }
    output.extend_from_slice(&source[range]);
    Ok(())
}

fn apply_planned_attribute_delta(
    source: &[u8],
    start_tag: &Range<usize>,
    name: &[u8],
    current: Option<&AttributeSource>,
    previous: Option<&String>,
    next: Option<&str>,
    output_len: &mut usize,
    owner_len: &mut usize,
    maximum_attribute_bytes: usize,
) -> Result<bool> {
    if previous.map(String::as_str) == next {
        return Ok(false);
    }
    let (removed, added) = match (current, next) {
        (Some(current), Some(next)) => (
            current.value.end.saturating_sub(current.value.start),
            escaped_attribute_len(next)?,
        ),
        (Some(current), None) => (current.whole.end.saturating_sub(current.whole.start), 0),
        (None, Some(next)) => {
            let end = start_tag.end;
            let insertion = if end >= 2 && source.get(end - 2..end) == Some(b"/>") {
                end - 2
            } else {
                end.checked_sub(1)
                    .ok_or_else(|| invalid("server-format start tag range underflow"))?
            };
            if source.get(insertion) != Some(&b'>')
                && source.get(insertion..insertion.saturating_add(2)) != Some(b"/>")
            {
                return Err(invalid(
                    "server-format attribute insertion point is ambiguous",
                ));
            }
            (
                0,
                name.len()
                    .saturating_add(escaped_attribute_len(next)?)
                    .saturating_add(4),
            )
        },
        (None, None) => return Ok(false),
    };
    if let Some(next) = next {
        let encoded_len = escaped_attribute_len(next)?;
        let attribute_bytes = name
            .len()
            .checked_add(encoded_len)
            .ok_or_else(|| invalid("server-format attribute length overflows"))?;
        if attribute_bytes > maximum_attribute_bytes {
            return Err(invalid("server-format attribute exceeds its text limit"));
        }
    }
    *output_len = output_len
        .checked_sub(removed)
        .and_then(|length| length.checked_add(added))
        .ok_or_else(|| invalid("server-format output length overflows"))?;
    *owner_len = owner_len
        .checked_sub(removed)
        .and_then(|length| length.checked_add(added))
        .ok_or_else(|| invalid("server-format owner length overflows"))?;
    Ok(true)
}

fn escaped_attribute_len(value: &str) -> Result<usize> {
    escaped_xstring_len(value)
}

// The generic attribute helper above intentionally refuses insertion because
// the attribute name is part of the owner contract.  This named wrapper keeps
// that decision explicit and avoids guessing from a missing semantic value.
fn push_named_attribute_edit(
    source: &[u8],
    start_tag: &Range<usize>,
    name: &[u8],
    current: Option<&AttributeSource>,
    previous: Option<&String>,
    next: Option<&str>,
    edits: &mut Vec<(Range<usize>, Vec<u8>)>,
) -> Result<()> {
    if previous.map(String::as_str) == next {
        return Ok(());
    }
    let encoded = next.map(try_escaped_xstring).transpose()?;
    match (current, encoded) {
        (Some(current), Some(encoded)) => {
            edits.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "PivotTable server-format edit list",
                source,
            })?;
            edits.push((current.value.clone(), encoded));
        },
        (Some(current), None) => {
            edits.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "PivotTable server-format edit list",
                source,
            })?;
            edits.push((current.whole.clone(), Vec::new()));
        },
        (None, Some(encoded)) => {
            let end = start_tag.end;
            let insertion = if end >= 2 && source.get(end - 2..end) == Some(b"/>") {
                end - 2
            } else {
                end.checked_sub(1)
                    .ok_or_else(|| invalid("server-format start tag range underflow"))?
            };
            if source.get(insertion) != Some(&b'>')
                && source.get(insertion..insertion.saturating_add(2)) != Some(b"/>")
            {
                return Err(invalid(
                    "server-format attribute insertion point is ambiguous",
                ));
            }
            let name = name.strip_prefix(b"@").unwrap_or(name);
            let mut replacement = Vec::new();
            replacement
                .try_reserve_exact(name.len().saturating_add(encoded.len()).saturating_add(4))
                .map_err(|source| Error::Allocation {
                    resource: "PivotTable server-format attribute",
                    source,
                })?;
            replacement.extend_from_slice(b" ");
            replacement.extend_from_slice(name);
            replacement.extend_from_slice(b"=\"");
            replacement.extend_from_slice(&encoded);
            replacement.extend_from_slice(b"\"");
            edits.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "PivotTable server-format edit list",
                source,
            })?;
            edits.push((insertion..insertion, replacement));
        },
        (None, None) => {},
    }
    Ok(())
}

fn load_connections(
    package: &OpcPackage,
    workbook: &dyn Part,
    relationship_index: &RelationshipIndex<'_>,
    need_model_ids: bool,
    allow_empty_names: bool,
) -> Result<Option<(SourcePart, ConnectionCatalog)>> {
    let mut matches = workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            CONNECTIONS_RELATIONSHIP | STRICT_CONNECTIONS_RELATIONSHIP
        )
    });
    let Some(relationship) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(invalid("workbook has multiple connections relationships"));
    }
    if relationship.is_external()
        || relationship.target_query().is_some()
        || relationship.target_fragment().is_some()
    {
        return Err(invalid(
            "workbook connections relationship must target an internal Part",
        ));
    }
    let uri = relationship.target_partname()?;
    let canonical_uri = relationship_index
        .canonical_part(&uri)?
        .ok_or_else(|| invalid("workbook connections target is not a package Part"))?;
    let part = package.get_part(canonical_uri)?;
    if part.content_type() != CONNECTIONS_CONTENT_TYPE {
        return Err(invalid(
            "workbook connections relationship targets the wrong content type",
        ));
    }
    if part.blob().len() > MAX_PART_BYTES
        || part.blob().len() > caller_part_limit(package.read_limits())
    {
        return Err(invalid("connections Part exceeds the caller's Part limit"));
    }
    let scan = scan_xml(part.blob(), "connections", package.read_limits())?;
    let root = scan
        .elements
        .iter()
        .find(|element| element.parent_index.is_none())
        .ok_or_else(|| invalid("connections Part has no root"))?;
    if root.ns.as_ref() != CORE_NS && root.ns.as_ref() != STRICT_CORE_NS
        || root.local != b"connections"
    {
        return Err(invalid("connections Part has an invalid root"));
    }
    let mut by_id = HashMap::new();
    let connection_count = scan
        .elements
        .iter()
        .filter(|element| {
            element.parent_index == Some(root.index)
                && element.ns == root.ns
                && element.local == b"connection"
        })
        .count();
    by_id
        .try_reserve(connection_count)
        .map_err(|source| Error::Allocation {
            resource: "workbook connection ID index",
            source,
        })?;
    let mut by_name = HashMap::new();
    by_name
        .try_reserve(connection_count)
        .map_err(|source| Error::Allocation {
            resource: "workbook connection name index",
            source,
        })?;
    let model_candidates = if need_model_ids {
        Some(scan_model_connection_candidates(
            &scan,
            root,
            connection_count,
        )?)
    } else {
        None
    };
    let mut model_ids = HashSet::new();
    if need_model_ids {
        model_ids
            .try_reserve(connection_count)
            .map_err(|source| Error::Allocation {
                resource: "workbook model connection index",
                source,
            })?;
    }
    for element in scan.elements.iter().filter(|element| {
        element.parent_index == Some(root.index)
            && element.ns == root.ns
            && element.local == b"connection"
    }) {
        let id = unique_unqualified_attr(element, b"id", "connection")?
            .map(|value| parse_u32(value, "connection id"))
            .transpose()?
            .ok_or_else(|| invalid("connection requires id"))?;
        if by_id.insert(id, String::new()).is_some() {
            return Err(invalid("connections contains a duplicate id"));
        }
        if let Some(name_attr) = unique_unqualified_attr(element, b"name", "connection")? {
            let name = decode_spreadsheet_text(name_attr)?;
            if name.is_empty() && !allow_empty_names {
                return Err(invalid("connection name cannot be empty"));
            }
            if by_name.insert(name.clone(), id).is_some() {
                return Err(invalid("connections contains a duplicate name"));
            }
            if let Some(stored) = by_id.get_mut(&id) {
                *stored = name;
            }
        }
        if need_model_ids
            && is_model_connection(
                &scan,
                root,
                element,
                model_candidates
                    .as_ref()
                    .and_then(|candidates| candidates.get(&element.index)),
            )?
        {
            model_ids.insert(id);
        }
    }
    if by_id.is_empty() {
        return Err(invalid("connections Part has no connection entries"));
    }
    Ok(Some((
        SourcePart::capture(relationship_index, part)?,
        ConnectionCatalog {
            by_id,
            by_name,
            model_ids,
        },
    )))
}

#[derive(Default)]
struct ModelConnectionCandidate {
    extension_count: u8,
    payload_count: u8,
    extension_index: Option<usize>,
    payload_index: Option<usize>,
}

/// Index the exact DE250 extension in one bounded pass over the already parsed
/// connections Part.  The C510 owner does not request this index, so malformed
/// optional model extensions cannot affect its ordinary connection scan.
fn scan_model_connection_candidates(
    scan: &XmlScan,
    root: &XmlElement,
    connection_count: usize,
) -> Result<HashMap<usize, ModelConnectionCandidate>> {
    let mut candidates: HashMap<usize, ModelConnectionCandidate> = HashMap::new();
    candidates
        .try_reserve(connection_count)
        .map_err(|source| Error::Allocation {
            resource: "workbook model connection candidate index",
            source,
        })?;
    for element in &scan.elements {
        if let Some(connection_index) = model_connection_owner(scan, root, element, false)
            && element.ns == root.ns
            && element.local == b"ext"
        {
            check_fragment_limit(element, "DE250 model extension fragment")?;
            let candidate = candidates.entry(connection_index).or_default();
            candidate.extension_count = candidate.extension_count.saturating_add(1);
            if candidate.extension_count == 1 {
                candidate.extension_index = Some(element.index);
            }
        }
        if let Some(connection_index) = model_connection_owner(scan, root, element, true) {
            check_fragment_limit(element, "DE250 model payload fragment")?;
            let candidate = candidates.entry(connection_index).or_default();
            candidate.payload_count = candidate.payload_count.saturating_add(1);
            if candidate.payload_count == 1 {
                candidate.payload_index = Some(element.index);
            }
        }
    }
    Ok(candidates)
}

/// Resolve a direct matching extension or payload to its owning root
/// connection without walking the complete scan.  Parent indices are the
/// parser's retained ancestry, so this is constant work per XML element.
fn model_connection_owner(
    scan: &XmlScan,
    root: &XmlElement,
    element: &XmlElement,
    payload: bool,
) -> Option<usize> {
    let parent_index = element.parent_index?;
    let parent = scan.elements.get(parent_index)?;
    let extension = if payload {
        if element.ns.as_ref() != EXT_NS || element.local != b"connection" {
            return None;
        }
        parent
    } else {
        if element.ns != root.ns || element.local != b"ext" {
            return None;
        }
        element
    };
    let ext_list = scan.elements.get(extension.parent_index?)?;
    let connection = scan.elements.get(ext_list.parent_index?)?;
    if ext_list.ns != root.ns
        || ext_list.local != b"extLst"
        || connection.parent_index != Some(root.index)
        || connection.ns != root.ns
        || connection.local != b"connection"
        || extension.ns != root.ns
        || extension.local != b"ext"
        || !candidate_attr_once(extension, b"uri")
            .is_some_and(|value| xml_token_eq(value, CONNECTION_MODEL_URI))
    {
        return None;
    }
    Some(connection.index)
}

/// Recognize one exact DE250 model-connection candidate after the one-pass
/// ancestry index has identified its unique extension and payload.
fn is_model_connection(
    scan: &XmlScan,
    _root: &XmlElement,
    connection: &XmlElement,
    candidate: Option<&ModelConnectionCandidate>,
) -> Result<bool> {
    let Some(candidate) = candidate else {
        return Ok(false);
    };
    if candidate.extension_count != 1 || candidate.payload_count != 1 {
        return Ok(false);
    }
    let Some(extension_index) = candidate.extension_index else {
        return Ok(false);
    };
    let Some(extension) = scan.elements.get(extension_index) else {
        return Ok(false);
    };
    let Some(payload_index) = candidate.payload_index else {
        return Ok(false);
    };
    let Some(payload) = scan.elements.get(payload_index) else {
        return Ok(false);
    };
    if connection.mce_context || extension.mce_context || payload.mce_context {
        return Ok(false);
    }
    if payload.has_element_child || payload.has_text || payload.has_cdata {
        return Ok(false);
    }
    let model = candidate_attr_once(payload, b"model")
        .map(|value| parse_bool(value, "connection model"))
        .transpose()
        .ok()
        .flatten()
        .unwrap_or(false);
    if !model {
        return Ok(false);
    }
    let Some(connection_type) = candidate_attr_once(connection, b"type")
        .map(|value| parse_u32(value, "connection type"))
        .transpose()
        .ok()
        .flatten()
    else {
        return Ok(false);
    };
    if connection_type != 5 {
        return Ok(false);
    }
    let Some(id) = candidate_attr_once(payload, b"id")
        .map(decode_spreadsheet_text)
        .transpose()
        .ok()
        .flatten()
    else {
        return Ok(false);
    };
    if !id.is_empty() {
        return Ok(false);
    }
    Ok(true)
}

fn candidate_attr_once<'a>(element: &'a XmlElement, local: &[u8]) -> Option<&'a str> {
    let mut found = None;
    for attribute in &element.attrs {
        if attribute.ns.is_empty() && attribute.local == local {
            if found.is_some() {
                return None;
            }
            found = Some(attribute.value.as_str());
        }
    }
    found
}

fn unique_unqualified_attr<'a>(
    element: &'a XmlElement,
    local: &[u8],
    owner: &str,
) -> Result<Option<&'a str>> {
    let mut found = None;
    for attribute in &element.attrs {
        if attribute.ns.is_empty() && attribute.local == local {
            if found.is_some() {
                return Err(invalid(format!("{owner} has a duplicate attribute")));
            }
            found = Some(attribute.value.as_str());
        }
    }
    Ok(found)
}

fn validate_table_incoming_closure(
    package: &OpcPackage,
    workbook: &dyn Part,
    references: &[Reference],
    relationship_index: &RelationshipIndex<'_>,
) -> Result<()> {
    for reference in references {
        if relationship_index
            .canonical_part(&reference.table_uri)?
            .is_none()
        {
            return Err(invalid(
                "PivotTable incoming relationship target is not a package Part",
            ));
        }
        let incoming = relationship_index.incoming(&reference.table_uri)?;
        let mut workbook_pivot_count = 0usize;
        for incoming in incoming {
            let relationship = incoming.relationship();
            if incoming.source() == workbook.partname().as_str() {
                if relationship.target_query().is_some() || relationship.target_fragment().is_some()
                {
                    return Err(invalid(
                        "PivotTable relationship cannot target a URI suffix",
                    ));
                }
                if matches!(
                    relationship.reltype(),
                    rt::PIVOT_TABLE | rt::STRICT_PIVOT_TABLE
                ) {
                    workbook_pivot_count =
                        workbook_pivot_count.checked_add(1).ok_or_else(|| {
                            invalid("PivotTable incoming relationship count overflows")
                        })?;
                } else {
                    return Err(invalid(
                        "PivotTable has an unexpected incoming workbook relationship",
                    ));
                }
                continue;
            }
            if incoming.source() == "/" {
                return Err(invalid(
                    "PivotTable cannot have an incoming package-root relationship",
                ));
            }
            let source = PackURI::new(incoming.source()).map_err(|error| {
                invalid(format!("incoming relationship source is invalid: {error}"))
            })?;
            let canonical_source = relationship_index
                .canonical_part(&source)?
                .ok_or_else(|| invalid("incoming relationship source is not a package Part"))?;
            let source_part = package.get_part(canonical_source)?;
            if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
                return Err(invalid(
                    "PivotTable relationship cannot target a URI suffix",
                ));
            }
            if source_part.content_type() == ct::SML_WORKSHEET
                || matches!(
                    relationship.reltype(),
                    rt::PIVOT_TABLE | rt::STRICT_PIVOT_TABLE
                )
            {
                return Err(invalid(
                    "Non-Worksheet PivotTable cannot have an incoming worksheet relationship",
                ));
            }
            return Err(invalid(
                "PivotTable has an unexpected incoming Part relationship",
            ));
        }
        if workbook_pivot_count != 1 {
            return Err(invalid(
                "PivotTable must have exactly one incoming workbook PivotTable relationship",
            ));
        }
    }
    Ok(())
}

fn parse_workbook_references(
    package: &OpcPackage,
    workbook: &dyn Part,
    relationship_index: &RelationshipIndex<'_>,
    parse_server_formats: bool,
    parse_mce_branches: bool,
) -> Result<WorkbookReferences> {
    let scan = if parse_mce_branches {
        scan_xml_with_mce(workbook.blob(), "workbook", package.read_limits())?
    } else {
        scan_xml(workbook.blob(), "workbook", package.read_limits())?
    };
    let root = scan
        .elements
        .iter()
        .find(|element| element.parent_index.is_none())
        .ok_or_else(|| invalid("workbook has no root"))?;
    let owner_exts = extension_exts(&scan, root, PIVOT_TABLE_REFERENCES_URI, parse_mce_branches)?;
    for candidate in &owner_exts {
        check_fragment_limit(candidate, "workbook pivotTableReferences owner fragment")?;
    }
    let mut diagnostic = false;
    if owner_exts.len() > 1 {
        if parse_server_formats {
            return Err(invalid("duplicate workbook pivotTableReferences extension"));
        }
        diagnostic = true;
    }
    let Some(ext) = owner_exts.first().copied() else {
        return Err(invalid("workbook has no pivotTableReferences extension"));
    };
    for attribute in &ext.attrs {
        if !attribute.ns.is_empty() || attribute.local.as_slice() != b"uri" {
            if parse_server_formats {
                return Err(invalid(
                    "pivotTableReferences extension has an unknown attribute",
                ));
            }
            diagnostic = true;
        }
    }
    for child in scan
        .elements
        .iter()
        .filter(|candidate| candidate.parent_index == Some(ext.index))
    {
        let allowed_payload =
            child.ns.as_ref() == EXT_NS && child.local.as_slice() == b"pivotTableReferences";
        let allowed_mce =
            child.ns.as_ref() == MCE_NS && child.local.as_slice() == b"AlternateContent";
        if !allowed_payload && !allowed_mce {
            if parse_server_formats {
                return Err(invalid(
                    "pivotTableReferences extension has an unexpected child",
                ));
            }
            diagnostic = true;
        }
    }
    let payloads = scan.elements.iter().filter(|candidate| {
        candidate.ns.as_ref() == EXT_NS
            && candidate.local == b"pivotTableReferences"
            && is_owned_payload(&scan, root, candidate, PIVOT_TABLE_REFERENCES_URI)
    });
    let mut references = None;
    let mut duplicate_payload = false;
    for candidate in payloads {
        check_fragment_limit(candidate, "workbook pivotTableReferences payload fragment")?;
        if references.is_some() {
            duplicate_payload = true;
        } else {
            references = Some(candidate);
        }
    }
    let references =
        references.ok_or_else(|| invalid("pivotTableReferences extension has no payload"))?;
    if duplicate_payload {
        if parse_server_formats {
            return Err(invalid("duplicate workbook pivotTableReferences payload"));
        }
        diagnostic = true;
    }
    if !references.attrs.is_empty() {
        if parse_server_formats {
            return Err(invalid(
                "pivotTableReferences payload has an unknown attribute",
            ));
        }
        diagnostic = true;
    }
    for child in scan
        .elements
        .iter()
        .filter(|candidate| candidate.parent_index == Some(references.index))
    {
        let allowed =
            child.ns.as_ref() == EXT_NS && child.local.as_slice() == b"pivotTableReference";
        if !allowed {
            if parse_server_formats {
                return Err(invalid(
                    "pivotTableReferences payload has an unexpected child",
                ));
            }
            diagnostic = true;
        }
    }
    let reference_count = scan
        .elements
        .iter()
        .filter(|candidate| {
            candidate.parent_index == Some(references.index)
                && candidate.ns.as_ref() == EXT_NS
                && candidate.local == b"pivotTableReference"
        })
        .count();
    if reference_count == 0 || reference_count > MAX_REFERENCE_COUNT {
        return Err(invalid(
            "pivotTableReferences collection cardinality is invalid",
        ));
    }
    let mut refs = Vec::new();
    refs.try_reserve_exact(reference_count)
        .map_err(|source| Error::Allocation {
            resource: "PivotTable reference collection",
            source,
        })?;
    let mut targets = HashSet::new();
    targets
        .try_reserve(reference_count)
        .map_err(|source| Error::Allocation {
            resource: "PivotTable reference targets",
            source,
        })?;
    for child in scan.elements.iter().filter(|candidate| {
        candidate.parent_index == Some(references.index)
            && candidate.ns.as_ref() == EXT_NS
            && candidate.local == b"pivotTableReference"
    }) {
        check_fragment_limit(child, "pivotTableReference fragment")?;
        if !validate_pivot_table_reference_shape(child, parse_server_formats)? {
            diagnostic = true;
        }
        let rid = required_rel_id(child, "pivotTableReference")?;
        let relation = workbook.rels().get(&rid).ok_or_else(|| {
            invalid("pivotTableReference r:id does not resolve in workbook relationships")
        })?;
        if relation.is_external()
            || relation.target_query().is_some()
            || relation.target_fragment().is_some()
            || !matches!(relation.reltype(), rt::PIVOT_TABLE | rt::STRICT_PIVOT_TABLE)
        {
            return Err(invalid(
                "pivotTableReference relationship is not an internal PivotTable relationship",
            ));
        }
        let uri = relation.target_partname()?;
        let canonical_uri = relationship_index
            .canonical_part(&uri)?
            .ok_or_else(|| invalid("pivotTableReference target is not a package Part"))?;
        if !targets.insert(canonical_uri.clone()) {
            if parse_server_formats {
                return Err(invalid(
                    "pivotTableReferences contains a duplicate PivotTable target",
                ));
            }
            diagnostic = true;
            continue;
        }
        let part = package.get_part(canonical_uri)?;
        if part.content_type() != ct::SML_PIVOT_TABLE {
            return Err(invalid(
                "pivotTableReference targets the wrong content type",
            ));
        }
        if part.blob().len() > MAX_PART_BYTES
            || part.blob().len() > caller_part_limit(package.read_limits())
        {
            return Err(invalid("PivotTable Part exceeds the caller's Part limit"));
        }
        let table = parse_table(package, part, relationship_index, parse_server_formats)?;
        refs.push(Reference {
            table_uri: canonical_uri.clone(),
            name: table.name.clone(),
            cache_id: table.cache_id,
            table,
        });
    }
    if refs.is_empty() || refs.len() > MAX_REFERENCE_COUNT {
        return Err(invalid(
            "pivotTableReferences collection cardinality is invalid",
        ));
    }
    let owner_range = ext.start.start..ext.end;
    let owner_source = workbook.blob();
    let root_context = owner_source
        .get(root.start.clone())
        .ok_or_else(|| invalid("workbook root context range is invalid"))?;
    if root_context.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("workbook root context exceeds limit"));
    }
    let mut context_copy = Vec::new();
    context_copy
        .try_reserve_exact(
            root_context
                .len()
                .saturating_add(root.ns.len())
                .saturating_add(2),
        )
        .map_err(|source| Error::Allocation {
            resource: "workbook root context",
            source,
        })?;
    context_copy.extend_from_slice(&root.ns);
    context_copy.push(0);
    context_copy.push(u8::from(root.mce_context));
    context_copy.extend_from_slice(root_context);
    let owner_bytes = owner_source
        .get(owner_range)
        .ok_or_else(|| invalid("workbook pivotTableReferences owner range is invalid"))?;
    if owner_bytes.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid(
            "workbook pivotTableReferences owner fragment exceeds limit",
        ));
    }
    let mut owner_copy = Vec::new();
    owner_copy
        .try_reserve_exact(owner_bytes.len())
        .map_err(|source| Error::Allocation {
            resource: "workbook pivotTableReferences owner source",
            source,
        })?;
    owner_copy.extend_from_slice(owner_bytes);
    let mce_ambiguous = scan.elements.iter().any(|element| {
        element.mce_context
            && element.ns.as_ref() == EXT_NS
            && element.local.as_slice() == b"pivotTableReferences"
    });
    Ok(WorkbookReferences {
        refs,
        owner: Arc::new(owner_copy),
        context: Arc::new(context_copy),
        mce_ambiguous,
        diagnostic,
    })
}

fn direct_exts<'a>(
    scan: &'a XmlScan,
    owner: &XmlElement,
    uri: &str,
) -> Result<Vec<&'a XmlElement>> {
    extension_exts(scan, owner, uri, false)
}

fn extension_exts<'a>(
    scan: &'a XmlScan,
    owner: &XmlElement,
    uri: &str,
    allow_mce: bool,
) -> Result<Vec<&'a XmlElement>> {
    let mut result = Vec::new();
    for ext in scan.elements.iter().filter(|ext| {
        ext.ns == owner.ns
            && ext.local == b"ext"
            && candidate_attr_once(ext, b"uri").is_some_and(|value| xml_token_eq(value, uri))
            && extension_owner_matches(scan, owner.index, ext, allow_mce)
    }) {
        result.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "PivotTable extension owner index",
            source,
        })?;
        result.push(ext);
    }
    Ok(result)
}

fn check_fragment_limit(element: &XmlElement, owner: &str) -> Result<()> {
    let observed = element.end.saturating_sub(element.start.start);
    if observed > MAX_FRAGMENT_BYTES {
        return Err(Error::ResourceLimit(ResourceLimit {
            resource: Resource::Memory,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            limit: u64::try_from(MAX_FRAGMENT_BYTES).unwrap_or(u64::MAX),
            scope: Arc::from(owner),
        }));
    }
    Ok(())
}

fn extension_owner_matches(
    scan: &XmlScan,
    owner_index: usize,
    ext: &XmlElement,
    allow_mce: bool,
) -> bool {
    let mut current = ext.parent_index;
    let mut through_mce = false;
    while let Some(index) = current {
        let Some(element) = scan.elements.get(index) else {
            return false;
        };
        if element.local == b"extLst" && element.ns == ext.ns {
            let mut parent = element.parent_index;
            while let Some(index) = parent {
                if index == owner_index {
                    return allow_mce || !through_mce;
                }
                let Some(element) = scan.elements.get(index) else {
                    return false;
                };
                if !allow_mce || element.ns.as_ref() != MCE_NS {
                    return false;
                }
                parent = element.parent_index;
            }
            return false;
        }
        if !allow_mce || element.ns.as_ref() != MCE_NS {
            return false;
        }
        through_mce = true;
        current = element.parent_index;
    }
    false
}

fn is_owned_payload(scan: &XmlScan, root: &XmlElement, candidate: &XmlElement, uri: &str) -> bool {
    let mut current = candidate.parent_index;
    let mut through_mce = false;
    while let Some(index) = current {
        let Some(parent) = scan.elements.get(index) else {
            return false;
        };
        if parent.local == b"ext" && parent.ns == root.ns {
            return extension_owner_matches(scan, root.index, parent, true)
                && candidate_attr_once(parent, b"uri")
                    .is_some_and(|value| xml_token_eq(value, uri))
                && (through_mce || candidate.parent_index == Some(parent.index));
        }
        if parent.ns.as_ref() != MCE_NS {
            return false;
        }
        through_mce = true;
        current = parent.parent_index;
    }
    false
}

fn is_owned_cache_source_payload(
    scan: &XmlScan,
    root: &XmlElement,
    source: &XmlElement,
    candidate: &XmlElement,
) -> bool {
    let mut current = candidate.parent_index;
    let mut through_mce = false;
    while let Some(index) = current {
        let Some(parent) = scan.elements.get(index) else {
            return false;
        };
        if parent.local == b"ext" && parent.ns == root.ns {
            return extension_owner_matches(scan, source.index, parent, true)
                && candidate_attr_once(parent, b"uri")
                    .is_some_and(|value| xml_token_eq(value, CACHE_SOURCE_URI))
                && (through_mce || candidate.parent_index == Some(parent.index));
        }
        if parent.ns.as_ref() != MCE_NS {
            return false;
        }
        through_mce = true;
        current = parent.parent_index;
    }
    false
}

/// Validate the bounded shape of a recognized extension container.  Unknown
/// extension siblings remain opaque, but a recognized owner cannot silently
/// absorb an extra direct child or attribute.  C444 reports a bad closure as
/// a diagnostic so the source can still be read; the older C510 graph keeps
/// its historical strict refusal.
fn validate_known_extension_shape(
    scan: &XmlScan,
    ext: &XmlElement,
    payload_ns: &[u8],
    payload_local: &[u8],
    strict: bool,
) -> Result<bool> {
    let mut valid = true;
    for attribute in &ext.attrs {
        if !attribute.ns.is_empty() || attribute.local.as_slice() != b"uri" {
            if strict {
                return Err(invalid(
                    "recognized extension owner has an unknown attribute",
                ));
            }
            valid = false;
        }
    }
    for child in scan
        .elements
        .iter()
        .filter(|element| element.parent_index == Some(ext.index))
    {
        let allowed_payload = child.ns.as_ref() == payload_ns && child.local == payload_local;
        let allowed_mce = child.ns.as_ref() == MCE_NS && child.local == b"AlternateContent";
        if !allowed_payload && !allowed_mce {
            if strict {
                return Err(invalid(
                    "recognized extension owner has an unexpected child",
                ));
            }
            valid = false;
        }
    }
    Ok(valid)
}

#[cfg(test)]
mod tests;

fn required_rel_id(element: &XmlElement, owner: &str) -> Result<String> {
    let mut value = None;
    for attr in &element.attrs {
        if (attr.ns.as_ref() == REL_NS || attr.ns.as_ref() == STRICT_REL_NS) && attr.local == b"id"
        {
            if value.is_some() {
                return Err(invalid(format!(
                    "{owner} has multiple strict/transitional r:id attributes"
                )));
            }
            value = Some(attr.value.clone());
        }
    }
    value
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid(format!("{owner} requires r:id")))
}

fn validate_pivot_table_reference_shape(element: &XmlElement, strict: bool) -> Result<bool> {
    let mut valid = true;
    if element.has_cdata || element.has_text {
        if strict {
            return Err(invalid(
                "pivotTableReference must not contain non-whitespace text",
            ));
        }
        valid = false;
    }
    let mut relation_ids = 0usize;
    for attribute in &element.attrs {
        if attribute.local.as_slice() == b"id"
            && (attribute.ns.as_ref() == REL_NS || attribute.ns.as_ref() == STRICT_REL_NS)
        {
            relation_ids = relation_ids
                .checked_add(1)
                .ok_or_else(|| invalid("pivotTableReference r:id count overflows"))?;
            if relation_ids > 1 {
                if strict {
                    return Err(invalid(
                        "pivotTableReference has multiple strict/transitional r:id attributes",
                    ));
                }
                valid = false;
            }
            continue;
        }
        if strict {
            return Err(invalid("pivotTableReference has an unknown attribute"));
        }
        valid = false;
    }
    Ok(valid)
}

fn parse_table(
    package: &OpcPackage,
    part: &dyn Part,
    relationship_index: &RelationshipIndex<'_>,
    parse_server_formats: bool,
) -> Result<TableInfo> {
    let scan = scan_xml(part.blob(), "pivotTableDefinition", package.read_limits())?;
    let meta = scan_table_metadata_from_scan(&scan)?;
    let mut matching_cache = part.rels().iter().filter(|relation| {
        matches!(
            relation.reltype(),
            rt::PIVOT_CACHE_DEFINITION | rt::STRICT_PIVOT_CACHE_DEFINITION
        )
    });
    let relation = matching_cache
        .next()
        .ok_or_else(|| invalid("PivotTable is missing its cache-definition relationship"))?;
    if matching_cache.next().is_some()
        || relation.is_external()
        || relation.target_query().is_some()
        || relation.target_fragment().is_some()
    {
        return Err(invalid(
            "PivotTable cache-definition relationship is ambiguous",
        ));
    }
    let cache_uri = relation.target_partname()?;
    let canonical_cache_uri = relationship_index
        .canonical_part(&cache_uri)?
        .ok_or_else(|| invalid("PivotTable cache target is not a package Part"))?;
    let cache_part = package.get_part(canonical_cache_uri)?;
    if cache_part.content_type() != ct::SML_PIVOT_CACHE_DEFINITION {
        return Err(invalid(
            "PivotTable cache relationship targets the wrong content type",
        ));
    }
    let payload = if parse_server_formats {
        Some(parse_payload(&scan, part.blob())?)
    } else {
        None
    };
    Ok(TableInfo {
        name: meta.name,
        cache_id: meta.cache_id,
        cache_uri: canonical_cache_uri.clone(),
        source: SourcePart::capture(relationship_index, part)?,
        owner: meta.owner,
        payload,
    })
}

struct TableMeta {
    name: String,
    cache_id: u32,
    owner: Range<usize>,
}

fn scan_table_metadata(bytes: &[u8], limits: ReadLimits) -> Result<TableMeta> {
    let scan = scan_xml(bytes, "pivotTableDefinition", limits)?;
    scan_table_metadata_from_scan(&scan)
}

fn scan_table_metadata_from_scan(scan: &XmlScan) -> Result<TableMeta> {
    let root = scan
        .elements
        .iter()
        .find(|element| element.parent_index.is_none())
        .ok_or_else(|| invalid("PivotTable Part has no root"))?;
    if root.ns.as_ref() != CORE_NS && root.ns.as_ref() != STRICT_CORE_NS
        || root.local != b"pivotTableDefinition"
    {
        return Err(invalid("PivotTable Part has an invalid root"));
    }
    let name = unique_unqualified_attr(root, b"name", "PivotTable definition")?
        .map(decode_spreadsheet_text)
        .transpose()?
        .ok_or_else(|| invalid("PivotTable definition requires name"))?;
    if name.is_empty() {
        return Err(invalid("PivotTable name cannot be empty"));
    }
    let cache_id = unique_unqualified_attr(root, b"cacheId", "PivotTable definition")?
        .map(|value| parse_u32(value, "PivotTable cacheId"))
        .transpose()?
        .ok_or_else(|| invalid("PivotTable definition requires cacheId"))?;
    if let Some(enable_edit) =
        unique_unqualified_attr(root, b"enableEdit", "PivotTable definition")?
    {
        if parse_bool(enable_edit, "PivotTable enableEdit")? {
            return Err(invalid("Non-Worksheet PivotTable enableEdit must be false"));
        }
    }
    let locations = scan.elements.iter().filter(|element| {
        element.parent_index == Some(root.index)
            && element.ns == root.ns
            && element.local == b"location"
    });
    let mut locations = locations;
    let location = locations
        .next()
        .ok_or_else(|| invalid("Non-Worksheet PivotTable requires location"))?;
    if locations.next().is_some() {
        return Err(invalid(
            "PivotTable definition has multiple location elements",
        ));
    }
    let location_ref = unique_unqualified_attr(location, b"ref", "PivotTable location")?
        .ok_or_else(|| invalid("Non-Worksheet PivotTable requires location@ref"))?;
    if !location_ref.starts_with("A1") {
        return Err(invalid(
            "Non-Worksheet PivotTable location@ref must begin with A1",
        ));
    }
    for child in scan
        .elements
        .iter()
        .filter(|element| element.parent_index == Some(root.index))
    {
        if child.ns == root.ns
            && (child.local == b"pivotEdits"
                || child.local == b"pivotChanges"
                || child.local == b"conditionalFormats")
        {
            return Err(invalid(
                "PivotTable is not eligible for pivotTableReferences",
            ));
        }
    }
    Ok(TableMeta {
        name,
        cache_id,
        owner: root.start.start..root.end,
    })
}

fn cache_requires_connection_route(
    scan: &XmlScan,
    root: &XmlElement,
    source: &XmlElement,
) -> Result<bool> {
    if unique_unqualified_attr(source, b"connectionId", "cacheSource")?.is_some() {
        return Ok(true);
    }
    for extension in scan.elements.iter().filter(|element| {
        element.ns == root.ns
            && element.local == b"ext"
            && extension_owner_matches(scan, source.index, element, true)
    }) {
        if unique_unqualified_attr(extension, b"uri", "cacheSource ext")?
            .is_some_and(|value| xml_token_eq(value, CACHE_SOURCE_URI))
        {
            return Ok(true);
        }
    }
    Ok(false)
}

fn validate_cache_closure(
    scan: &XmlScan,
    expected_cache_id: u32,
    connections: Option<&ConnectionCatalog>,
    require_cache_definition_id: bool,
    allow_empty_connection_names: bool,
) -> Result<CacheClosureStatus> {
    let root = scan
        .elements
        .iter()
        .find(|element| element.parent_index.is_none())
        .ok_or_else(|| invalid("PivotCache Part has no root"))?;
    if root.ns.as_ref() != CORE_NS && root.ns.as_ref() != STRICT_CORE_NS
        || root.local != b"pivotCacheDefinition"
    {
        return Err(invalid("PivotCache Part has an invalid root"));
    }
    if root
        .attrs
        .iter()
        .find(|attr| {
            attr.ns.is_empty() && (attr.local == b"pivotCacheId" || attr.local == b"cacheId")
        })
        .is_some()
    {
        return Err(invalid("core pivotCacheDefinition cannot carry a cache ID"));
    }
    let definition_exts = extension_exts(
        scan,
        root,
        PIVOT_CACHE_DEFINITION_URI,
        require_cache_definition_id,
    )?;
    for candidate in &definition_exts {
        check_fragment_limit(candidate, "pivotCacheDefinition owner fragment")?;
    }
    let mut diagnostic = false;
    if definition_exts.len() > 1 {
        if !require_cache_definition_id {
            return Err(invalid("duplicate pivotCacheDefinition extension"));
        }
        diagnostic = true;
    }
    let mut definition_cache_id = None;
    if let Some(definition_ext) = definition_exts.first().copied() {
        if !validate_known_extension_shape(
            scan,
            definition_ext,
            X14_NS,
            b"pivotCacheDefinition",
            !require_cache_definition_id,
        )? {
            diagnostic = true;
        }
        let definitions = scan.elements.iter().filter(|element| {
            element.ns.as_ref() == X14_NS
                && element.local == b"pivotCacheDefinition"
                && is_owned_payload(scan, root, element, PIVOT_CACHE_DEFINITION_URI)
        });
        for definition in scan.elements.iter().filter(|element| {
            element.ns.as_ref() == X14_NS
                && element.local == b"pivotCacheDefinition"
                && is_owned_payload(scan, root, element, PIVOT_CACHE_DEFINITION_URI)
        }) {
            check_fragment_limit(definition, "pivotCacheDefinition payload fragment")?;
        }
        let mut definitions = definitions;
        if let Some(definition) = definitions.next() {
            if definitions.next().is_some() {
                if !require_cache_definition_id {
                    return Err(invalid("duplicate pivotCacheDefinition extension payload"));
                }
                diagnostic = true;
            }
            if definition.has_element_child || definition.has_cdata || definition.has_text {
                if !require_cache_definition_id {
                    return Err(invalid(
                        "pivotCacheDefinition extension payload must be empty",
                    ));
                }
                diagnostic = true;
            }
            let mut cache_id = None;
            for attr in &definition.attrs {
                if !attr.ns.is_empty() || attr.local.as_slice() != b"pivotCacheId" {
                    if !require_cache_definition_id {
                        return Err(invalid(
                            "pivotCacheDefinition extension has an unknown attribute",
                        ));
                    }
                    diagnostic = true;
                    continue;
                }
                if cache_id.is_some() {
                    if !require_cache_definition_id {
                        return Err(invalid(
                            "pivotCacheDefinition extension has duplicate pivotCacheId",
                        ));
                    }
                    diagnostic = true;
                    continue;
                }
                match parse_u32(&attr.value, "pivotCacheDefinition pivotCacheId") {
                    Ok(value) => cache_id = Some(value),
                    Err(_error) if require_cache_definition_id => diagnostic = true,
                    Err(error) => return Err(error),
                }
            }
            if let Some(cache_id) = cache_id {
                if cache_id != expected_cache_id {
                    if !require_cache_definition_id {
                        return Err(invalid(
                            "pivotCacheDefinition pivotCacheId does not match the semantic workbook cache ID",
                        ));
                    }
                    diagnostic = true;
                } else {
                    definition_cache_id = Some(cache_id);
                }
            } else if !require_cache_definition_id {
                return Err(invalid(
                    "pivotCacheDefinition extension is missing pivotCacheId",
                ));
            } else {
                diagnostic = true;
            }
        } else if !require_cache_definition_id {
            return Err(invalid("pivotCacheDefinition extension has no payload"));
        } else {
            // The rest of the external-cache closure is still validated below.
            // A missing 725 payload is a readable C444 diagnostic; it is not a
            // reason to substitute an OPC relationship ID or to skip ABF5/F057.
            diagnostic = true;
        }
    }
    if require_cache_definition_id && definition_cache_id.is_none() {
        diagnostic = true;
    }
    let sources = scan.elements.iter().filter(|element| {
        element.parent_index == Some(root.index)
            && element.ns == root.ns
            && element.local == b"cacheSource"
    });
    let mut sources = sources;
    let source = sources
        .next()
        .ok_or_else(|| invalid("PivotCache has no cacheSource"))?;
    if sources.next().is_some() {
        return Err(invalid("PivotCache has multiple cacheSource elements"));
    }
    let source_type = unique_unqualified_attr(source, b"type", "cacheSource")?;
    if source_type != Some("external") {
        return Err(invalid(
            "pivotTableServerFormats requires cacheSource type=external",
        ));
    }
    if !validate_external_source_connection(
        scan,
        root,
        source,
        connections,
        allow_empty_connection_names,
        !require_cache_definition_id,
    )? {
        diagnostic = true;
    }
    let owner_exts = extension_exts(
        scan,
        root,
        PIVOT_CACHE_ID_VERSION_URI,
        require_cache_definition_id,
    )?;
    for candidate in &owner_exts {
        check_fragment_limit(candidate, "pivotCacheIdVersion owner fragment")?;
    }
    if owner_exts.len() > 1 {
        return Err(invalid("duplicate pivotCacheIdVersion extension"));
    }
    let Some(version_ext) = owner_exts.first().copied() else {
        return Err(invalid(
            "external PivotCache is missing pivotCacheIdVersion extension",
        ));
    };
    if !validate_known_extension_shape(
        scan,
        version_ext,
        EXT_NS,
        b"pivotCacheIdVersion",
        !require_cache_definition_id,
    )? {
        diagnostic = true;
    }
    let versions = scan.elements.iter().filter(|element| {
        element.ns.as_ref() == EXT_NS
            && element.local == b"pivotCacheIdVersion"
            && is_owned_payload(scan, root, element, PIVOT_CACHE_ID_VERSION_URI)
    });
    for version in scan.elements.iter().filter(|element| {
        element.ns.as_ref() == EXT_NS
            && element.local == b"pivotCacheIdVersion"
            && is_owned_payload(scan, root, element, PIVOT_CACHE_ID_VERSION_URI)
    }) {
        check_fragment_limit(version, "pivotCacheIdVersion payload fragment")?;
    }
    let mut versions = versions;
    let version = versions
        .next()
        .ok_or_else(|| invalid("pivotCacheIdVersion extension has no payload"))?;
    if versions.next().is_some() {
        return Err(invalid("duplicate pivotCacheIdVersion payload"));
    }
    if version.has_element_child || version.has_cdata || version.has_text {
        return Err(invalid("pivotCacheIdVersion payload must be empty"));
    }
    for attr in [
        b"cacheIdSupportedVersion".as_slice(),
        b"cacheIdCreatedVersion".as_slice(),
    ] {
        let value = unique_unqualified_attr(version, attr, "pivotCacheIdVersion")?
            .ok_or_else(|| invalid("pivotCacheIdVersion is missing a required attribute"))?;
        let parsed = parse_u32(value, "pivotCacheIdVersion attribute")?;
        if parsed > u8::MAX as u32 {
            return Err(invalid(
                "pivotCacheIdVersion attribute exceeds unsignedByte",
            ));
        }
    }
    for attribute in &version.attrs {
        if !attribute.ns.is_empty()
            || !matches!(
                attribute.local.as_slice(),
                b"cacheIdSupportedVersion" | b"cacheIdCreatedVersion"
            )
        {
            return Err(invalid("pivotCacheIdVersion has an unknown attribute"));
        }
    }
    let mce_ambiguous = require_cache_definition_id
        && scan.elements.iter().any(|element| {
            if !element.mce_context {
                return false;
            }
            if element.ns.as_ref() == X14_NS && element.local.as_slice() == b"pivotCacheDefinition"
            {
                return is_owned_payload(scan, root, element, PIVOT_CACHE_DEFINITION_URI);
            }
            if element.ns.as_ref() == EXT_NS && element.local.as_slice() == b"pivotCacheIdVersion" {
                return is_owned_payload(scan, root, element, PIVOT_CACHE_ID_VERSION_URI);
            }
            element.ns.as_ref() == X14_NS
                && element.local.as_slice() == b"sourceConnection"
                && is_owned_cache_source_payload(scan, root, source, element)
        });
    Ok(CacheClosureStatus {
        mce_ambiguous,
        diagnostic,
    })
}

fn validate_external_source_connection(
    scan: &XmlScan,
    root: &XmlElement,
    source: &XmlElement,
    connections: Option<&ConnectionCatalog>,
    allow_empty_connection_names: bool,
    strict: bool,
) -> Result<bool> {
    let mut valid = true;
    let connection_id = unique_unqualified_attr(source, b"connectionId", "cacheSource")?
        .map(|value| parse_u32(value, "cacheSource connectionId"))
        .transpose()?;
    if let Some(connection_id) = connection_id {
        if !connections.is_some_and(|connections| connections.by_id.contains_key(&connection_id)) {
            return Err(invalid(
                "cacheSource connectionId does not resolve to a workbook connection",
            ));
        }
    }
    let mut f057_ext = None;
    for ext in scan.elements.iter().filter(|element| {
        element.ns == root.ns
            && element.local == b"ext"
            && extension_owner_matches(scan, source.index, element, true)
    }) {
        let uri = unique_unqualified_attr(ext, b"uri", "cacheSource ext")?;
        if uri.is_some_and(|value| xml_token_eq(value, CACHE_SOURCE_URI)) {
            check_fragment_limit(ext, "cacheSource F057 owner fragment")?;
            if f057_ext.replace(ext).is_some() {
                if strict {
                    return Err(invalid("cacheSource has duplicate F057 extensions"));
                }
                valid = false;
            }
        }
    }
    let Some(ext) = f057_ext else {
        return Ok(valid);
    };
    for attribute in &ext.attrs {
        if !attribute.ns.is_empty() || attribute.local.as_slice() != b"uri" {
            if strict {
                return Err(invalid("F057 extension has an unknown attribute"));
            }
            valid = false;
        }
    }
    for child in scan
        .elements
        .iter()
        .filter(|element| element.parent_index == Some(ext.index))
    {
        let allowed = child.ns.as_ref() == X14_NS && child.local.as_slice() == b"sourceConnection";
        let allowed_mce =
            child.ns.as_ref() == MCE_NS && child.local.as_slice() == b"AlternateContent";
        if !allowed && !allowed_mce {
            if strict {
                return Err(invalid("F057 extension has an unexpected child"));
            }
            valid = false;
        }
    }
    let source_connections = scan.elements.iter().filter(|element| {
        element.ns.as_ref() == X14_NS
            && element.local == b"sourceConnection"
            && is_owned_cache_source_payload(scan, root, source, element)
    });
    for source_connection in scan.elements.iter().filter(|element| {
        element.ns.as_ref() == X14_NS
            && element.local == b"sourceConnection"
            && is_owned_cache_source_payload(scan, root, source, element)
    }) {
        check_fragment_limit(source_connection, "cacheSource F057 payload fragment")?;
    }
    let mut source_connections = source_connections;
    let source_connection = source_connections
        .next()
        .ok_or_else(|| invalid("F057 extension has no sourceConnection"))?;
    if source_connections.next().is_some() {
        if strict {
            return Err(invalid(
                "F057 extension has multiple sourceConnection elements",
            ));
        }
        valid = false;
    }
    for attribute in &source_connection.attrs {
        if !attribute.ns.is_empty() || attribute.local.as_slice() != b"name" {
            if strict {
                return Err(invalid("sourceConnection has an unknown attribute"));
            }
            valid = false;
        }
    }
    if source_connection.has_element_child
        || source_connection.has_cdata
        || source_connection.has_text
    {
        if strict {
            return Err(invalid("sourceConnection must be an empty element"));
        }
        valid = false;
    }
    let name = unique_unqualified_attr(source_connection, b"name", "sourceConnection")?
        .map(decode_spreadsheet_text)
        .transpose()?
        .ok_or_else(|| invalid("sourceConnection requires name"))?;
    if name.is_empty() && !allow_empty_connection_names {
        return Err(invalid("sourceConnection name cannot be empty"));
    }
    if name.encode_utf16().take(65_536).count() >= 65_536 {
        return Err(invalid("sourceConnection name exceeds its text limit"));
    }
    let resolved_id = connections
        .and_then(|connections| connections.by_name.get(&name).copied())
        .ok_or_else(|| {
            invalid("sourceConnection name does not resolve to a workbook connection")
        })?;
    if let Some(connection_id) = connection_id
        && connection_id != resolved_id
    {
        return Err(invalid(
            "cacheSource connectionId disagrees with sourceConnection name",
        ));
    }
    Ok(valid)
}

struct WorksheetNameIndex {
    pivot_names: Vec<String>,
    ordinary_names: HashSet<String>,
}

fn worksheet_name_index(
    package: &OpcPackage,
    catalog: &raw::Catalog,
    relationship_index: &RelationshipIndex<'_>,
    include_ordinary_names: bool,
) -> Result<WorksheetNameIndex> {
    let workbook = package.main_document_part()?;
    let mut pivot_names = Vec::new();
    let mut pivot_name_set = HashSet::<String>::new();
    let mut ordinary_names = HashSet::<String>::new();
    let mut table_targets = HashSet::<PackURI>::new();
    let mut ordinary_targets = HashSet::<PackURI>::new();
    for sheet in &catalog.sheets {
        let Some(relation) = workbook.rels().get(&sheet.relationship_id) else {
            continue;
        };
        if !matches!(relation.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET)
            || relation.is_external()
        {
            continue;
        }
        let uri = relation.target_partname()?;
        let canonical_uri = relationship_index
            .canonical_part(&uri)?
            .ok_or_else(|| invalid("worksheet target is not a package Part"))?;
        let worksheet = package.get_part(canonical_uri)?;
        for relation in worksheet.rels().iter().filter(|relation| {
            matches!(relation.reltype(), rt::PIVOT_TABLE | rt::STRICT_PIVOT_TABLE)
        }) {
            if relation.is_external()
                || relation.target_query().is_some()
                || relation.target_fragment().is_some()
            {
                return Err(invalid(
                    "worksheet PivotTable relationship cannot be external",
                ));
            }
            let table_uri = relation.target_partname()?;
            let canonical_table_uri = relationship_index
                .canonical_part(&table_uri)?
                .ok_or_else(|| invalid("worksheet PivotTable target is not a package Part"))?;
            if !admit_worksheet_table_target(&mut table_targets, canonical_table_uri)? {
                continue;
            }
            let table = package.get_part(canonical_table_uri)?;
            if table.content_type() != ct::SML_PIVOT_TABLE {
                return Err(invalid(
                    "worksheet PivotTable relationship targets the wrong content type",
                ));
            }
            if table.blob().len() > MAX_PART_BYTES
                || table.blob().len() > caller_part_limit(package.read_limits())
            {
                return Err(invalid(
                    "worksheet PivotTable Part exceeds the caller's Part limit",
                ));
            }
            let name = scan_table_metadata(table.blob(), package.read_limits())?.name;
            pivot_name_set
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "worksheet PivotTable name uniqueness index",
                    source,
                })?;
            if !pivot_name_set.insert(name.clone()) {
                return Err(invalid("worksheet PivotTable name is not unique"));
            }
            pivot_names
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "worksheet PivotTable name index",
                    source,
                })?;
            pivot_names.push(name);
        }
        if !include_ordinary_names {
            continue;
        }
        for relation in worksheet
            .rels()
            .iter()
            .filter(|relation| matches!(relation.reltype(), rt::TABLE | rt::STRICT_TABLE))
        {
            if relation.is_external()
                || relation.target_query().is_some()
                || relation.target_fragment().is_some()
            {
                return Err(invalid(
                    "worksheet table relationship cannot be external or target a URI suffix",
                ));
            }
            let table_uri = relation.target_partname()?;
            let canonical_table_uri = relationship_index
                .canonical_part(&table_uri)?
                .ok_or_else(|| invalid("worksheet table target is not a package Part"))?;
            if !admit_worksheet_table_target(&mut ordinary_targets, canonical_table_uri)? {
                continue;
            }
            let table = package.get_part(canonical_table_uri)?;
            if table.content_type() != ct::SML_TABLE {
                return Err(invalid(
                    "worksheet table relationship targets the wrong content type",
                ));
            }
            if table.blob().len() > MAX_PART_BYTES
                || table.blob().len() > caller_part_limit(package.read_limits())
            {
                return Err(invalid(
                    "worksheet table Part exceeds the caller's Part limit",
                ));
            }
            let table = crate::table::parse_table_xml(table.blob())?
                .ok_or_else(|| invalid("worksheet table Part has no table root"))?;
            for name in [table.name, table.display_name] {
                ordinary_names
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "ordinary worksheet table selector index",
                        source,
                    })?;
                ordinary_names.insert(name);
            }
        }
    }
    Ok(WorksheetNameIndex {
        pivot_names,
        ordinary_names,
    })
}

#[cfg(test)]
fn worksheet_pivot_names(
    package: &OpcPackage,
    catalog: &raw::Catalog,
    relationship_index: &RelationshipIndex<'_>,
) -> Result<Vec<String>> {
    Ok(worksheet_name_index(package, catalog, relationship_index, true)?.pivot_names)
}

fn admit_worksheet_table_target(
    table_targets: &mut HashSet<PackURI>,
    canonical_uri: &PackURI,
) -> Result<bool> {
    if table_targets.contains(canonical_uri) {
        // A worksheet may expose the same physical PivotTable through more
        // than one relationship ID.  The workbook-name check needs one
        // physical target; the relationship index retains every edge for
        // closure and readset validation.
        return Ok(false);
    }
    table_targets
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "worksheet PivotTable target cache",
            source,
        })?;
    table_targets.insert(canonical_uri.clone());
    Ok(true)
}

fn estimate_server_format_parser_bytes(scan: &XmlScan) -> Result<usize> {
    let root = scan
        .elements
        .iter()
        .find(|element| element.parent_index.is_none())
        .ok_or_else(|| invalid("PivotTable Part has no root"))?;
    let mut estimate = size_of::<PayloadInfo>();
    let mut payload_count = 0usize;
    let mut entry_count = 0usize;
    let mut reference_count = 0usize;
    for element in &scan.elements {
        if element.ns.as_ref() == EXT_NS
            && element.local.as_slice() == b"pivotTableServerFormats"
            && is_owned_payload(scan, root, element, PIVOT_TABLE_SERVER_FORMATS_URI)
        {
            payload_count = payload_count
                .checked_add(1)
                .ok_or_else(|| invalid("pivotTableServerFormats payload count overflows"))?;
            estimate = estimate
                .checked_add(size_of::<&XmlElement>())
                .ok_or_else(|| invalid("pivotTableServerFormats retained bytes overflow"))?;
            for child in scan
                .elements
                .iter()
                .filter(|candidate| candidate.parent_index == Some(element.index))
            {
                entry_count = entry_count
                    .checked_add(1)
                    .ok_or_else(|| invalid("pivotTableServerFormats entry count overflows"))?;
                estimate = estimate
                    .checked_add(size_of::<ParsedEntry>())
                    .and_then(|value| value.checked_add(size_of::<EntrySource>()))
                    .and_then(|value| value.checked_add(child.start.len()))
                    .ok_or_else(|| invalid("pivotTableServerFormats retained bytes overflow"))?;
                for attribute in &child.attrs {
                    if attribute.ns.is_empty()
                        && (attribute.local.as_slice() == b"culture"
                            || attribute.local.as_slice() == b"format")
                    {
                        estimate =
                            estimate.checked_add(attribute.value.len()).ok_or_else(|| {
                                invalid("pivotTableServerFormats retained text overflows")
                            })?;
                    }
                }
            }
        }
        if element.ns.as_ref() == EXT_NS && element.local.as_slice() == b"x" {
            let has_in = element
                .attrs
                .iter()
                .any(|attribute| attribute.local.as_slice() == b"in");
            if has_in {
                reference_count = reference_count.checked_add(1).ok_or_else(|| {
                    invalid("pivotTableData server-format reference count overflows")
                })?;
            }
        }
    }
    estimate = estimate
        .checked_add(size_of::<&XmlElement>().saturating_mul(payload_count))
        .and_then(|value| {
            value.checked_add(size_of::<IndexReference>().saturating_mul(reference_count))
        })
        .and_then(|value| value.checked_add(size_of::<ParsedEntry>().saturating_mul(entry_count)))
        .and_then(|value| value.checked_mul(2))
        .ok_or_else(|| invalid("pivotTableData server-format retained bytes overflow"))?;
    Ok(estimate)
}

fn parse_payload(scan: &XmlScan, source: &[u8]) -> Result<PayloadInfo> {
    let mut payloads = Vec::new();
    let root = scan
        .elements
        .iter()
        .find(|element| element.parent_index.is_none())
        .ok_or_else(|| invalid("PivotTable Part has no root"))?;
    for element in &scan.elements {
        if element.ns.as_ref() != EXT_NS || element.local != b"pivotTableServerFormats" {
            continue;
        }
        if !is_owned_payload(scan, root, element, PIVOT_TABLE_SERVER_FORMATS_URI) {
            continue;
        }
        check_fragment_limit(element, "pivotTableServerFormats owner fragment")?;
        payloads
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "PivotTable server-format payloads",
                source,
            })?;
        payloads.push(element);
    }
    if payloads.len() != 1 {
        return Err(invalid(if payloads.is_empty() {
            "PivotTable has no pivotTableServerFormats payload"
        } else {
            "PivotTable has duplicate pivotTableServerFormats payloads"
        }));
    }
    let payload = payloads[0];
    if payload.has_non_whitespace_cdata || payload.has_non_whitespace_text {
        return Err(invalid(
            "pivotTableServerFormats cannot contain non-whitespace text",
        ));
    }
    let count = payload
        .attrs
        .iter()
        .find(|attr| attr.ns.is_empty() && attr.local == b"count")
        .map(|attr| parse_u32(&attr.value, "pivotTableServerFormats count"))
        .transpose()?
        .ok_or_else(|| invalid("pivotTableServerFormats requires count"))?;
    let count_source = attr_source(payload, b"count")?
        .ok_or_else(|| invalid("pivotTableServerFormats requires count"))?;
    if count == 0 || count as usize > MAX_SERVER_FORMATS {
        return Err(invalid(
            "pivotTableServerFormats count is outside its bounded domain",
        ));
    }
    let mut entries = Vec::new();
    for element in scan
        .elements
        .iter()
        .filter(|candidate| candidate.parent_index == Some(payload.index))
    {
        if element.ns.as_ref() != EXT_NS || element.local != b"serverFormat" {
            return Err(invalid(
                "pivotTableServerFormats contains an unexpected child",
            ));
        }
        if entries.len() >= MAX_SERVER_FORMATS || entries.len() >= count as usize {
            return Err(invalid("pivotTableServerFormats child count exceeds count"));
        }
        validate_server_format_leaf(element)?;
        let culture = parse_optional_xstring(element, b"culture")?;
        let format = parse_optional_xstring(element, b"format")?;
        entries.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "PivotTable server-format values",
            source,
        })?;
        entries.push(ParsedEntry {
            value: ServerFormat { culture, format },
            source: EntrySource {
                culture: attr_source(element, b"culture")?,
                format: attr_source(element, b"format")?,
                start_tag: element.start.clone(),
                element: element.start.start..element.end,
                qname: source_element_qname(source, &element.start, scan.limits.name_bytes)?,
                namespace_decl: source_element_namespace_decl(source, &element.start, scan.limits)?,
            },
        });
    }
    if entries.is_empty() || entries.len() != count as usize {
        return Err(invalid(
            "pivotTableServerFormats count does not equal child count",
        ));
    }
    let mut index_refs = Vec::new();
    let mut opaque_index_refs = false;
    let mut diagnostic = false;
    for element in &scan.elements {
        if element.ns.as_ref() != EXT_NS || element.local != b"x" {
            continue;
        }
        let Some((data_index, mce_context)) = owned_pivot_data_ancestor(scan, root, element) else {
            continue;
        };
        if mce_context || element.mce_context {
            opaque_index_refs = true;
            continue;
        }
        let mut in_attribute = None;
        let mut in_count = 0usize;
        let mut namespaced_in = false;
        for attribute in &element.attrs {
            if attribute.local != b"in" {
                continue;
            }
            in_count += 1;
            if attribute.ns.is_empty() {
                in_attribute = Some(attribute);
            } else {
                namespaced_in = true;
            }
        }
        if in_count == 0 {
            continue;
        }
        let exact = exact_pivot_value_cell_chain(scan, element, data_index);
        if !exact || in_count != 1 || namespaced_in {
            opaque_index_refs = true;
            continue;
        }
        let Some(attribute) = in_attribute else {
            opaque_index_refs = true;
            continue;
        };
        let index = parse_u32(&attribute.value, "pivotValueCellExtra in")? as usize;
        if index > entries.len() {
            return Err(invalid(
                "pivotValueCellExtra@in exceeds server-format count",
            ));
        }
        if index == entries.len() {
            diagnostic = true;
        }
        index_refs
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "PivotTable server-format index references",
                source,
            })?;
        index_refs.push(IndexReference {
            value: attr_source(element, b"in")?
                .ok_or_else(|| invalid("pivotValueCellExtra in source range is missing"))?
                .value,
            index,
        });
    }
    Ok(PayloadInfo {
        entries,
        owner: payload.start.start..payload.end,
        count: count_source,
        index_refs,
        opaque_index_refs,
        diagnostic_index_boundary: diagnostic,
        mce_ambiguous: payload.mce_context,
    })
}

fn owned_pivot_data_ancestor(
    scan: &XmlScan,
    root: &XmlElement,
    element: &XmlElement,
) -> Option<(usize, bool)> {
    let mut current = element.parent_index;
    while let Some(index) = current {
        let data = scan.elements.get(index)?;
        if data.ns.as_ref() == EXT_NS
            && data.local == b"pivotTableData"
            && is_owned_payload(scan, root, data, PIVOT_TABLE_DATA_URI)
        {
            return Some((index, data.mce_context));
        }
        current = data.parent_index;
    }
    None
}

fn exact_pivot_value_cell_chain(scan: &XmlScan, element: &XmlElement, data: usize) -> bool {
    let Some(cell) = element
        .parent_index
        .and_then(|index| scan.elements.get(index))
    else {
        return false;
    };
    if cell.ns.as_ref() != EXT_NS || cell.local != b"c" {
        return false;
    }
    let Some(row) = cell.parent_index.and_then(|index| scan.elements.get(index)) else {
        return false;
    };
    if row.ns.as_ref() != EXT_NS || row.local != b"pivotRow" {
        return false;
    }
    row.parent_index == Some(data)
}

fn validate_server_format_leaf(element: &XmlElement) -> Result<()> {
    // CT_ServerFormat is an attribute-only empty complex type.  Its parent
    // collection may carry formatting whitespace, but the leaf itself has no
    // text particle, so even whitespace text/CDATA is schema-invalid.
    if element.has_element_child || element.has_cdata || element.has_text {
        return Err(invalid(
            "serverFormat must be an attribute-only empty element",
        ));
    }
    for attribute in &element.attrs {
        if !attribute.ns.is_empty()
            || (attribute.local.as_slice() != b"culture" && attribute.local.as_slice() != b"format")
        {
            return Err(invalid(
                "serverFormat has an attribute outside its CT_ServerFormat contract",
            ));
        }
    }
    Ok(())
}

fn parse_optional_xstring(element: &XmlElement, name: &[u8]) -> Result<Option<String>> {
    let mut found = None;
    for attr in &element.attrs {
        if attr.ns.is_empty() && attr.local == name {
            if found.is_some() {
                return Err(invalid("serverFormat has a duplicate optional attribute"));
            }
            found = Some(attr);
        }
    }
    found
        .map(|attr| decode_spreadsheet_text(&attr.value))
        .transpose()
}

fn attr_source(element: &XmlElement, name: &[u8]) -> Result<Option<AttributeSource>> {
    let mut found = None;
    for attr in &element.attrs {
        if attr.ns.is_empty() && attr.local == name {
            if found.is_some() {
                return Err(invalid("serverFormat has a duplicate optional attribute"));
            }
            found = Some(attr);
        }
    }
    Ok(found.map(|attr| AttributeSource {
        value: attr.value_range.clone(),
        whole: attr.whole_range.clone(),
    }))
}

fn source_element_qname(
    source: &[u8],
    start: &Range<usize>,
    maximum_name_bytes: usize,
) -> Result<Vec<u8>> {
    let name_start = match source.get(start.start) {
        Some(b'<') => start
            .start
            .checked_add(1)
            .ok_or_else(|| invalid("serverFormat QName range overflows"))?,
        Some(_) => start.start,
        None => return Err(invalid("serverFormat QName range is outside its source")),
    };
    let raw = source
        .get(name_start..start.end)
        .ok_or_else(|| invalid("serverFormat QName range is outside its source"))?;
    let name_len = raw
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>'))
        .ok_or_else(|| invalid("serverFormat start tag has no QName"))?;
    if name_len == 0 || name_len > MAX_NAME_BYTES || name_len > maximum_name_bytes {
        return Err(invalid("serverFormat QName exceeds its source limit"));
    }
    let mut qname = Vec::new();
    qname
        .try_reserve_exact(name_len)
        .map_err(|source| Error::Allocation {
            resource: "PivotTable server-format QName",
            source,
        })?;
    qname.extend_from_slice(&raw[..name_len]);
    Ok(qname)
}

fn source_element_namespace_decl(
    source: &[u8],
    start: &Range<usize>,
    limits: XmlScanLimits,
) -> Result<Option<Vec<u8>>> {
    let qname = source_element_qname(source, start, limits.name_bytes)?;
    let prefix = qname
        .iter()
        .position(|byte| *byte == b':')
        .map(|index| &qname[..index]);
    let name_start = if source.get(start.start) == Some(&b'<') {
        start.start.saturating_add(1)
    } else {
        start.start
    };
    let mut cursor = name_start.saturating_add(qname.len());
    while cursor < start.end {
        while cursor < start.end && is_xml_space(source[cursor]) {
            cursor += 1;
        }
        if cursor >= start.end || matches!(source[cursor], b'/' | b'>') {
            break;
        }
        let attr_start = cursor;
        while cursor < start.end
            && !is_xml_space(source[cursor])
            && !matches!(source[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let attr_name = source.get(attr_start..cursor).unwrap_or_default();
        while cursor < start.end && is_xml_space(source[cursor]) {
            cursor += 1;
        }
        if source.get(cursor) != Some(&b'=') {
            return Err(invalid("serverFormat namespace declaration is malformed"));
        }
        cursor += 1;
        while cursor < start.end && is_xml_space(source[cursor]) {
            cursor += 1;
        }
        let quote = *source
            .get(cursor)
            .ok_or_else(|| invalid("serverFormat namespace declaration has no value"))?;
        if !matches!(quote, b'\'' | b'"') {
            return Err(invalid("serverFormat namespace declaration is not quoted"));
        }
        cursor += 1;
        while cursor < start.end && source[cursor] != quote {
            cursor += 1;
        }
        if cursor >= start.end {
            return Err(invalid(
                "serverFormat namespace declaration is unterminated",
            ));
        }
        cursor += 1;
        let matches_prefix = match prefix {
            Some(prefix) => {
                attr_name.len() == 6 + prefix.len()
                    && attr_name.starts_with(b"xmlns:")
                    && &attr_name[6..] == prefix
            },
            None => attr_name == b"xmlns",
        };
        if matches_prefix {
            let mut declaration = Vec::new();
            let declaration_len = 1usize
                .checked_add(attr_name.len())
                .and_then(|length| length.checked_add(3))
                .and_then(|length| length.checked_add(EXT_NS.len()))
                .ok_or_else(|| invalid("serverFormat namespace declaration overflows"))?;
            if declaration_len > MAX_NAMESPACE_BYTES
                || declaration_len > limits.namespace_bytes
                || declaration_len > limits.attribute_bytes
            {
                return Err(invalid(
                    "serverFormat namespace declaration exceeds its source limit",
                ));
            }
            declaration
                .try_reserve_exact(declaration_len)
                .map_err(|source| Error::Allocation {
                    resource: "PivotTable server-format namespace declaration",
                    source,
                })?;
            declaration.push(b' ');
            declaration.extend_from_slice(attr_name);
            declaration.extend_from_slice(b"=\"");
            declaration.extend_from_slice(EXT_NS);
            declaration.push(b'"');
            return Ok(Some(declaration));
        }
    }
    Ok(None)
}

fn is_xml_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n')
}

/// Compare a recognized OOXML extension URI using the XML Schema `token`
/// whitespace facet.  The source value itself remains untouched in
/// `XmlAttribute::value` and its lexical range; only this semantic comparison
/// trims XML-S padding.  All recognized constants are URI tokens without
/// internal whitespace, so allocation-free edge trimming is sufficient here.
fn xml_token_eq(value: &str, expected: &str) -> bool {
    value.trim_matches(|character| matches!(character, ' ' | '\t' | '\r' | '\n')) == expected
}

fn parse_u32(value: &str, owner: &str) -> Result<u32> {
    let value = value.trim_matches(|character| matches!(character, ' ' | '\t' | '\r' | '\n'));
    let bytes = value.as_bytes();
    let (negative, digits) = match bytes.first().copied() {
        Some(b'+') => (false, &bytes[1..]),
        Some(b'-') => (true, &bytes[1..]),
        _ => (false, bytes),
    };
    if digits.is_empty() || !digits.iter().all(u8::is_ascii_digit) {
        return Err(invalid(format!("{owner} is not an unsignedInt")));
    }
    // XML Schema's unsignedInt lexical space admits a leading sign. A
    // negative lexical form denotes zero only when every digit is zero.
    if negative && digits.iter().any(|digit| *digit != b'0') {
        return Err(invalid(format!("{owner} is not an unsignedInt")));
    }
    std::str::from_utf8(digits)
        .expect("unsignedInt digits are ASCII")
        .parse::<u32>()
        .map_err(|_| invalid(format!("{owner} is not an unsignedInt")))
}

fn parse_bool(value: &str, owner: &str) -> Result<bool> {
    let value = value.trim_matches([' ', '\t', '\r', '\n']);
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("{owner} is not an XML Schema boolean"))),
    }
}

fn decode_spreadsheet_text(value: &str) -> Result<String> {
    raw::strings::decode_spreadsheet_text(value)
}

#[derive(Debug)]
struct XmlScan {
    elements: Vec<XmlElement>,
    limits: XmlScanLimits,
    ignorable_scopes: Vec<IgnorableScope>,
}

type NamespaceRef = Arc<[u8]>;

#[derive(Debug, Default)]
struct NamespaceInterner {
    values: HashSet<NamespaceRef>,
}

impl NamespaceInterner {
    fn intern(&mut self, value: &[u8], limits: XmlScanLimits) -> Result<NamespaceRef> {
        if value.len() > limits.namespace_bytes {
            return Err(invalid("PivotTable XML namespace URI exceeds caller limit"));
        }
        validate_xml_characters(value)?;
        if let Some(existing) = self.values.get(value) {
            return Ok(existing.clone());
        }
        self.values
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "PivotTable XML namespace interner",
                source,
            })?;
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(value.len())
            .map_err(|source| Error::Allocation {
                resource: "PivotTable XML namespace",
                source,
            })?;
        owned.extend_from_slice(value);
        let shared: NamespaceRef = Arc::from(owned);
        self.values.insert(shared.clone());
        Ok(shared)
    }
}

#[derive(Debug)]
struct IgnorableScope {
    parent: Option<usize>,
    namespaces: Vec<NamespaceRef>,
}

#[derive(Debug)]
struct XmlElement {
    index: usize,
    parent_index: Option<usize>,
    ns: NamespaceRef,
    local: Vec<u8>,
    start: Range<usize>,
    end: usize,
    attrs: Vec<XmlAttribute>,
    /// MCE branch metadata is retained only so extension owners can apply the
    /// first-supported Choice/Fallback rule without rescanning the Part.  The
    /// branch grammar is parsed by the cache-field owner; the server-format
    /// owner does not consult it.
    mce_branch: Option<cached_unique_names::MceBranch>,
    /// `mc:Ignorable` declarations made directly on this element.  The
    /// cache-field owner walks the already-built ancestor chain when it needs
    /// the effective set, so this scanner does not copy the inherited list at
    /// every depth level.
    ignorable_scope: usize,
    mce_context: bool,
    has_element_child: bool,
    has_cdata: bool,
    has_text: bool,
    has_non_whitespace_cdata: bool,
    has_non_whitespace_text: bool,
}

#[derive(Debug)]
struct XmlAttribute {
    ns: NamespaceRef,
    local: Vec<u8>,
    value: String,
    value_range: Range<usize>,
    whole_range: Range<usize>,
}

#[derive(Debug)]
struct OpenElement {
    index: usize,
    ignorable_scope: usize,
    mce_context: bool,
}

fn scan_xml(bytes: &[u8], expected_root: &str, read_limits: ReadLimits) -> Result<XmlScan> {
    scan_xml_config(bytes, expected_root, read_limits, false)
}

/// Scan one Part with the cache-field owner's opt-in MCE branch metadata.
/// Ordinary C510/workbook/connection scans intentionally keep their prior
/// compatibility behavior and do not validate unrelated MCE branches.
fn scan_xml_with_mce(
    bytes: &[u8],
    expected_root: &str,
    read_limits: ReadLimits,
) -> Result<XmlScan> {
    scan_xml_config(bytes, expected_root, read_limits, true)
}

fn scan_xml_config(
    bytes: &[u8],
    expected_root: &str,
    read_limits: ReadLimits,
    parse_mce_branches: bool,
) -> Result<XmlScan> {
    let limits = XmlScanLimits::from_read_limits(read_limits);
    if bytes.is_empty() || bytes.len() > limits.part_bytes {
        return Err(invalid("PivotTable XML exceeds its Part limit"));
    }
    validate_xml_characters(bytes)?;
    // Use the plain reader here instead of `NsReader`.  `NsReader` resolves
    // namespace values before the owner can perform XML attribute-value
    // normalization, which rejects a legal reserved `xml` binding written
    // with character references.  The bounded resolver below applies the
    // declarations after decoding their namespace URI and retains the same
    // source offsets from the reader.
    let mut reader = Reader::from_reader(bytes);
    let origin = ReaderOrigin::of(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut elements = Vec::<XmlElement>::new();
    let mut stack = Vec::<OpenElement>::new();
    let mut namespace_interner = NamespaceInterner::default();
    let mut resolver = NamespaceResolver::new(limits, &mut namespace_interner)?;
    let mut ignorable_scopes = Vec::new();
    ignorable_scopes
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "PivotTable MCE scope table",
            source,
        })?;
    ignorable_scopes.push(IgnorableScope {
        parent: None,
        namespaces: Vec::new(),
    });
    let mut root_seen = false;
    let mut root_closed = false;
    let mut nodes = 0usize;
    let mut events = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("PivotTable XML event count overflows"))?;
        if events > limits.events {
            return Err(invalid("PivotTable XML event count exceeds caller limit"));
        }
        let before = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("PivotTable XML offset exceeds usize"))?;
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?;
        let after = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("PivotTable XML offset exceeds usize"))?;
        if !matches!(event, Event::Eof) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| invalid("PivotTable XML node count overflow"))?;
            if nodes > MAX_NODES {
                return Err(invalid("PivotTable XML node count exceeds limit"));
            }
        }
        let is_start = matches!(&event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                // Reject impossible tree states and depth exhaustion before
                // resolving any names or namespace declarations.  Besides
                // preserving the event/depth accounting for empty elements,
                // this keeps a second root from making unbounded namespace
                // work before it is rejected.
                if root_seen && (root_closed || stack.is_empty()) {
                    return Err(invalid("PivotTable XML has multiple roots"));
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("PivotTable XML depth exceeds limit"));
                }
                if stack.len() >= limits.depth {
                    return Err(invalid("PivotTable XML depth exceeds caller limit"));
                }
                begin_namespace_scope(&mut resolver, &element, limits, &mut namespace_interner)?;
                let (ns, local) = resolved_name(&resolver, element.name(), limits)?;
                if !root_seen {
                    root_seen = true;
                    if local.as_slice() != expected_root.as_bytes() {
                        return Err(invalid(format!(
                            "PivotTable XML root must be {expected_root}"
                        )));
                    }
                }
                let parent = stack.last();
                let index = elements.len();
                let start = event_start_range(bytes, before, after, &element)?;
                let element_end = start.end;
                let attrs =
                    parse_attrs(bytes, &element, &resolver, reader.decoder(), &start, limits)?;
                let local_ignorable = if parse_mce_branches {
                    parse_ignorable_namespaces(&element, &resolver, reader.decoder(), limits)?
                } else {
                    Vec::new()
                };
                let parent_scope = parent
                    .map(|parent| parent.ignorable_scope)
                    .unwrap_or_default();
                let ignorable_scope = if parse_mce_branches {
                    extend_ignorable_scope(&mut ignorable_scopes, parent_scope, local_ignorable)?
                } else {
                    0
                };
                let parses_mce_branch = parse_mce_branches
                    && ns.as_ref() == MCE_NS
                    && matches!(local.as_slice(), b"Choice" | b"Fallback");
                if parse_mce_branches
                    && ns.as_ref() == MCE_NS
                    && local.as_slice() == b"AlternateContent"
                {
                    cached_unique_names::validate_mce_alternate_content_attributes(
                        &element,
                        &resolver,
                        ignorable_scope,
                        &ignorable_scopes,
                    )?;
                }
                let mce_branch = if parses_mce_branch {
                    cached_unique_names::scan_mce_branch(
                        ns.as_ref(),
                        &element,
                        &resolver,
                        reader.decoder(),
                        limits,
                        ignorable_scope,
                        &ignorable_scopes,
                    )?
                } else {
                    None
                };
                if let Some(parent) = parent {
                    if let Some(parent_element) = elements.get_mut(parent.index) {
                        parent_element.has_element_child = true;
                    }
                }
                let mce_context = parent.map_or(ns.as_ref() == MCE_NS, |parent| {
                    parent.mce_context
                        || elements
                            .get(parent.index)
                            .is_some_and(|parent| parent.ns.as_ref() == MCE_NS)
                        || ns.as_ref() == MCE_NS
                });
                let info = XmlElement {
                    index,
                    parent_index: parent.map(|parent| parent.index),
                    ns: ns.clone(),
                    local,
                    start,
                    end: element_end,
                    attrs,
                    mce_branch,
                    ignorable_scope,
                    mce_context,
                    has_element_child: false,
                    has_cdata: false,
                    has_text: false,
                    has_non_whitespace_cdata: false,
                    has_non_whitespace_text: false,
                };
                elements
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "PivotTable XML element index",
                        source,
                    })?;
                elements.push(info);
                if is_start {
                    stack.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "PivotTable XML element stack",
                        source,
                    })?;
                    stack.push(OpenElement {
                        index,
                        ignorable_scope,
                        mce_context,
                    });
                } else if stack.is_empty() {
                    root_closed = true;
                }
                if !is_start {
                    resolver.pop();
                }
            },
            Event::End(end_name) => {
                validate_qname(
                    end_name.name().as_ref(),
                    limits,
                    "PivotTable XML end element",
                )?;
                if end_name
                    .name()
                    .prefix()
                    .is_some_and(|prefix| prefix.is_xmlns())
                {
                    return Err(invalid(
                        "PivotTable XML end element uses the reserved xmlns prefix",
                    ));
                }
                let Some(open) = stack.pop() else {
                    return Err(invalid("PivotTable XML has an unmatched end"));
                };
                let end = event_end_position(bytes, before, after);
                if let Some(element) = elements.get_mut(open.index) {
                    element.end = end;
                }
                if stack.is_empty() {
                    root_closed = true;
                }
                resolver.pop();
            },
            Event::Eof => break,
            Event::PI(_) | Event::DocType(_) => {
                // PIs are preserved source material; DTDs are not accepted by
                // the bounded OOXML reader.  Predefined and numeric character
                // references are validated in the GeneralRef arm below.
                if matches!(event, Event::DocType(_)) {
                    return Err(invalid("PivotTable XML DTD is not supported"));
                }
            },
            Event::GeneralRef(reference) => {
                validate_general_reference(reference.as_ref())?;
                if let Some(open) = stack.last()
                    && let Some(element) = elements.get_mut(open.index)
                {
                    element.has_text = true;
                    element.has_non_whitespace_text |=
                        xml_reference_has_non_whitespace(reference.as_ref())?;
                }
            },
            Event::Text(_) => {
                if bytes
                    .get(before..after)
                    .is_some_and(|raw| raw.windows(3).any(|window| window == b"]]>"))
                {
                    return Err(invalid("PivotTable XML text contains raw ]]>"));
                }
                if let Some(open) = stack.last()
                    && let Some(element) = elements.get_mut(open.index)
                {
                    element.has_text = true;
                    element.has_non_whitespace_text |=
                        xml_text_has_non_whitespace(bytes.get(before..after).unwrap_or_default())?;
                }
            },
            Event::CData(data) => {
                if let Some(open) = stack.last()
                    && let Some(element) = elements.get_mut(open.index)
                {
                    element.has_cdata = true;
                    element.has_non_whitespace_cdata |=
                        raw_xml_cdata_has_non_whitespace(data.as_ref())?;
                }
            },
            Event::Comment(_) | Event::Decl(_) => {},
        }
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("PivotTable XML is unterminated"));
    }
    Ok(XmlScan {
        elements,
        limits,
        ignorable_scopes,
    })
}

fn event_end_position(bytes: &[u8], before: usize, after: usize) -> usize {
    if after <= bytes.len() && after >= before {
        if let Some(offset) = bytes[before..after].iter().position(|byte| *byte == b'>') {
            return before.saturating_add(offset).saturating_add(1);
        }
        return after;
    }
    bytes.len()
}

fn validate_general_reference(reference: &[u8]) -> Result<()> {
    if matches!(reference, b"amp" | b"lt" | b"gt" | b"apos" | b"quot") {
        return Ok(());
    }
    let (radix, digits) = if let Some(hex) = reference
        .strip_prefix(b"#x")
        .or_else(|| reference.strip_prefix(b"#X"))
    {
        (16, hex)
    } else if let Some(decimal) = reference.strip_prefix(b"#") {
        (10, decimal)
    } else {
        return Err(invalid(
            "PivotTable XML has an unsupported entity reference",
        ));
    };
    if digits.is_empty() {
        return Err(invalid("PivotTable XML character reference has no digits"));
    }
    let mut value = 0u32;
    for digit in digits {
        let digit = match digit {
            b'0'..=b'9' => u32::from(digit - b'0'),
            b'a'..=b'f' if radix == 16 => u32::from(digit - b'a' + 10),
            b'A'..=b'F' if radix == 16 => u32::from(digit - b'A' + 10),
            _ => return Err(invalid("PivotTable XML character reference is invalid")),
        };
        value = value
            .checked_mul(radix)
            .and_then(|value| value.checked_add(digit))
            .ok_or_else(|| invalid("PivotTable XML character reference overflows"))?;
    }
    let character = char::from_u32(value)
        .ok_or_else(|| invalid("PivotTable XML character reference is not a scalar"))?;
    if matches!(
        character,
        '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}'
            | '\u{10000}'..='\u{10ffff}'
    ) {
        Ok(())
    } else {
        Err(invalid(
            "PivotTable XML character reference is not an XML character",
        ))
    }
}

fn xml_reference_has_non_whitespace(reference: &[u8]) -> Result<bool> {
    let value = if let Some(hex) = reference
        .strip_prefix(b"#x")
        .or_else(|| reference.strip_prefix(b"#X"))
    {
        u32::from_str_radix(std::str::from_utf8(hex).unwrap_or_default(), 16).ok()
    } else if let Some(decimal) = reference.strip_prefix(b"#") {
        std::str::from_utf8(decimal)
            .ok()
            .and_then(|value| value.parse::<u32>().ok())
    } else {
        None
    };
    if let Some(value) = value.and_then(char::from_u32) {
        return Ok(!is_xml_space_char(value));
    }
    Ok(true)
}

fn xml_text_has_non_whitespace(raw: &[u8]) -> Result<bool> {
    let text = std::str::from_utf8(raw)
        .map_err(|error| invalid(format!("PivotTable XML text is not UTF-8: {error}")))?;
    let bytes = text.as_bytes();
    let mut cursor = 0usize;
    while cursor < bytes.len() {
        if bytes[cursor] != b'&' {
            let character = text[cursor..]
                .chars()
                .next()
                .ok_or_else(|| invalid("PivotTable XML text cursor is invalid"))?;
            if !is_xml_space_char(character) {
                return Ok(true);
            }
            cursor = cursor
                .checked_add(character.len_utf8())
                .ok_or_else(|| invalid("PivotTable XML text cursor overflows"))?;
            continue;
        }
        let end = bytes[cursor..]
            .iter()
            .position(|byte| *byte == b';')
            .and_then(|offset| cursor.checked_add(offset))
            .ok_or_else(|| invalid("PivotTable XML text entity is unterminated"))?;
        if xml_reference_has_non_whitespace(&bytes[cursor + 1..end])? {
            return Ok(true);
        }
        cursor = end
            .checked_add(1)
            .ok_or_else(|| invalid("PivotTable XML text entity cursor overflows"))?;
    }
    Ok(false)
}

/// CDATA is character data, not an XML entity-bearing text token.  In
/// particular, the literal bytes `&#x20;` inside CDATA are non-whitespace
/// content and must not be normalized to a space for element-only checks.
fn raw_xml_cdata_has_non_whitespace(raw: &[u8]) -> Result<bool> {
    let text = std::str::from_utf8(raw)
        .map_err(|error| invalid(format!("PivotTable XML CDATA is not UTF-8: {error}")))?;
    Ok(text.chars().any(|character| !is_xml_space_char(character)))
}

const fn is_xml_space_char(character: char) -> bool {
    matches!(character, ' ' | '\t' | '\r' | '\n')
}

const XML_NAMESPACE_URI: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE_URI: &[u8] = b"http://www.w3.org/2000/xmlns/";

#[derive(Debug)]
struct NamespaceBinding {
    prefix: Arc<[u8]>,
    namespace: NamespaceRef,
    level: usize,
}

#[derive(Debug)]
enum NamespaceResolution {
    Bound(NamespaceRef),
    Unknown(Vec<u8>),
}

/// A fallible, caller-bounded namespace resolver for the source scanner.
///
/// `quick_xml::name::NamespaceResolver` stores namespace bytes in private
/// `Vec`s and appends with infallible `extend_from_slice`.  That is acceptable
/// for ordinary parsing, but this owner promises caller-bounded allocations.
/// Keeping the small equivalent here lets every index/binding growth go
/// through the crate's fallible allocation path while retaining quick-xml's
/// expanded-name semantics.  The active prefix index restores shadowed
/// bindings on scope pop, and URI values are already interned by declaration.
#[derive(Debug)]
struct NamespaceResolver {
    bindings: Vec<NamespaceBinding>,
    by_prefix: HashMap<Arc<[u8]>, Vec<usize>>,
    empty: NamespaceRef,
    nesting_level: usize,
    max_bindings: usize,
}

impl NamespaceResolver {
    fn new(limits: XmlScanLimits, interner: &mut NamespaceInterner) -> Result<Self> {
        let empty = interner.intern(&[], limits)?;
        let xml_namespace = interner.intern(XML_NAMESPACE_URI, limits)?;
        let xmlns_namespace = interner.intern(XMLNS_NAMESPACE_URI, limits)?;
        let mut bindings = Vec::new();
        bindings
            .try_reserve_exact(2)
            .map_err(|source| Error::Allocation {
                resource: "PivotTable XML namespace bindings",
                source,
            })?;
        let mut by_prefix = HashMap::new();
        by_prefix
            .try_reserve(2)
            .map_err(|source| Error::Allocation {
                resource: "PivotTable XML namespace prefix index",
                source,
            })?;
        let xml_prefix = copy_namespace_prefix(b"xml")?;
        let mut xml_stack = Vec::new();
        xml_stack
            .try_reserve_exact(1)
            .map_err(|source| Error::Allocation {
                resource: "PivotTable XML namespace prefix stack",
                source,
            })?;
        xml_stack.push(0);
        by_prefix.insert(xml_prefix.clone(), xml_stack);
        bindings.push(NamespaceBinding {
            prefix: xml_prefix,
            namespace: xml_namespace,
            level: 0,
        });

        let xmlns_prefix = copy_namespace_prefix(b"xmlns")?;
        let mut xmlns_stack = Vec::new();
        xmlns_stack
            .try_reserve_exact(1)
            .map_err(|source| Error::Allocation {
                resource: "PivotTable XML namespace prefix stack",
                source,
            })?;
        xmlns_stack.push(1);
        by_prefix.insert(xmlns_prefix.clone(), xmlns_stack);
        bindings.push(NamespaceBinding {
            prefix: xmlns_prefix,
            namespace: xmlns_namespace,
            level: 0,
        });

        Ok(Self {
            bindings,
            by_prefix,
            empty,
            nesting_level: 0,
            max_bindings: limits
                .part_bytes
                .checked_add(2)
                .ok_or_else(|| invalid("PivotTable XML namespace binding count overflows"))?,
        })
    }

    fn level(&self) -> usize {
        self.nesting_level
    }

    fn set_level(&mut self, level: usize) {
        self.nesting_level = level;
    }

    fn add(
        &mut self,
        prefix: PrefixDeclaration<'_>,
        namespace: NamespaceRef,
        limits: XmlScanLimits,
    ) -> Result<()> {
        let namespace_bytes = namespace.as_ref();
        match prefix {
            PrefixDeclaration::Default => {},
            PrefixDeclaration::Named(prefix) if prefix == b"xml" => {
                if namespace_bytes != XML_NAMESPACE_URI {
                    return Err(invalid(
                        "the namespace prefix 'xml' cannot be rebound to another URI",
                    ));
                }
                // `xml` is already a fixed level-zero binding.
                return Ok(());
            },
            PrefixDeclaration::Named(prefix) if prefix == b"xmlns" => {
                return Err(invalid("the namespace prefix 'xmlns' cannot be declared"));
            },
            PrefixDeclaration::Named(_) if namespace_bytes == XML_NAMESPACE_URI => {
                return Err(invalid(
                    "a non-xml namespace prefix cannot bind the XML namespace",
                ));
            },
            PrefixDeclaration::Named(_) if namespace_bytes == XMLNS_NAMESPACE_URI => {
                return Err(invalid(
                    "a namespace prefix cannot bind the XMLNS namespace",
                ));
            },
            PrefixDeclaration::Named(_) => {},
        }

        if self.bindings.len() >= self.max_bindings {
            return Err(invalid(
                "PivotTable XML namespace binding count exceeds caller limit",
            ));
        }
        let prefix_bytes = match prefix {
            PrefixDeclaration::Default => &[][..],
            PrefixDeclaration::Named(prefix) => prefix,
        };
        if prefix_bytes.len() > limits.name_bytes {
            return Err(invalid(
                "PivotTable XML namespace prefix exceeds caller limit",
            ));
        }
        self.bindings
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "PivotTable XML namespace bindings",
                source,
            })?;
        let existing_prefix = self
            .by_prefix
            .get_key_value(prefix_bytes)
            .map(|(prefix, _)| Arc::clone(prefix));
        let (prefix, is_new) = if let Some(prefix) = existing_prefix {
            (prefix, false)
        } else {
            self.by_prefix
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "PivotTable XML namespace prefix index",
                    source,
                })?;
            (copy_namespace_prefix(prefix_bytes)?, true)
        };
        let binding_index = self.bindings.len();
        if is_new {
            let mut stack = Vec::new();
            stack
                .try_reserve_exact(1)
                .map_err(|source| Error::Allocation {
                    resource: "PivotTable XML namespace prefix stack",
                    source,
                })?;
            stack.push(binding_index);
            self.by_prefix.insert(prefix.clone(), stack);
        } else {
            let stack = self
                .by_prefix
                .get_mut(prefix_bytes)
                .ok_or_else(|| invalid("PivotTable XML namespace prefix index is inconsistent"))?;
            stack.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "PivotTable XML namespace prefix stack",
                source,
            })?;
            stack.push(binding_index);
        }
        self.bindings.push(NamespaceBinding {
            prefix,
            namespace,
            level: self.nesting_level,
        });
        Ok(())
    }

    fn pop(&mut self) {
        self.nesting_level = self.nesting_level.saturating_sub(1);
        let current_level = self.nesting_level;
        while self.bindings.len() > 2
            && self
                .bindings
                .last()
                .is_some_and(|binding| binding.level > current_level)
        {
            let binding = self
                .bindings
                .pop()
                .expect("namespace binding exists after length check");
            let remove_prefix = if let Some(stack) = self.by_prefix.get_mut(binding.prefix.as_ref())
            {
                let _ = stack.pop();
                stack.is_empty()
            } else {
                false
            };
            if remove_prefix {
                self.by_prefix.remove(binding.prefix.as_ref());
            }
        }
    }

    fn resolve_element<'name>(
        &self,
        name: QName<'name>,
    ) -> Result<(NamespaceResolution, LocalName<'name>)> {
        self.resolve(name, true)
    }

    fn resolve_attribute<'name>(
        &self,
        name: QName<'name>,
    ) -> Result<(NamespaceResolution, LocalName<'name>)> {
        self.resolve(name, false)
    }

    fn resolve<'name>(
        &self,
        name: QName<'name>,
        use_default: bool,
    ) -> Result<(NamespaceResolution, LocalName<'name>)> {
        let (local, prefix) = name.decompose();
        Ok((self.resolve_prefix(prefix, use_default)?, local))
    }

    fn resolve_prefix(
        &self,
        prefix: Option<Prefix<'_>>,
        use_default: bool,
    ) -> Result<NamespaceResolution> {
        if prefix.is_none() && !use_default {
            return Ok(NamespaceResolution::Bound(Arc::clone(&self.empty)));
        }
        let prefix_bytes = match prefix {
            Some(prefix) => prefix.into_inner(),
            None => &[][..],
        };
        let Some(binding_index) = self
            .by_prefix
            .get(prefix_bytes)
            .and_then(|stack| stack.last())
            .copied()
        else {
            return if let Some(prefix) = prefix {
                Ok(NamespaceResolution::Unknown(copy_unknown_prefix(prefix)?))
            } else {
                Ok(NamespaceResolution::Bound(Arc::clone(&self.empty)))
            };
        };
        let binding = self
            .bindings
            .get(binding_index)
            .ok_or_else(|| invalid("PivotTable XML namespace binding index is invalid"))?;
        if let Some(prefix) = prefix
            && binding.namespace.is_empty()
        {
            return Ok(NamespaceResolution::Unknown(copy_unknown_prefix(prefix)?));
        }
        Ok(NamespaceResolution::Bound(Arc::clone(&binding.namespace)))
    }
}

fn copy_namespace_prefix(value: &[u8]) -> Result<Arc<[u8]>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation {
            resource: "PivotTable XML namespace prefix",
            source,
        })?;
    owned.extend_from_slice(value);
    Ok(Arc::from(owned))
}

fn copy_unknown_prefix(prefix: Prefix<'_>) -> Result<Vec<u8>> {
    let mut unknown = Vec::new();
    unknown
        .try_reserve_exact(prefix.as_ref().len())
        .map_err(|source| Error::Allocation {
            resource: "PivotTable XML unknown namespace prefix",
            source,
        })?;
    unknown.extend_from_slice(prefix.as_ref());
    Ok(unknown)
}

fn begin_namespace_scope(
    resolver: &mut NamespaceResolver,
    element: &BytesStart<'_>,
    limits: XmlScanLimits,
    interner: &mut NamespaceInterner,
) -> Result<()> {
    let level = resolver
        .level()
        .checked_add(1)
        .ok_or_else(|| invalid("PivotTable XML namespace depth overflows"))?;
    resolver.set_level(level);
    let mut declarations = 0usize;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let key = attribute.key.as_ref();
        validate_qname(key, limits, "PivotTable XML namespace declaration")?;
        let Some(prefix) = attribute.key.as_namespace_binding() else {
            continue;
        };
        declarations = declarations
            .checked_add(1)
            .ok_or_else(|| invalid("PivotTable XML namespace declaration count overflows"))?;
        if declarations > limits.namespace_declarations {
            return Err(invalid(
                "PivotTable XML namespace declaration count exceeds limit",
            ));
        }
        let prefix_bytes = match prefix {
            PrefixDeclaration::Default => &[][..],
            PrefixDeclaration::Named(prefix) => prefix,
        };
        if prefix_bytes.len() > limits.name_bytes {
            return Err(invalid(
                "PivotTable XML namespace prefix exceeds caller limit",
            ));
        }
        if matches!(prefix, PrefixDeclaration::Named(prefix) if prefix.is_empty()) {
            return Err(invalid(
                "PivotTable XML has an empty prefixed namespace declaration",
            ));
        }
        if attribute.value.as_ref().contains(&b'<') {
            return Err(invalid("PivotTable XML namespace URI contains raw <"));
        }
        let namespace = decode_namespace_uri(attribute.value.as_ref(), limits)?;
        if matches!(prefix, PrefixDeclaration::Named(_)) && namespace.is_empty() {
            return Err(invalid(
                "PivotTable XML cannot undeclare a prefixed namespace",
            ));
        }
        if matches!(prefix, PrefixDeclaration::Default)
            && (namespace.as_slice() == XML_NAMESPACE_URI
                || namespace.as_slice() == XMLNS_NAMESPACE_URI)
        {
            return Err(invalid(
                "PivotTable XML reserved namespace cannot be the default namespace",
            ));
        }
        let namespace = interner.intern(&namespace, limits)?;
        resolver.add(prefix, namespace, limits)?;
    }
    Ok(())
}

fn validate_qname(raw: &[u8], limits: XmlScanLimits, what: &str) -> Result<()> {
    if raw.len() > limits.name_bytes {
        return Err(invalid(format!("{what} name exceeds caller limit")));
    }
    let value = std::str::from_utf8(raw)
        .map_err(|error| invalid(format!("{what} name is not UTF-8: {error}")))?;
    if !is_qualified_name(value) {
        return Err(invalid(format!("{what} name is not a valid QName")));
    }
    Ok(())
}

fn parse_ignorable_namespaces(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    limits: XmlScanLimits,
) -> Result<Vec<NamespaceRef>> {
    let mut namespaces: Vec<NamespaceRef> = Vec::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let (resolved, local) = resolver.resolve_attribute(attribute.key)?;
        if local.as_ref() != b"Ignorable" {
            continue;
        }
        let NamespaceResolution::Bound(namespace) = resolved else {
            continue;
        };
        if namespace.as_ref() != MCE_NS {
            continue;
        }
        if attribute.value.len() > limits.attribute_bytes {
            return Err(invalid("mc:Ignorable exceeds its text limit"));
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?;
        if value.len() > limits.attribute_bytes {
            return Err(invalid("mc:Ignorable exceeds its text limit"));
        }
        for prefix in value.split([' ', '\t', '\r', '\n']) {
            if prefix.is_empty() {
                continue;
            }
            if !litchi_ooxml_common::xml_name::is_ncname(prefix) {
                return Err(invalid("mc:Ignorable contains an invalid prefix"));
            }
            let mut qualified = Vec::new();
            let qualified_len = prefix
                .len()
                .checked_add(2)
                .ok_or_else(|| invalid("mc:Ignorable prefix length overflows"))?;
            qualified
                .try_reserve_exact(qualified_len)
                .map_err(|source| Error::Allocation {
                    resource: "PivotTable MCE ignorable prefix",
                    source,
                })?;
            qualified.extend_from_slice(prefix.as_bytes());
            qualified.extend_from_slice(b":x");
            let resolved = resolver.resolve_element(QName(&qualified))?.0;
            let namespace = match resolved {
                NamespaceResolution::Bound(value) => value,
                NamespaceResolution::Unknown(_) => {
                    return Err(invalid("mc:Ignorable contains an unbound prefix"));
                },
            };
            if namespace.as_ref() == MCE_NS {
                return Err(invalid("mc:Ignorable cannot name the MCE namespace"));
            }
            if !namespaces
                .iter()
                .any(|known| known.as_ref() == namespace.as_ref())
            {
                if namespaces.len() >= limits.namespace_declarations {
                    return Err(invalid(
                        "PivotTable MCE ignorable namespace count exceeds limit",
                    ));
                }
                namespaces
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "PivotTable MCE ignorable namespaces",
                        source,
                    })?;
                namespaces.push(namespace);
            }
        }
    }
    Ok(namespaces)
}

fn extend_ignorable_scope(
    scopes: &mut Vec<IgnorableScope>,
    parent: usize,
    namespaces: Vec<NamespaceRef>,
) -> Result<usize> {
    if namespaces.is_empty() {
        return Ok(parent);
    }
    if parent >= scopes.len() {
        return Err(invalid("PivotTable MCE parent scope is invalid"));
    }
    scopes.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "PivotTable MCE scope table",
        source,
    })?;
    scopes.push(IgnorableScope {
        parent: Some(parent),
        namespaces,
    });
    Ok(scopes.len().saturating_sub(1))
}

fn scope_contains(scopes: &[IgnorableScope], mut scope: usize, namespace: &[u8]) -> bool {
    while let Some(current) = scopes.get(scope) {
        if current
            .namespaces
            .iter()
            .any(|known| known.as_ref() == namespace)
        {
            return true;
        }
        let Some(parent) = current.parent else {
            return false;
        };
        scope = parent;
    }
    false
}

fn resolved_name(
    resolver: &NamespaceResolver,
    name: QName<'_>,
    limits: XmlScanLimits,
) -> Result<(NamespaceRef, Vec<u8>)> {
    validate_qname(name.as_ref(), limits, "PivotTable XML element")?;
    if name.prefix().is_some_and(|prefix| prefix.is_xmlns()) {
        return Err(invalid(
            "PivotTable XML element uses the reserved xmlns prefix",
        ));
    }
    let (resolved, local_name) = resolver.resolve_element(name)?;
    if local_name.as_ref().len() > limits.name_bytes {
        return Err(invalid("PivotTable XML local name exceeds caller limit"));
    }
    let ns = match resolved {
        NamespaceResolution::Bound(value) => value,
        NamespaceResolution::Unknown(prefix) => {
            if prefix.len() > limits.name_bytes {
                return Err(invalid(
                    "PivotTable XML namespace prefix exceeds caller limit",
                ));
            }
            return Err(invalid("unbound XML namespace prefix"));
        },
    };
    let mut local = Vec::new();
    local
        .try_reserve_exact(local_name.as_ref().len())
        .map_err(|source| Error::Allocation {
            resource: "PivotTable XML element local name",
            source,
        })?;
    local.extend_from_slice(local_name.as_ref());
    Ok((ns, local))
}

#[allow(
    deprecated,
    reason = "ST_Xstring attributes retain XML whitespace semantics"
)]
fn parse_attrs(
    bytes: &[u8],
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    start: &Range<usize>,
    limits: XmlScanLimits,
) -> Result<Vec<XmlAttribute>> {
    let mut attrs = Vec::<XmlAttribute>::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let key_bytes = attribute.key.as_ref();
        validate_qname(key_bytes, limits, "PivotTable XML attribute")?;
        if key_bytes.len() > limits.name_bytes {
            return Err(invalid(
                "PivotTable XML qualified attribute name exceeds caller limit",
            ));
        }
        let raw_attribute_bytes = key_bytes
            .len()
            .checked_add(attribute.value.len())
            .ok_or_else(|| invalid("PivotTable XML attribute length overflows"))?;
        if raw_attribute_bytes > MAX_ATTRIBUTE_TEXT_BYTES
            || raw_attribute_bytes > limits.attribute_bytes
        {
            return Err(invalid("PivotTable attribute exceeds its text limit"));
        }
        if let Some(prefix_declaration) = attribute.key.as_namespace_binding() {
            let prefix = match prefix_declaration {
                PrefixDeclaration::Default => &[][..],
                PrefixDeclaration::Named(prefix) => prefix,
            };
            if prefix.len() > limits.name_bytes {
                return Err(invalid(
                    "PivotTable XML namespace prefix exceeds caller limit",
                ));
            }
            if matches!(
                prefix_declaration,
                PrefixDeclaration::Named(prefix) if prefix.is_empty()
            ) {
                return Err(invalid(
                    "PivotTable XML has an empty prefixed namespace declaration",
                ));
            }
            if attribute.value.len() > limits.namespace_bytes {
                return Err(invalid("PivotTable XML namespace URI exceeds caller limit"));
            }
            continue;
        }
        if attribute
            .key
            .prefix()
            .is_some_and(|prefix| prefix.is_xmlns())
        {
            return Err(invalid(
                "PivotTable XML attribute uses the reserved xmlns prefix",
            ));
        }
        if attrs.len() >= MAX_NAMESPACE_DECLARATIONS {
            return Err(invalid("PivotTable XML attribute count exceeds limit"));
        }
        let (resolved, local_name) = resolver.resolve_attribute(attribute.key)?;
        if local_name.as_ref().len() > limits.name_bytes {
            return Err(invalid(
                "PivotTable XML attribute local name exceeds caller limit",
            ));
        }
        let mut owned_local_name = Vec::new();
        owned_local_name
            .try_reserve_exact(local_name.as_ref().len())
            .map_err(|source| Error::Allocation {
                resource: "PivotTable XML attribute local name",
                source,
            })?;
        owned_local_name.extend_from_slice(local_name.as_ref());
        let local_name = owned_local_name;
        let ns = match resolved {
            NamespaceResolution::Bound(value) => value,
            NamespaceResolution::Unknown(prefix) => {
                if prefix.len() > limits.name_bytes {
                    return Err(invalid(
                        "PivotTable XML attribute namespace prefix exceeds caller limit",
                    ));
                }
                return Err(invalid("unbound XML attribute prefix"));
            },
        };
        if attrs
            .iter()
            .any(|known| known.ns == ns && known.local == local_name)
        {
            return Err(invalid(
                "PivotTable XML contains duplicate expanded attributes",
            ));
        }
        if attribute.value.as_ref().contains(&b'<') {
            return Err(invalid("PivotTable XML attribute contains raw <"));
        }
        let value = attribute
            .decode_and_unescape_value(decoder)
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        validate_xml_characters(value.as_bytes())?;
        if value.len() > MAX_ATTRIBUTE_TEXT_BYTES || value.len() > limits.attribute_bytes {
            return Err(invalid("PivotTable attribute exceeds its text limit"));
        }
        let value_range = value_span(bytes, attribute.value.as_ref())?;
        let key_start = ptr_offset(bytes, attribute.key.0.as_ptr())?;
        let value_end = value_range.end;
        if key_start < start.start || value_end > start.end {
            return Err(invalid(
                "PivotTable attribute source range is outside its start tag",
            ));
        }
        let mut whole_start = key_start;
        while whole_start > start.start
            && matches!(bytes[whole_start - 1], b' ' | b'\t' | b'\r' | b'\n')
        {
            whole_start -= 1;
        }
        attrs.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "PivotTable XML attributes",
            source,
        })?;
        attrs.push(XmlAttribute {
            ns,
            local: local_name,
            value,
            value_range,
            whole_range: whole_start..value_end.saturating_add(1),
        });
    }
    Ok(attrs)
}

fn decode_namespace_uri(value: &[u8], limits: XmlScanLimits) -> Result<Vec<u8>> {
    if value.len() > limits.namespace_bytes {
        return Err(invalid("PivotTable XML namespace URI exceeds caller limit"));
    }
    let value = std::str::from_utf8(value).map_err(|error| {
        invalid(format!(
            "PivotTable XML namespace URI is not UTF-8: {error}"
        ))
    })?;
    let decoded = quick_xml::escape::unescape(value)
        .map_err(|error| invalid(format!("PivotTable XML namespace URI is invalid: {error}")))?;
    validate_xml_characters(decoded.as_bytes())?;
    if decoded.len() > limits.namespace_bytes {
        return Err(invalid("PivotTable XML namespace URI exceeds caller limit"));
    }
    let mut result = Vec::new();
    result
        .try_reserve_exact(decoded.len())
        .map_err(|source| Error::Allocation {
            resource: "PivotTable XML namespace URI",
            source,
        })?;
    result.extend_from_slice(decoded.as_bytes());
    Ok(result)
}

fn ptr_offset(bytes: &[u8], ptr: *const u8) -> Result<usize> {
    let start = bytes.as_ptr() as usize;
    let ptr = ptr as usize;
    ptr.checked_sub(start)
        .filter(|offset| *offset <= bytes.len())
        .ok_or_else(|| invalid("PivotTable XML source pointer is not backed by the Part"))
}

fn event_start_range(
    bytes: &[u8],
    before: usize,
    after: usize,
    element: &BytesStart<'_>,
) -> Result<Range<usize>> {
    let raw = element.as_ref();
    for extra in 0..=2 {
        let Some(start) = after.checked_sub(raw.len().saturating_add(extra)) else {
            continue;
        };
        let Some(end) = start.checked_add(raw.len()) else {
            continue;
        };
        if end <= bytes.len() && bytes.get(start..end) == Some(raw) {
            let candidate_end = if bytes.get(end) == Some(&b'>') {
                end + 1
            } else if bytes.get(end..end.saturating_add(2)) == Some(b"/>") {
                end + 2
            } else {
                end
            };
            if start >= before.saturating_sub(2) && candidate_end <= after.saturating_add(1) {
                return Ok(start..candidate_end);
            }
        }
    }
    // The reader's buffer position is after the closing `>`; fall back to the
    // nearest `<` in the event window while still checking the raw token.
    let lower = before.saturating_sub(raw.len().saturating_add(3));
    for start in (lower..after.min(bytes.len())).rev() {
        if bytes.get(start) == Some(&b'<')
            && bytes.get(start..start.saturating_add(raw.len())) == Some(raw)
        {
            let end = start
                + raw.len()
                + if bytes.get(start + raw.len()) == Some(&b'>') {
                    1
                } else if bytes.get(start + raw.len()..start + raw.len() + 2) == Some(b"/>") {
                    2
                } else {
                    0
                };
            return Ok(start..end);
        }
    }
    Err(invalid("PivotTable XML event has no source range"))
}
