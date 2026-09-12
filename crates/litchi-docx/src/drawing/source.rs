//! Bounded, source-backed inspection of direct WordprocessingML pictures.
//!
//! The ordinary drawing inventory is intentionally lossy.  This module keeps
//! the immutable story source in every independently usable record and stores
//! only the ranges, namespace contexts, and relationship metadata required by
//! the future SVG host transaction.  It does not mutate a package.

use std::borrow::Cow;
use std::fmt;
use std::hash::{Hash, Hasher};
use std::sync::Arc;

use litchi_drawingml::svg_blip::{self, Namespace as SvgNamespace, SvgBlip};
use litchi_ooxml_common::{xml::xsd_token_atom, xml_name::is_ncname};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesDecl, BytesStart, Event, attributes::Attribute};
use quick_xml::name::{Namespace, PrefixDeclaration, QName, ResolveResult};
use quick_xml::reader::{NsReader, Reader};

use crate::error::{Error, Result};
use crate::namespace::{STRICT_WORDPROCESSINGML_NAMESPACE, WORDPROCESSINGML_NAMESPACE};

const STRICT_WORDPROCESSING_DRAWING: &[u8] =
    b"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
const WORDPROCESSING_DRAWING: &[u8] =
    b"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const STRICT_DRAWINGML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const DRAWINGML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_PICTURE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/picture";
const PICTURE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/picture";
const RELATIONSHIPS: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const SVG_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const SVG_EXTENSION_URI: &[u8] = b"{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &[u8] = b"http://www.w3.org/2000/xmlns/";

/// Namespace identity used by the scanner.  Resolver output is reduced to a
/// small static tag before any mutable reader operation; this keeps the event
/// and resolver borrows live without cloning arbitrary namespace bytes for
/// every XML node.  Ranges only retain URI bytes for admitted/known elements.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Resolved {
    Unbound,
    Known(&'static [u8]),
    Other,
    Unknown,
}

fn resolved_tag(value: ResolveResult<'_>, decoder: Decoder) -> Result<Resolved> {
    match value {
        ResolveResult::Unbound => Ok(Resolved::Unbound),
        ResolveResult::Unknown(_) => Ok(Resolved::Unknown),
        ResolveResult::Bound(Namespace(value)) => {
            let value = decode_namespace_uri(value, decoder)?;
            Ok(match known_namespace(value.as_bytes()) {
                Some(value) => Resolved::Known(value),
                None => Resolved::Other,
            })
        },
    }
}

/// Decode a namespace declaration or resolver value using XML attribute
/// normalization. `quick_xml` intentionally leaves namespace entity
/// references lexical, while XML namespace identity is defined after entity
/// expansion; this keeps the resolver's borrowed source and bounds intact.
fn decode_namespace_uri<'a>(value: &'a [u8], decoder: Decoder) -> Result<Cow<'a, str>> {
    Attribute {
        key: QName(b"xmlns"),
        value: Cow::Borrowed(value),
    }
    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
    .map_err(xml_error)
}

fn known_namespace(value: &[u8]) -> Option<&'static [u8]> {
    [
        WORDPROCESSINGML_NAMESPACE,
        STRICT_WORDPROCESSINGML_NAMESPACE,
        WORDPROCESSING_DRAWING,
        STRICT_WORDPROCESSING_DRAWING,
        DRAWINGML,
        STRICT_DRAWINGML,
        PICTURE,
        STRICT_PICTURE,
        RELATIONSHIPS,
        STRICT_RELATIONSHIPS,
        SVG_NAMESPACE,
        MCE,
    ]
    .into_iter()
    .find(|candidate| *candidate == value)
}

/// Maximum source XML bytes inspected by this helper.
pub const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
/// Maximum XML events inspected by one scan.
pub const MAX_XML_NODES: usize = 1_000_000;
/// Maximum XML nesting depth inspected by one scan.
pub const MAX_XML_DEPTH: usize = 256;
/// Maximum direct pictures retained by one scan.
pub const MAX_PICTURES: usize = 100_000;
/// Maximum relationship attributes retained for one picture.
pub const MAX_RELATIONSHIP_REFERENCES: usize = 4_096;
/// Maximum relationship attributes retained for the whole story source.
pub const MAX_GLOBAL_RELATIONSHIP_REFERENCES: usize = 1_000_000;
/// Maximum qualified-name bytes inspected before storage.
pub const MAX_NAME_BYTES: usize = 4 * 1024;
/// Maximum namespace declaration URI/prefix bytes per scan context.
pub const MAX_NAMESPACE_BYTES: usize = 64 * 1024;
/// Maximum active namespace declarations per element scope.
pub const MAX_ACTIVE_NAMESPACE_BINDINGS: usize = 256;
/// Maximum declarations on one element.
pub const MAX_NAMESPACE_DECLARATIONS: usize = 256;
/// Maximum attributes on one element.
pub const MAX_ATTRIBUTES: usize = 256;
/// Maximum raw attribute value bytes inspected.
pub const MAX_ATTRIBUTE_VALUE_BYTES: usize = 1024 * 1024;
/// Maximum relationship identifier bytes after XML decoding.
pub const MAX_RELATIONSHIP_ID_BYTES: usize = 256;
/// Maximum standalone SVG fragment output bytes.
pub const MAX_FRAGMENT_BYTES: usize = 16 * 1024 * 1024;

/// A half-open byte range into the immutable source story member.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[must_use]
pub struct ByteRange {
    /// Inclusive start offset.
    pub start: usize,
    /// Exclusive end offset.
    pub end: usize,
}

impl ByteRange {
    /// Construct a range.  Bounds are checked when it is borrowed from source.
    pub const fn new(start: usize, end: usize) -> Self {
        Self { start, end }
    }

    /// Return the range length if its endpoints are ordered.
    pub fn len(self) -> Result<usize> {
        self.end
            .checked_sub(self.start)
            .ok_or_else(|| invalid("DOCX drawing source range has descending endpoints"))
    }

    /// Whether the range has equal ordered endpoints.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.start == self.end
    }

    /// Borrow the range from a source member.
    pub fn slice(self, source: &[u8]) -> Result<&[u8]> {
        source
            .get(self.start..self.end)
            .ok_or_else(|| invalid("DOCX drawing source range lies outside its source"))
    }
}

/// An element range retaining the source that owns it.
#[derive(Clone)]
#[must_use]
pub struct ElementRange<'a> {
    source: &'a [u8],
    range: ByteRange,
    start_end: usize,
    close_start: Option<usize>,
    prefix: Box<[u8]>,
    namespace_uri: Box<[u8]>,
}

fn same_source(left: &[u8], right: &[u8]) -> bool {
    std::ptr::eq(left.as_ptr(), right.as_ptr()) && left.len() == right.len()
}

fn hash_source<H: Hasher>(source: &[u8], state: &mut H) {
    state.write_usize(source.as_ptr() as usize);
    state.write_usize(source.len());
}

impl fmt::Debug for ElementRange<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ElementRange")
            .field("source_len", &self.source.len())
            .field("range", &self.range)
            .field("start_end", &self.start_end)
            .field("close_start", &self.close_start)
            .field("prefix", &self.prefix)
            .field("namespace_uri", &self.namespace_uri)
            .finish()
    }
}

impl PartialEq for ElementRange<'_> {
    fn eq(&self, other: &Self) -> bool {
        same_source(self.source, other.source)
            && self.range == other.range
            && self.start_end == other.start_end
            && self.close_start == other.close_start
            && self.prefix == other.prefix
            && self.namespace_uri == other.namespace_uri
    }
}

impl Eq for ElementRange<'_> {}

impl Hash for ElementRange<'_> {
    fn hash<H: Hasher>(&self, state: &mut H) {
        hash_source(self.source, state);
        self.range.hash(state);
        self.start_end.hash(state);
        self.close_start.hash(state);
        self.prefix.hash(state);
        self.namespace_uri.hash(state);
    }
}

impl<'a> ElementRange<'a> {
    /// Complete element range.
    pub const fn range(&self) -> ByteRange {
        self.range
    }

    /// Opening-tag end offset.
    #[must_use]
    pub const fn start_end(&self) -> usize {
        self.start_end
    }

    /// Closing-tag start offset for paired elements.
    #[must_use]
    pub const fn close_start(&self) -> Option<usize> {
        self.close_start
    }

    /// Prefix bytes in the source element name.
    #[must_use]
    pub fn prefix(&self) -> &[u8] {
        &self.prefix
    }

    /// Expanded namespace URI bytes.
    #[must_use]
    pub fn namespace_uri(&self) -> &[u8] {
        &self.namespace_uri
    }

    /// Borrow exact element bytes.
    pub fn bytes(&self) -> Result<&'a [u8]> {
        self.range.slice(self.source)
    }

    /// Whether the source used a self-closing element.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.close_start.is_none()
    }
}

/// Core WordprocessingML/DrawingML namespace profile detected in the source.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[non_exhaustive]
pub enum DrawingDialect {
    /// Transitional ECMA-376 namespace family.
    Transitional,
    /// ISO/IEC 29500 Strict namespace family.
    Strict,
}

/// Physical relationship namespace observed on a source `r:*` attribute.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[non_exhaustive]
pub enum RelationshipDialect {
    /// Transitional Office relationship namespace.
    Transitional,
    /// Strict Office relationship namespace.
    Strict,
}

impl RelationshipDialect {
    /// Namespace required for a newly authored MS-ODRAWXML SVG attribute.
    ///
    /// The unmodified local §5.24 schema imports the Transitional `AG_Blob`
    /// relationship vocabulary even when the host's core markup is Strict.
    #[must_use]
    pub const fn authored_svg_attribute_namespace() -> &'static str {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    }

    /// Physical relationship namespace for this host profile.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Transitional => {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            },
            Self::Strict => "http://purl.oclc.org/ooxml/officeDocument/relationships",
        }
    }
}

/// A persistent namespace scope shared by source-backed SVG owners.
pub use litchi_drawingml::svg_blip::NamespaceContext;

/// A relationship-bearing source attribute.
#[derive(Clone)]
#[must_use]
pub struct RelationshipReference<'a> {
    source: &'a [u8],
    id: Arc<str>,
    local_name: Arc<str>,
    range: ByteRange,
    dialect: RelationshipDialect,
}

impl fmt::Debug for RelationshipReference<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("RelationshipReference")
            .field("source_len", &self.source.len())
            .field("id", &self.id)
            .field("local_name", &self.local_name)
            .field("range", &self.range)
            .field("dialect", &self.dialect)
            .finish()
    }
}

impl PartialEq for RelationshipReference<'_> {
    fn eq(&self, other: &Self) -> bool {
        same_source(self.source, other.source)
            && self.id == other.id
            && self.local_name == other.local_name
            && self.range == other.range
            && self.dialect == other.dialect
    }
}

impl Eq for RelationshipReference<'_> {}

impl Hash for RelationshipReference<'_> {
    fn hash<H: Hasher>(&self, state: &mut H) {
        hash_source(self.source, state);
        self.id.hash(state);
        self.local_name.hash(state);
        self.range.hash(state);
        self.dialect.hash(state);
    }
}

impl<'a> RelationshipReference<'a> {
    /// Borrow the source member retaining this reference.
    #[must_use]
    pub const fn source(&self) -> &'a [u8] {
        self.source
    }

    /// Decoded relationship identifier.
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Attribute local name, usually `embed` or `link`.
    #[must_use]
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Source opening-tag range containing the relationship attribute.
    pub const fn range(&self) -> ByteRange {
        self.range
    }

    /// Relationship namespace dialect used by the attribute.
    #[must_use]
    pub const fn dialect(&self) -> RelationshipDialect {
        self.dialect
    }
}

/// A direct SVG extension owner with exact source ranges.
#[derive(Clone)]
#[must_use]
pub struct SvgOwner<'a> {
    source: &'a [u8],
    extension: ElementRange<'a>,
    svg_blip: ElementRange<'a>,
    uri_lexical: Box<[u8]>,
    value: SvgBlip,
    relationship_dialect: RelationshipDialect,
    namespace_context: NamespaceContext,
}

impl fmt::Debug for SvgOwner<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SvgOwner")
            .field("source_len", &self.source.len())
            .field("extension", &self.extension)
            .field("svg_blip", &self.svg_blip)
            .field("uri_lexical", &self.uri_lexical)
            .field("value", &self.value)
            .field("relationship_dialect", &self.relationship_dialect)
            .finish()
    }
}

impl PartialEq for SvgOwner<'_> {
    fn eq(&self, other: &Self) -> bool {
        same_source(self.source, other.source)
            && self.extension == other.extension
            && self.svg_blip == other.svg_blip
            && self.uri_lexical == other.uri_lexical
            && self.value == other.value
            && self.relationship_dialect == other.relationship_dialect
            && self.namespace_context == other.namespace_context
    }
}

impl Eq for SvgOwner<'_> {}

impl<'a> SvgOwner<'a> {
    /// Borrow the source member retaining this owner.
    #[must_use]
    pub const fn source(&self) -> &'a [u8] {
        self.source
    }

    /// Exact `a:ext` range.
    pub const fn extension_range(&self) -> ByteRange {
        self.extension.range()
    }

    /// Complete `a:ext` source metadata.
    pub const fn extension_element(&self) -> &ElementRange<'a> {
        &self.extension
    }

    /// Exact `asvg:svgBlip` source metadata.
    pub const fn svg_blip_element(&self) -> &ElementRange<'a> {
        &self.svg_blip
    }

    /// Raw lexical bytes of the recognized `uri` attribute value.
    #[must_use]
    pub fn uri_lexical(&self) -> &[u8] {
        &self.uri_lexical
    }

    /// Parsed shared SVG-blip value.
    pub const fn value(&self) -> &SvgBlip {
        &self.value
    }

    /// Embedded relationship ID, if the owner uses `r:embed`.
    #[must_use]
    pub fn embedded_relationship_id(&self) -> Option<&str> {
        self.value.embedded().map(|id| id.as_str())
    }

    /// Linked relationship ID, if the owner uses `r:link`.
    #[must_use]
    pub fn linked_relationship_id(&self) -> Option<&str> {
        self.value.linked().map(|id| id.as_str())
    }

    /// Relationship namespace used by the physical SVG attribute.
    #[must_use]
    pub const fn relationship_dialect(&self) -> RelationshipDialect {
        self.relationship_dialect
    }

    /// Captured active namespace bindings at the extension owner.
    pub fn namespace_context(&self) -> &NamespaceContext {
        &self.namespace_context
    }

    /// Return a standalone fragment with inherited namespace declarations
    /// injected before the root's closing delimiter.  The shared DrawingML
    /// codec retains the raw contextual fragment and performs this completion
    /// lazily, so the host never stores a flattened copy per owner.
    pub fn namespace_complete(&self, max_output_bytes: usize) -> Result<Vec<u8>> {
        svg_blip::write_contextual(&self.value, max_output_bytes).map_err(Into::into)
    }
}

/// Typed status of the direct extension list on one picture.
#[derive(Clone, Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum SvgOwnerState<'a> {
    /// No direct extension was found.
    None,
    /// One valid embedded SVG owner was found.
    Embedded(SvgOwner<'a>),
    /// One valid linked SVG owner was found; mutation must remain inert.
    Linked(SvgOwner<'a>),
    /// More than one admitted owner was found.
    Ambiguous,
    /// An admitted owner is malformed, MCE-wrapped, or otherwise unsafe.
    Refused,
    /// Unknown extension content is present but no owner was inferred.
    Opaque,
}

impl<'a> SvgOwnerState<'a> {
    /// Borrow an embedded owner suitable for a future host mutation.
    #[must_use]
    pub fn owner(&self) -> Option<&SvgOwner<'a>> {
        match self {
            Self::Embedded(owner) => Some(owner),
            Self::None | Self::Linked(_) | Self::Ambiguous | Self::Refused | Self::Opaque => None,
        }
    }

    /// Whether a host mutation must refuse this state.
    #[must_use]
    pub const fn is_refused(&self) -> bool {
        matches!(self, Self::Linked(_) | Self::Ambiguous | Self::Refused)
    }
}

/// Placement of one direct WordprocessingML drawing.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[must_use]
pub enum DrawingPlacement {
    /// `wp:inline` placement.
    Inline,
    /// Floating `wp:anchor` placement.
    Floating,
}

/// One borrowed direct WordprocessingML picture record.
#[derive(Clone)]
#[must_use]
pub struct PictureSource<'a> {
    source: &'a [u8],
    drawing_ordinal: usize,
    picture_ordinal: usize,
    drawing_range: ElementRange<'a>,
    anchor_range: ElementRange<'a>,
    placement: DrawingPlacement,
    graphic_range: ElementRange<'a>,
    graphic_data_range: ElementRange<'a>,
    picture_range: ElementRange<'a>,
    blip_range: Option<ElementRange<'a>>,
    ext_list_range: Option<ElementRange<'a>>,
    c_nv_pr_id: Option<Box<str>>,
    c_nv_pr_range: Option<ElementRange<'a>>,
    raster_relationship_id: Option<Box<str>>,
    raster_relationship_is_link: bool,
    raster_relationship_dialect: Option<RelationshipDialect>,
    svg_owner: SvgOwnerState<'a>,
    relationship_references: Vec<RelationshipReference<'a>>,
    namespace_context: NamespaceContext,
    blip_namespace_context: Option<NamespaceContext>,
    ext_list_namespace_context: Option<NamespaceContext>,
    opaque_svg_extension: bool,
}

impl fmt::Debug for PictureSource<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PictureSource")
            .field("source_len", &self.source.len())
            .field("drawing_ordinal", &self.drawing_ordinal)
            .field("picture_ordinal", &self.picture_ordinal)
            .field("drawing_range", &self.drawing_range)
            .field("anchor_range", &self.anchor_range)
            .field("placement", &self.placement)
            .field("graphic_range", &self.graphic_range)
            .field("graphic_data_range", &self.graphic_data_range)
            .field("picture_range", &self.picture_range)
            .field("blip_range", &self.blip_range)
            .field("ext_list_range", &self.ext_list_range)
            .field("c_nv_pr_id", &self.c_nv_pr_id)
            .field("c_nv_pr_range", &self.c_nv_pr_range)
            .field("raster_relationship_id", &self.raster_relationship_id)
            .field(
                "raster_relationship_is_link",
                &self.raster_relationship_is_link,
            )
            .field(
                "raster_relationship_dialect",
                &self.raster_relationship_dialect,
            )
            .field("svg_owner", &self.svg_owner)
            .field("opaque_svg_extension", &self.opaque_svg_extension)
            .field("relationship_references", &self.relationship_references)
            .finish()
    }
}

impl PartialEq for PictureSource<'_> {
    fn eq(&self, other: &Self) -> bool {
        same_source(self.source, other.source)
            && self.drawing_ordinal == other.drawing_ordinal
            && self.picture_ordinal == other.picture_ordinal
            && self.drawing_range == other.drawing_range
            && self.anchor_range == other.anchor_range
            && self.placement == other.placement
            && self.graphic_range == other.graphic_range
            && self.graphic_data_range == other.graphic_data_range
            && self.picture_range == other.picture_range
            && self.blip_range == other.blip_range
            && self.ext_list_range == other.ext_list_range
            && self.c_nv_pr_id == other.c_nv_pr_id
            && self.c_nv_pr_range == other.c_nv_pr_range
            && self.raster_relationship_id == other.raster_relationship_id
            && self.raster_relationship_is_link == other.raster_relationship_is_link
            && self.raster_relationship_dialect == other.raster_relationship_dialect
            && self.svg_owner == other.svg_owner
            && self.relationship_references == other.relationship_references
            && self.namespace_context == other.namespace_context
            && self.blip_namespace_context == other.blip_namespace_context
            && self.ext_list_namespace_context == other.ext_list_namespace_context
            && self.opaque_svg_extension == other.opaque_svg_extension
    }
}

impl Eq for PictureSource<'_> {}

impl<'a> PictureSource<'a> {
    /// Borrow the source member retaining this picture.
    #[must_use]
    pub const fn source(&self) -> &'a [u8] {
        self.source
    }

    /// Drawing ordinal supplied to the scanner.
    #[must_use]
    pub const fn drawing_ordinal(&self) -> usize {
        self.drawing_ordinal
    }

    /// Source-order ordinal among direct pictures.
    #[must_use]
    pub const fn picture_ordinal(&self) -> usize {
        self.picture_ordinal
    }

    /// Inline or floating placement.
    pub const fn placement(&self) -> DrawingPlacement {
        self.placement
    }

    /// Complete `w:drawing` range.
    pub const fn drawing_range(&self) -> ByteRange {
        self.drawing_range.range()
    }

    /// Complete `wp:inline` or `wp:anchor` range.
    pub const fn anchor_range(&self) -> ByteRange {
        self.anchor_range.range()
    }

    /// Complete `a:graphic` range.
    pub const fn graphic_range(&self) -> ByteRange {
        self.graphic_range.range()
    }

    /// Complete `a:graphicData` range.
    pub const fn graphic_data_range(&self) -> ByteRange {
        self.graphic_data_range.range()
    }

    /// Complete direct `pic:pic` range.
    pub const fn picture_range(&self) -> ByteRange {
        self.picture_range.range()
    }

    /// Direct raster blip range, when present.
    #[must_use]
    pub const fn blip_range(&self) -> Option<&ElementRange<'a>> {
        self.blip_range.as_ref()
    }

    /// Direct extension-list range, when present.
    #[must_use]
    pub const fn ext_list_range(&self) -> Option<&ElementRange<'a>> {
        self.ext_list_range.as_ref()
    }

    /// Optional nonvisual picture ID.
    #[must_use]
    pub fn c_nv_pr_id(&self) -> Option<&str> {
        self.c_nv_pr_id.as_deref()
    }

    /// Optional complete `pic:cNvPr` range.
    #[must_use]
    pub const fn c_nv_pr_range(&self) -> Option<&ElementRange<'a>> {
        self.c_nv_pr_range.as_ref()
    }

    /// Existing raster relationship ID, if the direct blip has one.
    #[must_use]
    pub fn raster_relationship_id(&self) -> Option<&str> {
        self.raster_relationship_id.as_deref()
    }

    /// Whether the direct raster blip uses `r:link` instead of `r:embed`.
    #[must_use]
    pub const fn raster_relationship_is_link(&self) -> bool {
        self.raster_relationship_is_link
    }

    /// Physical relationship namespace used by the direct raster blip.
    #[must_use]
    pub const fn raster_relationship_dialect(&self) -> Option<RelationshipDialect> {
        self.raster_relationship_dialect
    }

    /// SVG owner state, including opaque and refusal statuses.
    #[must_use]
    pub const fn svg_owner(&self) -> &SvgOwnerState<'a> {
        &self.svg_owner
    }

    /// All relationship attributes below this direct picture.
    pub fn relationship_references(&self) -> &[RelationshipReference<'a>] {
        &self.relationship_references
    }

    /// Borrow exact picture bytes.
    pub fn picture_bytes(&self) -> Result<&'a [u8]> {
        self.picture_range.bytes()
    }

    /// Borrow the exact containing `w:drawing` bytes.
    pub fn drawing_bytes(&self) -> Result<&'a [u8]> {
        self.drawing_range.bytes()
    }

    /// Borrow exact anchor bytes.
    pub fn anchor_bytes(&self) -> Result<&'a [u8]> {
        self.anchor_range.bytes()
    }

    /// Borrow the exact containing `a:graphic` bytes.
    pub fn graphic_bytes(&self) -> Result<&'a [u8]> {
        self.graphic_range.bytes()
    }

    /// Borrow the exact containing `a:graphicData` bytes.
    pub fn graphic_data_bytes(&self) -> Result<&'a [u8]> {
        self.graphic_data_range.bytes()
    }

    /// Borrow the exact direct raster `a:blip` bytes.
    pub fn blip_bytes(&self) -> Result<Option<&'a [u8]>> {
        self.blip_range
            .as_ref()
            .map(ElementRange::bytes)
            .transpose()
    }

    /// Borrow the exact direct `a:extLst` bytes.
    pub fn ext_list_bytes(&self) -> Result<Option<&'a [u8]>> {
        self.ext_list_range
            .as_ref()
            .map(ElementRange::bytes)
            .transpose()
    }

    /// Borrow the complete source range.
    pub fn source_bytes(&self, range: ByteRange) -> Result<&'a [u8]> {
        range.slice(self.source)
    }

    /// Captured namespace context at the picture root.
    pub fn namespace_context(&self) -> &NamespaceContext {
        &self.namespace_context
    }

    /// Captured namespace context at the direct raster blip.
    #[must_use]
    pub fn blip_namespace_context(&self) -> Option<&NamespaceContext> {
        self.blip_namespace_context.as_ref()
    }

    /// Captured namespace context at the direct extension list.
    #[must_use]
    pub fn ext_list_namespace_context(&self) -> Option<&NamespaceContext> {
        self.ext_list_namespace_context.as_ref()
    }

    /// Whether this picture retains an unknown direct SVG extension alongside
    /// a recognized owner.  A detach therefore leaves an opaque owner state
    /// instead of manufacturing an empty `None` state without rescanning the
    /// source story.
    #[must_use]
    pub const fn has_opaque_svg_extension(&self) -> bool {
        self.opaque_svg_extension
    }
}

/// Complete source-backed main-story drawing inventory.
///
/// Every returned picture retains a borrow of the scanned source member.  A
/// caller therefore cannot mutate or reuse the input allocation while a
/// picture record remains live:
///
/// ```compile_fail
/// use litchi_docx::drawing::SourceDrawing;
///
/// fn main() {
///     let mut xml = Vec::new();
///     let drawing = SourceDrawing::scan(&xml).unwrap();
///     let picture = drawing.picture(0).unwrap().clone();
///     drop(drawing);
///     xml.clear();
///     let _ = picture.picture_bytes();
/// }
/// ```
#[derive(Clone)]
#[must_use]
pub struct SourceDrawing<'a> {
    source: &'a [u8],
    drawing_ordinal: usize,
    dialect: DrawingDialect,
    relationship_dialect: RelationshipDialect,
    pictures: Vec<PictureSource<'a>>,
    relationship_references: Vec<RelationshipReference<'a>>,
}

impl fmt::Debug for SourceDrawing<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceDrawing")
            .field("source_len", &self.source.len())
            .field("drawing_ordinal", &self.drawing_ordinal)
            .field("dialect", &self.dialect)
            .field("relationship_dialect", &self.relationship_dialect)
            .field("pictures", &self.pictures)
            .field("relationship_references", &self.relationship_references)
            .finish()
    }
}

impl PartialEq for SourceDrawing<'_> {
    fn eq(&self, other: &Self) -> bool {
        same_source(self.source, other.source)
            && self.drawing_ordinal == other.drawing_ordinal
            && self.dialect == other.dialect
            && self.relationship_dialect == other.relationship_dialect
            && self.pictures == other.pictures
            && self.relationship_references == other.relationship_references
    }
}

impl Eq for SourceDrawing<'_> {}

impl<'a> SourceDrawing<'a> {
    /// Scan a complete main-story XML member.
    pub fn scan(source: &'a [u8]) -> Result<Self> {
        Self::scan_with_limits(source, 0, ScanLimits::default())
    }

    /// Scan with a caller-supplied semantic drawing ordinal.
    pub fn scan_with_ordinal(source: &'a [u8], drawing_ordinal: usize) -> Result<Self> {
        Self::scan_with_limits(source, drawing_ordinal, ScanLimits::default())
    }

    /// Scan with explicit finite bounds.
    pub fn scan_with_limits(
        source: &'a [u8],
        drawing_ordinal: usize,
        limits: ScanLimits,
    ) -> Result<Self> {
        Scanner::new(source, drawing_ordinal, limits)?.run()
    }

    /// Borrow the exact source member.
    #[must_use]
    pub const fn source(&self) -> &'a [u8] {
        self.source
    }

    /// Drawing ordinal supplied by the caller.
    #[must_use]
    pub const fn drawing_ordinal(&self) -> usize {
        self.drawing_ordinal
    }

    /// Core namespace dialect detected from the document/drawing elements.
    #[must_use]
    pub const fn dialect(&self) -> DrawingDialect {
        self.dialect
    }

    /// Host physical relationship dialect selected from the core document
    /// profile.  Individual SVG child attributes expose their own physical
    /// dialect through [`SvgOwner::relationship_dialect`].
    #[must_use]
    pub const fn relationship_dialect(&self) -> RelationshipDialect {
        self.relationship_dialect
    }

    /// Direct source pictures in document order.
    pub fn pictures(&self) -> &[PictureSource<'a>] {
        &self.pictures
    }

    /// Every relationship-bearing attribute in the source story.
    pub fn relationship_references(&self) -> &[RelationshipReference<'a>] {
        &self.relationship_references
    }

    /// Resolve a checked source-order picture selector.
    pub fn picture(&self, ordinal: usize) -> Result<&PictureSource<'a>> {
        self.pictures.get(ordinal).ok_or(Error::OutOfBounds {
            object: "DOCX drawing picture",
            index: ordinal,
            len: self.pictures.len(),
        })
    }
}

/// Explicit finite source-scan limits.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[must_use]
pub struct ScanLimits {
    /// Maximum source XML bytes.
    pub max_xml_bytes: usize,
    /// Maximum XML events.
    pub max_nodes: usize,
    /// Maximum XML depth.
    pub max_depth: usize,
    /// Maximum direct pictures.
    pub max_pictures: usize,
    /// Maximum relationship attributes per picture.
    pub max_relationship_references: usize,
    /// Maximum relationship attributes in the complete source.
    pub max_global_relationship_references: usize,
    /// Maximum active namespace binding entries.
    pub max_active_namespace_bindings: usize,
    /// Maximum aggregate prefix/URI bytes admitted to persistent namespace
    /// context nodes. Each declaration is charged once when its scope node is
    /// constructed; retained owner handles share ancestor nodes and do not
    /// charge those bytes again.
    pub max_namespace_bytes: usize,
    /// Maximum raw SVG fragment bytes admitted into contextual owner storage.
    /// Standalone namespace completion takes a separate explicit output cap.
    pub max_fragment_bytes: usize,
}

impl Default for ScanLimits {
    fn default() -> Self {
        Self {
            max_xml_bytes: MAX_XML_BYTES,
            max_nodes: MAX_XML_NODES,
            max_depth: MAX_XML_DEPTH,
            max_pictures: MAX_PICTURES,
            max_relationship_references: MAX_RELATIONSHIP_REFERENCES,
            max_global_relationship_references: MAX_GLOBAL_RELATIONSHIP_REFERENCES,
            max_active_namespace_bindings: MAX_ACTIVE_NAMESPACE_BINDINGS,
            max_namespace_bytes: MAX_NAMESPACE_BYTES,
            max_fragment_bytes: MAX_FRAGMENT_BYTES,
        }
    }
}

impl ScanLimits {
    fn validate(self) -> Result<Self> {
        for (name, value, hard) in [
            ("max_xml_bytes", self.max_xml_bytes, MAX_XML_BYTES),
            ("max_nodes", self.max_nodes, MAX_XML_NODES),
            ("max_depth", self.max_depth, MAX_XML_DEPTH),
            ("max_pictures", self.max_pictures, MAX_PICTURES),
            (
                "max_relationship_references",
                self.max_relationship_references,
                MAX_RELATIONSHIP_REFERENCES,
            ),
            (
                "max_global_relationship_references",
                self.max_global_relationship_references,
                MAX_GLOBAL_RELATIONSHIP_REFERENCES,
            ),
            (
                "max_active_namespace_bindings",
                self.max_active_namespace_bindings,
                MAX_ACTIVE_NAMESPACE_BINDINGS,
            ),
            (
                "max_namespace_bytes",
                self.max_namespace_bytes,
                MAX_NAMESPACE_BYTES,
            ),
            (
                "max_fragment_bytes",
                self.max_fragment_bytes,
                MAX_FRAGMENT_BYTES,
            ),
        ] {
            if value == 0 || value > hard {
                return Err(limit(name, value, hard));
            }
        }
        Ok(self)
    }
}

#[derive(Clone, Debug)]
struct Frame<'a> {
    kind: Kind,
    source_start: usize,
    range: ElementRange<'a>,
    namespace_declarations: usize,
    parent_namespace_context: NamespaceContext,
    picture_index: Option<usize>,
    extension_index: Option<usize>,
    blocks_drawing: bool,
}

#[derive(Clone, Debug)]
struct PendingExtension<'a> {
    range: ElementRange<'a>,
    admitted: bool,
    malformed: bool,
    uri_lexical: Box<[u8]>,
    svg_blip: Option<ElementRange<'a>>,
    svg_context: Option<NamespaceContext>,
    svg_count: usize,
    mce_ancestor: bool,
}

#[derive(Clone, Debug)]
struct PendingPicture<'a> {
    source: &'a [u8],
    drawing_range: ElementRange<'a>,
    anchor_range: ElementRange<'a>,
    placement: DrawingPlacement,
    graphic_range: ElementRange<'a>,
    graphic_data_range: ElementRange<'a>,
    picture_range: ElementRange<'a>,
    blip_range: Option<ElementRange<'a>>,
    ext_list_range: Option<ElementRange<'a>>,
    c_nv_pr_id: Option<Box<str>>,
    c_nv_pr_range: Option<ElementRange<'a>>,
    nv_pic_pr_seen: bool,
    c_nv_pic_pr_seen: bool,
    blip_fill_seen: bool,
    sp_pr_seen: bool,
    style_seen: bool,
    picture_ext_list_seen: bool,
    raster_relationship_id: Option<Box<str>>,
    raster_relationship_is_link: bool,
    raster_relationship_dialect: Option<RelationshipDialect>,
    extensions: Vec<PendingExtension<'a>>,
    relationship_references: Vec<RelationshipReference<'a>>,
    namespace_context: NamespaceContext,
    blip_namespace_context: Option<NamespaceContext>,
    ext_list_namespace_context: Option<NamespaceContext>,
    ext_list_malformed: bool,
    mce_ancestor: bool,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Kind {
    Root,
    Other,
    Mce,
    LegacyPict,
    Object,
    Run,
    Drawing,
    Anchor(DrawingPlacement),
    Graphic,
    GraphicDataPicture,
    GraphicDataOther,
    Picture,
    PictureNvPr,
    CnvPr,
    CnvPicPr,
    BlipFill,
    PictureSpPr,
    PictureStyle,
    PictureExtList,
    Blip,
    ExtList,
    Ext,
    SvgBlip,
}

struct Scanner<'a> {
    source: &'a [u8],
    drawing_ordinal: usize,
    limits: ScanLimits,
    frames: Vec<Frame<'a>>,
    namespace_context: NamespaceContext,
    pictures: Vec<PendingPicture<'a>>,
    relationship_references: Vec<RelationshipReference<'a>>,
    active_namespace_bindings: usize,
    namespace_scope_bytes: usize,
    root_seen: bool,
    root_closed: bool,
    declaration_seen: bool,
    prolog: bool,
    nodes: usize,
    dialect: Option<DrawingDialect>,
    relationship_dialect: Option<RelationshipDialect>,
    active_picture: Option<usize>,
}

impl<'a> Scanner<'a> {
    fn new(source: &'a [u8], drawing_ordinal: usize, limits: ScanLimits) -> Result<Self> {
        let limits = limits.validate()?;
        if source.len() > limits.max_xml_bytes {
            return Err(limit(
                "DOCX drawing XML bytes",
                source.len(),
                limits.max_xml_bytes,
            ));
        }
        let mut frames = Vec::new();
        frames
            .try_reserve(16.min(limits.max_depth))
            .map_err(|source| allocation("DOCX drawing XML stack", source))?;
        let mut pictures = Vec::new();
        pictures
            .try_reserve(4.min(limits.max_pictures))
            .map_err(|source| allocation("DOCX drawing picture index", source))?;
        let mut relationship_references = Vec::new();
        relationship_references
            .try_reserve(8.min(limits.max_global_relationship_references))
            .map_err(|source| allocation("DOCX drawing relationship index", source))?;
        Ok(Self {
            source,
            drawing_ordinal,
            limits,
            frames,
            namespace_context: NamespaceContext::empty(),
            pictures,
            relationship_references,
            active_namespace_bindings: 0,
            namespace_scope_bytes: 0,
            root_seen: false,
            root_closed: false,
            declaration_seen: false,
            prolog: true,
            nodes: 0,
            dialect: None,
            relationship_dialect: None,
            active_picture: None,
        })
    }

    fn recognized_element_only_context(&self) -> bool {
        let Some(frame) = self.frames.last() else {
            return false;
        };
        match frame.kind {
            Kind::Blip | Kind::ExtList | Kind::SvgBlip => true,
            Kind::Ext => frame
                .picture_index
                .and_then(|index| {
                    frame.extension_index.and_then(|extension_index| {
                        self.pictures
                            .get(index)
                            .and_then(|picture| picture.extensions.get(extension_index))
                    })
                })
                .is_some_and(|extension| extension.admitted),
            _ => false,
        }
    }

    fn active_picture_mut(&mut self) -> Result<&mut PendingPicture<'a>> {
        let index = self
            .active_picture
            .ok_or_else(|| invalid("DOCX picture state is not active"))?;
        self.pictures
            .get_mut(index)
            .ok_or_else(|| invalid("DOCX picture index is outside source index"))
    }

    fn mark_picture_ancestor_frames(&mut self, index: usize) -> Result<()> {
        if self.frames.iter().any(|frame| {
            matches!(
                frame.kind,
                Kind::Drawing | Kind::Anchor(_) | Kind::Graphic | Kind::GraphicDataPicture
            ) && frame.picture_index.is_some()
        }) {
            return Err(invalid(
                "DOCX drawing anchor contains more than one direct picture",
            ));
        }
        for frame in &mut self.frames {
            if matches!(
                frame.kind,
                Kind::Drawing | Kind::Anchor(_) | Kind::Graphic | Kind::GraphicDataPicture
            ) {
                frame.picture_index = Some(index);
            }
        }
        Ok(())
    }

    fn validate_picture_child_start(&self, parent: Option<Kind>, kind: Kind) -> Result<()> {
        if parent == Some(Kind::GraphicDataPicture) && kind != Kind::Picture {
            return Err(invalid("picture graphicData has a non-pic direct child"));
        }
        let Some(index) = self.active_picture else {
            return Ok(());
        };
        let picture = self
            .pictures
            .get(index)
            .ok_or_else(|| invalid("DOCX picture index is outside source index"))?;
        match parent {
            Some(Kind::Picture) => {
                if !matches!(
                    kind,
                    Kind::PictureNvPr
                        | Kind::BlipFill
                        | Kind::PictureSpPr
                        | Kind::PictureStyle
                        | Kind::PictureExtList
                ) {
                    return Err(invalid("DOCX pic:pic has an unsupported direct child"));
                }
                match kind {
                    Kind::PictureNvPr if picture.nv_pic_pr_seen => {
                        Err(invalid("DOCX pic:pic has duplicate nvPicPr"))
                    },
                    Kind::PictureNvPr if picture.blip_fill_seen || picture.sp_pr_seen => {
                        Err(invalid("DOCX pic:pic children are out of order: nvPicPr"))
                    },
                    Kind::BlipFill if !picture.nv_pic_pr_seen || picture.blip_fill_seen => {
                        Err(invalid("DOCX pic:pic children are out of order: blipFill"))
                    },
                    Kind::PictureSpPr if !picture.blip_fill_seen || picture.sp_pr_seen => {
                        Err(invalid("DOCX pic:pic children are out of order: spPr"))
                    },
                    Kind::PictureStyle | Kind::PictureExtList if !picture.sp_pr_seen => {
                        Err(invalid("DOCX pic:pic optional child appears before spPr"))
                    },
                    Kind::PictureStyle if picture.picture_ext_list_seen => {
                        Err(invalid("DOCX pic:pic style appears after extLst"))
                    },
                    Kind::PictureStyle if picture.style_seen => {
                        Err(invalid("DOCX pic:pic has duplicate style"))
                    },
                    Kind::PictureExtList if picture.picture_ext_list_seen => {
                        Err(invalid("DOCX pic:pic has duplicate extLst"))
                    },
                    _ => Ok(()),
                }
            },
            Some(Kind::PictureNvPr) => match kind {
                Kind::CnvPr if picture.c_nv_pr_range.is_some() => {
                    Err(invalid("DOCX nvPicPr has duplicate cNvPr"))
                },
                Kind::CnvPicPr if picture.c_nv_pr_range.is_none() => {
                    Err(invalid("DOCX nvPicPr requires cNvPr before cNvPicPr"))
                },
                Kind::CnvPicPr if picture.c_nv_pic_pr_seen => {
                    Err(invalid("DOCX nvPicPr has duplicate cNvPicPr"))
                },
                Kind::CnvPr | Kind::CnvPicPr => Ok(()),
                _ => Err(invalid("DOCX nvPicPr has an unsupported direct child")),
            },
            _ => Ok(()),
        }
    }

    fn run(mut self) -> Result<SourceDrawing<'a>> {
        let mut reader = NsReader::from_reader(self.source);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        reader.config_mut().check_comments = true;
        let mut buffer = Vec::new();

        loop {
            let event_start = position(&reader)?;
            let decoder = reader.decoder();
            let (resolved, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(xml_error)?;
            // Keep the event borrowed from the reader.  `resolved_tag` stores
            // only a static namespace identity, so no resolver-owned URI or
            // event clone is needed before the caller limits are checked.
            let resolved = resolved_tag(resolved, decoder)?;
            let event_end = position(&reader)?;
            self.nodes = self.nodes.checked_add(1).ok_or_else(|| {
                limit("DOCX drawing XML nodes", self.nodes, self.limits.max_nodes)
            })?;
            if self.nodes > self.limits.max_nodes {
                return Err(limit(
                    "DOCX drawing XML nodes",
                    self.nodes,
                    self.limits.max_nodes,
                ));
            }

            match event {
                Event::Start(element) => {
                    let frame = self.start_element(
                        &mut reader,
                        resolved,
                        &element,
                        event_start,
                        event_end,
                    )?;
                    self.frames
                        .try_reserve(1)
                        .map_err(|source| allocation("DOCX drawing XML stack", source))?;
                    self.frames.push(frame);
                    self.prolog = false;
                },
                Event::Empty(element) => {
                    let frame = self.start_element(
                        &mut reader,
                        resolved,
                        &element,
                        event_start,
                        event_end,
                    )?;
                    let root = frame.kind == Kind::Root;
                    let declarations = frame.namespace_declarations;
                    let parent_namespace_context = frame.parent_namespace_context.clone();
                    self.finish_element(frame, event_start, event_end, true)?;
                    self.active_namespace_bindings = self
                        .active_namespace_bindings
                        .checked_sub(declarations)
                        .ok_or_else(|| invalid("DOCX namespace scope underflow"))?;
                    self.namespace_context = parent_namespace_context;
                    if root {
                        self.root_closed = true;
                    }
                    self.prolog = false;
                },
                Event::End(element) => {
                    validate_qname(element.name().as_ref(), "end element")?;
                    let frame = self
                        .frames
                        .pop()
                        .ok_or_else(|| invalid("DOCX drawing XML has an unmatched end element"))?;
                    let parent_namespace_context = frame.parent_namespace_context.clone();
                    self.finish_element(frame.clone(), event_start, event_end, false)?;
                    self.active_namespace_bindings = self
                        .active_namespace_bindings
                        .checked_sub(frame.namespace_declarations)
                        .ok_or_else(|| invalid("DOCX namespace scope underflow"))?;
                    self.namespace_context = parent_namespace_context;
                    if frame.kind == Kind::Root {
                        self.root_closed = true;
                    }
                },
                Event::Text(text) => {
                    validate_xml_text(text.as_ref())?;
                    if self.recognized_element_only_context() && !is_xml_whitespace(text.as_ref()) {
                        return Err(invalid(
                            "recognized DOCX picture extension container contains non-whitespace text",
                        ));
                    }
                    if self.frames.is_empty()
                        && (!self.root_seen || self.root_closed)
                        && !is_xml_whitespace(text.as_ref())
                    {
                        return Err(invalid("non-whitespace DOCX text appears outside the root"));
                    }
                },
                Event::CData(data) => {
                    validate_xml_characters(data.as_ref())?;
                    if self.frames.is_empty() {
                        return Err(invalid("DOCX CDATA appears outside the root"));
                    }
                    if self.recognized_element_only_context() && !is_xml_whitespace(data.as_ref()) {
                        return Err(invalid(
                            "recognized DOCX picture extension container contains non-whitespace CDATA",
                        ));
                    }
                },
                Event::GeneralRef(reference) => {
                    if self.frames.is_empty() {
                        return Err(invalid("DOCX entity reference appears outside the root"));
                    }
                    let decoded = litchi_ooxml_common::xml::decode_xml_reference(&reference)
                        .map_err(|error| Error::Invalid(error.to_string()))?;
                    validate_xml_characters(decoded.as_bytes())?;
                    if self.recognized_element_only_context()
                        && !is_xml_whitespace(decoded.as_bytes())
                    {
                        return Err(invalid(
                            "recognized DOCX picture extension container contains non-whitespace entity text",
                        ));
                    }
                },
                Event::DocType(_) => {
                    return Err(invalid("DOCTYPE is forbidden in DOCX drawing scan"));
                },
                Event::Decl(declaration) => {
                    if self.declaration_seen || !self.prolog || self.root_seen {
                        return Err(invalid(
                            "DOCX XML declaration is not at the beginning of the main story",
                        ));
                    }
                    validate_declaration(&declaration)?;
                    self.declaration_seen = true;
                },
                Event::Comment(comment) => validate_xml_characters(comment.as_ref())?,
                Event::PI(instruction) => validate_processing_instruction(instruction.as_ref())?,
                Event::Eof => break,
            }
            buffer.clear();
        }

        if !self.root_seen || !self.root_closed || !self.frames.is_empty() {
            return Err(invalid("DOCX main story XML has an incomplete root"));
        }
        let dialect = self
            .dialect
            .ok_or_else(|| invalid("DOCX main story has no WordprocessingML dialect"))?;
        let relationship_dialect = self.relationship_dialect.unwrap_or(match dialect {
            DrawingDialect::Transitional => RelationshipDialect::Transitional,
            DrawingDialect::Strict => RelationshipDialect::Strict,
        });

        let mut pictures = Vec::new();
        pictures
            .try_reserve_exact(self.pictures.len())
            .map_err(|source| allocation("DOCX drawing picture projection", source))?;
        for (picture_ordinal, pending) in self.pictures.into_iter().enumerate() {
            let opaque_svg_extension = pending
                .extensions
                .iter()
                .any(|extension| !extension.admitted);
            let svg_owner = project_svg_owner(
                self.source,
                &pending,
                relationship_dialect,
                self.limits.max_fragment_bytes,
            )?;
            pictures.push(PictureSource {
                source: pending.source,
                drawing_ordinal: self.drawing_ordinal,
                picture_ordinal,
                drawing_range: pending.drawing_range,
                anchor_range: pending.anchor_range,
                placement: pending.placement,
                graphic_range: pending.graphic_range,
                graphic_data_range: pending.graphic_data_range,
                picture_range: pending.picture_range,
                blip_range: pending.blip_range,
                ext_list_range: pending.ext_list_range,
                c_nv_pr_id: pending.c_nv_pr_id,
                c_nv_pr_range: pending.c_nv_pr_range,
                raster_relationship_id: pending.raster_relationship_id,
                raster_relationship_is_link: pending.raster_relationship_is_link,
                raster_relationship_dialect: pending.raster_relationship_dialect,
                svg_owner,
                relationship_references: pending.relationship_references,
                namespace_context: pending.namespace_context,
                blip_namespace_context: pending.blip_namespace_context,
                ext_list_namespace_context: pending.ext_list_namespace_context,
                opaque_svg_extension,
            });
        }
        Ok(SourceDrawing {
            source: self.source,
            drawing_ordinal: self.drawing_ordinal,
            dialect,
            relationship_dialect,
            pictures,
            relationship_references: self.relationship_references,
        })
    }

    fn start_element(
        &mut self,
        reader: &mut NsReader<&'a [u8]>,
        resolved: Resolved,
        element: &BytesStart<'_>,
        event_start: usize,
        event_end: usize,
    ) -> Result<Frame<'a>> {
        let depth =
            self.frames.len().checked_add(1).ok_or_else(|| {
                limit("DOCX drawing XML depth", usize::MAX, self.limits.max_depth)
            })?;
        if depth > self.limits.max_depth {
            return Err(limit(
                "DOCX drawing XML depth",
                depth,
                self.limits.max_depth,
            ));
        }
        if self.root_closed {
            return Err(invalid("DOCX main story has content after its root"));
        }
        validate_qname(element.name().as_ref(), "element")?;
        if matches!(resolved, Resolved::Unknown) && element.name().prefix().is_some() {
            return Err(invalid("DOCX XML element uses an unbound prefix"));
        }
        let parent_namespace_context = self.namespace_context.clone();
        let declarations = validate_attributes(element, reader, self.limits)?;
        let active = self
            .active_namespace_bindings
            .checked_add(declarations)
            .ok_or_else(|| {
                limit(
                    "active namespace bindings",
                    usize::MAX,
                    self.limits.max_active_namespace_bindings,
                )
            })?;
        if active > self.limits.max_active_namespace_bindings {
            return Err(limit(
                "active namespace bindings",
                active,
                self.limits.max_active_namespace_bindings,
            ));
        }
        self.active_namespace_bindings = active;
        let namespace_bytes = namespace_declaration_bytes(element, reader.decoder())?;
        let namespace_scope_bytes = self
            .namespace_scope_bytes
            .checked_add(namespace_bytes)
            .ok_or_else(|| {
                limit(
                    "namespace context bytes",
                    usize::MAX,
                    self.limits.max_namespace_bytes,
                )
            })?;
        if namespace_scope_bytes > self.limits.max_namespace_bytes {
            return Err(limit(
                "namespace context bytes",
                namespace_scope_bytes,
                self.limits.max_namespace_bytes,
            ));
        }
        let context =
            extend_namespace_context(&parent_namespace_context, element, reader.decoder())?;
        self.namespace_scope_bytes = namespace_scope_bytes;
        self.namespace_context = context.clone();

        let parent = self.frames.last().map(|frame| frame.kind);
        let mce_ancestor = self.frames.iter().any(|frame| frame.kind == Kind::Mce);
        let mut kind = classify(
            &resolved,
            element.name(),
            parent,
            !self.root_seen,
            element,
            reader.decoder(),
        )?;
        if kind == Kind::Root && !self.frames.is_empty() {
            return Err(invalid("DOCX main story has a nested w:document root"));
        }
        if !self.root_seen && kind != Kind::Root {
            return Err(invalid(
                "DOCX main story has a non-root element before w:document",
            ));
        }
        if kind == Kind::Drawing && self.frames.iter().any(|frame| frame.blocks_drawing) {
            // Legacy `w:pict`/`w:object` payloads are opaque to this ordinary
            // drawing owner.  A nested w:drawing must not become a second
            // effective owner inside that payload.
            kind = Kind::Other;
        }
        let extension_attributes = if kind == Kind::Ext {
            Some(extension_uri(element, reader.decoder())?)
        } else {
            None
        };
        if parent == Some(Kind::ExtList) && kind != Kind::Ext {
            return Err(invalid(
                "direct DOCX a:extLst child is not an a:ext element",
            ));
        }
        if parent == Some(Kind::Blip) && kind != Kind::ExtList {
            return Err(invalid(
                "direct DOCX a:blip child is not an a:extLst element",
            ));
        }
        let admitted_parent = if kind == Kind::SvgBlip {
            self.active_picture
                .and_then(|index| {
                    self.frames
                        .iter()
                        .rev()
                        .find(|frame| frame.kind == Kind::Ext)
                        .and_then(|frame| frame.extension_index)
                        .and_then(|extension_index| {
                            self.pictures
                                .get(index)
                                .and_then(|picture| picture.extensions.get(extension_index))
                        })
                })
                .is_some_and(|extension| extension.admitted)
        } else {
            false
        };
        if kind == Kind::SvgBlip && !admitted_parent {
            kind = Kind::Other;
        }
        self.validate_picture_child_start(parent, kind)?;
        self.collect_relationship_attributes(reader, element, event_start..event_end)?;

        match kind {
            Kind::Root => {
                if self.root_seen {
                    return Err(invalid("DOCX main story has more than one root"));
                }
                self.root_seen = true;
                self.dialect = Some(
                    if resolved == Resolved::Known(STRICT_WORDPROCESSINGML_NAMESPACE) {
                        DrawingDialect::Strict
                    } else {
                        DrawingDialect::Transitional
                    },
                );
                self.relationship_dialect = Some(match self.dialect {
                    Some(DrawingDialect::Strict) => RelationshipDialect::Strict,
                    Some(DrawingDialect::Transitional) | None => RelationshipDialect::Transitional,
                });
            },
            Kind::LegacyPict | Kind::Object | Kind::Run => {},
            Kind::Drawing => {},
            Kind::Anchor(_) => {},
            Kind::Picture => {
                let anchor = self
                    .frames
                    .iter()
                    .rev()
                    .find(|frame| matches!(frame.kind, Kind::Anchor(_)))
                    .ok_or_else(|| invalid("direct picture appears outside a Word anchor"))?;
                let drawing = self
                    .frames
                    .iter()
                    .rev()
                    .find(|frame| frame.kind == Kind::Drawing)
                    .ok_or_else(|| invalid("direct picture appears outside w:drawing"))?;
                if self.pictures.len() >= self.limits.max_pictures {
                    return Err(limit(
                        "DOCX drawing pictures",
                        self.pictures.len() + 1,
                        self.limits.max_pictures,
                    ));
                }
                let anchor_frame = self
                    .frames
                    .iter()
                    .rev()
                    .find(|frame| matches!(frame.kind, Kind::Anchor(_)))
                    .ok_or_else(|| invalid("picture anchor frame is missing"))?;
                let graphic_frame = self
                    .frames
                    .iter()
                    .rev()
                    .find(|frame| frame.kind == Kind::Graphic)
                    .ok_or_else(|| invalid("picture graphic frame is missing"))?;
                let graphic_data_frame = self
                    .frames
                    .last()
                    .filter(|frame| frame.kind == Kind::GraphicDataPicture)
                    .ok_or_else(|| invalid("picture graphicData frame is missing"))?;
                let anchor_placement = match anchor.kind {
                    Kind::Anchor(placement) => placement,
                    _ => return Err(invalid("picture anchor has invalid placement")),
                };
                let anchor_range = anchor_frame.range.clone();
                let drawing_range = drawing.range.clone();
                let graphic_range = graphic_frame.range.clone();
                let graphic_data_range = graphic_data_frame.range.clone();
                let picture_range =
                    element_range(self.source, event_start, event_end, element, &resolved)?;
                let picture = PendingPicture {
                    source: self.source,
                    drawing_range,
                    anchor_range,
                    placement: anchor_placement,
                    graphic_range,
                    graphic_data_range,
                    picture_range,
                    blip_range: None,
                    ext_list_range: None,
                    c_nv_pr_id: None,
                    c_nv_pr_range: None,
                    nv_pic_pr_seen: false,
                    c_nv_pic_pr_seen: false,
                    blip_fill_seen: false,
                    sp_pr_seen: false,
                    style_seen: false,
                    picture_ext_list_seen: false,
                    raster_relationship_id: None,
                    raster_relationship_is_link: false,
                    raster_relationship_dialect: None,
                    extensions: Vec::new(),
                    relationship_references: Vec::new(),
                    namespace_context: context.clone(),
                    blip_namespace_context: None,
                    ext_list_namespace_context: None,
                    ext_list_malformed: false,
                    mce_ancestor,
                };
                self.pictures
                    .try_reserve(1)
                    .map_err(|source| allocation("DOCX drawing picture index", source))?;
                self.pictures.push(picture);
                let index = self.pictures.len() - 1;
                self.active_picture = Some(index);
                self.mark_picture_ancestor_frames(index)?;
                if let Some(picture) = self.pictures.get_mut(index) {
                    picture.nv_pic_pr_seen = false;
                }
            },
            Kind::CnvPr => {
                let range = element_range(self.source, event_start, event_end, element, &resolved)?;
                let value = unqualified_attribute(element, b"id", reader.decoder())?
                    .ok_or_else(|| invalid("DOCX picture cNvPr is missing its id"))?;
                if unqualified_attribute(element, b"name", reader.decoder())?.is_none() {
                    return Err(invalid("DOCX picture cNvPr is missing its name"));
                }
                let text = value.trim();
                if text.is_empty()
                    || text.len() > 10
                    || !text.bytes().all(|byte| byte.is_ascii_digit())
                    || text.parse::<u32>().is_err()
                {
                    return Err(invalid(
                        "DOCX picture cNvPr id is not a bounded unsigned integer",
                    ));
                }
                let picture = self.active_picture_mut()?;
                if picture.c_nv_pr_range.is_some() {
                    return Err(invalid("direct picture has duplicate cNvPr"));
                }
                let mut id = String::new();
                id.try_reserve(text.len())
                    .map_err(|source| allocation("DOCX picture cNvPr ID", source))?;
                id.push_str(text);
                picture.c_nv_pr_range = Some(range);
                picture.c_nv_pr_id = Some(id.into_boxed_str());
            },
            Kind::CnvPicPr => {
                self.active_picture_mut()?.c_nv_pic_pr_seen = true;
            },
            Kind::PictureNvPr => {
                self.active_picture_mut()?.nv_pic_pr_seen = true;
            },
            Kind::BlipFill => {
                self.active_picture_mut()?.blip_fill_seen = true;
            },
            Kind::PictureSpPr => {
                self.active_picture_mut()?.sp_pr_seen = true;
            },
            Kind::PictureStyle => {
                self.active_picture_mut()?.style_seen = true;
            },
            Kind::PictureExtList => {
                self.active_picture_mut()?.picture_ext_list_seen = true;
            },
            Kind::Blip => {
                let range = element_range(self.source, event_start, event_end, element, &resolved)?;
                let embed = relationship_attribute(element, reader, b"embed")?;
                let link = relationship_attribute(element, reader, b"link")?;
                if embed.is_some() && link.is_some() {
                    return Err(invalid(
                        "DOCX raster blip cannot carry both r:embed and r:link",
                    ));
                }
                let picture = self.active_picture_mut()?;
                if picture.blip_range.is_some() {
                    return Err(invalid("direct picture has duplicate raster blip"));
                }
                let (relationship, is_link) = match (embed, link) {
                    (Some(value), None) => (Some(value), false),
                    (None, Some(value)) => (Some(value), true),
                    (None, None) => (None, false),
                    (Some(_), Some(_)) => unreachable!(),
                };
                picture.blip_range = Some(range);
                picture.blip_namespace_context = Some(context.clone());
                picture.raster_relationship_dialect =
                    relationship.as_ref().map(|(_, dialect)| *dialect);
                picture.raster_relationship_is_link = is_link;
                picture.raster_relationship_id = relationship.map(|(id, _)| id);
            },
            Kind::ExtList => {
                if let Some(index) = self.active_picture {
                    let range =
                        element_range(self.source, event_start, event_end, element, &resolved)?;
                    let picture = self
                        .pictures
                        .get_mut(index)
                        .ok_or_else(|| invalid("picture index is outside source index"))?;
                    if picture.ext_list_range.is_some() {
                        return Err(invalid("direct picture has duplicate extLst"));
                    }
                    picture.ext_list_range = Some(range);
                    picture.ext_list_namespace_context = Some(context.clone());
                }
            },
            Kind::Ext => {
                if let Some(index) = self.active_picture {
                    let (admitted, uri_lexical, malformed) = extension_attributes
                        .ok_or_else(|| invalid("DOCX extension URI was not preflighted"))?;
                    let range =
                        element_range(self.source, event_start, event_end, element, &resolved)?;
                    let mut extensions = PendingExtension {
                        range,
                        admitted,
                        malformed,
                        uri_lexical,
                        svg_blip: None,
                        svg_context: None,
                        svg_count: 0,
                        mce_ancestor,
                    };
                    if mce_ancestor {
                        extensions.mce_ancestor = true;
                    }
                    let picture = self
                        .pictures
                        .get_mut(index)
                        .ok_or_else(|| invalid("picture index is outside source index"))?;
                    picture
                        .extensions
                        .try_reserve(1)
                        .map_err(|source| allocation("DOCX picture extensions", source))?;
                    picture.extensions.push(extensions);
                }
            },
            Kind::SvgBlip => {
                if let Some(index) = self.active_picture {
                    let extension_index = self
                        .frames
                        .last()
                        .and_then(|frame| frame.extension_index)
                        .ok_or_else(|| invalid("SVG blip extension index is missing"))?;
                    let extension = self
                        .pictures
                        .get_mut(index)
                        .and_then(|picture| picture.extensions.get_mut(extension_index))
                        .ok_or_else(|| invalid("SVG extension index is outside source index"))?;
                    extension.svg_count = extension
                        .svg_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("SVG blip count overflow"))?;
                    extension.svg_blip = Some(element_range(
                        self.source,
                        event_start,
                        event_end,
                        element,
                        &resolved,
                    )?);
                    extension.svg_context = Some(context.clone());
                }
            },
            Kind::Other
            | Kind::Mce
            | Kind::Graphic
            | Kind::GraphicDataPicture
            | Kind::GraphicDataOther => {},
        }

        if kind == Kind::Other {
            match parent {
                Some(Kind::Ext) => {
                    if let Some(index) = self.active_picture
                        && let Some(extension_index) = self
                            .frames
                            .iter()
                            .rev()
                            .find(|frame| frame.kind == Kind::Ext)
                            .and_then(|frame| frame.extension_index)
                        && let Some(extension) = self
                            .pictures
                            .get_mut(index)
                            .and_then(|picture| picture.extensions.get_mut(extension_index))
                    {
                        extension.malformed = true;
                    }
                },
                Some(Kind::ExtList) => {
                    if let Some(index) = self.active_picture
                        && let Some(picture) = self.pictures.get_mut(index)
                    {
                        picture.ext_list_malformed = true;
                    }
                },
                _ => {},
            }
        }

        let picture_index = self.active_picture;
        let extension_index = if kind == Kind::Ext {
            picture_index.and_then(|index| {
                self.pictures
                    .get(index)
                    .map(|picture| picture.extensions.len().saturating_sub(1))
            })
        } else {
            None
        };
        let blocks_drawing = match kind {
            Kind::LegacyPict
            | Kind::Object
            | Kind::Drawing
            | Kind::Anchor(_)
            | Kind::Graphic
            | Kind::GraphicDataPicture
            | Kind::GraphicDataOther
            | Kind::Picture
            | Kind::PictureNvPr
            | Kind::CnvPr
            | Kind::CnvPicPr
            | Kind::BlipFill
            | Kind::PictureSpPr
            | Kind::PictureStyle
            | Kind::PictureExtList
            | Kind::Blip
            | Kind::ExtList
            | Kind::Ext
            | Kind::SvgBlip => true,
            Kind::Other => !is_safe_word_container(&resolved, element.name()),
            Kind::Root | Kind::Mce => false,
            Kind::Run => parent == Some(Kind::Run),
        };
        Ok(Frame {
            kind,
            source_start: event_start,
            range: element_range(self.source, event_start, event_end, element, &resolved)?,
            namespace_declarations: declarations,
            parent_namespace_context,
            picture_index,
            extension_index,
            blocks_drawing,
        })
    }

    fn finish_element(
        &mut self,
        frame: Frame<'a>,
        event_start: usize,
        event_end: usize,
        empty: bool,
    ) -> Result<()> {
        match frame.kind {
            Kind::Drawing | Kind::Graphic | Kind::GraphicDataPicture => {
                if let Some(index) = frame.picture_index
                    && let Some(picture) = self.pictures.get_mut(index)
                {
                    let range = match frame.kind {
                        Kind::Drawing => &mut picture.drawing_range,
                        Kind::Graphic => &mut picture.graphic_range,
                        Kind::GraphicDataPicture => &mut picture.graphic_data_range,
                        _ => unreachable!(),
                    };
                    if range.range.start == frame.source_start {
                        range.range.end = event_end;
                        if !empty {
                            range.close_start = Some(event_start);
                        }
                    }
                }
            },
            Kind::Anchor(_) => {
                if let Some(index) = frame.picture_index
                    && let Some(picture) = self.pictures.get_mut(index)
                    && picture.anchor_range.range.start == frame.source_start
                {
                    picture.anchor_range.range.end = event_end;
                    if !empty {
                        picture.anchor_range.close_start = Some(event_start);
                    }
                }
            },
            Kind::Picture => {
                if let Some(index) = frame.picture_index {
                    if let Some(picture) = self.pictures.get_mut(index) {
                        if !picture.nv_pic_pr_seen
                            || picture.c_nv_pr_range.is_none()
                            || !picture.c_nv_pic_pr_seen
                            || !picture.blip_fill_seen
                            || picture.blip_range.is_none()
                            || !picture.sp_pr_seen
                        {
                            return Err(invalid(
                                "DOCX pic:pic is missing a required child or raster relationship",
                            ));
                        }
                        picture.picture_range.range.end = event_end;
                        if !empty {
                            picture.picture_range.close_start = Some(event_start);
                        }
                    }
                    self.active_picture = None;
                }
            },
            Kind::PictureNvPr => {
                if let Some(index) = frame.picture_index
                    && let Some(picture) = self.pictures.get(index)
                    && (picture.c_nv_pr_range.is_none() || !picture.c_nv_pic_pr_seen)
                {
                    return Err(invalid("DOCX nvPicPr is missing cNvPr or cNvPicPr"));
                }
            },
            Kind::BlipFill => {
                if let Some(index) = frame.picture_index
                    && let Some(picture) = self.pictures.get(index)
                    && picture.blip_range.is_none()
                {
                    return Err(invalid("DOCX blipFill is missing its direct a:blip"));
                }
            },
            Kind::Blip => {
                if let Some(index) = frame.picture_index {
                    if let Some(picture) = self.pictures.get_mut(index)
                        && let Some(range) = picture.blip_range.as_mut()
                        && range.range.start == frame.source_start
                    {
                        range.range.end = event_end;
                        if !empty {
                            range.close_start = Some(event_start);
                        }
                    }
                }
            },
            Kind::ExtList => {
                if let Some(index) = frame.picture_index {
                    if let Some(picture) = self.pictures.get_mut(index)
                        && let Some(range) = picture.ext_list_range.as_mut()
                        && range.range.start == frame.source_start
                    {
                        range.range.end = event_end;
                        if !empty {
                            range.close_start = Some(event_start);
                        }
                    }
                }
            },
            Kind::Ext => {
                if let (Some(index), Some(extension_index)) =
                    (frame.picture_index, frame.extension_index)
                    && let Some(extension) = self
                        .pictures
                        .get_mut(index)
                        .and_then(|picture| picture.extensions.get_mut(extension_index))
                    && extension.range.range.start == frame.source_start
                {
                    extension.range.range.end = event_end;
                    if !empty {
                        extension.range.close_start = Some(event_start);
                    }
                }
            },
            Kind::SvgBlip => {
                if let (Some(index), Some(extension_index)) =
                    (frame.picture_index, frame.extension_index)
                    && let Some(extension) = self
                        .pictures
                        .get_mut(index)
                        .and_then(|picture| picture.extensions.get_mut(extension_index))
                    && let Some(range) = extension.svg_blip.as_mut()
                    && range.range.start == frame.source_start
                {
                    range.range.end = event_end;
                    if !empty {
                        range.close_start = Some(event_start);
                    }
                }
            },
            Kind::CnvPr => {
                if let Some(index) = frame.picture_index
                    && let Some(picture) = self.pictures.get_mut(index)
                    && let Some(range) = picture.c_nv_pr_range.as_mut()
                    && range.range.start == frame.source_start
                {
                    range.range.end = event_end;
                    if !empty {
                        range.close_start = Some(event_start);
                    }
                }
            },
            Kind::Root
            | Kind::Other
            | Kind::Mce
            | Kind::LegacyPict
            | Kind::Object
            | Kind::Run
            | Kind::GraphicDataOther
            | Kind::CnvPicPr
            | Kind::PictureSpPr
            | Kind::PictureStyle
            | Kind::PictureExtList => {},
        }
        Ok(())
    }

    fn collect_relationship_attributes(
        &mut self,
        reader: &NsReader<&'a [u8]>,
        element: &BytesStart<'_>,
        range: std::ops::Range<usize>,
    ) -> Result<()> {
        for attribute in element.attributes().with_checks(true) {
            let attribute = attribute.map_err(xml_error)?;
            let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
            let dialect = match namespace {
                ResolveResult::Bound(Namespace(value)) => {
                    let value = decode_namespace_uri(value, reader.decoder())?;
                    if value.as_bytes() == RELATIONSHIPS {
                        RelationshipDialect::Transitional
                    } else if value.as_bytes() == STRICT_RELATIONSHIPS {
                        RelationshipDialect::Strict
                    } else {
                        continue;
                    }
                },
                _ => continue,
            };
            if self.relationship_references.len() >= self.limits.max_global_relationship_references
            {
                return Err(limit(
                    "DOCX global relationship references",
                    self.relationship_references.len() + 1,
                    self.limits.max_global_relationship_references,
                ));
            }
            if let Some(index) = self.active_picture {
                let count = self
                    .pictures
                    .get(index)
                    .ok_or_else(|| invalid("DOCX picture index is outside source index"))?
                    .relationship_references
                    .len();
                if count >= self.limits.max_relationship_references {
                    return Err(limit(
                        "DOCX picture relationship references",
                        count + 1,
                        self.limits.max_relationship_references,
                    ));
                }
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(xml_error)?;
            if value.len() > MAX_RELATIONSHIP_ID_BYTES || value.is_empty() || !is_ncname(&value) {
                return Err(invalid("DOCX drawing relationship ID is invalid"));
            }
            let picture_index = self.active_picture;
            if let Some(index) = picture_index {
                let picture = self
                    .pictures
                    .get_mut(index)
                    .ok_or_else(|| invalid("DOCX picture index is outside source index"))?;
                picture
                    .relationship_references
                    .try_reserve(1)
                    .map_err(|source| allocation("DOCX picture relationship references", source))?;
            }
            self.relationship_references
                .try_reserve(1)
                .map_err(|source| allocation("DOCX global relationship references", source))?;
            let id: Arc<str> = Arc::from(value.into_owned());
            let local_name: Arc<str> = Arc::from(
                std::str::from_utf8(attribute.key.local_name().as_ref())
                    .map_err(|error| Error::Invalid(error.to_string()))?
                    .to_owned(),
            );
            let reference = RelationshipReference {
                source: self.source,
                id,
                local_name,
                range: ByteRange::new(range.start, range.end),
                dialect,
            };
            self.relationship_references.push(reference.clone());
            if let Some(index) = picture_index {
                let picture = self
                    .pictures
                    .get_mut(index)
                    .ok_or_else(|| invalid("DOCX picture index is outside source index"))?;
                picture.relationship_references.push(reference);
            }
        }
        Ok(())
    }
}

fn classify(
    resolved: &Resolved,
    name: QName<'_>,
    parent: Option<Kind>,
    first: bool,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<Kind> {
    if is_exact_namespace(resolved, MCE) {
        return Ok(Kind::Mce);
    }
    if first
        && parent.is_none()
        && is_name(
            resolved,
            name,
            b"document",
            WORDPROCESSINGML_NAMESPACE,
            STRICT_WORDPROCESSINGML_NAMESPACE,
        )
    {
        return Ok(Kind::Root);
    }
    if parent == Some(Kind::Run)
        && is_name(
            resolved,
            name,
            b"drawing",
            WORDPROCESSINGML_NAMESPACE,
            STRICT_WORDPROCESSINGML_NAMESPACE,
        )
    {
        return Ok(Kind::Drawing);
    }
    if is_name(
        resolved,
        name,
        b"pict",
        WORDPROCESSINGML_NAMESPACE,
        STRICT_WORDPROCESSINGML_NAMESPACE,
    ) {
        return Ok(Kind::LegacyPict);
    }
    if is_name(
        resolved,
        name,
        b"object",
        WORDPROCESSINGML_NAMESPACE,
        STRICT_WORDPROCESSINGML_NAMESPACE,
    ) {
        return Ok(Kind::Object);
    }
    if is_name(
        resolved,
        name,
        b"r",
        WORDPROCESSINGML_NAMESPACE,
        STRICT_WORDPROCESSINGML_NAMESPACE,
    ) {
        return Ok(Kind::Run);
    }
    if parent == Some(Kind::Drawing)
        && is_name(
            resolved,
            name,
            b"inline",
            WORDPROCESSING_DRAWING,
            STRICT_WORDPROCESSING_DRAWING,
        )
    {
        return Ok(Kind::Anchor(DrawingPlacement::Inline));
    }
    if parent == Some(Kind::Drawing)
        && is_name(
            resolved,
            name,
            b"anchor",
            WORDPROCESSING_DRAWING,
            STRICT_WORDPROCESSING_DRAWING,
        )
    {
        return Ok(Kind::Anchor(DrawingPlacement::Floating));
    }
    if matches!(parent, Some(Kind::Anchor(_)))
        && is_name(resolved, name, b"graphic", DRAWINGML, STRICT_DRAWINGML)
    {
        return Ok(Kind::Graphic);
    }
    if parent == Some(Kind::Graphic)
        && is_name(resolved, name, b"graphicData", DRAWINGML, STRICT_DRAWINGML)
    {
        let uri = unqualified_attribute(element, b"uri", decoder)?;
        return Ok(
            if uri
                .as_deref()
                .and_then(xsd_token_atom)
                .is_some_and(|uri| uri.as_bytes() == PICTURE || uri.as_bytes() == STRICT_PICTURE)
            {
                Kind::GraphicDataPicture
            } else {
                Kind::GraphicDataOther
            },
        );
    }
    if parent == Some(Kind::GraphicDataPicture)
        && is_name(resolved, name, b"pic", PICTURE, STRICT_PICTURE)
    {
        return Ok(Kind::Picture);
    }
    if parent == Some(Kind::Picture) && is_name(resolved, name, b"nvPicPr", PICTURE, STRICT_PICTURE)
    {
        return Ok(Kind::PictureNvPr);
    }
    if parent == Some(Kind::PictureNvPr)
        && is_name(resolved, name, b"cNvPr", PICTURE, STRICT_PICTURE)
    {
        return Ok(Kind::CnvPr);
    }
    if parent == Some(Kind::PictureNvPr)
        && is_name(resolved, name, b"cNvPicPr", PICTURE, STRICT_PICTURE)
    {
        return Ok(Kind::CnvPicPr);
    }
    if parent == Some(Kind::Picture)
        && is_name(resolved, name, b"blipFill", PICTURE, STRICT_PICTURE)
    {
        return Ok(Kind::BlipFill);
    }
    if parent == Some(Kind::Picture) && is_name(resolved, name, b"spPr", PICTURE, STRICT_PICTURE) {
        return Ok(Kind::PictureSpPr);
    }
    if parent == Some(Kind::Picture) && is_name(resolved, name, b"style", PICTURE, STRICT_PICTURE) {
        return Ok(Kind::PictureStyle);
    }
    if parent == Some(Kind::Picture) && is_name(resolved, name, b"extLst", PICTURE, STRICT_PICTURE)
    {
        return Ok(Kind::PictureExtList);
    }
    if parent == Some(Kind::BlipFill)
        && is_name(resolved, name, b"blip", DRAWINGML, STRICT_DRAWINGML)
    {
        return Ok(Kind::Blip);
    }
    if parent == Some(Kind::Blip) && is_name(resolved, name, b"extLst", DRAWINGML, STRICT_DRAWINGML)
    {
        return Ok(Kind::ExtList);
    }
    if parent == Some(Kind::ExtList) && is_name(resolved, name, b"ext", DRAWINGML, STRICT_DRAWINGML)
    {
        return Ok(Kind::Ext);
    }
    if parent == Some(Kind::Ext)
        && is_name(resolved, name, b"svgBlip", SVG_NAMESPACE, SVG_NAMESPACE)
    {
        return Ok(Kind::SvgBlip);
    }
    Ok(Kind::Other)
}

fn is_safe_word_container(namespace: &Resolved, name: QName<'_>) -> bool {
    if !is_namespace(
        namespace,
        WORDPROCESSINGML_NAMESPACE,
        STRICT_WORDPROCESSINGML_NAMESPACE,
    ) {
        return false;
    }
    matches!(
        name.local_name().as_ref(),
        b"body"
            | b"p"
            | b"pPr"
            | b"rPr"
            | b"hyperlink"
            | b"tbl"
            | b"tblPr"
            | b"tblGrid"
            | b"tr"
            | b"trPr"
            | b"tc"
            | b"tcPr"
            | b"sdt"
            | b"sdtPr"
            | b"sdtContent"
            | b"customXml"
            | b"smartTag"
            | b"fldSimple"
            | b"ins"
            | b"del"
            | b"moveFrom"
            | b"moveTo"
            | b"txbxContent"
            | b"proofErr"
            | b"permStart"
            | b"permEnd"
            | b"bookmarkStart"
            | b"bookmarkEnd"
            | b"commentRangeStart"
            | b"commentRangeEnd"
            | b"footnote"
            | b"endnote"
            | b"comment"
            | b"hdr"
            | b"ftr"
            | b"sectPr"
            | b"altChunk"
            | b"contentPart"
            | b"dir"
    )
}

fn is_name(
    namespace: &Resolved,
    name: QName<'_>,
    local: &[u8],
    transitional: &[u8],
    strict: &[u8],
) -> bool {
    name.local_name().as_ref() == local && is_namespace(namespace, transitional, strict)
}

fn is_namespace(namespace: &Resolved, transitional: &[u8], strict: &[u8]) -> bool {
    matches!(namespace, Resolved::Known(value) if *value == transitional || *value == strict)
}

fn is_exact_namespace(namespace: &Resolved, expected: &[u8]) -> bool {
    matches!(namespace, Resolved::Known(value) if *value == expected)
}

fn project_svg_owner<'a>(
    source: &'a [u8],
    picture: &PendingPicture<'a>,
    host_relationship_dialect: RelationshipDialect,
    max_fragment_bytes: usize,
) -> Result<SvgOwnerState<'a>> {
    if picture.mce_ancestor {
        return Ok(SvgOwnerState::Refused);
    }
    // The future SVG lifecycle can only replace an embedded raster payload.
    // Keep the source inventory, but expose link-backed rasters as an
    // explicit refusal even when no SVG extension is present.
    if picture.raster_relationship_is_link {
        return Ok(SvgOwnerState::Refused);
    }
    let mut admitted = Vec::new();
    let mut opaque = false;
    for extension in &picture.extensions {
        if extension.admitted {
            admitted
                .try_reserve(1)
                .map_err(|source| allocation("DOCX admitted SVG owner list", source))?;
            admitted.push(extension);
        } else {
            opaque = true;
        }
    }
    if admitted.len() > 1 {
        return Ok(SvgOwnerState::Ambiguous);
    }
    let Some(extension) = admitted.first().copied() else {
        return Ok(if opaque {
            SvgOwnerState::Opaque
        } else {
            SvgOwnerState::None
        });
    };
    if extension.malformed || extension.mce_ancestor {
        return Ok(SvgOwnerState::Refused);
    }
    if picture.ext_list_malformed
        || picture.raster_relationship_id.is_none()
        || picture.raster_relationship_is_link
    {
        return Ok(SvgOwnerState::Refused);
    }
    let Some(svg_blip) = extension.svg_blip.as_ref() else {
        return Ok(SvgOwnerState::Refused);
    };
    if extension.svg_count != 1 {
        return Ok(SvgOwnerState::Ambiguous);
    }
    let context = extension
        .svg_context
        .as_ref()
        .ok_or_else(|| invalid("admitted SVG owner has no namespace context"))?;
    let raw = svg_blip.bytes()?;
    let maximum = max_fragment_bytes.min(MAX_FRAGMENT_BYTES);
    if raw.len() > maximum {
        return Err(limit("DOCX SVG fragment bytes", raw.len(), maximum));
    }
    let value = match svg_blip::read_contextual(raw, context) {
        Ok(value) => value,
        Err(_) => return Ok(SvgOwnerState::Refused),
    };
    let embedded = value.embedded().is_some();
    let linked = value.linked().is_some();
    if embedded == linked {
        return Ok(SvgOwnerState::Refused);
    }
    let relationship_dialect = svg_relationship_dialect(raw, context, host_relationship_dialect)?;
    let owner = SvgOwner {
        source,
        extension: extension.range.clone(),
        svg_blip: svg_blip.clone(),
        uri_lexical: extension.uri_lexical.clone(),
        value,
        relationship_dialect,
        namespace_context: context.clone(),
    };
    Ok(if embedded {
        SvgOwnerState::Embedded(owner)
    } else {
        SvgOwnerState::Linked(owner)
    })
}

fn svg_relationship_dialect(
    fragment: &[u8],
    context: &NamespaceContext,
    fallback: RelationshipDialect,
) -> Result<RelationshipDialect> {
    let mut reader = Reader::from_reader(fragment);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(xml_error)?
            .into_owned();
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let mut dialect = None;
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(xml_error)?;
                    let local = attribute.key.local_name();
                    if local.as_ref() != b"embed" && local.as_ref() != b"link" {
                        continue;
                    }
                    let Some(prefix) = attribute.key.prefix() else {
                        continue;
                    };
                    let prefix = std::str::from_utf8(prefix.as_ref())
                        .map_err(|error| Error::Invalid(error.to_string()))?;
                    let Some(candidate) = relationship_dialect_for_prefix(
                        &element,
                        prefix,
                        reader.decoder(),
                        context,
                    )?
                    else {
                        continue;
                    };
                    if dialect.replace(candidate).is_some() {
                        return Err(invalid(
                            "SVG blip carries relationship attributes from two dialects",
                        ));
                    }
                }
                return Ok(dialect.unwrap_or(fallback));
            },
            Event::Eof => return Err(invalid("SVG blip fragment has no root element")),
            _ => {},
        }
        buffer.clear();
    }
}

fn relationship_dialect_for_prefix(
    element: &BytesStart<'_>,
    prefix: &str,
    decoder: Decoder,
    context: &NamespaceContext,
) -> Result<Option<RelationshipDialect>> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let Some(binding) = attribute.key.as_namespace_binding() else {
            continue;
        };
        let binding_prefix = std::str::from_utf8(prefix_bytes(binding).unwrap_or_default())
            .map_err(|error| Error::Invalid(error.to_string()))?;
        if binding_prefix != prefix {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        return Ok(relationship_dialect_from_uri(value.as_bytes()));
    }
    let mut candidate = None;
    context.visit_visible(|candidate_prefix, uri| {
        if candidate.is_none() && candidate_prefix == Some(prefix) {
            candidate = relationship_dialect_from_uri(uri.as_bytes());
        }
    })?;
    Ok(candidate)
}

fn relationship_dialect_from_uri(uri: &[u8]) -> Option<RelationshipDialect> {
    if uri == RELATIONSHIPS {
        Some(RelationshipDialect::Transitional)
    } else if uri == STRICT_RELATIONSHIPS {
        Some(RelationshipDialect::Strict)
    } else {
        None
    }
}

fn extend_namespace_context(
    parent: &NamespaceContext,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<NamespaceContext> {
    let mut declarations = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let Some(prefix) = attribute.key.as_namespace_binding() else {
            continue;
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        validate_namespace_binding(prefix, value.as_bytes())?;
        let prefix = prefix_bytes(prefix)
            .map(|prefix| {
                std::str::from_utf8(prefix).map_err(|error| Error::Invalid(error.to_string()))
            })
            .transpose()?;
        let namespace = SvgNamespace::new(prefix, value.as_ref())
            .map_err(|error| invalid(format!("invalid DOCX namespace binding: {error}")))?;
        declarations
            .try_reserve(1)
            .map_err(|source| allocation("DOCX namespace context declarations", source))?;
        declarations.push(namespace);
    }
    let context = parent
        .child(declarations)
        .map_err(|error| invalid(format!("invalid DOCX namespace context: {error}")))?;
    Ok(context)
}

fn namespace_declaration_bytes(element: &BytesStart<'_>, decoder: Decoder) -> Result<usize> {
    let mut bytes = 0usize;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let Some(prefix) = attribute.key.as_namespace_binding() else {
            continue;
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        validate_namespace_binding(prefix, value.as_bytes())?;
        let prefix_len = prefix_bytes(prefix).map_or(0, <[u8]>::len);
        let value_len = value.len();
        bytes = bytes
            .checked_add(prefix_len)
            .and_then(|length| length.checked_add(value_len))
            .ok_or_else(|| invalid("DOCX namespace context byte count overflow"))?;
    }
    Ok(bytes)
}

fn validate_attributes(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    _limits: ScanLimits,
) -> Result<usize> {
    let mut count = 0usize;
    let mut declarations = 0usize;
    let mut expanded_names: Vec<(Box<[u8]>, Box<[u8]>)> = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        count = count
            .checked_add(1)
            .ok_or_else(|| limit("DOCX XML attributes", usize::MAX, MAX_ATTRIBUTES))?;
        if count > MAX_ATTRIBUTES {
            return Err(limit("DOCX XML attributes", count, MAX_ATTRIBUTES));
        }
        validate_qname(attribute.key.as_ref(), "attribute")?;
        if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(limit(
                "DOCX XML attribute value bytes",
                attribute.value.len(),
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        validate_raw_attribute(attribute.value.as_ref())?;
        let mut namespace_uri = Vec::new();
        let local_qname = attribute.key.local_name();
        let mut local_name = local_qname.as_ref();
        if let Some(prefix) = attribute.key.as_namespace_binding() {
            declarations = declarations.checked_add(1).ok_or_else(|| {
                limit(
                    "DOCX namespace declarations",
                    usize::MAX,
                    MAX_NAMESPACE_DECLARATIONS,
                )
            })?;
            if declarations > MAX_NAMESPACE_DECLARATIONS {
                return Err(limit(
                    "DOCX namespace declarations",
                    declarations,
                    MAX_NAMESPACE_DECLARATIONS,
                ));
            }
            if attribute.value.len() > MAX_NAMESPACE_BYTES {
                return Err(limit(
                    "DOCX namespace URI bytes",
                    attribute.value.len(),
                    MAX_NAMESPACE_BYTES,
                ));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(xml_error)?;
            validate_namespace_binding(prefix, value.as_bytes())?;
            namespace_uri.extend_from_slice(XMLNS_NAMESPACE);
            local_name = prefix_bytes(prefix).unwrap_or_default();
        } else {
            match reader.resolver().resolve_attribute(attribute.key) {
                (ResolveResult::Unknown(prefix), _) if !prefix.is_empty() => {
                    return Err(invalid("DOCX XML attribute uses an unbound prefix"));
                },
                (ResolveResult::Bound(Namespace([])), _) => {
                    return Err(invalid(
                        "DOCX prefixed XML attribute uses an empty namespace binding",
                    ));
                },
                (ResolveResult::Bound(Namespace(value)), _) => {
                    let value = decode_namespace_uri(value, reader.decoder())?;
                    if value.is_empty() {
                        return Err(invalid(
                            "DOCX prefixed XML attribute uses an empty namespace binding",
                        ));
                    }
                    namespace_uri.extend_from_slice(value.as_bytes());
                },
                _ => {},
            }
        }
        if namespace_uri.len() > MAX_NAMESPACE_BYTES {
            return Err(limit(
                "DOCX expanded attribute namespace URI bytes",
                namespace_uri.len(),
                MAX_NAMESPACE_BYTES,
            ));
        }
        if expanded_names.iter().any(|(namespace, local)| {
            namespace.as_ref() == namespace_uri.as_slice() && local.as_ref() == local_name
        }) {
            return Err(invalid("DOCX element has duplicate expanded attributes"));
        }
        let mut namespace_copy = Vec::new();
        namespace_copy
            .try_reserve_exact(namespace_uri.len())
            .map_err(|source| allocation("DOCX expanded attribute namespace", source))?;
        namespace_copy.extend_from_slice(&namespace_uri);
        let mut local_copy = Vec::new();
        local_copy
            .try_reserve_exact(local_name.len())
            .map_err(|source| allocation("DOCX expanded attribute local name", source))?;
        local_copy.extend_from_slice(local_name);
        expanded_names
            .try_reserve(1)
            .map_err(|source| allocation("DOCX expanded attribute names", source))?;
        expanded_names.push((
            namespace_copy.into_boxed_slice(),
            local_copy.into_boxed_slice(),
        ));
    }
    Ok(declarations)
}

fn validate_namespace_binding(prefix: PrefixDeclaration<'_>, uri: &[u8]) -> Result<()> {
    let prefix = prefix_bytes(prefix).unwrap_or_default();
    if prefix.len() > MAX_NAME_BYTES {
        return Err(limit(
            "DOCX namespace prefix bytes",
            prefix.len(),
            MAX_NAME_BYTES,
        ));
    }
    let uri_text = std::str::from_utf8(uri).map_err(|error| Error::Invalid(error.to_string()))?;
    if prefix == b"xmlns" || uri_text == std::str::from_utf8(XMLNS_NAMESPACE).unwrap_or_default() {
        return Err(invalid(
            "DOCX namespace binding uses the reserved XMLNS URI",
        ));
    }
    if uri_text == std::str::from_utf8(XML_NAMESPACE).unwrap_or_default() && prefix != b"xml" {
        return Err(invalid("DOCX XML namespace URI must use the xml prefix"));
    }
    if prefix == b"xml" && uri != XML_NAMESPACE {
        return Err(invalid("DOCX xml prefix has the wrong namespace URI"));
    }
    if !prefix.is_empty() && uri.is_empty() {
        return Err(invalid("DOCX prefixed namespace binding cannot be empty"));
    }
    Ok(())
}

fn extension_uri(element: &BytesStart<'_>, decoder: Decoder) -> Result<(bool, Box<[u8]>, bool)> {
    let mut found = false;
    let mut admitted = false;
    let mut malformed = false;
    let mut lexical = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        if attribute.key.as_ref() != b"uri" {
            // An admitted `a:ext` has a closed attribute vocabulary.  Keep
            // unknown-URI extension payload opaque, but refuse typed
            // inference when a recognized owner carries an unexpected attr.
            malformed = true;
            continue;
        }
        if found {
            malformed = true;
            continue;
        }
        found = true;
        if attribute.value.len() > MAX_NAMESPACE_BYTES {
            return Err(limit(
                "DOCX extension URI lexical bytes",
                attribute.value.len(),
                MAX_NAMESPACE_BYTES,
            ));
        }
        lexical
            .try_reserve_exact(attribute.value.len())
            .map_err(|source| allocation("DOCX extension URI lexical bytes", source))?;
        lexical.extend_from_slice(attribute.value.as_ref());
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        let Some(token) = xsd_token_atom(&value) else {
            malformed = true;
            continue;
        };
        admitted = token.as_bytes() == SVG_EXTENSION_URI;
    }
    if !found {
        malformed = true;
    }
    Ok((admitted, lexical.into_boxed_slice(), malformed && admitted))
}

fn relationship_attribute(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    local: &[u8],
) -> Result<Option<(Box<str>, RelationshipDialect)>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.local_name().as_ref() != local {
            continue;
        }
        let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
        let dialect = match namespace {
            ResolveResult::Bound(Namespace(value)) => {
                let value = decode_namespace_uri(value, reader.decoder())?;
                if value.as_bytes() == RELATIONSHIPS {
                    RelationshipDialect::Transitional
                } else if value.as_bytes() == STRICT_RELATIONSHIPS {
                    RelationshipDialect::Strict
                } else {
                    continue;
                }
            },
            _ => continue,
        };
        if result.is_some() {
            return Err(invalid(
                "DOCX drawing has duplicate relationship attributes",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?;
        if value.len() > MAX_RELATIONSHIP_ID_BYTES || value.is_empty() || !is_ncname(&value) {
            return Err(invalid("DOCX drawing relationship ID is invalid"));
        }
        result = Some((value.into_owned().into_boxed_str(), dialect));
    }
    Ok(result)
}

fn unqualified_attribute(
    element: &BytesStart<'_>,
    local: &[u8],
    decoder: Decoder,
) -> Result<Option<String>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == local {
            if result.is_some() {
                return Err(invalid("DOCX drawing has duplicate unqualified attributes"));
            }
            result = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(xml_error)?
                    .into_owned(),
            );
        }
    }
    Ok(result)
}

fn element_range<'a>(
    source: &'a [u8],
    start: usize,
    end: usize,
    element: &BytesStart<'_>,
    resolved: &Resolved,
) -> Result<ElementRange<'a>> {
    let prefix = element
        .name()
        .prefix()
        .map(|prefix| prefix.as_ref().to_vec())
        .unwrap_or_default()
        .into_boxed_slice();
    let namespace_uri = match resolved {
        Resolved::Known(uri) => Box::from(*uri),
        _ => Box::from([]),
    };
    Ok(ElementRange {
        source,
        range: ByteRange::new(start, end),
        start_end: end,
        close_start: None,
        prefix,
        namespace_uri,
    })
}

fn prefix_bytes(prefix: PrefixDeclaration<'_>) -> Option<&[u8]> {
    match prefix {
        PrefixDeclaration::Default => None,
        PrefixDeclaration::Named(prefix) => Some(prefix),
    }
}

fn validate_qname(value: &[u8], kind: &str) -> Result<()> {
    if value.len() > MAX_NAME_BYTES {
        return Err(invalid(format!(
            "DOCX XML {kind} QName exceeds {MAX_NAME_BYTES} bytes"
        )));
    }
    let text = std::str::from_utf8(value).map_err(|error| Error::Invalid(error.to_string()))?;
    let mut pieces = text.split(':');
    let first = pieces.next().unwrap_or_default();
    let second = pieces.next();
    if pieces.next().is_some() || !is_ncname(first) || second.is_some_and(|value| !is_ncname(value))
    {
        return Err(invalid(format!("invalid DOCX XML {kind} QName")));
    }
    Ok(())
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn is_xml_char(value: char) -> bool {
    matches!(
        value as u32,
        0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    )
}

fn validate_xml_text(bytes: &[u8]) -> Result<()> {
    let value = std::str::from_utf8(bytes).map_err(|error| Error::Invalid(error.to_string()))?;
    if bytes.windows(3).any(|window| window == b"]]>") {
        return Err(invalid(
            "DOCX XML text contains the forbidden ]]> delimiter",
        ));
    }
    if value.chars().all(is_xml_char) {
        Ok(())
    } else {
        Err(invalid("DOCX XML text contains an invalid XML character"))
    }
}

fn validate_xml_characters(bytes: &[u8]) -> Result<()> {
    let value = std::str::from_utf8(bytes).map_err(|error| Error::Invalid(error.to_string()))?;
    if value.chars().all(is_xml_char) {
        Ok(())
    } else {
        Err(invalid("DOCX XML value contains an invalid XML character"))
    }
}

fn validate_raw_attribute(bytes: &[u8]) -> Result<()> {
    if bytes.contains(&b'<') {
        return Err(invalid(
            "DOCX XML attribute contains a forbidden raw delimiter",
        ));
    }
    Ok(())
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let version = declaration.version().map_err(xml_error)?;
    if version.as_ref() != b"1.0" {
        return Err(invalid("DOCX XML declaration must use version 1.0"));
    }
    if let Some(encoding) = declaration.encoding() {
        let encoding = encoding.map_err(xml_error)?;
        if !encoding.as_ref().eq_ignore_ascii_case(b"utf-8") {
            return Err(invalid("DOCX XML declaration must use UTF-8"));
        }
    }
    Ok(())
}

fn validate_processing_instruction(instruction: &[u8]) -> Result<()> {
    validate_xml_characters(instruction)?;
    let end = instruction
        .iter()
        .position(u8::is_ascii_whitespace)
        .unwrap_or(instruction.len());
    let target = instruction
        .get(..end)
        .ok_or_else(|| invalid("empty DOCX processing-instruction target"))?;
    let target = std::str::from_utf8(target).map_err(|error| Error::Invalid(error.to_string()))?;
    if !litchi_ooxml_common::xml_name::is_xml_name(target) || target.eq_ignore_ascii_case("xml") {
        return Err(invalid("invalid DOCX processing-instruction target"));
    }
    Ok(())
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("DOCX drawing XML offset exceeds usize"))
}

fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

fn limit(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::InvalidFormat(format!(
        "DOCX drawing {resource} exceeds {maximum}: {actual}"
    ))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
