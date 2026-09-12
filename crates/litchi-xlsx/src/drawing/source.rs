//! Bounded, source-position preserving inspection of worksheet drawings.
//!
//! The ordinary [`super::parse`] inventory is intentionally a semantic view.
//! It is not a preservation index: a drawing may inherit namespace bindings
//! from its root, contain opaque extension payloads, or use a non-canonical
//! lexical spelling which must survive an edit.  This module scans one
//! `SpreadsheetDrawing` source member once and records only the ranges and
//! relationship metadata needed by the SVG owner.  The source bytes remain
//! owned by the package/`PartData` caller; no drawing-sized copy is retained.

use std::collections::HashMap;
use std::sync::Arc;

use litchi_drawingml::svg_blip::{self, Namespace, NamespaceContext, SvgBlip};
use litchi_spreadsheet_drawing::shape::{
    Anchor as DrawingAnchor, CellMarker, EditAs, Emu, EmuExtent, EmuOffset,
};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesDecl, BytesRef, BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::reader::Reader;

use super::model::Drawing;
use crate::error::{Error, Result, allocation, invalid as xlsx_invalid};
use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value, xsd_token_atom};

const SPREADSHEET_DRAWING: &[u8] =
    b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_SPREADSHEET_DRAWING: &[u8] =
    b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const DRAWINGML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const RELATIONSHIPS: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const SVG_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/drawing/2016/SVG/main";
const SVG_EXTENSION_URI: &[u8] = b"{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &[u8] = b"http://www.w3.org/2000/xmlns/";

/// The source member's bounded XML ceiling.
pub const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
/// The maximum number of XML events inspected by the source scanner.
pub const MAX_XML_NODES: usize = 1_000_000;
/// The maximum XML nesting depth inspected by the source scanner.
pub const MAX_XML_DEPTH: usize = 256;
/// The maximum direct picture candidates retained by one scan.
pub const MAX_PICTURES: usize = 100_000;
/// The maximum direct SpreadsheetDrawing content-part candidates retained by
/// one scan.
pub const MAX_CONTENT_PARTS: usize = 100_000;
/// The maximum relationship references retained per direct picture.
pub const MAX_RELATIONSHIP_REFERENCES: usize = 4_096;
/// The maximum qualified-name or namespace lexical length retained by a scan.
pub const MAX_NAME_BYTES: usize = 4 * 1024;
/// The maximum lexical bytes of one namespace URI or extension URI value.
pub const MAX_NAMESPACE_BYTES: usize = 16 * 1024;
/// The maximum decoded relationship identifier retained by one source index.
pub const MAX_RELATIONSHIP_ID_BYTES: usize = 256;
/// The maximum attributes on one element.
pub const MAX_ATTRIBUTES: usize = 256;
/// The maximum raw value bytes decoded from one non-namespace attribute.
pub const MAX_ATTRIBUTE_VALUE_BYTES: usize = 1024 * 1024;
/// The maximum namespace declarations on one element.
pub const MAX_NAMESPACE_DECLARATIONS: usize = 256;
/// The maximum active namespace bindings.
pub const MAX_ACTIVE_NAMESPACE_BINDINGS: usize = 16 * 1024;
/// The maximum marker text accepted by an anchor geometry parser.
pub const MAX_MARKER_TEXT_BYTES: usize = 64;
/// The maximum output used to namespace-complete one standalone fragment.
pub const MAX_FRAGMENT_BYTES: usize = 16 * 1024 * 1024;

/// A half-open range into the immutable source member.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[must_use]
pub struct ByteRange {
    /// Inclusive byte offset.
    pub start: usize,
    /// Exclusive byte offset.
    pub end: usize,
}

impl ByteRange {
    /// Construct a checked source range.
    pub const fn new(start: usize, end: usize) -> Self {
        Self { start, end }
    }

    /// Return the range length when the endpoints are ordered.
    pub fn len(self) -> Result<usize> {
        self.end
            .checked_sub(self.start)
            .ok_or_else(|| invalid("source range has descending endpoints"))
    }

    /// Return whether this range has no bytes.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.start == self.end
    }

    /// Borrow this range from a source member.
    pub fn slice(self, source: &[u8]) -> Result<&[u8]> {
        if self.start > self.end || self.end > source.len() {
            return Err(invalid("source range lies outside the drawing member"));
        }
        Ok(&source[self.start..self.end])
    }
}

/// A non-owning source lease carried by every escaped range projection.
///
/// The custom equality and debug implementations deliberately avoid walking
/// the complete drawing member once per picture.  Identity and length are
/// enough for source provenance because the lifetime keeps the borrowed
/// member alive and prevents safe in-place mutation while a projection is
/// retained.
#[derive(Clone, Copy)]
struct SourceLease<'a> {
    bytes: &'a [u8],
}

impl<'a> SourceLease<'a> {
    const fn new(bytes: &'a [u8]) -> Self {
        Self { bytes }
    }

    const fn bytes(self) -> &'a [u8] {
        self.bytes
    }
}

impl std::fmt::Debug for SourceLease<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SourceLease")
            .field("len", &self.bytes.len())
            .finish()
    }
}

impl PartialEq for SourceLease<'_> {
    fn eq(&self, other: &Self) -> bool {
        std::ptr::eq(self.bytes.as_ptr(), other.bytes.as_ptr())
            && self.bytes.len() == other.bytes.len()
    }
}

impl Eq for SourceLease<'_> {}

/// Source range for an element whose opening and closing tags may be edited
/// independently. `close_start` is `None` for a self-closing element.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
#[must_use]
pub struct ElementRange {
    range: ByteRange,
    start_end: usize,
    close_start: Option<usize>,
    prefix: Box<[u8]>,
    namespace_uri: Box<[u8]>,
}

impl ElementRange {
    /// Complete element range.
    pub const fn range(&self) -> ByteRange {
        self.range
    }

    /// End offset immediately after the opening tag.
    #[must_use]
    pub const fn start_end(&self) -> usize {
        self.start_end
    }

    /// Offset where a non-self-closing element's closing tag begins.
    #[must_use]
    pub const fn close_start(&self) -> Option<usize> {
        self.close_start
    }

    /// Prefix bytes used by the source element's qualified name.
    #[must_use]
    pub fn prefix(&self) -> &[u8] {
        &self.prefix
    }

    /// Expanded namespace URI resolved at the element's opening tag.
    #[must_use]
    pub fn namespace_uri(&self) -> &[u8] {
        &self.namespace_uri
    }

    /// Whether the source element uses the self-closing lexical form.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.close_start.is_none()
    }
}

/// Core OOXML dialect detected from the drawing namespace.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum DrawingDialect {
    /// Transitional DrawingML namespaces are in use.
    Transitional,
    /// Strict DrawingML namespaces are in use.
    Strict,
}

/// Physical relationship dialect observed on the host drawing member.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum RelationshipDialect {
    /// Transitional relationship namespace.
    Transitional,
    /// Strict relationship namespace.
    Strict,
}

/// SpreadsheetDrawing content-part profile recognized by the source owner.
///
/// The Office 2010 group extension is deliberately represented by the enum
/// for future compatibility, but this read slice publishes only
/// [`Self::CoreAnchor`].  Its relationship profile remains unresolved by the
/// local evidence and is therefore not guessed here.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ContentPartProfile {
    /// Core `xdr:contentPart` directly hosted by an anchor.
    CoreAnchor,
    /// Office 2010 `xdr14:contentPart` hosted by a group.
    GroupExtension,
}

impl RelationshipDialect {
    /// Return the relationship namespace used by new `asvg:svgBlip` XML.
    ///
    /// The MS-SVG extension schema imports the Transitional `AG_Blob`
    /// vocabulary even when the host drawing is Strict.  This is therefore
    /// deliberately independent from the physical host dialect.
    #[must_use]
    pub const fn authored_svg_attribute_namespace() -> &'static str {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    }

    /// Return the physical image relationship namespace for this host.
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

/// A relationship-bearing attribute found anywhere below one direct picture.
///
/// Opaque descendants are included.  A lifecycle owner can consequently use
/// this set for conservative shared-edge cleanup without searching XML text
/// for relationship-ID substrings.
#[derive(Clone, Debug, PartialEq, Eq)]
#[must_use]
pub struct RelationshipReference {
    id: Box<str>,
    local_name: Box<str>,
    range: ByteRange,
    dialect: RelationshipDialect,
}

impl RelationshipReference {
    /// Borrow the decoded relationship ID.
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Borrow the relationship attribute's local name (`embed`, `link`, or
    /// another producer-defined relationship attribute).
    #[must_use]
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Source span of the enclosing start tag.
    pub const fn range(&self) -> ByteRange {
        self.range
    }

    /// Dialect of the relationship attribute namespace.
    #[must_use]
    pub const fn dialect(&self) -> RelationshipDialect {
        self.dialect
    }

    fn try_clone_for_source(&self) -> Result<Self> {
        let mut id = String::new();
        id.try_reserve_exact(self.id.len())
            .map_err(|source| allocation("drawing source relationship ID", source))?;
        id.push_str(&self.id);
        let mut local_name = String::new();
        local_name
            .try_reserve_exact(self.local_name.len())
            .map_err(|source| allocation("drawing source relationship name", source))?;
        local_name.push_str(&self.local_name);
        Ok(Self {
            id: id.into_boxed_str(),
            local_name: local_name.into_boxed_str(),
            range: self.range,
            dialect: self.dialect,
        })
    }
}

/// A direct recognized SVG owner and its exact source spans.
#[derive(Clone, Debug, PartialEq, Eq)]
#[must_use]
pub struct SvgOwner<'a> {
    /// The package-owned member borrowed by every range in this projection.
    ///
    /// Keeping this borrow in escaped records is part of the provenance
    /// contract: a cloned owner cannot outlive the bytes it indexes, so a
    /// caller cannot safely mutate a same-length source behind the record.
    source: SourceLease<'a>,
    extension: ElementRange,
    svg_blip: ByteRange,
    uri_lexical: Box<[u8]>,
    value: SvgBlip,
    relationship_dialect: RelationshipDialect,
    namespace_context: Arc<NamespaceContext>,
}

impl<'a> SvgOwner<'a> {
    /// The complete `a:ext` source range.
    pub const fn extension_range(&self) -> ByteRange {
        self.extension.range()
    }

    /// Exact source element metadata for the admitted `a:ext` owner.
    pub const fn extension_element(&self) -> &ElementRange {
        &self.extension
    }

    /// Opening-tag range accepted by source XML splice operations.
    pub const fn extension_opening_range(&self) -> ByteRange {
        ByteRange::new(self.extension.range().start, self.extension.start_end())
    }

    /// The complete `asvg:svgBlip` source range.
    pub const fn svg_blip_range(&self) -> ByteRange {
        self.svg_blip
    }

    /// Return the exact lexical bytes of the admitted `uri` attribute value.
    #[must_use]
    pub fn uri_lexical(&self) -> &[u8] {
        &self.uri_lexical
    }

    /// Return the parsed, bounded shared SVG-blip projection.
    pub const fn value(&self) -> &SvgBlip {
        &self.value
    }

    /// Return the embedded relationship ID, if present.
    #[must_use]
    pub fn embedded_relationship_id(&self) -> Option<&str> {
        self.value.embedded().map(|value| value.as_str())
    }

    /// Return the linked relationship ID, if present.
    #[must_use]
    pub fn linked_relationship_id(&self) -> Option<&str> {
        self.value.linked().map(|value| value.as_str())
    }

    /// Return the physical relationship dialect observed on the SVG
    /// extension attribute.
    #[must_use]
    pub const fn relationship_dialect(&self) -> RelationshipDialect {
        self.relationship_dialect
    }

    /// Return the fragment with inherited ancestor namespace declarations
    /// added to its root, suitable for an independent shared-codec read.
    pub fn namespace_complete(&self, source: &[u8], max_output_bytes: usize) -> Result<Vec<u8>> {
        self.check_source(source)?;
        let fragment = self.svg_blip.slice(source)?;
        namespace_complete_element_fragment(
            fragment,
            &namespace_context_bindings(&self.namespace_context)?,
            max_output_bytes,
        )
    }

    /// Return the parsed SVG blip's exact raw source fragment as retained by
    /// the shared codec.  Contextual values deliberately do not expose this
    /// fragment through `SvgBlip::source`, because inherited bindings are not
    /// standalone declarations.  The complete owner source range remains
    /// available via [`Self::svg_blip_range`] for byte-preserving host edits.
    #[must_use]
    pub fn parsed_source(&self) -> Option<&[u8]> {
        self.value.raw_source()
    }

    fn check_source(&self, source: &[u8]) -> Result<()> {
        let expected = self.source.bytes();
        if !std::ptr::eq(source.as_ptr(), expected.as_ptr()) || source.len() != expected.len() {
            return Err(invalid(
                "SVG owner source does not match its scanned member",
            ));
        }
        Ok(())
    }
}

/// Projection status for one picture's direct extension list.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum SvgOwnerState<'a> {
    /// No extension was observed.
    None,
    /// One valid embedded `asvg:svgBlip` owner was observed.
    Embedded(SvgOwner<'a>),
    /// One valid linked `asvg:svgBlip` owner was observed.  The first edit
    /// profile treats this as inert and refuses to mutate it.
    Linked(SvgOwner<'a>),
    /// More than one admitted owner was observed.
    Ambiguous,
    /// An admitted extension was malformed, MCE-wrapped, or otherwise unsafe
    /// to project.  The source remains available to an explicit raw editor.
    Refused,
    /// A non-admitted extension was retained as opaque content.
    Opaque,
}

impl<'a> SvgOwnerState<'a> {
    /// Return a valid direct owner, excluding linked and refused states.
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

/// One source-backed direct `xdr:pic` candidate.
///
/// The source lifetime is retained by cloned candidates as well as by the
/// containing [`SourceDrawing`].  This prevents a caller from mutating the
/// original member in place after dropping the index while an escaped range
/// projection is still alive:
///
/// ```compile_fail
/// use litchi_xlsx::drawing::SourceDrawing;
///
/// let mut xml = Vec::new();
/// let drawing = SourceDrawing::scan(&xml).unwrap();
/// let picture = drawing.picture(0).unwrap().clone();
/// drop(drawing);
/// xml.push(0);
/// let _ = picture.namespace_complete_picture(&xml, 1024);
/// ```
#[derive(Clone, Debug, PartialEq, Eq)]
#[must_use]
pub struct PictureSource<'a> {
    /// The package-owned member borrowed by every range in this projection.
    source: SourceLease<'a>,
    drawing_ordinal: usize,
    picture_ordinal: usize,
    anchor: DrawingAnchor,
    anchor_range: ByteRange,
    picture_range: ByteRange,
    blip_range: ElementRange,
    ext_list_range: Option<ElementRange>,
    c_nv_pr_id: Option<Box<str>>,
    c_nv_pr_range: Option<ByteRange>,
    raster_relationship_id: Box<str>,
    svg_owner: SvgOwnerState<'a>,
    relationship_references: Vec<RelationshipReference>,
    namespace_context: Arc<NamespaceContext>,
    blip_namespace_context: Arc<NamespaceContext>,
    ext_list_namespace_context: Option<Arc<NamespaceContext>>,
}

impl<'a> PictureSource<'a> {
    /// The drawing position used by a checked drawing selector.
    #[must_use]
    pub const fn drawing_ordinal(&self) -> usize {
        self.drawing_ordinal
    }

    /// The source-order direct picture position within this drawing.
    #[must_use]
    pub const fn picture_ordinal(&self) -> usize {
        self.picture_ordinal
    }

    /// Complete typed anchor geometry.
    #[must_use]
    pub const fn anchor(&self) -> &DrawingAnchor {
        &self.anchor
    }

    /// Exact anchor range in the original drawing member.
    pub const fn anchor_range(&self) -> ByteRange {
        self.anchor_range
    }

    /// Exact direct picture range in the original drawing member.
    pub const fn picture_range(&self) -> ByteRange {
        self.picture_range
    }

    /// Exact source range of the direct raster `a:blip` element.
    pub const fn blip_range(&self) -> &ElementRange {
        &self.blip_range
    }

    /// Resolved DrawingML namespace URI of the direct raster `a:blip`.
    ///
    /// This is a bounded token-level value, so an author can bind the source
    /// prefix on a standalone authored fragment without cloning the active
    /// namespace scope.
    #[must_use]
    pub fn blip_namespace_uri(&self) -> &[u8] {
        self.blip_range.namespace_uri()
    }

    /// Exact source range of the direct `a:extLst`, when present.
    #[must_use]
    pub const fn ext_list_range(&self) -> Option<&ElementRange> {
        self.ext_list_range.as_ref()
    }

    /// Resolved DrawingML namespace URI of the direct `a:extLst`, when one is
    /// present.
    #[must_use]
    pub fn ext_list_namespace_uri(&self) -> Option<&[u8]> {
        self.ext_list_range
            .as_ref()
            .map(ElementRange::namespace_uri)
    }

    /// Optional exact lexical `xdr:cNvPr@id` value.
    #[must_use]
    pub fn c_nv_pr_id(&self) -> Option<&str> {
        self.c_nv_pr_id.as_deref()
    }

    /// Optional complete `xdr:cNvPr` source range.
    #[must_use]
    pub const fn c_nv_pr_range(&self) -> Option<ByteRange> {
        self.c_nv_pr_range
    }

    /// Existing raster fallback relationship ID.
    #[must_use]
    pub fn raster_relationship_id(&self) -> &str {
        &self.raster_relationship_id
    }

    /// SVG owner projection, including an inert/refused status.
    #[must_use]
    pub const fn svg_owner(&self) -> &SvgOwnerState<'a> {
        &self.svg_owner
    }

    /// Every relationship-ID attribute below the direct picture, including
    /// opaque extension descendants.
    pub fn relationship_references(&self) -> &[RelationshipReference] {
        &self.relationship_references
    }

    /// Borrow the direct picture bytes from the package-owned source member.
    pub fn picture_bytes<'b>(&self, source: &'b [u8]) -> Result<&'b [u8]> {
        self.checked_slice(source, self.picture_range)
    }

    /// Borrow the exact raster picture XML range.
    pub fn anchor_bytes<'b>(&self, source: &'b [u8]) -> Result<&'b [u8]> {
        self.checked_slice(source, self.anchor_range)
    }

    /// Resolve a selected source range against the package-owned bytes.
    pub fn source_bytes<'b>(&self, source: &'b [u8], range: ByteRange) -> Result<&'b [u8]> {
        self.checked_slice(source, range)
    }

    /// Complete the direct picture fragment with inherited namespace
    /// declarations before passing it to a standalone parser.
    pub fn namespace_complete_picture(
        &self,
        source: &[u8],
        max_output_bytes: usize,
    ) -> Result<Vec<u8>> {
        self.check_source(source)?;
        namespace_complete_element_fragment(
            self.picture_bytes(source)?,
            &namespace_context_bindings(&self.namespace_context)?,
            max_output_bytes,
        )
    }

    /// Complete the direct raster `a:blip` fragment with the namespace
    /// bindings active at that element.  The result is bounded and suitable
    /// for a standalone shared-codec parse; the original bytes remain
    /// available through [`Self::blip_range`].
    pub fn namespace_complete_blip(
        &self,
        source: &[u8],
        max_output_bytes: usize,
    ) -> Result<Vec<u8>> {
        self.check_source(source)?;
        namespace_complete_element_fragment(
            self.blip_range.range().slice(source)?,
            &namespace_context_bindings(&self.blip_namespace_context)?,
            max_output_bytes,
        )
    }

    /// Complete the direct `a:extLst` fragment with the namespace bindings
    /// active at that element, when an extension list was present.
    pub fn namespace_complete_ext_list(
        &self,
        source: &[u8],
        max_output_bytes: usize,
    ) -> Result<Option<Vec<u8>>> {
        self.check_source(source)?;
        let Some(range) = self.ext_list_range.as_ref() else {
            return Ok(None);
        };
        let context = self
            .ext_list_namespace_context
            .as_ref()
            .ok_or_else(|| invalid("extLst namespace context is unavailable"))?;
        Ok(Some(namespace_complete_element_fragment(
            range.range().slice(source)?,
            &namespace_context_bindings(context)?,
            max_output_bytes,
        )?))
    }

    /// Return whether this owner has the unique, direct, embedded profile
    /// required by the first XLSX lifecycle transaction.
    #[must_use]
    pub fn is_direct_embedded_svg(&self) -> bool {
        matches!(self.svg_owner, SvgOwnerState::Embedded(_))
    }

    fn checked_slice<'b>(&self, source: &'b [u8], range: ByteRange) -> Result<&'b [u8]> {
        self.check_source(source)?;
        range.slice(source)
    }

    fn check_source(&self, source: &[u8]) -> Result<()> {
        let expected = self.source.bytes();
        if !std::ptr::eq(source.as_ptr(), expected.as_ptr()) || source.len() != expected.len() {
            return Err(invalid("picture source does not match its scanned member"));
        }
        Ok(())
    }
}

/// One source-backed core SpreadsheetDrawing `contentPart` directly hosted by
/// an anchor.
///
/// The owner retains the borrowed drawing source and immutable namespace
/// provenance alongside every range.  A selector therefore cannot outlive or
/// accidentally pair with a different drawing member.  Relationship and
/// target validation belongs to the package-owned worksheet facade; this
/// record only captures the source-authoritative owner edge and geometry.
#[derive(Clone, Debug, PartialEq, Eq)]
#[must_use]
pub struct ContentPartSource<'a> {
    source: SourceLease<'a>,
    drawing_ordinal: usize,
    content_part_ordinal: usize,
    profile: ContentPartProfile,
    anchor: DrawingAnchor,
    anchor_range: ByteRange,
    owner: ElementRange,
    relationship_id: Box<str>,
    relationship_dialect: RelationshipDialect,
    mce_ancestor: bool,
    namespace_context: Arc<NamespaceContext>,
}

impl<'a> ContentPartSource<'a> {
    /// Semantic drawing ordinal supplied by the worksheet selector.
    #[must_use]
    pub const fn drawing_ordinal(&self) -> usize {
        self.drawing_ordinal
    }

    /// Source-order ordinal among direct core content-part owners.
    #[must_use]
    pub const fn content_part_ordinal(&self) -> usize {
        self.content_part_ordinal
    }

    /// The recognized SpreadsheetDrawing content-part profile.
    #[must_use]
    pub const fn profile(&self) -> ContentPartProfile {
        self.profile
    }

    /// Complete anchor geometry containing the owner.
    #[must_use]
    pub const fn anchor(&self) -> &DrawingAnchor {
        &self.anchor
    }

    /// Exact enclosing anchor range in the source drawing member.
    pub const fn anchor_range(&self) -> ByteRange {
        self.anchor_range
    }

    /// Exact owner element range, including its opening and closing markup.
    pub const fn owner_element(&self) -> &ElementRange {
        &self.owner
    }

    /// Alias for [`Self::owner_element`] used by range-oriented callers.
    pub const fn owner_range(&self) -> ByteRange {
        self.owner.range()
    }

    /// Required `r:id` value authored by the owner.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Relationship namespace dialect used by the required `r:id`.
    #[must_use]
    pub const fn relationship_dialect(&self) -> RelationshipDialect {
        self.relationship_dialect
    }

    /// Whether the owner was encountered below markup-compatibility content.
    ///
    /// The direct core read profile refuses such ownership as ambiguous, so a
    /// published source owner currently always returns `false`.  Keeping the
    /// provenance bit makes that refusal explicit and leaves room for a later
    /// branch-aware owner without changing the range model.
    #[must_use]
    pub const fn has_mce_ancestor(&self) -> bool {
        self.mce_ancestor
    }

    /// Borrow the exact owner element bytes after checking source identity.
    pub fn owner_bytes<'b>(&self, source: &'b [u8]) -> Result<&'b [u8]> {
        self.checked_slice(source, self.owner.range())
    }

    /// Borrow the complete enclosing anchor bytes after checking source
    /// identity.
    pub fn anchor_bytes<'b>(&self, source: &'b [u8]) -> Result<&'b [u8]> {
        self.checked_slice(source, self.anchor_range)
    }

    /// Borrow the exact drawing member used to create this source record.
    #[must_use]
    pub const fn source(&self) -> &'a [u8] {
        self.source.bytes()
    }

    fn checked_slice<'b>(&self, source: &'b [u8], range: ByteRange) -> Result<&'b [u8]> {
        let expected = self.source.bytes();
        if !std::ptr::eq(source.as_ptr(), expected.as_ptr()) || source.len() != expected.len() {
            return Err(invalid(
                "content-part source does not match its scanned member",
            ));
        }
        range.slice(source)
    }
}

/// A complete source-backed drawing inventory.
#[derive(Clone, Debug, PartialEq, Eq)]
#[must_use]
pub struct SourceDrawing<'a> {
    source: SourceLease<'a>,
    drawing_ordinal: usize,
    dialect: DrawingDialect,
    relationship_dialect: RelationshipDialect,
    pictures: Vec<PictureSource<'a>>,
    content_parts: Vec<ContentPartSource<'a>>,
    relationship_references: Vec<RelationshipReference>,
}

impl<'a> SourceDrawing<'a> {
    /// Scan a worksheet drawing in source order.
    pub fn scan(source: &'a [u8]) -> Result<Self> {
        Self::scan_with_ordinal(source, 0)
    }

    /// Scan a worksheet drawing with the caller's semantic drawing ordinal.
    pub fn scan_with_ordinal(source: &'a [u8], drawing_ordinal: usize) -> Result<Self> {
        Self::scan_with_limits(source, drawing_ordinal, ScanLimits::default())
    }

    /// Scan with explicit finite source/output limits.
    pub fn scan_with_limits(
        source: &'a [u8],
        drawing_ordinal: usize,
        limits: ScanLimits,
    ) -> Result<Self> {
        Scanner::new(source, drawing_ordinal, limits)?.run()
    }

    /// Return the semantic drawing ordinal supplied by the host selector.
    #[must_use]
    pub const fn drawing_ordinal(&self) -> usize {
        self.drawing_ordinal
    }

    /// Return the detected core DrawingML dialect.
    #[must_use]
    pub const fn dialect(&self) -> DrawingDialect {
        self.dialect
    }

    /// Return the physical host relationship dialect.
    #[must_use]
    pub const fn relationship_dialect(&self) -> RelationshipDialect {
        self.relationship_dialect
    }

    /// Return direct pictures in source order.
    pub fn pictures(&self) -> &[PictureSource<'a>] {
        &self.pictures
    }

    /// Return direct core SpreadsheetDrawing content parts in source order.
    ///
    /// Group-extension `xdr14:contentPart` elements are intentionally absent:
    /// their relationship profile is unresolved and they remain opaque source
    /// markup until that profile is closed.
    pub fn content_parts(&self) -> &[ContentPartSource<'a>] {
        &self.content_parts
    }

    /// Borrow the exact source member used to build this index.
    #[must_use]
    pub const fn source(&self) -> &'a [u8] {
        self.source.bytes()
    }

    /// Every relationship-ID attribute in the complete drawing member,
    /// including opaque extension payloads and non-picture anchors.
    pub fn relationship_references(&self) -> &[RelationshipReference] {
        &self.relationship_references
    }

    /// Resolve a checked source-order picture selector.
    pub fn picture(&self, picture_ordinal: usize) -> Result<&PictureSource<'a>> {
        self.pictures.get(picture_ordinal).ok_or_else(|| {
            Error::Invalid(format!(
                "drawing picture ordinal {picture_ordinal} is outside {} pictures",
                self.pictures.len()
            ))
        })
    }

    /// Resolve a checked source-order direct core content-part selector.
    pub fn content_part(&self, content_part_ordinal: usize) -> Result<&ContentPartSource<'a>> {
        self.content_parts.get(content_part_ordinal).ok_or_else(|| {
            Error::Invalid(format!(
                "drawing content-part ordinal {content_part_ordinal} is outside {} content parts",
                self.content_parts.len()
            ))
        })
    }

    /// Match a source candidate to a typed inventory picture by source-order
    /// ordinal.  The model is intentionally borrowed; it never becomes a
    /// preservation authority.
    pub fn match_typed_picture<'b>(
        &self,
        drawing: &'b Drawing,
        picture_ordinal: usize,
    ) -> Result<&'b super::model::Picture> {
        let source = self.picture(picture_ordinal)?;
        let typed = drawing
            .pictures()
            .nth(source.picture_ordinal)
            .ok_or_else(|| invalid("source picture has no matching typed inventory picture"))?;
        if typed.relationship_id.as_str() != source.raster_relationship_id.as_ref()
            || typed.drawing_anchor != source.anchor
        {
            return Err(invalid(
                "source picture does not match the typed drawing inventory",
            ));
        }
        Ok(typed)
    }

    /// Borrow a source range after checking it against the scanned member
    /// length.  The caller remains responsible for keeping the same source
    /// member alive.
    pub fn source_range<'b>(&self, source: &'b [u8], range: ByteRange) -> Result<&'b [u8]> {
        let expected = self.source.bytes();
        if !std::ptr::eq(source.as_ptr(), expected.as_ptr()) || source.len() != expected.len() {
            return Err(invalid(
                "source bytes differ from the scanned drawing member",
            ));
        }
        range.slice(source)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum AnchorKind {
    TwoCell,
    OneCell,
    Absolute,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum MarkerTarget {
    From,
    To,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum MarkerField {
    Column,
    ColumnOffset,
    Row,
    RowOffset,
}

impl MarkerField {
    const fn ordinal(self) -> u8 {
        match self {
            Self::Column => 0,
            Self::ColumnOffset => 1,
            Self::Row => 2,
            Self::RowOffset => 3,
        }
    }
}

#[derive(Default)]
struct Marker {
    column: Option<u32>,
    column_offset: Option<i64>,
    row: Option<u32>,
    row_offset: Option<i64>,
    next_field: u8,
}

impl Marker {
    fn finish(&self, label: &str) -> Result<CellMarker> {
        let column_offset = self
            .column_offset
            .ok_or_else(|| invalid(format!("{label} is missing its column offset")))?;
        let row_offset = self
            .row_offset
            .ok_or_else(|| invalid(format!("{label} is missing its row offset")))?;
        check_coordinate(column_offset, "drawing column offset")?;
        check_coordinate(row_offset, "drawing row offset")?;
        let marker = CellMarker {
            column: self
                .column
                .ok_or_else(|| invalid(format!("{label} is missing its column")))?,
            column_offset: Emu(column_offset),
            row: self
                .row
                .ok_or_else(|| invalid(format!("{label} is missing its row")))?,
            row_offset: Emu(row_offset),
        };
        check_marker_bounds(marker)?;
        Ok(marker)
    }

    fn set(&mut self, field: MarkerField, value: i64, label: &str) -> Result<()> {
        if field.ordinal() != self.next_field {
            return Err(invalid(format!("{label} fields are out of order")));
        }
        self.next_field = self
            .next_field
            .checked_add(1)
            .ok_or_else(|| limit("drawing marker fields", MAX_MARKER_TEXT_BYTES))?;
        match field {
            MarkerField::Column => {
                if value < 0 || value > u32::MAX as i64 {
                    return Err(invalid("drawing column is outside its numeric range"));
                }
                self.column = Some(value as u32);
            },
            MarkerField::ColumnOffset => self.column_offset = Some(value),
            MarkerField::Row => {
                if value < 0 || value > u32::MAX as i64 {
                    return Err(invalid("drawing row is outside its numeric range"));
                }
                self.row = Some(value as u32);
            },
            MarkerField::RowOffset => self.row_offset = Some(value),
        }
        Ok(())
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ObjectKind {
    Picture,
    Group,
    ContentPart,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum AnchorChild {
    From,
    To,
    Position,
    Extent,
    Object,
    ClientData,
}

struct AnchorCapture {
    kind: AnchorKind,
    range: ByteRange,
    edit_as: EditAs,
    from: Option<Marker>,
    to: Option<Marker>,
    position: Option<EmuOffset>,
    extent: Option<EmuExtent>,
    object_kind: Option<ObjectKind>,
    picture_index: Option<usize>,
    content_part_index: Option<usize>,
    phase: u8,
    client_data_seen: bool,
}

impl AnchorCapture {
    fn new(kind: AnchorKind, range: ByteRange, edit_as: EditAs) -> Self {
        Self {
            kind,
            range,
            edit_as,
            from: None,
            to: None,
            position: None,
            extent: None,
            object_kind: None,
            picture_index: None,
            content_part_index: None,
            phase: 0,
            client_data_seen: false,
        }
    }

    fn take_child(&mut self, child: AnchorChild) -> Result<()> {
        let expected = match (self.kind, self.phase) {
            (AnchorKind::TwoCell, 0) => AnchorChild::From,
            (AnchorKind::TwoCell, 1) => AnchorChild::To,
            (AnchorKind::TwoCell, 2) => AnchorChild::Object,
            (AnchorKind::TwoCell, 3) => AnchorChild::ClientData,
            (AnchorKind::OneCell, 0) => AnchorChild::From,
            (AnchorKind::OneCell, 1) => AnchorChild::Extent,
            (AnchorKind::OneCell, 2) => AnchorChild::Object,
            (AnchorKind::OneCell, 3) => AnchorChild::ClientData,
            (AnchorKind::Absolute, 0) => AnchorChild::Position,
            (AnchorKind::Absolute, 1) => AnchorChild::Extent,
            (AnchorKind::Absolute, 2) => AnchorChild::Object,
            (AnchorKind::Absolute, 3) => AnchorChild::ClientData,
            _ => return Err(invalid("drawing anchor has duplicate or trailing children")),
        };
        if expected != child {
            return Err(invalid("drawing anchor children are out of order"));
        }
        self.phase = self
            .phase
            .checked_add(1)
            .ok_or_else(|| limit("drawing anchor children", MAX_XML_DEPTH))?;
        Ok(())
    }

    fn finish(&self) -> Result<DrawingAnchor> {
        if self.phase != 4 || !self.client_data_seen {
            return Err(invalid("drawing anchor is missing required children"));
        }
        match self.kind {
            AnchorKind::TwoCell => {
                let from = self
                    .from
                    .as_ref()
                    .ok_or_else(|| invalid("drawing anchor is missing from marker"))?
                    .finish("drawing from marker")?;
                let to = self
                    .to
                    .as_ref()
                    .ok_or_else(|| invalid("drawing anchor is missing to marker"))?
                    .finish("drawing to marker")?;
                if to.row < from.row
                    || to.column < from.column
                    || (to.row == from.row && to.row_offset < from.row_offset)
                    || (to.column == from.column && to.column_offset < from.column_offset)
                {
                    return Err(invalid("drawing anchor has descending markers"));
                }
                Ok(DrawingAnchor::TwoCell {
                    from,
                    to,
                    edit_as: self.edit_as,
                })
            },
            AnchorKind::OneCell => Ok(DrawingAnchor::OneCell {
                from: self
                    .from
                    .as_ref()
                    .ok_or_else(|| invalid("one-cell anchor is missing from marker"))?
                    .finish("drawing from marker")?,
                extent: self
                    .extent
                    .ok_or_else(|| invalid("one-cell anchor is missing its extent"))?,
            }),
            AnchorKind::Absolute => Ok(DrawingAnchor::Absolute {
                position: self
                    .position
                    .ok_or_else(|| invalid("absolute anchor is missing its position"))?,
                extent: self
                    .extent
                    .ok_or_else(|| invalid("absolute anchor is missing its extent"))?,
            }),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Kind {
    Root,
    Anchor(AnchorKind),
    From(MarkerTarget),
    To(MarkerTarget),
    Marker(MarkerTarget, MarkerField),
    Position,
    Extent,
    Picture,
    PictureNvPr,
    CnvPr,
    BlipFill,
    Blip,
    ExtList,
    ClientData,
    SvgExt,
    OpaqueExt,
    SvgBlip,
    Group,
    ContentPart,
    Mce,
    Unknown,
}

#[derive(Clone, Debug)]
struct Frame {
    kind: Kind,
    namespace_declarations: usize,
    context_before: Arc<NamespaceContext>,
    source_start: usize,
    picture_index: Option<usize>,
    extension_index: Option<usize>,
}

#[derive(Clone, Debug)]
struct PendingExtension {
    range: ElementRange,
    admitted: bool,
    uri_lexical: Vec<u8>,
    svg_blip: Option<ByteRange>,
    svg_context: Option<Arc<NamespaceContext>>,
    svg_relationship_dialect: Option<RelationshipDialect>,
    svg_count: usize,
    malformed: bool,
    mce_ancestor: bool,
}

#[derive(Clone, Debug)]
struct PendingPicture {
    picture_range: ByteRange,
    anchor_range: ByteRange,
    anchor: Option<DrawingAnchor>,
    blip: Option<ElementRange>,
    ext_list: Option<ElementRange>,
    c_nv_pr_id: Option<Box<str>>,
    c_nv_pr_range: Option<ByteRange>,
    raster_relationship_id: Option<Box<str>>,
    extensions: Vec<PendingExtension>,
    relationship_references: Vec<RelationshipReference>,
    namespace_context: Arc<NamespaceContext>,
    blip_namespace_context: Option<Arc<NamespaceContext>>,
    ext_list_namespace_context: Option<Arc<NamespaceContext>>,
    mce_ancestor: bool,
    grouped_or_foreign_descendant: bool,
}

#[derive(Clone, Debug)]
struct PendingContentPart {
    owner: ElementRange,
    anchor_range: ByteRange,
    anchor: Option<DrawingAnchor>,
    relationship_id: Box<str>,
    relationship_dialect: RelationshipDialect,
    namespace_context: Arc<NamespaceContext>,
    mce_ancestor: bool,
}

struct Scanner<'a> {
    source: &'a [u8],
    drawing_ordinal: usize,
    limits: ScanLimits,
    namespaces: Namespaces,
    frames: Vec<Frame>,
    anchor: Option<AnchorCapture>,
    marker_text: String,
    pictures: Vec<PendingPicture>,
    content_parts: Vec<PendingContentPart>,
    global_relationship_references: Vec<RelationshipReference>,
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
        if source.len() > limits.max_xml_bytes || source.len() > MAX_XML_BYTES {
            return Err(limit(
                "drawing source XML bytes",
                limits.max_xml_bytes.min(MAX_XML_BYTES),
            ));
        }
        if limits.max_nodes == 0
            || limits.max_depth == 0
            || limits.max_pictures == 0
            || limits.max_content_parts == 0
            || limits.max_relationship_references == 0
            || limits.max_fragment_bytes == 0
        {
            return Err(invalid("drawing source scan limits must be nonzero"));
        }
        if limits.max_nodes > MAX_XML_NODES {
            return Err(limit("drawing source XML nodes", MAX_XML_NODES));
        }
        if limits.max_pictures > MAX_PICTURES {
            return Err(limit("drawing source direct pictures", MAX_PICTURES));
        }
        if limits.max_content_parts > MAX_CONTENT_PARTS {
            return Err(limit("drawing source content parts", MAX_CONTENT_PARTS));
        }
        if limits.max_relationship_references > MAX_RELATIONSHIP_REFERENCES {
            return Err(limit(
                "drawing source relationship references",
                MAX_RELATIONSHIP_REFERENCES,
            ));
        }
        let mut frames = Vec::new();
        frames
            .try_reserve(16)
            .map_err(|source| allocation("drawing source XML stack", source))?;
        let mut pictures = Vec::new();
        pictures
            .try_reserve(8.min(limits.max_pictures))
            .map_err(|source| allocation("drawing source picture index", source))?;
        let mut content_parts = Vec::new();
        content_parts
            .try_reserve(8.min(limits.max_content_parts))
            .map_err(|source| allocation("drawing source content-part index", source))?;
        let mut global_relationship_references = Vec::new();
        global_relationship_references
            .try_reserve(8)
            .map_err(|source| allocation("drawing source relationship index", source))?;
        Ok(Self {
            source,
            drawing_ordinal,
            limits,
            namespaces: Namespaces::default(),
            frames,
            anchor: None,
            marker_text: String::new(),
            pictures,
            content_parts,
            global_relationship_references,
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

    fn run(mut self) -> Result<SourceDrawing<'a>> {
        let mut reader = Reader::from_reader(self.source);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        reader.config_mut().check_comments = true;
        let source_prefix = if self.source.starts_with(b"\xEF\xBB\xBF") {
            3
        } else {
            0
        };
        let mut buffer = Vec::new();

        loop {
            let event_start = position(&reader, source_prefix)?;
            let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
            let event_end = position(&reader, source_prefix)?;
            if event_end < event_start || event_end > self.source.len() {
                return Err(invalid("XML event range is outside the drawing source"));
            }
            self.nodes = self
                .nodes
                .checked_add(1)
                .ok_or_else(|| limit("drawing source XML nodes", self.limits.max_nodes))?;
            if self.nodes > self.limits.max_nodes {
                return Err(limit("drawing source XML nodes", self.limits.max_nodes));
            }
            match event {
                Event::Start(element) => {
                    let frame =
                        self.start_element(&element, event_start, event_end, reader.decoder())?;
                    self.frames
                        .try_reserve(1)
                        .map_err(|source| allocation("drawing source XML stack", source))?;
                    self.frames.push(frame);
                    self.prolog = false;
                },
                Event::Empty(element) => {
                    let frame =
                        self.start_element(&element, event_start, event_end, reader.decoder())?;
                    let root = frame.kind == Kind::Root;
                    let depth =
                        self.frames.len().checked_add(1).ok_or_else(|| {
                            limit("drawing source XML depth", self.limits.max_depth)
                        })?;
                    let declarations = frame.namespace_declarations;
                    let context_before = Arc::clone(&frame.context_before);
                    self.finish_element(frame, event_start, event_end, true)?;
                    self.namespaces.pop(depth, declarations);
                    self.namespaces.context = context_before;
                    if root {
                        self.root_closed = true;
                    }
                    self.prolog = false;
                },
                Event::End(element) => {
                    let depth = self.frames.len();
                    if depth == 0 {
                        return Err(invalid("drawing XML contains an unmatched end element"));
                    }
                    validate_qname(element.name().as_ref(), "end element")?;
                    let frame = self
                        .frames
                        .pop()
                        .ok_or_else(|| invalid("drawing source XML stack underflow"))?;
                    let resolved = self.namespaces.resolve_element(element.name())?;
                    let root_close_matches = is_name(
                        resolved,
                        element.name(),
                        b"wsDr",
                        SPREADSHEET_DRAWING,
                        STRICT_SPREADSHEET_DRAWING,
                    );
                    self.finish_element(frame.clone(), event_start, event_end, false)?;
                    if frame.kind == Kind::Root {
                        if !root_close_matches {
                            return Err(invalid("drawing root close does not match xdr:wsDr"));
                        }
                        self.root_closed = true;
                    }
                    self.namespaces.pop(depth, frame.namespace_declarations);
                    self.namespaces.context = frame.context_before;
                },
                Event::Decl(declaration) => {
                    if self.declaration_seen || !self.prolog || self.root_seen {
                        return Err(invalid(
                            "XML declaration is not at the beginning of drawing XML",
                        ));
                    }
                    validate_declaration(&declaration)?;
                    self.declaration_seen = true;
                },
                Event::DocType(_) => return Err(invalid("DOCTYPE is forbidden in drawing XML")),
                Event::Text(text) => {
                    validate_xml_text(text.as_ref())?;
                    let outside = self.frames.is_empty();
                    let non_whitespace = !is_xml_whitespace(text.as_ref());
                    if outside && (!self.root_seen || self.root_closed) && non_whitespace {
                        return Err(invalid("non-whitespace text appears outside xdr:wsDr"));
                    }
                    if matches!(
                        self.frames.last().map(|frame| frame.kind),
                        Some(Kind::Marker(..))
                    ) {
                        self.append_marker_text(text.as_ref())?;
                    } else {
                        self.validate_text_context(outside, non_whitespace)?;
                    }
                },
                Event::CData(data) => {
                    validate_xml_characters(data.as_ref())?;
                    if self.frames.is_empty() {
                        return Err(invalid("CDATA appears outside drawing XML"));
                    }
                    if matches!(
                        self.frames.last().map(|frame| frame.kind),
                        Some(Kind::Marker(..))
                    ) {
                        self.append_marker_text(data.as_ref())?;
                    } else {
                        self.validate_text_context(false, !is_xml_whitespace(data.as_ref()))?;
                    }
                },
                Event::GeneralRef(reference) => {
                    validate_reference(&reference)?;
                    if self.frames.is_empty() {
                        return Err(invalid("entity reference appears outside drawing XML"));
                    }
                    let decoded_reference = decode_xml_reference(&reference)
                        .map_err(|error| Error::Invalid(error.to_string()))?;
                    let non_whitespace = !is_xml_whitespace(decoded_reference.as_bytes());
                    if matches!(
                        self.frames.last().map(|frame| frame.kind),
                        Some(Kind::Marker(..))
                    ) {
                        self.append_marker_text(decoded_reference.as_bytes())?;
                    } else {
                        self.validate_text_context(false, non_whitespace)?;
                    }
                },
                Event::Comment(comment) => validate_xml_characters(comment.as_ref())?,
                Event::PI(pi) => validate_processing_instruction(pi.as_ref())?,
                Event::Eof => break,
            }
            buffer.clear();
        }

        if !self.root_seen || !self.root_closed || !self.frames.is_empty() {
            return Err(invalid("drawing XML has an incomplete xdr:wsDr root"));
        }
        if self.anchor.is_some() || self.active_picture.is_some() {
            return Err(invalid("drawing XML has an incomplete anchor or picture"));
        }
        let dialect = self
            .dialect
            .ok_or_else(|| invalid("drawing XML has no detected core dialect"))?;
        let relationship_dialect = self.relationship_dialect.unwrap_or(match dialect {
            DrawingDialect::Transitional => RelationshipDialect::Transitional,
            DrawingDialect::Strict => RelationshipDialect::Strict,
        });
        let mut pictures = Vec::new();
        pictures
            .try_reserve_exact(self.pictures.len())
            .map_err(|source| allocation("drawing source picture projection", source))?;
        for (picture_ordinal, pending) in self.pictures.into_iter().enumerate() {
            let anchor = pending
                .anchor
                .ok_or_else(|| invalid("source picture has no complete anchor geometry"))?;
            let raster_relationship_id = pending
                .raster_relationship_id
                .as_ref()
                .cloned()
                .ok_or_else(|| invalid("direct picture has no raster r:embed fallback"))?;
            let svg_owner = project_svg_owner(
                self.source,
                &pending,
                relationship_dialect,
                self.limits.max_fragment_bytes,
            )?;
            pictures.push(PictureSource {
                source: SourceLease::new(self.source),
                drawing_ordinal: self.drawing_ordinal,
                picture_ordinal,
                anchor,
                anchor_range: pending.anchor_range,
                picture_range: pending.picture_range,
                blip_range: pending
                    .blip
                    .ok_or_else(|| invalid("direct picture has no a:blip range"))?,
                ext_list_range: pending.ext_list,
                c_nv_pr_id: pending.c_nv_pr_id,
                c_nv_pr_range: pending.c_nv_pr_range,
                raster_relationship_id,
                svg_owner,
                relationship_references: pending.relationship_references,
                namespace_context: pending.namespace_context,
                blip_namespace_context: pending
                    .blip_namespace_context
                    .ok_or_else(|| invalid("direct picture has no blip namespace context"))?,
                ext_list_namespace_context: pending.ext_list_namespace_context,
            });
        }
        let mut content_parts = Vec::new();
        content_parts
            .try_reserve_exact(self.content_parts.len())
            .map_err(|source| allocation("drawing source content-part projection", source))?;
        for (content_part_ordinal, pending) in self.content_parts.into_iter().enumerate() {
            let anchor = pending
                .anchor
                .ok_or_else(|| invalid("source content part has no complete anchor geometry"))?;
            content_parts.push(ContentPartSource {
                source: SourceLease::new(self.source),
                drawing_ordinal: self.drawing_ordinal,
                content_part_ordinal,
                profile: ContentPartProfile::CoreAnchor,
                anchor,
                anchor_range: pending.anchor_range,
                owner: pending.owner,
                relationship_id: pending.relationship_id,
                relationship_dialect: pending.relationship_dialect,
                mce_ancestor: pending.mce_ancestor,
                namespace_context: pending.namespace_context,
            });
        }
        Ok(SourceDrawing {
            source: SourceLease::new(self.source),
            drawing_ordinal: self.drawing_ordinal,
            dialect,
            relationship_dialect,
            pictures,
            content_parts,
            relationship_references: self.global_relationship_references,
        })
    }

    fn start_element(
        &mut self,
        element: &BytesStart<'_>,
        event_start: usize,
        event_end: usize,
        decoder: Decoder,
    ) -> Result<Frame> {
        let depth = self
            .frames
            .len()
            .checked_add(1)
            .ok_or_else(|| limit("drawing source XML depth", self.limits.max_depth))?;
        if depth > self.limits.max_depth || depth > MAX_XML_DEPTH {
            return Err(limit(
                "drawing source XML depth",
                self.limits.max_depth.min(MAX_XML_DEPTH),
            ));
        }
        if self.root_closed {
            return Err(invalid("drawing XML contains a second root element"));
        }
        let declarations = self.namespaces.preflight(element, decoder)?;
        let context_before = Arc::clone(&self.namespaces.context);
        self.namespaces
            .push(element, decoder, depth, declarations)?;
        let resolved = self.namespaces.resolve_element(element.name())?;
        validate_start_attributes(element, &self.namespaces, decoder)?;

        if !self.root_seen
            && self.frames.is_empty()
            && !is_name(
                resolved,
                element.name(),
                b"wsDr",
                SPREADSHEET_DRAWING,
                STRICT_SPREADSHEET_DRAWING,
            )
        {
            return Err(invalid("drawing XML root must be xdr:wsDr"));
        }

        let parent = self.frames.last().map(|frame| frame.kind);
        let mce_ancestor = self.frames.iter().any(|frame| frame.kind == Kind::Mce);
        let mut kind = classify(resolved, element.name(), parent);
        let core_content_part_name = is_spreadsheet(resolved, element.name(), b"contentPart");
        let direct_anchor = matches!(parent, Some(Kind::Anchor(_)));
        if parent == Some(Kind::ContentPart) {
            return Err(invalid(
                "core SpreadsheetDrawing contentPart must be childless",
            ));
        }
        if direct_anchor
            && element.name().local_name().as_ref() == b"contentPart"
            && kind != Kind::ContentPart
        {
            return Err(invalid(
                "direct anchor contentPart must use the core SpreadsheetDrawing QName",
            ));
        }
        if core_content_part_name && mce_ancestor {
            return Err(invalid(
                "direct core contentPart ownership under MCE is ambiguous",
            ));
        }
        if core_content_part_name && parent == Some(Kind::Group) {
            return Err(invalid(
                "core SpreadsheetDrawing contentPart has no group placement",
            ));
        }
        if core_content_part_name
            && !direct_anchor
            && !matches!(parent, Some(Kind::Unknown | Kind::OpaqueExt))
        {
            return Err(invalid(
                "core SpreadsheetDrawing contentPart must be a direct anchor object",
            ));
        }
        if kind == Kind::ContentPart {
            let expected = match self.dialect {
                Some(DrawingDialect::Strict) => STRICT_SPREADSHEET_DRAWING,
                Some(DrawingDialect::Transitional) => SPREADSHEET_DRAWING,
                None => return Err(invalid("contentPart appears before the drawing root")),
            };
            if resolved != Some(expected) {
                return Err(invalid(
                    "core contentPart uses a different SpreadsheetDrawing dialect than its root",
                ));
            }
        }
        if kind == Kind::Unknown && parent == Some(Kind::SvgExt) {
            // A recognized extension has a narrow grammar: one direct
            // `asvg:svgBlip` child, with comments/whitespace tolerated by the
            // XML layer.  Keep the raw extension intact but refuse to project
            // it when an unrecognized direct child appears beside the owner.
            if let Some(index) = self.active_picture {
                if let Some(extension_index) = self
                    .frames
                    .iter()
                    .rev()
                    .find(|frame| frame.kind == Kind::SvgExt)
                    .and_then(|frame| frame.extension_index)
                {
                    if let Some(extension) = self
                        .pictures
                        .get_mut(index)
                        .and_then(|picture| picture.extensions.get_mut(extension_index))
                    {
                        extension.malformed = true;
                    }
                }
            }
        }
        if kind == Kind::Unknown
            && parent == Some(Kind::ExtList)
            && is_name(
                resolved,
                element.name(),
                b"ext",
                DRAWINGML,
                STRICT_DRAWINGML,
            )
        {
            let (admitted, uri_lexical, malformed_attributes) = extension_uri(element, decoder)?;
            kind = if admitted {
                Kind::SvgExt
            } else {
                Kind::OpaqueExt
            };
            if let Some(picture_index) = self.active_picture {
                let pending = PendingExtension {
                    range: ElementRange {
                        range: ByteRange::new(event_start, event_end),
                        start_end: event_end,
                        close_start: None,
                        prefix: qname_prefix(element.name().as_ref())?.into_boxed_slice(),
                        namespace_uri: namespace_uri_copy(resolved)?,
                    },
                    admitted,
                    uri_lexical,
                    svg_blip: None,
                    svg_context: None,
                    svg_relationship_dialect: None,
                    svg_count: 0,
                    malformed: malformed_attributes,
                    mce_ancestor,
                };
                let picture = self
                    .pictures
                    .get_mut(picture_index)
                    .ok_or_else(|| invalid("active picture index is outside the source index"))?;
                picture
                    .extensions
                    .try_reserve(1)
                    .map_err(|source| allocation("drawing source extension candidates", source))?;
                picture.extensions.push(pending);
                let extension_index = picture.extensions.len() - 1;
                self.collect_relationship_attributes(element, event_start, event_end, decoder)?;
                return Ok(Frame {
                    kind,
                    namespace_declarations: declarations,
                    context_before,
                    source_start: event_start,
                    picture_index: Some(picture_index),
                    extension_index: Some(extension_index),
                });
            }
        }

        let mut picture_index = self.active_picture;
        let mut extension_index = None;
        match kind {
            Kind::Root => {
                if self.root_seen
                    || !is_name(
                        resolved,
                        element.name(),
                        b"wsDr",
                        SPREADSHEET_DRAWING,
                        STRICT_SPREADSHEET_DRAWING,
                    )
                {
                    return Err(invalid(
                        "drawing XML must contain exactly one xdr:wsDr root",
                    ));
                }
                self.root_seen = true;
                self.dialect = Some(if resolved == Some(STRICT_SPREADSHEET_DRAWING) {
                    DrawingDialect::Strict
                } else {
                    DrawingDialect::Transitional
                });
            },
            Kind::Anchor(anchor_kind) => {
                if self.anchor.is_some() {
                    return Err(invalid("nested spreadsheet drawing anchors"));
                }
                let edit_as = if anchor_kind == AnchorKind::TwoCell {
                    unqualified_attribute(element, b"editAs", decoder)?
                        .as_deref()
                        .map_or(Ok(EditAs::TwoCell), |value| {
                            let value = xsd_token_atom(value)
                                .ok_or_else(|| invalid("drawing editAs is not one token"))?;
                            value.parse().map_err(|_| invalid("invalid drawing editAs"))
                        })?
                } else {
                    EditAs::TwoCell
                };
                self.anchor = Some(AnchorCapture::new(
                    anchor_kind,
                    ByteRange::new(event_start, event_end),
                    edit_as,
                ));
            },
            Kind::From(target) | Kind::To(target) => {
                let anchor = self
                    .anchor
                    .as_mut()
                    .ok_or_else(|| invalid("drawing marker appears outside an anchor"))?;
                anchor.take_child(match target {
                    MarkerTarget::From => AnchorChild::From,
                    MarkerTarget::To => AnchorChild::To,
                })?;
                let marker = match target {
                    MarkerTarget::From => &mut anchor.from,
                    MarkerTarget::To => &mut anchor.to,
                };
                if marker.replace(Marker::default()).is_some() {
                    return Err(invalid("drawing anchor has duplicate from/to markers"));
                }
            },
            Kind::Marker(target, field) => {
                if !matches!(
                    self.frames.last().map(|frame| frame.kind),
                    Some(Kind::From(_) | Kind::To(_))
                ) {
                    return Err(invalid("drawing marker field appears outside from/to"));
                }
                let marker = match target {
                    MarkerTarget::From => {
                        self.anchor.as_ref().and_then(|anchor| anchor.from.as_ref())
                    },
                    MarkerTarget::To => self.anchor.as_ref().and_then(|anchor| anchor.to.as_ref()),
                }
                .ok_or_else(|| invalid("drawing marker field has no marker state"))?;
                if marker.next_field != field.ordinal() {
                    return Err(invalid("drawing marker fields are out of order"));
                }
                self.marker_text.clear();
            },
            Kind::Position => {
                self.anchor_mut()?.take_child(AnchorChild::Position)?;
                let position = EmuOffset {
                    x: Emu(coordinate_attribute(
                        element,
                        b"x",
                        decoder,
                        "drawing position x",
                    )?),
                    y: Emu(coordinate_attribute(
                        element,
                        b"y",
                        decoder,
                        "drawing position y",
                    )?),
                };
                if self.anchor_mut()?.position.replace(position).is_some() {
                    return Err(invalid("drawing anchor has duplicate positions"));
                }
            },
            Kind::Extent => {
                self.anchor_mut()?.take_child(AnchorChild::Extent)?;
                let extent = EmuExtent {
                    width: Emu(positive_coordinate_attribute(
                        element,
                        b"cx",
                        decoder,
                        "drawing extent width",
                    )?),
                    height: Emu(positive_coordinate_attribute(
                        element,
                        b"cy",
                        decoder,
                        "drawing extent height",
                    )?),
                };
                if self.anchor_mut()?.extent.replace(extent).is_some() {
                    return Err(invalid("drawing anchor has duplicate extents"));
                }
            },
            Kind::Picture => {
                self.anchor_mut()?.take_child(AnchorChild::Object)?;
                if self
                    .anchor_mut()?
                    .object_kind
                    .replace(ObjectKind::Picture)
                    .is_some()
                {
                    return Err(invalid("drawing anchor has duplicate objects"));
                }
                if self.pictures.len() >= self.limits.max_pictures {
                    return Err(limit(
                        "drawing source direct pictures",
                        self.limits.max_pictures,
                    ));
                }
                self.pictures
                    .try_reserve(1)
                    .map_err(|source| allocation("drawing source direct picture index", source))?;
                let context = Arc::clone(&self.namespaces.context);
                self.pictures.push(PendingPicture {
                    picture_range: ByteRange::new(event_start, event_end),
                    anchor_range: self
                        .anchor
                        .as_ref()
                        .map_or(ByteRange::new(event_start, event_end), |anchor| {
                            anchor.range
                        }),
                    anchor: None,
                    blip: None,
                    ext_list: None,
                    c_nv_pr_id: None,
                    c_nv_pr_range: None,
                    raster_relationship_id: None,
                    extensions: Vec::new(),
                    relationship_references: Vec::new(),
                    namespace_context: context,
                    blip_namespace_context: None,
                    ext_list_namespace_context: None,
                    mce_ancestor,
                    grouped_or_foreign_descendant: false,
                });
                let index = self.pictures.len() - 1;
                self.active_picture = Some(index);
                picture_index = Some(index);
                self.anchor_mut()?.picture_index = Some(index);
            },
            Kind::ContentPart => {
                let owner_namespace = namespace_uri_copy(resolved)?;
                self.anchor_mut()?.take_child(AnchorChild::Object)?;
                if self
                    .anchor_mut()?
                    .object_kind
                    .replace(ObjectKind::ContentPart)
                    .is_some()
                {
                    return Err(invalid("drawing anchor has duplicate objects"));
                }
                if self.content_parts.len() >= self.limits.max_content_parts {
                    return Err(limit(
                        "drawing source content parts",
                        self.limits.max_content_parts,
                    ));
                }
                self.content_parts.try_reserve(1).map_err(|source| {
                    allocation("drawing source direct content-part index", source)
                })?;
                let (relationship_id, relationship_dialect) =
                    required_core_content_part_relationship(
                        element,
                        &self.namespaces,
                        decoder,
                        self.dialect,
                    )?;
                let owner = ElementRange {
                    range: ByteRange::new(event_start, event_end),
                    start_end: event_end,
                    close_start: None,
                    prefix: qname_prefix(element.name().as_ref())?.into_boxed_slice(),
                    namespace_uri: owner_namespace,
                };
                let anchor_range = self
                    .anchor
                    .as_ref()
                    .map_or(ByteRange::new(event_start, event_end), |anchor| {
                        anchor.range
                    });
                let pending = PendingContentPart {
                    owner,
                    anchor_range,
                    anchor: None,
                    relationship_id: relationship_id.into_boxed_str(),
                    relationship_dialect,
                    namespace_context: Arc::clone(&self.namespaces.context),
                    mce_ancestor,
                };
                self.content_parts.push(pending);
                let index = self.content_parts.len() - 1;
                self.anchor_mut()?.content_part_index = Some(index);
                picture_index = None;
            },
            Kind::Group => {
                if parent == Some(Kind::Anchor(AnchorKind::TwoCell))
                    || parent == Some(Kind::Anchor(AnchorKind::OneCell))
                    || parent == Some(Kind::Anchor(AnchorKind::Absolute))
                {
                    self.anchor_mut()?.take_child(AnchorChild::Object)?;
                    if self
                        .anchor_mut()?
                        .object_kind
                        .replace(ObjectKind::Group)
                        .is_some()
                    {
                        return Err(invalid("drawing anchor has duplicate objects"));
                    }
                }
                if let Some(index) = self.active_picture {
                    if let Some(picture) = self.pictures.get_mut(index) {
                        picture.grouped_or_foreign_descendant = true;
                    }
                }
                picture_index = None;
            },
            Kind::Blip => {
                if parent != Some(Kind::BlipFill) {
                    return Err(invalid("a:blip is not a direct blipFill child"));
                }
                let index = self
                    .active_picture
                    .ok_or_else(|| invalid("a:blip is outside a direct picture"))?;
                let relation =
                    relationship_attribute(element, &self.namespaces, b"embed", decoder)?
                        .ok_or_else(|| invalid("raster a:blip is missing r:embed"))?;
                if relationship_attribute(element, &self.namespaces, b"link", decoder)?.is_some() {
                    return Err(invalid(
                        "raster a:blip cannot carry both r:embed and r:link",
                    ));
                }
                self.relationship_dialect.get_or_insert(relation.1);
                let picture = self
                    .pictures
                    .get_mut(index)
                    .ok_or_else(|| invalid("active picture index is outside the source index"))?;
                if picture
                    .raster_relationship_id
                    .replace(relation.0.into_boxed_str())
                    .is_some()
                {
                    return Err(invalid(
                        "direct picture has duplicate raster a:blip elements",
                    ));
                }
                picture.blip = Some(ElementRange {
                    range: ByteRange::new(event_start, event_end),
                    start_end: event_end,
                    close_start: None,
                    prefix: qname_prefix(element.name().as_ref())?.into_boxed_slice(),
                    namespace_uri: namespace_uri_copy(resolved)?,
                });
                picture.blip_namespace_context = Some(Arc::clone(&self.namespaces.context));
            },
            Kind::ExtList => {
                let index = self
                    .active_picture
                    .ok_or_else(|| invalid("a:extLst is outside a direct picture"))?;
                let picture = self
                    .pictures
                    .get_mut(index)
                    .ok_or_else(|| invalid("active picture index is outside the source index"))?;
                if picture.ext_list.is_some() {
                    return Err(invalid("direct picture has duplicate a:extLst elements"));
                }
                picture.ext_list = Some(ElementRange {
                    range: ByteRange::new(event_start, event_end),
                    start_end: event_end,
                    close_start: None,
                    prefix: qname_prefix(element.name().as_ref())?.into_boxed_slice(),
                    namespace_uri: namespace_uri_copy(resolved)?,
                });
                picture.ext_list_namespace_context = Some(Arc::clone(&self.namespaces.context));
            },
            Kind::ClientData => {
                self.anchor_mut()?.take_child(AnchorChild::ClientData)?;
                if self.anchor_mut()?.client_data_seen {
                    return Err(invalid("drawing anchor has duplicate clientData"));
                }
                self.anchor_mut()?.client_data_seen = true;
            },
            Kind::SvgBlip => {
                let index = self
                    .active_picture
                    .ok_or_else(|| invalid("svgBlip is outside a direct picture"))?;
                let ext_frame = self
                    .frames
                    .iter()
                    .rev()
                    .find(|frame| matches!(frame.kind, Kind::SvgExt));
                let ext_index = ext_frame
                    .and_then(|frame| frame.extension_index)
                    .ok_or_else(|| invalid("svgBlip has no direct admitted extension owner"))?;
                let embedded = relationship_dialect_attribute(element, &self.namespaces, b"embed")?;
                let linked = relationship_dialect_attribute(element, &self.namespaces, b"link")?;
                let pending = self
                    .pictures
                    .get_mut(index)
                    .and_then(|picture| picture.extensions.get_mut(ext_index))
                    .ok_or_else(|| invalid("SVG extension index is outside the source index"))?;
                pending.svg_count = pending
                    .svg_count
                    .checked_add(1)
                    .ok_or_else(|| limit("drawing source SVG owner count", 2))?;
                if pending.svg_count == 1 {
                    pending.svg_blip = Some(ByteRange::new(event_start, event_end));
                    pending.svg_context = Some(Arc::clone(&context_before));
                    pending.svg_relationship_dialect = embedded.or(linked);
                } else {
                    pending.malformed = true;
                }
            },
            Kind::CnvPr => {
                let index = self
                    .active_picture
                    .ok_or_else(|| invalid("cNvPr is outside a direct picture"))?;
                let picture = self
                    .pictures
                    .get_mut(index)
                    .ok_or_else(|| invalid("active picture index is outside the source index"))?;
                if picture.c_nv_pr_range.is_some() {
                    return Err(invalid("direct picture has duplicate cNvPr elements"));
                }
                picture.c_nv_pr_range = Some(ByteRange::new(event_start, event_end));
                if let Some(value) = bounded_unqualified_attribute(
                    element,
                    b"id",
                    decoder,
                    MAX_NAME_BYTES,
                    "cNvPr id",
                )? {
                    if value.is_empty() {
                        return Err(invalid("cNvPr id is empty"));
                    }
                    picture.c_nv_pr_id = Some(value.into_boxed_str());
                }
            },
            Kind::SvgExt | Kind::OpaqueExt => {
                // The extension candidate was inserted in the early branch
                // above.  Keep its index on the stack for source-range close
                // updates and direct svgBlip ownership.
                extension_index = self
                    .frames
                    .iter()
                    .rev()
                    .find(|frame| matches!(frame.kind, Kind::ExtList))
                    .and(self.active_picture)
                    .and_then(|index| self.pictures.get(index))
                    .and_then(|picture| picture.extensions.len().checked_sub(1));
            },
            Kind::Mce => {
                if self.active_picture.is_some() {
                    if let Some(index) = self.active_picture {
                        if let Some(picture) = self.pictures.get_mut(index) {
                            picture.mce_ancestor = true;
                        }
                    }
                }
            },
            Kind::PictureNvPr | Kind::BlipFill | Kind::Unknown => {},
        }
        self.collect_relationship_attributes(element, event_start, event_end, decoder)?;
        Ok(Frame {
            kind,
            namespace_declarations: declarations,
            context_before,
            source_start: event_start,
            picture_index,
            extension_index,
        })
    }

    fn finish_element(
        &mut self,
        frame: Frame,
        event_start: usize,
        event_end: usize,
        empty: bool,
    ) -> Result<()> {
        match frame.kind {
            Kind::Marker(target, field) => {
                let value = trim_xml_schema_whitespace(&self.marker_text);
                let parsed = value
                    .parse::<i64>()
                    .map_err(|_| invalid(format!("invalid drawing marker value '{value}'")))?;
                let marker = match target {
                    MarkerTarget::From => {
                        self.anchor.as_mut().and_then(|anchor| anchor.from.as_mut())
                    },
                    MarkerTarget::To => self.anchor.as_mut().and_then(|anchor| anchor.to.as_mut()),
                }
                .ok_or_else(|| invalid("drawing marker value has no marker state"))?;
                marker.set(field, parsed, "drawing marker")?;
            },
            Kind::From(target) | Kind::To(target) => {
                let marker = match target {
                    MarkerTarget::From => {
                        self.anchor.as_ref().and_then(|anchor| anchor.from.as_ref())
                    },
                    MarkerTarget::To => self.anchor.as_ref().and_then(|anchor| anchor.to.as_ref()),
                }
                .ok_or_else(|| invalid("drawing marker container has no marker state"))?;
                marker.finish(match target {
                    MarkerTarget::From => "drawing from marker",
                    MarkerTarget::To => "drawing to marker",
                })?;
            },
            Kind::Anchor(_) => {
                let anchor = self
                    .anchor
                    .take()
                    .ok_or_else(|| invalid("drawing anchor close has no pending anchor"))?;
                let anchor_value = anchor.finish()?;
                let picture_index = anchor.picture_index;
                let content_part_index = anchor.content_part_index;
                if let Some(index) = picture_index {
                    let picture = self
                        .pictures
                        .get_mut(index)
                        .ok_or_else(|| invalid("anchor picture index is outside source index"))?;
                    picture.anchor = Some(anchor_value);
                    picture.anchor_range = ByteRange::new(anchor.range.start, event_end);
                    if picture.raster_relationship_id.is_none() {
                        return Err(invalid("direct picture has no raster r:embed fallback"));
                    }
                } else if let Some(index) = content_part_index {
                    let content_part = self
                        .content_parts
                        .get_mut(index)
                        .ok_or_else(|| invalid("anchor content-part index is outside source"))?;
                    content_part.anchor = Some(anchor_value);
                    content_part.anchor_range = ByteRange::new(anchor.range.start, event_end);
                }
            },
            Kind::Picture => {
                if let Some(index) = frame.picture_index {
                    let picture = self
                        .pictures
                        .get_mut(index)
                        .ok_or_else(|| invalid("picture index is outside source index"))?;
                    picture.picture_range.end = event_end;
                    self.active_picture = None;
                }
            },
            Kind::Blip => {
                if let Some(index) = frame.picture_index {
                    let picture = self
                        .pictures
                        .get_mut(index)
                        .ok_or_else(|| invalid("blip picture index is outside source index"))?;
                    let range = picture
                        .blip
                        .as_mut()
                        .ok_or_else(|| invalid("direct a:blip close has no opening range"))?;
                    if range.range.start == frame.source_start && !empty {
                        range.range.end = event_end;
                        range.close_start = Some(event_start);
                    }
                }
            },
            Kind::ExtList => {
                if let Some(index) = frame.picture_index {
                    let picture = self
                        .pictures
                        .get_mut(index)
                        .ok_or_else(|| invalid("extLst picture index is outside source index"))?;
                    let range = picture
                        .ext_list
                        .as_mut()
                        .ok_or_else(|| invalid("direct extLst close has no opening range"))?;
                    if range.range.start == frame.source_start && !empty {
                        range.range.end = event_end;
                        range.close_start = Some(event_start);
                    }
                }
            },
            Kind::SvgExt | Kind::OpaqueExt => {
                if let (Some(index), Some(extension_index)) =
                    (frame.picture_index, frame.extension_index)
                {
                    let extension = self
                        .pictures
                        .get_mut(index)
                        .and_then(|picture| picture.extensions.get_mut(extension_index))
                        .ok_or_else(|| invalid("extension index is outside source index"))?;
                    if extension.range.range.start == frame.source_start && !empty {
                        extension.range.range.end = event_end;
                        extension.range.close_start = Some(event_start);
                    }
                }
            },
            Kind::SvgBlip => {
                if let Some(index) = frame.picture_index {
                    let extension = self.pictures.get_mut(index).and_then(|picture| {
                        picture.extensions.iter_mut().rev().find(|extension| {
                            extension
                                .svg_blip
                                .is_some_and(|range| range.start == frame.source_start)
                        })
                    });
                    if let Some(extension) = extension {
                        if let Some(range) = extension.svg_blip.as_mut() {
                            range.end = event_end;
                        }
                    }
                }
            },
            Kind::CnvPr => {
                if let Some(index) = frame.picture_index {
                    if let Some(range) = self
                        .pictures
                        .get_mut(index)
                        .and_then(|picture| picture.c_nv_pr_range.as_mut())
                    {
                        range.end = event_end;
                    }
                }
            },
            Kind::ContentPart => {
                let index = self
                    .content_parts
                    .iter()
                    .position(|content_part| content_part.owner.range.start == frame.source_start)
                    .ok_or_else(|| invalid("contentPart has no pending source owner"))?;
                let pending = self
                    .content_parts
                    .get_mut(index)
                    .ok_or_else(|| invalid("contentPart source index is outside the scan"))?;
                if !empty {
                    pending.owner.range.end = event_end;
                    pending.owner.close_start = Some(event_start);
                }
            },
            Kind::Root
            | Kind::Position
            | Kind::Extent
            | Kind::Group
            | Kind::Mce
            | Kind::ClientData
            | Kind::Unknown => {},
            Kind::PictureNvPr | Kind::BlipFill => {},
        }
        Ok(())
    }

    fn collect_relationship_attributes(
        &mut self,
        element: &BytesStart<'_>,
        range_start: usize,
        range_end: usize,
        decoder: Decoder,
    ) -> Result<()> {
        let mut found = Vec::<RelationshipReference>::new();
        for attribute in element.attributes().with_checks(true) {
            let attribute = attribute.map_err(xml_error)?;
            if attribute.key.as_namespace_binding().is_some() {
                continue;
            }
            let namespace = self.namespaces.resolve_attribute(attribute.key)?;
            let Some(namespace) = namespace else { continue };
            let dialect = if namespace == STRICT_RELATIONSHIPS {
                RelationshipDialect::Strict
            } else if namespace == RELATIONSHIPS {
                RelationshipDialect::Transitional
            } else {
                continue;
            };
            if attribute.value.len() > MAX_RELATIONSHIP_ID_BYTES.saturating_mul(4) {
                return Err(limit(
                    "relationship ID lexical bytes",
                    MAX_RELATIONSHIP_ID_BYTES.saturating_mul(4),
                ));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?;
            if value.len() > MAX_RELATIONSHIP_ID_BYTES
                || value.is_empty()
                || !litchi_ooxml_common::xml_name::is_ncname(&value)
            {
                return Err(invalid("relationship attribute has an invalid ID"));
            }
            let local = attribute.key.local_name();
            let local = std::str::from_utf8(local.as_ref())
                .map_err(|error| Error::Invalid(error.to_string()))?;
            let mut id = String::new();
            id.try_reserve(value.len())
                .map_err(|source| allocation("drawing source relationship ID", source))?;
            id.push_str(&value);
            let mut local_copy = String::new();
            local_copy
                .try_reserve(local.len())
                .map_err(|source| allocation("drawing source relationship name", source))?;
            local_copy.push_str(local);
            found
                .try_reserve(1)
                .map_err(|source| allocation("drawing source relationship references", source))?;
            found.push(RelationshipReference {
                id: id.into_boxed_str(),
                local_name: local_copy.into_boxed_str(),
                range: ByteRange::new(range_start, range_end),
                dialect,
            });
        }
        if found.is_empty() {
            return Ok(());
        }
        let max_relationship_owners = self
            .limits
            .max_pictures
            .checked_add(self.limits.max_content_parts)
            .and_then(|count| count.checked_add(1))
            .ok_or_else(|| limit("drawing source relationship references", usize::MAX))?;
        let global_limit = self
            .limits
            .max_relationship_references
            .checked_mul(max_relationship_owners)
            .ok_or_else(|| limit("drawing source relationship references", usize::MAX))?;
        let global_len = self
            .global_relationship_references
            .len()
            .checked_add(found.len())
            .ok_or_else(|| limit("drawing source relationship references", global_limit))?;
        if global_len > global_limit {
            return Err(limit(
                "drawing source relationship references",
                global_limit,
            ));
        }
        let picture_new_len = if let Some(index) = self.active_picture {
            let current = self
                .pictures
                .get(index)
                .ok_or_else(|| invalid("active picture index is outside the source index"))?
                .relationship_references
                .len();
            Some(current.checked_add(found.len()).ok_or_else(|| {
                limit(
                    "picture relationship references",
                    self.limits.max_relationship_references,
                )
            })?)
        } else {
            None
        };
        if picture_new_len.is_some_and(|length| length > self.limits.max_relationship_references) {
            return Err(limit(
                "picture relationship references",
                self.limits.max_relationship_references,
            ));
        }
        self.global_relationship_references
            .try_reserve(found.len())
            .map_err(|source| allocation("drawing source relationship index", source))?;
        if let Some(index) = self.active_picture {
            let picture = self
                .pictures
                .get_mut(index)
                .ok_or_else(|| invalid("active picture index is outside source index"))?;
            picture
                .relationship_references
                .try_reserve(found.len())
                .map_err(|source| allocation("picture relationship references", source))?;
        }
        for value in &found {
            self.global_relationship_references
                .push(value.try_clone_for_source()?);
        }
        if let Some(index) = self.active_picture {
            let picture = self
                .pictures
                .get_mut(index)
                .ok_or_else(|| invalid("active picture index is outside source index"))?;
            for value in found {
                picture
                    .relationship_references
                    .push(value.try_clone_for_source()?);
            }
        }
        Ok(())
    }

    fn anchor_mut(&mut self) -> Result<&mut AnchorCapture> {
        self.anchor
            .as_mut()
            .ok_or_else(|| invalid("drawing anchor child appears outside an anchor"))
    }

    fn append_marker_text(&mut self, bytes: &[u8]) -> Result<()> {
        let text = std::str::from_utf8(bytes).map_err(|error| Error::Invalid(error.to_string()))?;
        let length = self
            .marker_text
            .len()
            .checked_add(text.len())
            .ok_or_else(|| limit("drawing marker text", MAX_MARKER_TEXT_BYTES))?;
        if length > MAX_MARKER_TEXT_BYTES {
            return Err(limit("drawing marker text", MAX_MARKER_TEXT_BYTES));
        }
        self.marker_text
            .try_reserve(text.len())
            .map_err(|source| allocation("drawing source marker text", source))?;
        self.marker_text.push_str(text);
        Ok(())
    }

    fn validate_text_context(&self, outside: bool, non_whitespace: bool) -> Result<()> {
        if outside || !non_whitespace {
            return Ok(());
        }
        let Some(frame) = self.frames.last() else {
            return Ok(());
        };
        if matches!(
            frame.kind,
            Kind::Root
                | Kind::Anchor(_)
                | Kind::From(_)
                | Kind::To(_)
                | Kind::Marker(..)
                | Kind::Position
                | Kind::Extent
                | Kind::Picture
                | Kind::PictureNvPr
                | Kind::CnvPr
                | Kind::BlipFill
                | Kind::Blip
                | Kind::ExtList
                | Kind::ContentPart
                | Kind::SvgExt
        ) {
            return Err(invalid(
                "known drawing container contains non-whitespace text",
            ));
        }
        Ok(())
    }
}

fn classify(namespace: Option<&[u8]>, name: QName<'_>, parent: Option<Kind>) -> Kind {
    if namespace.is_some_and(|value| value == MCE) {
        return Kind::Mce;
    }
    if parent.is_none()
        && is_name(
            namespace,
            name,
            b"wsDr",
            SPREADSHEET_DRAWING,
            STRICT_SPREADSHEET_DRAWING,
        )
    {
        return Kind::Root;
    }
    if parent == Some(Kind::Root) && is_spreadsheet_anchor(namespace, name) {
        return match name.local_name().as_ref() {
            b"twoCellAnchor" => Kind::Anchor(AnchorKind::TwoCell),
            b"oneCellAnchor" => Kind::Anchor(AnchorKind::OneCell),
            b"absoluteAnchor" => Kind::Anchor(AnchorKind::Absolute),
            _ => Kind::Unknown,
        };
    }
    if is_spreadsheet(namespace, name, b"from") && matches!(parent, Some(Kind::Anchor(_))) {
        return Kind::From(MarkerTarget::From);
    }
    if is_spreadsheet(namespace, name, b"to") && parent == Some(Kind::Anchor(AnchorKind::TwoCell)) {
        return Kind::To(MarkerTarget::To);
    }
    if matches!(parent, Some(Kind::From(_) | Kind::To(_)))
        && is_spreadsheet(namespace, name, b"col")
    {
        return Kind::Marker(parent_marker_target(parent), MarkerField::Column);
    }
    if matches!(parent, Some(Kind::From(_) | Kind::To(_)))
        && is_spreadsheet(namespace, name, b"colOff")
    {
        return Kind::Marker(parent_marker_target(parent), MarkerField::ColumnOffset);
    }
    if matches!(parent, Some(Kind::From(_) | Kind::To(_)))
        && is_spreadsheet(namespace, name, b"row")
    {
        return Kind::Marker(parent_marker_target(parent), MarkerField::Row);
    }
    if matches!(parent, Some(Kind::From(_) | Kind::To(_)))
        && is_spreadsheet(namespace, name, b"rowOff")
    {
        return Kind::Marker(parent_marker_target(parent), MarkerField::RowOffset);
    }
    if parent == Some(Kind::Anchor(AnchorKind::Absolute)) && is_spreadsheet(namespace, name, b"pos")
    {
        return Kind::Position;
    }
    if matches!(
        parent,
        Some(Kind::Anchor(AnchorKind::OneCell | AnchorKind::Absolute))
    ) && is_spreadsheet(namespace, name, b"ext")
    {
        return Kind::Extent;
    }
    if matches!(parent, Some(Kind::Anchor(_))) && is_spreadsheet(namespace, name, b"clientData") {
        return Kind::ClientData;
    }
    if matches!(parent, Some(Kind::Anchor(_))) && is_spreadsheet(namespace, name, b"pic") {
        return Kind::Picture;
    }
    if matches!(parent, Some(Kind::Anchor(_))) && is_spreadsheet(namespace, name, b"contentPart") {
        return Kind::ContentPart;
    }
    if matches!(parent, Some(Kind::Anchor(_) | Kind::Group))
        && is_spreadsheet(namespace, name, b"grpSp")
    {
        return Kind::Group;
    }
    if parent == Some(Kind::Picture) && is_spreadsheet(namespace, name, b"nvPicPr") {
        return Kind::PictureNvPr;
    }
    if parent == Some(Kind::Picture) && is_spreadsheet(namespace, name, b"blipFill") {
        return Kind::BlipFill;
    }
    if parent == Some(Kind::PictureNvPr) && is_spreadsheet(namespace, name, b"cNvPr") {
        return Kind::CnvPr;
    }
    if parent == Some(Kind::BlipFill)
        && is_name(namespace, name, b"blip", DRAWINGML, STRICT_DRAWINGML)
    {
        return Kind::Blip;
    }
    if parent == Some(Kind::Blip)
        && is_name(namespace, name, b"extLst", DRAWINGML, STRICT_DRAWINGML)
    {
        return Kind::ExtList;
    }
    if parent == Some(Kind::SvgExt)
        && is_name(namespace, name, b"svgBlip", SVG_NAMESPACE, SVG_NAMESPACE)
    {
        return Kind::SvgBlip;
    }
    Kind::Unknown
}

fn parent_marker_target(parent: Option<Kind>) -> MarkerTarget {
    match parent {
        Some(Kind::From(target) | Kind::To(target)) => target,
        _ => MarkerTarget::From,
    }
}

fn is_spreadsheet_anchor(namespace: Option<&[u8]>, name: QName<'_>) -> bool {
    matches!(
        name.local_name().as_ref(),
        b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor"
    ) && is_namespace(namespace, SPREADSHEET_DRAWING, STRICT_SPREADSHEET_DRAWING)
}

fn is_spreadsheet(namespace: Option<&[u8]>, name: QName<'_>, local: &[u8]) -> bool {
    is_name(
        namespace,
        name,
        local,
        SPREADSHEET_DRAWING,
        STRICT_SPREADSHEET_DRAWING,
    )
}

fn is_name(
    namespace: Option<&[u8]>,
    name: QName<'_>,
    local: &[u8],
    expected: &[u8],
    strict: &[u8],
) -> bool {
    name.local_name().as_ref() == local
        && namespace.is_some_and(|value| value == expected || value == strict)
}

fn is_namespace(namespace: Option<&[u8]>, expected: &[u8], strict: &[u8]) -> bool {
    namespace.is_some_and(|value| value == expected || value == strict)
}

fn project_svg_owner<'a>(
    source: &'a [u8],
    picture: &PendingPicture,
    host_relationship_dialect: RelationshipDialect,
    max_fragment_bytes: usize,
) -> Result<SvgOwnerState<'a>> {
    if picture.mce_ancestor {
        return Ok(SvgOwnerState::Refused);
    }
    let mut admitted = Vec::<&PendingExtension>::new();
    let mut opaque = false;
    for extension in &picture.extensions {
        if extension.admitted {
            admitted
                .try_reserve(1)
                .map_err(|source| allocation("drawing source admitted SVG owners", source))?;
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
    if extension.mce_ancestor || extension.malformed || picture.mce_ancestor {
        return Ok(SvgOwnerState::Refused);
    }
    let Some(svg_range) = extension.svg_blip else {
        return Ok(SvgOwnerState::Refused);
    };
    if extension.svg_count != 1 {
        return Ok(SvgOwnerState::Ambiguous);
    }
    let context = extension
        .svg_context
        .as_ref()
        .ok_or_else(|| invalid("admitted SVG owner has no namespace context"))?;
    let fragment = svg_range.slice(source)?;
    let maximum = max_fragment_bytes.min(MAX_FRAGMENT_BYTES);
    if fragment.len() > maximum {
        return Err(limit("drawing source SVG fragment bytes", maximum));
    }
    // Keep the host range authoritative and resolve inherited names through
    // the persistent context.  Namespace completion remains an explicit,
    // bounded export operation on the owner.
    let value = match svg_blip::read_contextual(fragment, context) {
        Ok(value) => value,
        Err(_) => return Ok(SvgOwnerState::Refused),
    };
    let (embedded, linked) = (value.embedded().is_some(), value.linked().is_some());
    if embedded == linked {
        return Ok(SvgOwnerState::Refused);
    }
    let relationship_dialect = extension
        .svg_relationship_dialect
        .unwrap_or(host_relationship_dialect);
    let mut uri_lexical = Vec::new();
    uri_lexical
        .try_reserve_exact(extension.uri_lexical.len())
        .map_err(|source| allocation("drawing source SVG URI lexical bytes", source))?;
    uri_lexical.extend_from_slice(&extension.uri_lexical);
    let owner = SvgOwner {
        source: SourceLease::new(source),
        extension: extension.range.clone(),
        svg_blip: svg_range,
        uri_lexical: uri_lexical.into_boxed_slice(),
        value,
        relationship_dialect,
        namespace_context: Arc::clone(context),
    };
    Ok(if embedded {
        SvgOwnerState::Embedded(owner)
    } else {
        SvgOwnerState::Linked(owner)
    })
}

/// Explicit finite source scan limits.  Every allocation in the scanner is
/// charged to one of these bounds before it is attempted.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[must_use]
pub struct ScanLimits {
    /// Maximum source XML bytes.
    pub max_xml_bytes: usize,
    /// Maximum parsed XML nodes/events.
    pub max_nodes: usize,
    /// Maximum XML depth.
    pub max_depth: usize,
    /// Maximum direct picture candidates.
    pub max_pictures: usize,
    /// Maximum direct core content-part candidates.
    pub max_content_parts: usize,
    /// Maximum relationship references per picture.
    pub max_relationship_references: usize,
    /// Maximum raw SVG owner bytes accepted during contextual projection.
    /// Lazy standalone namespace completion uses its own explicit output cap.
    pub max_fragment_bytes: usize,
}

impl Default for ScanLimits {
    fn default() -> Self {
        Self {
            max_xml_bytes: MAX_XML_BYTES,
            max_nodes: MAX_XML_NODES,
            max_depth: MAX_XML_DEPTH,
            max_pictures: MAX_PICTURES,
            max_content_parts: MAX_CONTENT_PARTS,
            max_relationship_references: MAX_RELATIONSHIP_REFERENCES,
            max_fragment_bytes: MAX_FRAGMENT_BYTES,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Binding {
    prefix: Vec<u8>,
    uri: Vec<u8>,
    depth: usize,
}

/// Active namespace bindings.  Bindings are pushed once and popped once; a
/// candidate stores an `Arc` chain instead of cloning the full active scope on
/// every XML event.
struct Namespaces {
    bindings: Vec<Arc<Binding>>,
    by_prefix: HashMap<Vec<u8>, Vec<Arc<Binding>>>,
    active: usize,
    context: Arc<NamespaceContext>,
}

impl Default for Namespaces {
    fn default() -> Self {
        Self {
            bindings: Vec::new(),
            by_prefix: HashMap::new(),
            active: 0,
            context: Arc::new(NamespaceContext::empty()),
        }
    }
}

impl Namespaces {
    fn preflight(&self, element: &BytesStart<'_>, decoder: Decoder) -> Result<usize> {
        let mut attributes = 0usize;
        let mut declarations = 0usize;
        for attribute in element.attributes().with_checks(true) {
            let attribute = attribute.map_err(xml_error)?;
            attributes = attributes
                .checked_add(1)
                .ok_or_else(|| limit("XML attributes", MAX_ATTRIBUTES))?;
            if attributes > MAX_ATTRIBUTES {
                return Err(limit("XML attributes", MAX_ATTRIBUTES));
            }
            validate_qname(attribute.key.as_ref(), "attribute")?;
            validate_raw_attribute(attribute.value.as_ref())?;
            if let Some(prefix) = attribute.key.as_namespace_binding() {
                declarations = declarations
                    .checked_add(1)
                    .ok_or_else(|| limit("namespace declarations", MAX_NAMESPACE_DECLARATIONS))?;
                if declarations > MAX_NAMESPACE_DECLARATIONS {
                    return Err(limit("namespace declarations", MAX_NAMESPACE_DECLARATIONS));
                }
                if attribute.value.len() > MAX_NAMESPACE_BYTES {
                    return Err(limit("namespace URI lexical bytes", MAX_NAMESPACE_BYTES));
                }
                let uri = attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(xml_error)?;
                validate_xml_characters(uri.as_bytes())?;
                validate_namespace_binding(prefix, uri.as_bytes())?;
            }
        }
        let active = self
            .active
            .checked_add(declarations)
            .ok_or_else(|| limit("active namespace bindings", MAX_ACTIVE_NAMESPACE_BINDINGS))?;
        if active > MAX_ACTIVE_NAMESPACE_BINDINGS {
            return Err(limit(
                "active namespace bindings",
                MAX_ACTIVE_NAMESPACE_BINDINGS,
            ));
        }
        Ok(declarations)
    }

    fn push(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        depth: usize,
        declarations: usize,
    ) -> Result<()> {
        if declarations == 0 {
            return Ok(());
        }
        self.bindings
            .try_reserve(declarations)
            .map_err(|source| allocation("drawing source namespace bindings", source))?;
        self.by_prefix
            .try_reserve(declarations)
            .map_err(|source| allocation("drawing source namespace prefix index", source))?;
        let mut shared_declarations = Vec::<Namespace>::new();
        shared_declarations
            .try_reserve_exact(declarations)
            .map_err(|source| allocation("drawing source shared namespace context", source))?;
        for attribute in element.attributes().with_checks(true) {
            let attribute = attribute.map_err(xml_error)?;
            let Some(prefix) = attribute.key.as_namespace_binding() else {
                continue;
            };
            let uri = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?;
            let prefix = match prefix {
                quick_xml::name::PrefixDeclaration::Default => &[][..],
                quick_xml::name::PrefixDeclaration::Named(prefix) => prefix,
            };
            let prefix_text = if prefix.is_empty() {
                None
            } else {
                Some(
                    std::str::from_utf8(prefix)
                        .map_err(|error| Error::Invalid(error.to_string()))?,
                )
            };
            let shared = Namespace::new(prefix_text, uri.as_ref())
                .map_err(|error| invalid(error.to_string()))?;
            shared_declarations.push(shared);
            let mut prefix_copy = Vec::new();
            prefix_copy
                .try_reserve_exact(prefix.len())
                .map_err(|source| allocation("drawing source namespace prefix", source))?;
            prefix_copy.extend_from_slice(prefix);
            let mut uri_copy = Vec::new();
            uri_copy
                .try_reserve_exact(uri.len())
                .map_err(|source| allocation("drawing source namespace URI", source))?;
            uri_copy.extend_from_slice(uri.as_bytes());
            let binding = Arc::new(Binding {
                prefix: prefix_copy,
                uri: uri_copy,
                depth,
            });
            self.bindings.push(Arc::clone(&binding));
            let mut prefix_key = Vec::new();
            prefix_key
                .try_reserve_exact(binding.prefix.len())
                .map_err(|source| allocation("drawing source namespace prefix key", source))?;
            prefix_key.extend_from_slice(&binding.prefix);
            let stack = self.by_prefix.entry(prefix_key).or_default();
            stack
                .try_reserve(1)
                .map_err(|source| allocation("drawing source namespace prefix stack", source))?;
            stack.push(Arc::clone(&binding));
        }
        // Publish only this element's declarations.  Descendants retain this
        // immutable node through `Arc` instead of flattening the active scope.
        let context = self.context.child(shared_declarations)?;
        self.context = Arc::new(context);
        self.active = self
            .active
            .checked_add(declarations)
            .ok_or_else(|| limit("active namespace bindings", MAX_ACTIVE_NAMESPACE_BINDINGS))?;
        Ok(())
    }

    fn pop(&mut self, depth: usize, declarations: usize) {
        if declarations == 0 {
            return;
        }
        let mut removed = 0usize;
        while removed < declarations {
            let Some(binding) = self.bindings.pop() else {
                break;
            };
            debug_assert_eq!(binding.depth, depth);
            let empty = if let Some(stack) = self.by_prefix.get_mut(&binding.prefix) {
                let _ = stack.pop();
                stack.is_empty()
            } else {
                false
            };
            if empty {
                self.by_prefix.remove(&binding.prefix);
            }
            removed += 1;
        }
        self.active = self.active.saturating_sub(removed);
    }

    fn resolve_element(&self, name: QName<'_>) -> Result<Option<&[u8]>> {
        let (_, prefix) = name.decompose();
        match prefix {
            Some(prefix) => self.resolve_prefix(prefix.as_ref(), true),
            None => Ok(self
                .by_prefix
                .get(&[][..])
                .and_then(|stack| stack.last())
                .and_then(|binding| (!binding.uri.is_empty()).then_some(binding.uri.as_slice()))),
        }
    }

    fn resolve_attribute(&self, name: QName<'_>) -> Result<Option<&[u8]>> {
        let (_, prefix) = name.decompose();
        match prefix {
            None => Ok(None),
            Some(prefix) => self.resolve_prefix(prefix.as_ref(), false),
        }
    }

    fn resolve_prefix(&self, prefix: &[u8], prefixed: bool) -> Result<Option<&[u8]>> {
        if prefix == b"xml" {
            return Ok(Some(XML_NAMESPACE));
        }
        let Some(binding) = self.by_prefix.get(prefix).and_then(|stack| stack.last()) else {
            return Err(invalid("XML name uses an unbound namespace prefix"));
        };
        if binding.uri.is_empty() && prefixed {
            return Err(invalid(
                "a prefixed XML name cannot use an empty namespace binding",
            ));
        }
        Ok((!binding.uri.is_empty()).then_some(binding.uri.as_slice()))
    }
}

fn namespace_context_bindings(context: &NamespaceContext) -> Result<Vec<(Vec<u8>, Vec<u8>)>> {
    let mut output = Vec::new();
    output
        .try_reserve(context.binding_count())
        .map_err(|source| allocation("drawing source inherited namespace bindings", source))?;
    let mut failure = None;
    context.visit_visible(|prefix, uri| {
        if failure.is_some() {
            return;
        }
        let prefix_bytes = prefix.map_or(&[][..], str::as_bytes);
        let mut prefix_copy = Vec::new();
        if let Err(source) = prefix_copy.try_reserve_exact(prefix_bytes.len()) {
            failure = Some(allocation(
                "drawing source inherited namespace prefix",
                source,
            ));
            return;
        }
        prefix_copy.extend_from_slice(prefix_bytes);
        let mut uri_copy = Vec::new();
        if let Err(source) = uri_copy.try_reserve_exact(uri.len()) {
            failure = Some(allocation("drawing source inherited namespace URI", source));
            return;
        }
        uri_copy.extend_from_slice(uri.as_bytes());
        if let Err(source) = output.try_reserve(1) {
            failure = Some(allocation(
                "drawing source inherited namespace bindings",
                source,
            ));
            return;
        }
        output.push((prefix_copy, uri_copy));
    })?;
    if let Some(error) = failure {
        return Err(error);
    }
    Ok(output)
}

/// Add inherited namespace declarations to the root of a standalone XML
/// element fragment.  The fragment itself remains byte-for-byte unchanged
/// after the root opening tag; comments, PI, CDATA, and unknown descendants
/// are therefore carried through exactly.
fn namespace_complete_element_fragment(
    fragment: &[u8],
    inherited: &[(Vec<u8>, Vec<u8>)],
    max_output_bytes: usize,
) -> Result<Vec<u8>> {
    let max_output_bytes = max_output_bytes.min(MAX_FRAGMENT_BYTES);
    if fragment.len() > max_output_bytes {
        return Err(limit("namespace-complete fragment bytes", max_output_bytes));
    }
    if inherited.is_empty() {
        let mut output = Vec::new();
        output
            .try_reserve_exact(fragment.len())
            .map_err(|source| allocation("namespace-complete drawing fragment", source))?;
        output.extend_from_slice(fragment);
        return Ok(output);
    }
    let mut reader = Reader::from_reader(fragment);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        let empty_root = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let end = usize::try_from(reader.buffer_position())
                    .map_err(|_| invalid("fragment root end exceeds usize"))?;
                let mut declared = Vec::<Vec<u8>>::new();
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(xml_error)?;
                    let Some(prefix) = attribute.key.as_namespace_binding() else {
                        continue;
                    };
                    let prefix = match prefix {
                        quick_xml::name::PrefixDeclaration::Default => &[][..],
                        quick_xml::name::PrefixDeclaration::Named(prefix) => prefix,
                    };
                    declared
                        .try_reserve(1)
                        .map_err(|source| allocation("fragment namespace declarations", source))?;
                    let mut copy = Vec::new();
                    copy.try_reserve_exact(prefix.len())
                        .map_err(|source| allocation("fragment namespace prefix", source))?;
                    copy.extend_from_slice(prefix);
                    declared.push(copy);
                }
                let insertion = if empty_root {
                    end.checked_sub(2)
                        .ok_or_else(|| invalid("fragment self-closing root is truncated"))?
                } else {
                    end.checked_sub(1)
                        .ok_or_else(|| invalid("fragment root is truncated"))?
                };
                let mut additions_len = 0usize;
                for (prefix, uri) in inherited {
                    if declared
                        .iter()
                        .any(|old| old.as_slice() == prefix.as_slice())
                    {
                        continue;
                    }
                    let escaped_len = namespace_uri_escaped_len(uri)?;
                    let fixed = if prefix.is_empty() { 9 } else { 10 };
                    additions_len = additions_len
                        .checked_add(fixed)
                        .and_then(|length| length.checked_add(prefix.len()))
                        .and_then(|length| length.checked_add(escaped_len))
                        .ok_or_else(|| invalid("fragment namespace output overflows"))?;
                }
                let output_len = fragment
                    .len()
                    .checked_add(additions_len)
                    .ok_or_else(|| invalid("fragment output overflows"))?;
                if output_len > max_output_bytes {
                    return Err(limit("namespace-complete fragment bytes", max_output_bytes));
                }
                let mut output = Vec::new();
                output
                    .try_reserve_exact(output_len)
                    .map_err(|source| allocation("namespace-complete drawing fragment", source))?;
                output.extend_from_slice(&fragment[..insertion]);
                for (prefix, uri) in inherited {
                    if declared
                        .iter()
                        .any(|old| old.as_slice() == prefix.as_slice())
                    {
                        continue;
                    }
                    if prefix.is_empty() {
                        output.extend_from_slice(b" xmlns=\"");
                    } else {
                        output.extend_from_slice(b" xmlns:");
                        output.extend_from_slice(prefix);
                        output.extend_from_slice(b"=\"");
                    }
                    push_namespace_uri(&mut output, uri)?;
                    output.push(b'"');
                }
                output.extend_from_slice(&fragment[insertion..]);
                return Ok(output);
            },
            Event::Eof => return Err(invalid("fragment has no root element")),
            _ => {},
        }
        buffer.clear();
    }
}

fn extension_uri(element: &BytesStart<'_>, decoder: Decoder) -> Result<(bool, Vec<u8>, bool)> {
    let mut uri = None::<Vec<u8>>;
    let mut decoded = None::<String>;
    let mut malformed_attributes = false;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        if attribute.key.prefix().is_some() || attribute.key.as_ref() != b"uri" {
            // Unknown extension attributes are opaque.  An admitted URI with
            // an extra attribute is treated as malformed by the projection.
            malformed_attributes = true;
            continue;
        }
        if uri.is_some() {
            return Err(invalid("a:ext has duplicate uri attributes"));
        }
        if attribute.value.len() > MAX_NAMESPACE_BYTES {
            return Err(limit(
                "SVG extension URI lexical bytes",
                MAX_NAMESPACE_BYTES,
            ));
        }
        let mut raw = Vec::new();
        raw.try_reserve_exact(attribute.value.len())
            .map_err(|source| allocation("SVG extension URI lexical bytes", source))?;
        raw.extend_from_slice(attribute.value.as_ref());
        decoded = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned(),
        );
        uri = Some(raw);
    }
    let raw = uri.ok_or_else(|| invalid("a:ext is missing its uri attribute"))?;
    let value = decoded.ok_or_else(|| invalid("a:ext URI cannot be decoded"))?;
    Ok((
        xsd_token_is(value.as_bytes(), SVG_EXTENSION_URI),
        raw,
        malformed_attributes,
    ))
}

fn xsd_token_is(value: &[u8], expected: &[u8]) -> bool {
    let mut atoms = value
        .split(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
        .filter(|atom| !atom.is_empty());
    atoms.next() == Some(expected) && atoms.next().is_none()
}

fn relationship_attribute(
    element: &BytesStart<'_>,
    namespaces: &Namespaces,
    local_name: &[u8],
    decoder: Decoder,
) -> Result<Option<(String, RelationshipDialect)>> {
    let mut found = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(limit(
                "XML attribute value bytes",
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        let namespace = namespaces.resolve_attribute(attribute.key)?;
        let Some(namespace) = namespace else { continue };
        let dialect = if namespace == STRICT_RELATIONSHIPS {
            RelationshipDialect::Strict
        } else if namespace == RELATIONSHIPS {
            RelationshipDialect::Transitional
        } else {
            continue;
        };
        if attribute.key.local_name().as_ref() != local_name {
            continue;
        }
        if found.is_some() {
            return Err(invalid("relationship attribute is duplicated"));
        }
        if attribute.value.len() > MAX_RELATIONSHIP_ID_BYTES.saturating_mul(4) {
            return Err(limit(
                "relationship ID lexical bytes",
                MAX_RELATIONSHIP_ID_BYTES.saturating_mul(4),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if value.len() > MAX_RELATIONSHIP_ID_BYTES
            || value.is_empty()
            || !litchi_ooxml_common::xml_name::is_ncname(&value)
        {
            return Err(invalid("relationship ID is not an XML NCName"));
        }
        found = Some((value, dialect));
    }
    Ok(found)
}

fn required_core_content_part_relationship(
    element: &BytesStart<'_>,
    namespaces: &Namespaces,
    decoder: Decoder,
    drawing_dialect: Option<DrawingDialect>,
) -> Result<(String, RelationshipDialect)> {
    let expected = match drawing_dialect {
        Some(DrawingDialect::Strict) => RelationshipDialect::Strict,
        Some(DrawingDialect::Transitional) | None => RelationshipDialect::Transitional,
    };
    let mut found = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let namespace = namespaces.resolve_attribute(attribute.key)?;
        let dialect = if namespace == Some(STRICT_RELATIONSHIPS) {
            RelationshipDialect::Strict
        } else if namespace == Some(RELATIONSHIPS) {
            RelationshipDialect::Transitional
        } else {
            return Err(invalid(
                "core SpreadsheetDrawing contentPart allows only a matching-dialect r:id",
            ));
        };
        if attribute.key.local_name().as_ref() != b"id" {
            return Err(invalid(
                "core SpreadsheetDrawing contentPart allows only a matching-dialect r:id",
            ));
        }
        if dialect != expected {
            return Err(invalid(
                "core SpreadsheetDrawing contentPart r:id uses the wrong relationship dialect",
            ));
        }
        if found.is_some() {
            return Err(invalid(
                "core SpreadsheetDrawing contentPart has duplicate r:id",
            ));
        }
        if attribute.value.len() > MAX_RELATIONSHIP_ID_BYTES.saturating_mul(4) {
            return Err(limit(
                "contentPart relationship ID lexical bytes",
                MAX_RELATIONSHIP_ID_BYTES.saturating_mul(4),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if value.len() > MAX_RELATIONSHIP_ID_BYTES
            || value.is_empty()
            || !litchi_ooxml_common::xml_name::is_ncname(&value)
        {
            return Err(invalid("contentPart relationship ID is not an XML NCName"));
        }
        found = Some((value, dialect));
    }
    found.ok_or_else(|| invalid("core SpreadsheetDrawing contentPart is missing r:id"))
}

fn relationship_dialect_attribute(
    element: &BytesStart<'_>,
    namespaces: &Namespaces,
    local_name: &[u8],
) -> Result<Option<RelationshipDialect>> {
    let mut found = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_namespace_binding().is_some()
            || attribute.key.local_name().as_ref() != local_name
        {
            continue;
        }
        let Some(namespace) = namespaces.resolve_attribute(attribute.key)? else {
            continue;
        };
        let dialect = if namespace == STRICT_RELATIONSHIPS {
            RelationshipDialect::Strict
        } else if namespace == RELATIONSHIPS {
            RelationshipDialect::Transitional
        } else {
            continue;
        };
        if found.is_none() {
            found = Some(dialect);
        }
    }
    Ok(found)
}

fn validate_start_attributes(
    element: &BytesStart<'_>,
    namespaces: &Namespaces,
    decoder: Decoder,
) -> Result<()> {
    let mut seen = Vec::<(Vec<u8>, Vec<u8>)>::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let namespace = namespaces.resolve_attribute(attribute.key)?;
        let uri = namespace.unwrap_or(&[]);
        if attribute.key.prefix().is_some() && uri.is_empty() {
            return Err(invalid("prefixed XML attribute has no namespace binding"));
        }
        let local = attribute.key.local_name();
        if seen.iter().any(|(old_uri, old_local)| {
            old_uri.as_slice() == uri && old_local.as_slice() == local.as_ref()
        }) {
            return Err(invalid("element has duplicate expanded attributes"));
        }
        seen.try_reserve(1)
            .map_err(|source| allocation("drawing source expanded attribute set", source))?;
        let mut uri_copy = Vec::new();
        uri_copy
            .try_reserve_exact(uri.len())
            .map_err(|source| allocation("drawing source expanded attribute namespace", source))?;
        uri_copy.extend_from_slice(uri);
        let mut local_copy = Vec::new();
        local_copy
            .try_reserve_exact(local.as_ref().len())
            .map_err(|source| allocation("drawing source expanded attribute name", source))?;
        local_copy.extend_from_slice(local.as_ref());
        seen.push((uri_copy, local_copy));
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        validate_xml_characters(value.as_bytes())?;
    }
    Ok(())
}

fn validate_qname(bytes: &[u8], kind: &str) -> Result<()> {
    if bytes.is_empty() || bytes.len() > MAX_NAME_BYTES {
        return Err(limit("XML qualified name bytes", MAX_NAME_BYTES));
    }
    let value = std::str::from_utf8(bytes).map_err(|error| Error::Invalid(error.to_string()))?;
    if !litchi_ooxml_common::xml_name::is_qualified_name(value) {
        return Err(invalid(format!("invalid XML {kind} qualified name")));
    }
    Ok(())
}

fn validate_namespace_binding(
    prefix: quick_xml::name::PrefixDeclaration<'_>,
    uri: &[u8],
) -> Result<()> {
    let prefix = match prefix {
        quick_xml::name::PrefixDeclaration::Default => &[][..],
        quick_xml::name::PrefixDeclaration::Named(prefix) => prefix,
    };
    if prefix.len() > MAX_NAME_BYTES || uri.len() > MAX_NAMESPACE_BYTES {
        return Err(limit("namespace prefix or URI bytes", MAX_NAMESPACE_BYTES));
    }
    if prefix.is_empty() {
        if uri == XML_NAMESPACE || uri == XMLNS_NAMESPACE {
            return Err(invalid("reserved XML URI cannot be the default binding"));
        }
        return Ok(());
    }
    if prefix == b"xmlns" {
        return Err(invalid("xmlns is not a bindable namespace prefix"));
    }
    if prefix == b"xml" {
        if uri != XML_NAMESPACE {
            return Err(invalid("xml prefix is bound to the wrong namespace URI"));
        }
        return Ok(());
    }
    if uri.is_empty() || uri == XML_NAMESPACE || uri == XMLNS_NAMESPACE {
        return Err(invalid("invalid reserved or empty namespace binding"));
    }
    Ok(())
}

fn validate_xml_text(bytes: &[u8]) -> Result<()> {
    let value = std::str::from_utf8(bytes).map_err(|error| Error::Invalid(error.to_string()))?;
    if bytes.windows(3).any(|window| window == b"]]>") {
        return Err(invalid("XML text contains the forbidden ]]> delimiter"));
    }
    if value.chars().all(is_xml_char) {
        Ok(())
    } else {
        Err(invalid("XML text contains an invalid XML character"))
    }
}

fn validate_xml_characters(bytes: &[u8]) -> Result<()> {
    let value = std::str::from_utf8(bytes).map_err(|error| Error::Invalid(error.to_string()))?;
    if value.chars().all(is_xml_char) {
        Ok(())
    } else {
        Err(invalid("XML value contains an invalid XML character"))
    }
}

fn validate_raw_attribute(bytes: &[u8]) -> Result<()> {
    if bytes.contains(&b'<') {
        return Err(invalid("XML attribute contains a forbidden raw delimiter"));
    }
    Ok(())
}

fn validate_reference(reference: &BytesRef<'_>) -> Result<()> {
    decode_xml_reference(reference).map(|value| validate_xml_characters(value.as_bytes()))?
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let version = declaration.version().map_err(xml_error)?;
    if version.as_ref() != b"1.0" {
        return Err(invalid("XML declaration must use version 1.0"));
    }
    if let Some(encoding) = declaration.encoding() {
        let encoding = encoding.map_err(xml_error)?;
        if !encoding.as_ref().eq_ignore_ascii_case(b"utf-8") {
            return Err(invalid("drawing XML declaration must use UTF-8"));
        }
    }
    Ok(())
}

fn validate_processing_instruction(bytes: &[u8]) -> Result<()> {
    validate_xml_characters(bytes)?;
    let end = bytes
        .iter()
        .position(u8::is_ascii_whitespace)
        .unwrap_or(bytes.len());
    let target = bytes.get(..end).ok_or_else(|| invalid("empty PI target"))?;
    let target = std::str::from_utf8(target).map_err(|error| Error::Invalid(error.to_string()))?;
    if !litchi_ooxml_common::xml_name::is_xml_name(target) || target.eq_ignore_ascii_case("xml") {
        return Err(invalid("invalid XML processing-instruction target"));
    }
    Ok(())
}

fn qname_prefix(name: &[u8]) -> Result<Vec<u8>> {
    let prefix = name
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&[][..], |index| &name[..index]);
    if prefix.len() > MAX_NAME_BYTES {
        return Err(limit("XML namespace prefix bytes", MAX_NAME_BYTES));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(prefix.len())
        .map_err(|source| allocation("drawing source XML prefix", source))?;
    output.extend_from_slice(prefix);
    Ok(output)
}

fn namespace_uri_copy(namespace: Option<&[u8]>) -> Result<Box<[u8]>> {
    let namespace = namespace.unwrap_or_default();
    if namespace.len() > MAX_NAMESPACE_BYTES {
        return Err(limit("XML namespace URI bytes", MAX_NAMESPACE_BYTES));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(namespace.len())
        .map_err(|source| allocation("drawing source namespace URI", source))?;
    output.extend_from_slice(namespace);
    Ok(output.into_boxed_slice())
}

fn unqualified_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<String>> {
    unqualified_attribute_value(element, name, decoder)
        .map_err(|error| Error::Invalid(error.to_string()))
}

fn bounded_unqualified_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    max_bytes: usize,
    label: &str,
) -> Result<Option<String>> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_namespace_binding().is_some()
            || attribute.key.prefix().is_some()
            || attribute.key.as_ref() != name
        {
            continue;
        }
        if attribute.value.len() > max_bytes.saturating_mul(4) {
            return Err(limit(label, max_bytes));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if value.len() > max_bytes {
            return Err(limit(label, max_bytes));
        }
        return Ok(Some(value));
    }
    Ok(None)
}

fn coordinate_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    label: &str,
) -> Result<i64> {
    let value = unqualified_attribute(element, name, decoder)?
        .ok_or_else(|| invalid(format!("{label} attribute is missing")))?;
    let value = value.trim_matches([' ', '\t', '\r', '\n']);
    let parsed = value
        .parse::<i64>()
        .map_err(|_| invalid(format!("invalid {label} '{value}'")))?;
    check_coordinate(parsed, label)?;
    Ok(parsed)
}

fn positive_coordinate_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    label: &str,
) -> Result<i64> {
    let value = coordinate_attribute(element, name, decoder, label)?;
    if value < 0 {
        return Err(invalid(format!("{label} is negative")));
    }
    Ok(value)
}

fn check_coordinate(value: i64, label: &str) -> Result<()> {
    const MIN: i64 = -27_273_042_329_600;
    const MAX: i64 = 27_273_042_316_900;
    if !(MIN..=MAX).contains(&value) {
        return Err(invalid(format!("{label} exceeds DrawingML bounds")));
    }
    Ok(())
}

fn check_marker_bounds(marker: CellMarker) -> Result<()> {
    if marker.column >= 16_384 || marker.row >= 1_048_576 {
        return Err(invalid("drawing anchor exceeds worksheet bounds"));
    }
    Ok(())
}

fn trim_xml_schema_whitespace(value: &str) -> &str {
    value.trim_matches([' ', '\t', '\r', '\n'])
}

fn is_xml_char(value: char) -> bool {
    matches!(
        value,
        '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
    )
}

fn is_xml_whitespace(bytes: &[u8]) -> bool {
    bytes
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn namespace_uri_escaped_len(value: &[u8]) -> Result<usize> {
    let value = std::str::from_utf8(value).map_err(|error| Error::Invalid(error.to_string()))?;
    value.chars().try_fold(0usize, |length, character| {
        let added = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' => 6,
            '\t' | '\n' | '\r' => 5,
            _ => character.len_utf8(),
        };
        length
            .checked_add(added)
            .ok_or_else(|| invalid("namespace URI escaped size overflows"))
    })
}

fn push_namespace_uri(output: &mut Vec<u8>, value: &[u8]) -> Result<()> {
    let value = std::str::from_utf8(value).map_err(|error| Error::Invalid(error.to_string()))?;
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0u8; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
    Ok(())
}

fn position(reader: &Reader<&[u8]>, source_prefix: usize) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("XML source position exceeds usize"))?
        .checked_add(source_prefix)
        .ok_or_else(|| invalid("XML source position overflows usize"))
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    invalid(format!("source-backed drawing XML: {error}"))
}

fn invalid(message: impl Into<String>) -> Error {
    xlsx_invalid(format!("source-backed XLSX drawing: {}", message.into()))
}

fn limit(resource: impl Into<String>, value: usize) -> Error {
    let resource = resource.into();
    xlsx_invalid(format!(
        "source-backed XLSX drawing {resource} exceeds the limit of {value}"
    ))
}
