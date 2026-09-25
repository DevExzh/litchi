//! Source-preserving `themeFamily` ownership inside a complete `a:theme` part.
//!
//! The fragment codec in [`super::codec`] owns the `themeFamily` grammar. This
//! module owns only the context needed to reach that fragment from a complete
//! DrawingML theme: the direct `theme/extLst/ext` path, the extension URI
//! profile, inherited namespace bindings, and source spans. Package
//! relationships, transactions, and durable patches remain with the format
//! that owns the Theme part.
//!
//! XML names, characters, references, and document boundaries are checked across
//! the entire source. Declarations must specify XML 1.0 and, when present, UTF-8.
//! The owned root/list/recognized-extension containers allow only whitespace
//! between elements. The direct extension list admits only DrawingML `ext`
//! children and MCE `AlternateContent` envelopes (subject to the MCE mutation
//! refusal policy). Other descendants remain opaque and may contain valid XML
//! text; this ownership scanner does not perform full Theme schema validation.
//!
//! Limited mutations check the prospective complete-part size before serializing
//! replacement buffers. Their cap also bounds the standalone family intermediate
//! used to preserve inherited namespaces; bounded scalar sizing and source
//! validation still occur before those output buffers are constructed.

use std::{
    collections::{HashMap, HashSet},
    fmt,
    ops::Range,
    sync::Arc,
};

use litchi_core::xml::ReaderOrigin;
use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml_name::is_qualified_name;
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use crate::{Error, Result};

use super::{
    DRAWINGML_NAMESPACE, DRAWINGML_NAMESPACE_STRICT, Family, MAX_ATTRIBUTE_VALUE_BYTES,
    MAX_ATTRIBUTES, MAX_DEPTH, MAX_NAMESPACE_BYTES, MAX_NAMESPACE_DECLARATIONS, MAX_NODES,
    NAMESPACE as FAMILY_NAMESPACE, XML_NAMESPACE, codec as family_codec,
};
use crate::theme::codec as theme_codec;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_ooxml_common::xml::attributes::first_wins;

/// The normative extension URI from `[MS-ODRAWXML]` §2.2.8.
pub const EXTENSION_URI: &str = FAMILY_NAMESPACE;

/// The Office native discriminator used by the vendored PPTX/XLSB fixtures.
pub const NATIVE_EXTENSION_URI: &str = "{05A4C25C-085E-4340-85A3-A5531E510DB2}";

const MCE_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Maximum user namespace declarations retained across the active element stack.
///
/// Shadowed declarations count until their element closes. Prefix lookup scans
/// this bounded scope; namespace handling is not claimed to be fully linear.
/// The scanner counts entries using the scope vector length and restores it on
/// scope exit, without folding the bindings.
///
/// quick-xml resolves each event before this scanner admits its declarations.
/// Its per-element cap therefore permits up to one additional event's
/// `MAX_NAMESPACE_DECLARATIONS` user declarations transiently (at most 512 user
/// declarations total), plus the resolver's built-in bindings. Oversized events
/// are rejected before classification or candidate scope copying.
pub const MAX_ACTIVE_NAMESPACE_DECLARATIONS: usize = MAX_NAMESPACE_DECLARATIONS;

// The manually maintained scope includes the built-in `xml` binding.
const MAX_ACTIVE_NAMESPACE_BINDINGS: usize = MAX_ACTIVE_NAMESPACE_DECLARATIONS + 1;

/// Maximum complete theme-part bytes inspected by this owner.
pub const MAX_XML_BYTES: usize = theme_codec::MAX_XML_BYTES;

/// The extension URI profile that owns a parsed family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[must_use]
pub enum ExtensionProfile {
    /// `[MS-ODRAWXML]` §2.2.8's namespace URI.
    Normative,
    /// The native Office discriminator used by the local fixture corpus.
    NativeDiscriminator,
}

impl ExtensionProfile {
    /// Return the exact URI used by this profile.
    #[must_use]
    pub const fn uri(self) -> &'static str {
        match self {
            Self::Normative => EXTENSION_URI,
            Self::NativeDiscriminator => NATIVE_EXTENSION_URI,
        }
    }

    fn from_uri(value: &str) -> Option<Self> {
        match value {
            EXTENSION_URI => Some(Self::Normative),
            NATIVE_EXTENSION_URI => Some(Self::NativeDiscriminator),
            _ => None,
        }
    }
}

/// An immutable, source-backed complete DrawingML Theme part.
#[derive(Debug, Clone)]
#[must_use]
pub struct Snapshot {
    source: Arc<[u8]>,
    family: Option<Family>,
    owner: Option<Arc<Owner>>,
}

impl Snapshot {
    /// Parse a complete Theme part and retain its source bytes.
    pub fn from_xml(xml: impl AsRef<[u8]>) -> Result<Self> {
        read(xml.as_ref())
    }

    /// Parse a complete Theme part while retaining an existing source
    /// allocation.
    pub fn from_shared_xml(xml: Arc<[u8]>) -> Result<Self> {
        read_shared(xml)
    }

    /// Borrow the exact source XML.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.source.as_ref()
    }

    /// Borrow the optional source-backed `themeFamily` projection. Inherited
    /// namespace bindings are injected as declarations in its standalone-ready
    /// [`Family::source`]. Use [`Self::family_range`] with [`Self::xml_bytes`]
    /// for the exact raw fragment before those declarations were injected.
    #[must_use]
    pub fn family(&self) -> Option<&Family> {
        self.family.as_ref()
    }

    /// Return the URI profile of the recognized owner, if one is present.
    #[must_use]
    pub fn family_profile(&self) -> Option<ExtensionProfile> {
        self.owner.as_ref().map(|owner| owner.profile)
    }

    /// Return the normalized recognized extension URI, if one is present.
    ///
    /// The original lexical spelling, including XML token whitespace or
    /// entity references, remains in [`Self::xml_bytes`].
    #[must_use]
    pub fn family_extension_uri(&self) -> Option<&str> {
        self.owner.as_ref().map(|owner| owner.uri.as_str())
    }

    /// Whether a direct recognized `themeFamily` owner is present.
    #[must_use]
    pub fn has_family(&self) -> bool {
        self.family.is_some()
    }

    /// Replace the existing family, or add it when the owner is absent.
    ///
    /// When an owner exists, only the three typed scalar values from `family`
    /// are applied to the source-backed current family. This deliberately
    /// ignores any opaque source carried by an incoming value, so a detached
    /// value is safe and a value from another document cannot replace unknown
    /// attributes or extension children by accident.
    pub fn replace_family(&self, family: &Family) -> Result<Vec<u8>> {
        replace_family_source(self.source.as_ref(), family)
    }

    /// Replace the family while enforcing a caller-specific complete-part
    /// output bound before allocating the patched source.
    pub fn replace_family_with_limit(
        &self,
        family: &Family,
        max_xml_bytes: usize,
    ) -> Result<Vec<u8>> {
        replace_family_source_with_limit(self.source.as_ref(), family, max_xml_bytes)
    }

    /// Add a family using the normative extension URI.
    ///
    /// A leading standalone UTF-8 BOM on the incoming fragment is stripped
    /// before embedding; an XML declaration remains invalid as child markup.
    pub fn add_family(&self, family: &Family) -> Result<Vec<u8>> {
        add_family_source(self.source.as_ref(), family, EXTENSION_URI)
    }

    /// Add a normative family with a caller-specific complete-part bound.
    pub fn add_family_with_limit(&self, family: &Family, max_xml_bytes: usize) -> Result<Vec<u8>> {
        add_family_source_with_limit(self.source.as_ref(), family, EXTENSION_URI, max_xml_bytes)
    }

    /// Add a family using an explicitly admitted extension URI profile.
    ///
    /// A leading standalone UTF-8 BOM is stripped before embedding; an XML
    /// declaration remains invalid as child markup.
    pub fn add_family_with_uri(&self, family: &Family, uri: &str) -> Result<Vec<u8>> {
        add_family_source(self.source.as_ref(), family, uri)
    }

    /// Add a family with an admitted URI and caller-specific complete-part
    /// output bound.
    pub fn add_family_with_uri_limit(
        &self,
        family: &Family,
        uri: &str,
        max_xml_bytes: usize,
    ) -> Result<Vec<u8>> {
        add_family_source_with_limit(self.source.as_ref(), family, uri, max_xml_bytes)
    }

    /// Remove the recognized direct family owner and its semantically empty
    /// extension/list containers. XML whitespace alone does not retain a
    /// container; comments, unrelated content, and foreign attributes do.
    pub fn remove_family(&self) -> Result<Vec<u8>> {
        remove_family_source(self.source.as_ref())
    }

    /// Remove the family while enforcing a caller-specific complete-part
    /// output bound before allocating the patched source.
    pub fn remove_family_with_limit(&self, max_xml_bytes: usize) -> Result<Vec<u8>> {
        remove_family_source_with_limit(self.source.as_ref(), max_xml_bytes)
    }

    /// Borrow the complete source-backed owner range, when present.
    #[must_use]
    pub fn family_range(&self) -> Option<Range<usize>> {
        self.owner.as_ref().map(|owner| owner.family_range.clone())
    }
}

/// Parse a complete DrawingML Theme part.
pub fn read(xml: &[u8]) -> Result<Snapshot> {
    // Validate the borrowed input before allocating the retained source.
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("theme XML bytes", MAX_XML_BYTES));
    }
    let scanned = scan(xml)?;
    let (family, owner) = parse_candidate(xml, &scanned)?;
    Ok(Snapshot {
        source: Arc::<[u8]>::from(xml),
        family,
        owner: owner.map(Arc::new),
    })
}

/// Parse a complete DrawingML Theme part from an existing source allocation.
pub fn read_shared(source: Arc<[u8]>) -> Result<Snapshot> {
    if source.len() > MAX_XML_BYTES {
        return Err(limit("theme XML bytes", MAX_XML_BYTES));
    }
    let scanned = scan(source.as_ref())?;
    let (family, owner) = parse_candidate(source.as_ref(), &scanned)?;
    Ok(Snapshot {
        source,
        family,
        owner: owner.map(Arc::new),
    })
}

/// Read only the optional family projection from borrowed Theme bytes.
///
/// This path does not retain or copy the complete Theme part and is
/// intended for package owners that already manage the part's source
/// allocation. When the family inherits namespace bindings from the
/// Theme, the returned [`Family::source`] includes injected namespace
/// declarations so the family codec can write a standalone value. Use
/// [`family_range`] with the original Theme bytes for the exact raw span.
pub fn read_family(xml: &[u8]) -> Result<Option<Family>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("theme XML bytes", MAX_XML_BYTES));
    }
    let scanned = scan(xml)?;
    parse_family(xml, &scanned)
}

/// Return the direct family owner span without copying the complete part.
pub fn family_range(xml: &[u8]) -> Result<Option<Range<usize>>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("theme XML bytes", MAX_XML_BYTES));
    }
    Ok(scan(xml)?.candidate.map(|candidate| candidate.family_range))
}

/// Return the profile of the direct recognized family owner without copying
/// the complete part.
pub fn family_profile(xml: &[u8]) -> Result<Option<ExtensionProfile>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("theme XML bytes", MAX_XML_BYTES));
    }
    let scanned = scan(xml)?;
    Ok(scanned
        .candidate
        .as_ref()
        .and_then(|candidate| ExtensionProfile::from_uri(&candidate.uri)))
}

/// Replace an existing family or add a new normative family in a complete
/// Theme part.
pub fn replace_family(xml: &[u8], family: &Family) -> Result<Vec<u8>> {
    replace_family_source(xml, family)
}

/// Replace an existing family or add one while enforcing a caller-specific
/// complete-part output bound before allocating the patched source.
pub fn replace_family_with_limit(
    xml: &[u8],
    family: &Family,
    max_xml_bytes: usize,
) -> Result<Vec<u8>> {
    replace_family_source_with_limit(xml, family, max_xml_bytes)
}

/// Add a normative family to a complete Theme part. A leading standalone
/// UTF-8 BOM on the incoming fragment is stripped before embedding; an XML
/// declaration remains invalid as child markup.
pub fn add_family(xml: &[u8], family: &Family) -> Result<Vec<u8>> {
    add_family_source(xml, family, EXTENSION_URI)
}

/// Add a normative family with a caller-specific complete-part bound.
pub fn add_family_with_limit(xml: &[u8], family: &Family, max_xml_bytes: usize) -> Result<Vec<u8>> {
    add_family_source_with_limit(xml, family, EXTENSION_URI, max_xml_bytes)
}

/// Add a family with an explicitly admitted extension URI. A leading
/// standalone UTF-8 BOM is stripped before embedding; an XML declaration
/// remains invalid as child markup.
pub fn add_family_with_uri(xml: &[u8], family: &Family, uri: &str) -> Result<Vec<u8>> {
    add_family_source(xml, family, uri)
}

/// Add a family with an admitted URI and caller-specific complete-part bound.
pub fn add_family_with_uri_limit(
    xml: &[u8],
    family: &Family,
    uri: &str,
    max_xml_bytes: usize,
) -> Result<Vec<u8>> {
    add_family_source_with_limit(xml, family, uri, max_xml_bytes)
}

/// Remove a direct recognized family and empty enclosing extension/list
/// containers. Preserve comments, unrelated content, and foreign attributes.
pub fn remove_family(xml: &[u8]) -> Result<Vec<u8>> {
    remove_family_source(xml)
}

/// Remove a direct recognized family with a caller-specific complete-part
/// output bound.
pub fn remove_family_with_limit(xml: &[u8], max_xml_bytes: usize) -> Result<Vec<u8>> {
    remove_family_source_with_limit(xml, max_xml_bytes)
}

fn replace_family_source(xml: &[u8], incoming: &Family) -> Result<Vec<u8>> {
    replace_family_source_with_limit(xml, incoming, MAX_XML_BYTES)
}

fn replace_family_source_with_limit(
    xml: &[u8],
    incoming: &Family,
    max_xml_bytes: usize,
) -> Result<Vec<u8>> {
    validate_output_limit(max_xml_bytes)?;
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("theme XML bytes", MAX_XML_BYTES));
    }
    let scanned = scan(xml)?;
    let Some(candidate) = scanned.candidate.as_ref() else {
        return add_family_from_scan(xml, incoming, &scanned, EXTENSION_URI, max_xml_bytes);
    };
    if xml.len() > max_xml_bytes {
        return Err(limit("patched Theme XML bytes", max_xml_bytes));
    }
    if scanned.hidden_owner {
        return Err(invalid(
            "Theme Family owner is also hidden in an MCE branch; mutation is ambiguous",
        ));
    }
    let (current, owner) = parse_candidate(xml, &scanned)?;
    let mut staged = current.ok_or_else(|| invalid("Theme Family candidate has no projection"))?;
    if staged == *incoming {
        return Ok(xml.to_vec());
    }
    // Copy only the supported scalar projection. The source-backed `staged`
    // value remains the owner of opaque attributes, children, and MCE markup.
    staged.set_name(incoming.name())?;
    staged.set_id(incoming.id().as_str())?;
    staged.set_variant_id(incoming.variant_id().as_str())?;
    let owner = owner.ok_or_else(|| invalid("Theme Family candidate owner is missing"))?;
    let replacement_len = family_child_len(&staged, &owner.inherited_bindings, max_xml_bytes)?;
    ensure_splice_output_limit(xml, &candidate.family_range, replacement_len, max_xml_bytes)?;
    let replacement = family_child_bytes(&staged, &owner.inherited_bindings)?;
    let output = splice_with_limit(
        xml,
        &[Replacement {
            range: candidate.family_range.clone(),
            bytes: replacement,
        }],
        max_xml_bytes,
    )?;
    let parsed = validate_result(&output)?;
    if parsed.candidate.is_none() {
        return Err(invalid(
            "updated Theme Family is not a direct recognized owner",
        ));
    }
    Ok(output)
}

fn add_family_source(xml: &[u8], family: &Family, uri: &str) -> Result<Vec<u8>> {
    add_family_source_with_limit(xml, family, uri, MAX_XML_BYTES)
}

fn add_family_source_with_limit(
    xml: &[u8],
    family: &Family,
    uri: &str,
    max_xml_bytes: usize,
) -> Result<Vec<u8>> {
    validate_output_limit(max_xml_bytes)?;
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("theme XML bytes", MAX_XML_BYTES));
    }
    let scanned = scan(xml)?;
    add_family_from_scan(xml, family, &scanned, uri, max_xml_bytes)
}

fn add_family_from_scan(
    xml: &[u8],
    family: &Family,
    scanned: &Scanned,
    uri: &str,
    max_xml_bytes: usize,
) -> Result<Vec<u8>> {
    if scanned.candidate.is_some() {
        return Err(invalid("Theme Family owner already exists"));
    }
    if scanned.hidden_owner {
        return Err(invalid(
            "Theme Family owner is hidden in an MCE branch; direct insertion is ambiguous",
        ));
    }
    let _requested = ExtensionProfile::from_uri(uri)
        .ok_or_else(|| invalid("unsupported Theme Family extension URI"))?;
    let family_len = family_child_len(family, &[], max_xml_bytes)?;
    if let Some(extension) = scanned.root.admitted_ext.as_ref() {
        ensure_extension_insertion_limit(xml, extension, family_len, max_xml_bytes)?;
    } else if let Some(list) = scanned.root.ext_list.as_ref() {
        let length = serialized_extension_len(qname_prefix(&list.qualified_name), uri, family_len)?;
        ensure_container_insertion_limit(xml, list, length, max_xml_bytes)?;
    } else {
        let length = serialized_ext_list_len(&scanned.root.element_prefix, uri, family_len)?;
        ensure_root_insertion_limit(xml, &scanned.root, length, max_xml_bytes)?;
    }
    let family = family_child_bytes(family, &[])?;
    let output = if let Some(ext) = scanned.root.admitted_ext.as_ref() {
        // Preserve a native discriminator already present in the Theme part.
        insert_into_ext(xml, ext, &family, max_xml_bytes)?
    } else if let Some(ext_list) = scanned.root.ext_list.as_ref() {
        // Use the extLst's own resolved prefix. The root may rebind another
        // prefix on the extLst, so reusing the root QName here could emit an
        // `ext` in a foreign namespace.
        let prefix = qname_prefix(&ext_list.qualified_name);
        let extension_len = serialized_extension_len(prefix, uri, family.len())?;
        ensure_container_insertion_limit(xml, ext_list, extension_len, max_xml_bytes)?;
        let extension = make_extension(prefix, uri, &family)?;
        insert_into_container(xml, ext_list, &extension, max_xml_bytes)?
    } else {
        let ext_list_len =
            serialized_ext_list_len(&scanned.root.element_prefix, uri, family.len())?;
        ensure_root_insertion_limit(xml, &scanned.root, ext_list_len, max_xml_bytes)?;
        let ext_list = make_ext_list(&scanned.root.element_prefix, uri, &family)?;
        insert_before_root_close(xml, &scanned.root, &ext_list, max_xml_bytes)?
    };
    let parsed = validate_result(&output)?;
    if parsed.candidate.is_none() {
        return Err(invalid(
            "inserted Theme Family is not a direct recognized owner",
        ));
    }
    Ok(output)
}

fn remove_family_source(xml: &[u8]) -> Result<Vec<u8>> {
    remove_family_source_with_limit(xml, MAX_XML_BYTES)
}

fn remove_family_source_with_limit(xml: &[u8], max_xml_bytes: usize) -> Result<Vec<u8>> {
    validate_output_limit(max_xml_bytes)?;
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("theme XML bytes", MAX_XML_BYTES));
    }
    let scanned = scan(xml)?;
    if scanned.hidden_owner {
        return Err(invalid(
            "Theme Family ownership is hidden in an MCE or unsupported branch; removal is ambiguous",
        ));
    }
    let Some(candidate) = scanned.candidate.as_ref() else {
        if xml.len() > max_xml_bytes {
            return Err(limit("patched Theme XML bytes", max_xml_bytes));
        }
        return Ok(xml.to_vec());
    };
    // Validate the recognized fragment before deleting its source span. A
    // malformed supported owner must refuse the edit rather than being
    // silently sanitized by removal.
    let _ = parse_family(xml, &scanned)?;
    let mut removal = candidate.family_range.clone();
    if let Some(extension) = scanned.root.admitted_ext.as_ref()
        && empty_container_after_removal(
            xml,
            &extension.range,
            extension.end_start,
            &removal,
            true,
        )?
    {
        removal = extension.range.clone();
        if let Some(list) = scanned.root.ext_list.as_ref()
            && empty_container_after_removal(xml, &list.range, list.end_start, &removal, false)?
        {
            removal = list.range.clone();
        }
    }
    let output = splice_with_limit(
        xml,
        &[Replacement {
            range: removal,
            bytes: Vec::new(),
        }],
        max_xml_bytes,
    )?;
    let parsed = validate_result(&output)?;
    if parsed.candidate.is_some() {
        return Err(invalid("removed Theme Family owner remains present"));
    }
    Ok(output)
}

fn empty_container_after_removal(
    xml: &[u8],
    container: &Range<usize>,
    end_start: Option<usize>,
    removed: &Range<usize>,
    extension: bool,
) -> Result<bool> {
    let open_end = container.start + open_tag_end(&xml[container.clone()])?;
    let close_start =
        end_start.ok_or_else(|| invalid("Theme container closing span is missing"))?;
    if removed.start < open_end || removed.end > close_start {
        return Err(invalid("Theme removal span lies outside its container"));
    }
    // Namespace declarations and the selected extension's discriminator belong
    // to the wrapper. Other attributes are opaque metadata and retain it.
    let mut reader = NsReader::from_reader(&xml[container.start..open_end]);
    let Event::Start(element) = reader.read_event().map_err(xml_error)? else {
        return Err(invalid("Theme container opening tag is missing"));
    };
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let name = attribute.key.as_ref();
        if namespace_declaration(name).is_none() && !(extension && name == b"uri") {
            return Ok(false);
        }
    }
    Ok(only_xml_whitespace(&xml[open_end..removed.start])?
        && only_xml_whitespace(&xml[removed.end..close_start])?)
}

fn only_xml_whitespace(xml: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        match reader.read_event().map_err(xml_error)? {
            Event::Text(value) if is_xml_whitespace(value.as_ref()) => {},
            Event::CData(value) if is_xml_whitespace(value.as_ref()) => {},
            Event::GeneralRef(value)
                if family_codec::validate_general_ref(&value)?
                    .is_some_and(is_xml_whitespace_character) => {},
            Event::Eof => return Ok(true),
            _ => return Ok(false),
        }
    }
}

#[derive(Debug, Clone)]
struct Owner {
    family_range: Range<usize>,
    profile: ExtensionProfile,
    uri: String,
    inherited_bindings: Vec<Binding>,
}

#[derive(Debug, Clone)]
struct Root {
    namespace: String,
    element_prefix: Vec<u8>,
    qualified_name: Vec<u8>,
    range: Range<usize>,
    end_start: Option<usize>,
    empty: bool,
    ext_list: Option<Container>,
    admitted_ext: Option<Extension>,
}

#[derive(Debug, Clone)]
struct Container {
    range: Range<usize>,
    end_start: Option<usize>,
    qualified_name: Vec<u8>,
    empty: bool,
}

#[derive(Debug, Clone)]
struct Extension {
    range: Range<usize>,
    end_start: Option<usize>,
    qualified_name: Vec<u8>,
    empty: bool,
}

#[derive(Debug, Clone)]
struct Candidate {
    family_range: Range<usize>,
    uri: String,
    inherited_bindings: Vec<Binding>,
}

#[derive(Debug, Clone)]
struct Binding {
    prefix: Vec<u8>,
    uri: Vec<u8>,
}

#[derive(Debug)]
struct Scanned {
    root: Root,
    candidate: Option<Candidate>,
    hidden_owner: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Root,
    ExtList,
    Ext(usize),
    Family(usize),
    MceOther,
    MceChoice,
    MceFallback,
    HiddenExtList,
    HiddenExt(bool),
    Other,
}

#[derive(Debug)]
struct Frame {
    local: Vec<u8>,
    namespace: Vec<u8>,
    kind: Kind,
    ns_added: usize,
    scope_bytes_before: usize,
}

#[derive(Debug)]
struct ExtensionRecord {
    range: Range<usize>,
    end_start: Option<usize>,
    qualified_name: Vec<u8>,
    uri: Option<String>,
    empty: bool,
}

#[derive(Debug)]
struct FamilyRecord {
    extension_index: usize,
    start: usize,
    end: Option<usize>,
    inherited_scope: Vec<Binding>,
    local_declarations: Vec<Vec<u8>>,
}

struct Classifier<'a> {
    root: &'a mut Option<Root>,
    ext_list_seen: &'a mut bool,
    extensions: &'a mut Vec<ExtensionRecord>,
    families: &'a mut Vec<FamilyRecord>,
    recognized_family_extensions: &'a mut HashSet<usize>,
    hidden_owner: &'a mut bool,
}

fn parse_family(xml: &[u8], scanned: &Scanned) -> Result<Option<Family>> {
    let Some(candidate) = scanned.candidate.as_ref() else {
        return Ok(None);
    };
    let decorated = decorate_family(xml, candidate)?;
    let family = family_codec::read(&decorated)?;
    Ok(Some(family))
}

fn parse_candidate(xml: &[u8], scanned: &Scanned) -> Result<(Option<Family>, Option<Owner>)> {
    let Some(candidate) = scanned.candidate.as_ref() else {
        return Ok((None, None));
    };
    let family = parse_family(xml, scanned)?;
    let profile = ExtensionProfile::from_uri(&candidate.uri)
        .ok_or_else(|| invalid("internal Theme Family URI profile is unsupported"))?;
    Ok((
        family,
        Some(Owner {
            family_range: candidate.family_range.clone(),
            profile,
            uri: candidate.uri.clone(),
            inherited_bindings: candidate.inherited_bindings.clone(),
        }),
    ))
}

fn scan(xml: &[u8]) -> Result<Scanned> {
    // The scan splits off the first byte-order mark itself; a second one,
    // which the reader would drop uncounted, is the parsed slice's origin.
    let bom_len = ReaderOrigin::of(xml).skipped();
    let parse_xml = &xml[bom_len..];
    let origin = ReaderOrigin::of(parse_xml);
    let mut reader = NsReader::from_reader(parse_xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    reader
        .resolver_mut()
        .set_max_declarations_per_element(MAX_NAMESPACE_DECLARATIONS);

    // This stack is mutated in place. Ordinary theme nodes therefore do not
    // clone the namespace environment; only the candidate family records
    // retain a bounded copy when their direct owner path is reached.
    let mut scope = vec![Binding {
        prefix: b"xml".to_vec(),
        uri: XML_NAMESPACE.as_bytes().to_vec(),
    }];
    // Account once for the built-in binding. Ordinary elements do not walk
    // inherited bindings merely to recompute their aggregate size.
    let mut scope_bytes = b"xml".len() + XML_NAMESPACE.len();
    let mut stack = Vec::<Frame>::new();
    let mut extensions = Vec::<ExtensionRecord>::new();
    let mut families = Vec::<FamilyRecord>::new();
    let mut recognized_family_extensions = HashSet::<usize>::new();
    let mut root: Option<Root> = None;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut pre_root_markup = false;
    let mut ext_list_seen = false;
    let mut hidden_owner = false;
    let mut nodes = 0usize;

    loop {
        let event_start = position(&reader, origin, bom_len)?;
        let (resolved, event) = reader.read_resolved_event().map_err(xml_error)?;
        match &event {
            Event::Start(element) | Event::Empty(element) => {
                validate_namespace_prefix(qname_prefix(element.name().as_ref()))?;
            },
            Event::End(element) => {
                validate_namespace_prefix(qname_prefix(element.name().as_ref()))?;
            },
            _ => {},
        }
        let namespace = resolved_namespace(&resolved)?;
        let event_end = position(&reader, origin, bom_len)?;
        if !root_seen && !matches!(&event, Event::Decl(_) | Event::Eof) {
            pre_root_markup = true;
        }
        match event {
            Event::Decl(declaration) => {
                if root_seen || declaration_seen || pre_root_markup {
                    return Err(invalid("Theme XML declaration is misplaced"));
                }
                family_codec::validate_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::Start(element) => {
                bump_nodes(&mut nodes)?;
                let family_extension = direct_recognized_family_extension(
                    &stack,
                    element.local_name().as_ref(),
                    &namespace,
                    &extensions,
                );
                // Only one direct supported family can be unambiguous. Refuse
                // another before cloning its inherited namespace environment,
                // even when it belongs to a different recognized extension.
                if family_extension.is_some() && !families.is_empty() {
                    return Err(invalid(
                        "Theme XML contains multiple direct Theme Family owners",
                    ));
                }
                let need_scope_copy = family_extension
                    .is_some_and(|index| !recognized_family_extensions.contains(&index));
                let scope_len_before = scope.len();
                let scope_bytes_before = scope_bytes;
                let added =
                    apply_namespace_declarations(&element, &mut scope, &mut scope_bytes, &reader)?;
                let inherited = need_scope_copy.then(|| scope[..scope_len_before].to_vec());
                validate_element_name(&element)?;
                validate_element_attributes(&element, &scope, &reader)?;
                let local = element.local_name().as_ref().to_vec();
                let qualified_name = element.name().as_ref().to_vec();
                let depth = stack.len();
                if depth >= MAX_DEPTH {
                    return Err(limit("Theme XML depth", MAX_DEPTH));
                }
                let kind = if !root_seen {
                    if local != b"theme" || !is_drawingml_namespace(&namespace) {
                        return Err(invalid("Theme XML root must be a DrawingML theme"));
                    }
                    root_seen = true;
                    root = Some(Root {
                        namespace: String::from_utf8(namespace.clone()).map_err(xml_error)?,
                        element_prefix: qname_prefix(&qualified_name).to_vec(),
                        qualified_name: qualified_name.clone(),
                        range: event_start..0,
                        end_start: None,
                        empty: false,
                        ext_list: None,
                        admitted_ext: None,
                    });
                    Kind::Root
                } else {
                    if root_closed || stack.is_empty() {
                        return Err(invalid("Theme XML has more than one root element"));
                    }
                    let mut classifier = Classifier {
                        root: &mut root,
                        ext_list_seen: &mut ext_list_seen,
                        extensions: &mut extensions,
                        families: &mut families,
                        recognized_family_extensions: &mut recognized_family_extensions,
                        hidden_owner: &mut hidden_owner,
                    };
                    classify_start(
                        &local,
                        &namespace,
                        &qualified_name,
                        &element,
                        inherited.as_deref().unwrap_or(&scope),
                        depth,
                        &mut classifier,
                        &stack,
                        event_start,
                        event_end,
                    )?
                };
                stack.push(Frame {
                    local,
                    namespace,
                    kind,
                    ns_added: added,
                    scope_bytes_before,
                });
            },
            Event::Empty(element) => {
                bump_nodes(&mut nodes)?;
                let family_extension = direct_recognized_family_extension(
                    &stack,
                    element.local_name().as_ref(),
                    &namespace,
                    &extensions,
                );
                // Only one direct supported family can be unambiguous. Refuse
                // another before cloning its inherited namespace environment,
                // even when it belongs to a different recognized extension.
                if family_extension.is_some() && !families.is_empty() {
                    return Err(invalid(
                        "Theme XML contains multiple direct Theme Family owners",
                    ));
                }
                let need_scope_copy = family_extension
                    .is_some_and(|index| !recognized_family_extensions.contains(&index));
                let scope_len_before = scope.len();
                let scope_bytes_before = scope_bytes;
                let added =
                    apply_namespace_declarations(&element, &mut scope, &mut scope_bytes, &reader)?;
                let inherited = need_scope_copy.then(|| scope[..scope_len_before].to_vec());
                validate_element_name(&element)?;
                validate_element_attributes(&element, &scope, &reader)?;
                let local = element.local_name().as_ref().to_vec();
                let qualified_name = element.name().as_ref().to_vec();
                let depth = stack.len();
                if depth >= MAX_DEPTH {
                    return Err(limit("Theme XML depth", MAX_DEPTH));
                }
                let kind = if !root_seen {
                    if local != b"theme" || !is_drawingml_namespace(&namespace) {
                        return Err(invalid("Theme XML root must be a DrawingML theme"));
                    }
                    root_seen = true;
                    root_closed = true;
                    root = Some(Root {
                        namespace: String::from_utf8(namespace.clone()).map_err(xml_error)?,
                        element_prefix: qname_prefix(&qualified_name).to_vec(),
                        qualified_name: qualified_name.clone(),
                        range: event_start..event_end,
                        end_start: None,
                        empty: true,
                        ext_list: None,
                        admitted_ext: None,
                    });
                    Kind::Root
                } else {
                    if root_closed || stack.is_empty() {
                        return Err(invalid("Theme XML has more than one root element"));
                    }
                    let mut classifier = Classifier {
                        root: &mut root,
                        ext_list_seen: &mut ext_list_seen,
                        extensions: &mut extensions,
                        families: &mut families,
                        recognized_family_extensions: &mut recognized_family_extensions,
                        hidden_owner: &mut hidden_owner,
                    };
                    classify_empty(
                        &local,
                        &namespace,
                        &qualified_name,
                        &element,
                        inherited.as_deref().unwrap_or(&scope),
                        depth,
                        &mut classifier,
                        &stack,
                        event_start,
                        event_end,
                    )?
                };
                if let Kind::Family(index) = kind {
                    families[index].end = Some(event_end);
                }
                if let Kind::Ext(index) = kind {
                    extensions[index].range.end = event_end;
                }
                scope.truncate(scope.len().saturating_sub(added));
                scope_bytes = scope_bytes_before;
            },
            Event::End(element) => {
                if !root_seen || root_closed {
                    return Err(invalid("Theme XML has markup outside its root"));
                }
                let Some(frame) = stack.pop() else {
                    return Err(invalid("Theme XML has an unexpected closing element"));
                };
                validate_end_name(&element)?;
                if frame.local.as_slice() != element.local_name().as_ref()
                    || frame.namespace != namespace
                {
                    return Err(invalid("Theme XML closing element does not match"));
                }
                match frame.kind {
                    Kind::Root => {
                        root_closed = true;
                        let root = root
                            .as_mut()
                            .ok_or_else(|| invalid("Theme root state is missing"))?;
                        root.end_start = Some(event_start);
                        root.range.end = event_end;
                    },
                    Kind::ExtList => {
                        let root = root
                            .as_mut()
                            .ok_or_else(|| invalid("Theme root state is missing"))?;
                        if let Some(container) = root.ext_list.as_mut() {
                            container.end_start = Some(event_start);
                            container.range.end = event_end;
                        }
                    },
                    Kind::Ext(index) => {
                        extensions[index].end_start = Some(event_start);
                        extensions[index].range.end = event_end;
                    },
                    Kind::Family(index) => families[index].end = Some(event_end),
                    Kind::Other
                    | Kind::MceOther
                    | Kind::MceChoice
                    | Kind::MceFallback
                    | Kind::HiddenExtList
                    | Kind::HiddenExt(_) => {},
                }
                scope.truncate(scope.len().saturating_sub(frame.ns_added));
                scope_bytes = frame.scope_bytes_before;
            },
            Event::Text(text) => {
                family_codec::validate_event_text(text.as_ref(), "Theme XML text")?;
                validate_raw_theme_text(text.as_ref())?;
                if element_only_context(&stack, &extensions) && !is_xml_whitespace(text.as_ref()) {
                    return Err(invalid(
                        "owned Theme container contains non-whitespace text",
                    ));
                }
                if (!root_seen || root_closed) && !is_xml_whitespace(text.as_ref()) {
                    return Err(invalid("Theme XML has text outside its root"));
                }
            },
            Event::CData(text) => {
                family_codec::validate_event_text(text.as_ref(), "Theme XML CDATA")?;
                if element_only_context(&stack, &extensions) && !is_xml_whitespace(text.as_ref()) {
                    return Err(invalid(
                        "owned Theme container contains non-whitespace CDATA",
                    ));
                }
                if !root_seen || root_closed {
                    return Err(invalid("Theme XML has CDATA outside its root"));
                }
            },
            Event::GeneralRef(reference) => {
                if !root_seen || root_closed {
                    return Err(invalid("Theme XML has a reference outside its root"));
                }
                let character = family_codec::validate_general_ref(&reference)?;
                if element_only_context(&stack, &extensions)
                    && !character.is_some_and(is_xml_whitespace_character)
                {
                    return Err(invalid(
                        "owned Theme container contains a non-whitespace reference",
                    ));
                }
            },
            Event::Comment(comment) => {
                family_codec::validate_event_text(comment.as_ref(), "Theme XML comment")?;
            },
            Event::PI(_) | Event::DocType(_) => {
                return Err(invalid("Theme XML contains forbidden document markup"));
            },
            Event::Eof => break,
        }
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("Theme XML is unterminated"));
    }
    let mut root = root.ok_or_else(|| invalid("Theme XML has no root"))?;
    let admitted = extensions
        .iter()
        .enumerate()
        .filter_map(|(index, extension)| {
            extension
                .uri
                .as_deref()
                .and_then(ExtensionProfile::from_uri)
                .map(|_| (index, extension))
        })
        .collect::<Vec<_>>();
    if admitted.len() > 1 {
        return Err(invalid(
            "Theme XML contains multiple recognized Theme Family extensions",
        ));
    }
    let mut candidate = None;
    if let Some((ext_index, extension)) = admitted.first() {
        root.admitted_ext = Some(Extension {
            range: extension.range.clone(),
            end_start: extension.end_start,
            qualified_name: extension.qualified_name.clone(),
            empty: extension.empty,
        });
        let direct = families
            .iter()
            .filter(|family| family.extension_index == *ext_index)
            .collect::<Vec<_>>();
        if direct.len() > 1 {
            return Err(invalid(
                "Theme XML contains duplicate direct Theme Family owners",
            ));
        }
        if let Some(family) = direct.first() {
            let end = family
                .end
                .ok_or_else(|| invalid("Theme Family is unterminated"))?;
            candidate = Some(Candidate {
                family_range: family.start..end,
                uri: extension.uri.clone().expect("admitted extension has URI"),
                inherited_bindings: effective_inherited_bindings(family),
            });
        }
    }
    Ok(Scanned {
        root,
        candidate,
        hidden_owner,
    })
}

fn validate_direct_list_child(
    stack: &[Frame],
    local: &[u8],
    namespace: &[u8],
    expected: &[u8],
) -> Result<()> {
    if matches!(stack.last().map(|frame| frame.kind), Some(Kind::ExtList))
        && !(local == b"ext" && namespace == expected)
        // MCE envelopes retain the existing opaque-read/ambiguous-edit policy.
        && !(local == b"AlternateContent" && namespace == MCE_NAMESPACE.as_bytes())
    {
        return Err(invalid(
            "direct Theme extLst contains a child other than ext or MCE AlternateContent",
        ));
    }
    Ok(())
}

fn classify_start(
    local: &[u8],
    namespace: &[u8],
    qualified_name: &[u8],
    element: &BytesStart<'_>,
    inherited: &[Binding],
    depth: usize,
    classifier: &mut Classifier<'_>,
    stack: &[Frame],
    start: usize,
    start_end: usize,
) -> Result<Kind> {
    let expected = root_namespace(classifier.root)?;
    validate_direct_list_child(stack, local, namespace, expected)?;
    let in_mce_branch = stack.iter().any(|frame| {
        matches!(
            frame.kind,
            Kind::MceOther | Kind::MceChoice | Kind::MceFallback
        )
    });
    let direct_recognized_extension = stack
        .last()
        .is_some_and(|frame| recognized_extension_kind(&frame.kind, classifier.extensions));
    let recognized_extension_ancestor = stack
        .iter()
        .any(|frame| recognized_extension_kind(&frame.kind, classifier.extensions));
    let mce_ownership_context = recognized_extension_ancestor || mce_container_context(stack);
    if in_mce_branch && mce_ownership_context && mce_owner_container(local, namespace, element)? {
        *classifier.hidden_owner = true;
    }
    if namespace == MCE_NAMESPACE.as_bytes() && local == b"AlternateContent" {
        if direct_recognized_extension {
            *classifier.hidden_owner = true;
        }
        return Ok(Kind::MceOther);
    }
    if matches!(stack.last().map(|frame| frame.kind), Some(Kind::MceOther))
        && namespace == MCE_NAMESPACE.as_bytes()
        && local == b"Choice"
    {
        return Ok(Kind::MceChoice);
    }
    if matches!(stack.last().map(|frame| frame.kind), Some(Kind::MceOther))
        && namespace == MCE_NAMESPACE.as_bytes()
        && local == b"Fallback"
    {
        return Ok(Kind::MceFallback);
    }
    if local == b"themeFamily"
        && namespace == FAMILY_NAMESPACE.as_bytes()
        && (in_mce_branch || !direct_recognized_extension)
    {
        if (in_mce_branch && mce_ownership_context) || recognized_extension_ancestor {
            *classifier.hidden_owner = true;
        }
        if !direct_recognized_extension {
            return Ok(Kind::Other);
        }
    }
    if local == b"themeFamily" && recognized_extension_ancestor && !direct_recognized_extension {
        *classifier.hidden_owner = true;
        return Ok(Kind::Other);
    }
    if depth == 1 && local == b"extLst" && namespace == expected {
        if *classifier.ext_list_seen {
            return Err(invalid(
                "Theme XML contains duplicate direct extLst elements",
            ));
        }
        *classifier.ext_list_seen = true;
        classifier
            .root
            .as_mut()
            .ok_or_else(|| invalid("Theme root state is missing"))?
            .ext_list = Some(Container {
            range: start..start_end,
            end_start: None,
            qualified_name: qualified_name.to_vec(),
            empty: false,
        });
        return Ok(Kind::ExtList);
    }
    if matches!(stack.last().map(|frame| frame.kind), Some(Kind::ExtList))
        && local == b"ext"
        && namespace == expected
    {
        let index = classifier.extensions.len();
        classifier.extensions.push(ExtensionRecord {
            range: start..start_end,
            end_start: None,
            qualified_name: qualified_name.to_vec(),
            uri: extension_uri(element)?,
            empty: false,
        });
        return Ok(Kind::Ext(index));
    }
    if let Some(Kind::Ext(extension_index)) = stack.last().map(|frame| frame.kind)
        && local == b"themeFamily"
        && namespace == FAMILY_NAMESPACE.as_bytes()
    {
        if !recognized_extension_kind(&Kind::Ext(extension_index), classifier.extensions) {
            return Ok(Kind::Other);
        }
        if !classifier
            .recognized_family_extensions
            .insert(extension_index)
        {
            return Err(invalid(
                "Theme XML contains duplicate direct Theme Family owners",
            ));
        }
        let index = classifier.families.len();
        classifier.families.push(FamilyRecord {
            extension_index,
            start,
            end: None,
            inherited_scope: inherited.to_vec(),
            local_declarations: namespace_declaration_names(element),
        });
        return Ok(Kind::Family(index));
    }
    let recognized_extension = match stack.last().map(|frame| frame.kind) {
        Some(Kind::Ext(index)) => classifier
            .extensions
            .get(index)
            .and_then(|extension| extension.uri.as_deref())
            .and_then(ExtensionProfile::from_uri)
            .is_some(),
        _ => false,
    };
    if recognized_extension && local == b"themeFamily" {
        return Err(invalid(
            "recognized Theme extension contains a Theme Family lookalike in another namespace",
        ));
    }
    if in_mce_branch
        && mce_ownership_context
        && local == b"extLst"
        && is_drawingml_namespace(namespace)
    {
        return Ok(Kind::HiddenExtList);
    }
    if matches!(
        stack.last().map(|frame| frame.kind),
        Some(Kind::HiddenExtList)
    ) && local == b"ext"
        && is_drawingml_namespace(namespace)
    {
        let recognized = extension_uri(element)?
            .as_deref()
            .and_then(ExtensionProfile::from_uri)
            .is_some();
        return Ok(Kind::HiddenExt(recognized));
    }
    if matches!(
        stack.last().map(|frame| frame.kind),
        Some(Kind::HiddenExt(true))
    ) && local == b"themeFamily"
        && namespace == FAMILY_NAMESPACE.as_bytes()
    {
        *classifier.hidden_owner = true;
    }
    Ok(Kind::Other)
}

fn classify_empty(
    local: &[u8],
    namespace: &[u8],
    qualified_name: &[u8],
    element: &BytesStart<'_>,
    inherited: &[Binding],
    depth: usize,
    classifier: &mut Classifier<'_>,
    stack: &[Frame],
    start: usize,
    end: usize,
) -> Result<Kind> {
    let expected = root_namespace(classifier.root)?;
    validate_direct_list_child(stack, local, namespace, expected)?;
    let in_mce_branch = stack.iter().any(|frame| {
        matches!(
            frame.kind,
            Kind::MceOther | Kind::MceChoice | Kind::MceFallback
        )
    });
    let direct_recognized_extension = stack
        .last()
        .is_some_and(|frame| recognized_extension_kind(&frame.kind, classifier.extensions));
    let recognized_extension_ancestor = stack
        .iter()
        .any(|frame| recognized_extension_kind(&frame.kind, classifier.extensions));
    let mce_ownership_context = recognized_extension_ancestor || mce_container_context(stack);
    if in_mce_branch && mce_ownership_context && mce_owner_container(local, namespace, element)? {
        *classifier.hidden_owner = true;
    }
    if namespace == MCE_NAMESPACE.as_bytes() && local == b"AlternateContent" {
        if direct_recognized_extension {
            *classifier.hidden_owner = true;
        }
        return Ok(Kind::MceOther);
    }
    if matches!(stack.last().map(|frame| frame.kind), Some(Kind::MceOther))
        && namespace == MCE_NAMESPACE.as_bytes()
        && local == b"Choice"
    {
        return Ok(Kind::MceChoice);
    }
    if matches!(stack.last().map(|frame| frame.kind), Some(Kind::MceOther))
        && namespace == MCE_NAMESPACE.as_bytes()
        && local == b"Fallback"
    {
        return Ok(Kind::MceFallback);
    }
    if local == b"themeFamily"
        && namespace == FAMILY_NAMESPACE.as_bytes()
        && (in_mce_branch || !direct_recognized_extension)
    {
        if (in_mce_branch && mce_ownership_context) || recognized_extension_ancestor {
            *classifier.hidden_owner = true;
        }
        if !direct_recognized_extension {
            return Ok(Kind::Other);
        }
    }
    if local == b"themeFamily" && recognized_extension_ancestor && !direct_recognized_extension {
        *classifier.hidden_owner = true;
        return Ok(Kind::Other);
    }
    if depth == 1 && local == b"extLst" && namespace == expected {
        if *classifier.ext_list_seen {
            return Err(invalid(
                "Theme XML contains duplicate direct extLst elements",
            ));
        }
        *classifier.ext_list_seen = true;
        classifier
            .root
            .as_mut()
            .ok_or_else(|| invalid("Theme root state is missing"))?
            .ext_list = Some(Container {
            range: start..end,
            end_start: None,
            qualified_name: qualified_name.to_vec(),
            empty: true,
        });
        return Ok(Kind::ExtList);
    }
    if matches!(stack.last().map(|frame| frame.kind), Some(Kind::ExtList))
        && local == b"ext"
        && namespace == expected
    {
        let index = classifier.extensions.len();
        classifier.extensions.push(ExtensionRecord {
            range: start..end,
            end_start: None,
            qualified_name: qualified_name.to_vec(),
            uri: extension_uri(element)?,
            empty: true,
        });
        return Ok(Kind::Ext(index));
    }
    if let Some(Kind::Ext(extension_index)) = stack.last().map(|frame| frame.kind)
        && local == b"themeFamily"
        && namespace == FAMILY_NAMESPACE.as_bytes()
    {
        if !recognized_extension_kind(&Kind::Ext(extension_index), classifier.extensions) {
            return Ok(Kind::Other);
        }
        if !classifier
            .recognized_family_extensions
            .insert(extension_index)
        {
            return Err(invalid(
                "Theme XML contains duplicate direct Theme Family owners",
            ));
        }
        let index = classifier.families.len();
        classifier.families.push(FamilyRecord {
            extension_index,
            start,
            end: Some(end),
            inherited_scope: inherited.to_vec(),
            local_declarations: namespace_declaration_names(element),
        });
        return Ok(Kind::Family(index));
    }
    let recognized_extension = match stack.last().map(|frame| frame.kind) {
        Some(Kind::Ext(index)) => classifier
            .extensions
            .get(index)
            .and_then(|extension| extension.uri.as_deref())
            .and_then(ExtensionProfile::from_uri)
            .is_some(),
        _ => false,
    };
    if recognized_extension && local == b"themeFamily" {
        return Err(invalid(
            "recognized Theme extension contains a Theme Family lookalike in another namespace",
        ));
    }
    if in_mce_branch
        && mce_ownership_context
        && local == b"extLst"
        && is_drawingml_namespace(namespace)
    {
        return Ok(Kind::HiddenExtList);
    }
    if matches!(
        stack.last().map(|frame| frame.kind),
        Some(Kind::HiddenExtList)
    ) && local == b"ext"
        && is_drawingml_namespace(namespace)
    {
        let recognized = extension_uri(element)?
            .as_deref()
            .and_then(ExtensionProfile::from_uri)
            .is_some();
        return Ok(Kind::HiddenExt(recognized));
    }
    if matches!(
        stack.last().map(|frame| frame.kind),
        Some(Kind::HiddenExt(true))
    ) && local == b"themeFamily"
        && namespace == FAMILY_NAMESPACE.as_bytes()
    {
        *classifier.hidden_owner = true;
    }
    Ok(Kind::Other)
}

fn mce_container_context(stack: &[Frame]) -> bool {
    // MCE can wrap the root-owned list/extension path. An intervening foreign
    // element or unrecognized extension keeps its subtree opaque.
    stack.iter().all(|frame| {
        matches!(
            frame.kind,
            Kind::Root
                | Kind::ExtList
                | Kind::HiddenExtList
                | Kind::MceOther
                | Kind::MceChoice
                | Kind::MceFallback
        )
    })
}

fn mce_owner_container(local: &[u8], namespace: &[u8], element: &BytesStart<'_>) -> Result<bool> {
    if !is_drawingml_namespace(namespace) {
        return Ok(false);
    }
    // Even an empty effective list/recognized extension can conflict with a
    // second direct owner after insertion. Do not wait for a family child.
    Ok(local == b"extLst"
        || (local == b"ext"
            && extension_uri(element)?
                .as_deref()
                .and_then(ExtensionProfile::from_uri)
                .is_some()))
}

fn root_namespace(root: &Option<Root>) -> Result<&[u8]> {
    root.as_ref()
        .map(|root| root.namespace.as_bytes())
        .ok_or_else(|| invalid("Theme root state is missing"))
}

fn recognized_extension_kind(kind: &Kind, extensions: &[ExtensionRecord]) -> bool {
    match kind {
        Kind::Ext(index) => extensions
            .get(*index)
            .and_then(|extension| extension.uri.as_deref())
            .and_then(ExtensionProfile::from_uri)
            .is_some(),
        Kind::HiddenExt(true) => true,
        _ => false,
    }
}

fn direct_recognized_family_extension(
    stack: &[Frame],
    local: &[u8],
    namespace: &[u8],
    extensions: &[ExtensionRecord],
) -> Option<usize> {
    if local != b"themeFamily" || namespace != FAMILY_NAMESPACE.as_bytes() {
        return None;
    }
    match stack.last().map(|frame| frame.kind) {
        Some(Kind::Ext(index)) if recognized_extension_kind(&Kind::Ext(index), extensions) => {
            Some(index)
        },
        _ => None,
    }
}

fn validate_element_name(element: &BytesStart<'_>) -> Result<()> {
    let qualified_name = element.name();
    let name = std::str::from_utf8(qualified_name.as_ref()).map_err(xml_error)?;
    if !is_qualified_name(name) {
        return Err(invalid("Theme XML element name is invalid"));
    }
    Ok(())
}

fn validate_end_name(element: &quick_xml::events::BytesEnd<'_>) -> Result<()> {
    let qualified_name = element.name();
    let name = std::str::from_utf8(qualified_name.as_ref()).map_err(xml_error)?;
    if !is_qualified_name(name) {
        return Err(invalid("Theme XML closing element name is invalid"));
    }
    Ok(())
}

fn extension_uri(element: &BytesStart<'_>) -> Result<Option<String>> {
    let mut uri = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"uri" {
            if uri.is_some() {
                return Err(invalid("Theme extension has duplicate uri attributes"));
            }
            uri = Some(normalize_token(
                &quick_xml::escape::unescape(
                    std::str::from_utf8(attribute.value.as_ref()).map_err(xml_error)?,
                )
                .map_err(xml_error)?,
            ));
        }
    }
    Ok(uri)
}

fn normalize_token(value: &str) -> String {
    let mut normalized = String::with_capacity(value.len());
    for token in value.split([' ', '\t', '\r', '\n']) {
        if token.is_empty() {
            continue;
        }
        if !normalized.is_empty() {
            normalized.push(' ');
        }
        normalized.push_str(token);
    }
    normalized
}

fn effective_inherited_bindings(record: &FamilyRecord) -> Vec<Binding> {
    // Keep every active inherited binding (apart from the built-in `xml`
    // binding), because opaque attribute values may contain QNames that are
    // not recoverable from their lexical value. Collapse shadowed ancestors
    // to the one binding that is actually in scope and omit prefixes the
    // family root redeclares locally.
    let local = record
        .local_declarations
        .iter()
        .map(Vec::as_slice)
        .collect::<HashSet<_>>();
    let mut seen = HashSet::<&[u8]>::new();
    let mut result = Vec::new();
    for binding in record.inherited_scope.iter().rev() {
        if binding.prefix == b"xml"
            || local.contains(binding.prefix.as_slice())
            || !seen.insert(binding.prefix.as_slice())
        {
            continue;
        }
        result.push(binding.clone());
    }
    result.reverse();
    result
}

fn decorate_family(source: &[u8], candidate: &Candidate) -> Result<Vec<u8>> {
    let fragment = source
        .get(candidate.family_range.clone())
        .ok_or_else(|| invalid("Theme Family source range is invalid"))?;
    if fragment.len() > super::MAX_XML_BYTES {
        return Err(limit(
            "decorated Theme Family XML bytes",
            super::MAX_XML_BYTES,
        ));
    }
    if candidate.inherited_bindings.is_empty() {
        let mut output = family_buffer(fragment.len())?;
        output.extend_from_slice(fragment);
        return Ok(output);
    }
    let tag_end = open_tag_end(fragment)?;
    let before_close = tag_end.saturating_sub(1);
    let insertion = if before_close > 0 && fragment[before_close - 1] == b'/' {
        before_close - 1
    } else {
        before_close
    };
    let mut declaration_bytes = 0usize;
    for binding in &candidate.inherited_bindings {
        let prefix = std::str::from_utf8(&binding.prefix).map_err(xml_error)?;
        let uri = std::str::from_utf8(&binding.uri).map_err(xml_error)?;
        let escaped_uri = escape_xml(uri);
        let qualified_prefix_bytes = if prefix.is_empty() {
            0
        } else {
            1usize
                .checked_add(prefix.len())
                .ok_or_else(|| invalid("Theme Family namespace declaration length overflows"))?
        };
        let one = 1usize
            .checked_add(5)
            .and_then(|value| value.checked_add(qualified_prefix_bytes))
            .and_then(|value| value.checked_add(3))
            .and_then(|value| value.checked_add(escaped_uri.len()))
            .ok_or_else(|| invalid("Theme Family namespace declaration length overflows"))?;
        declaration_bytes = declaration_bytes
            .checked_add(one)
            .ok_or_else(|| invalid("Theme Family namespace declaration length overflows"))?;
    }
    let output_len = fragment
        .len()
        .checked_add(declaration_bytes)
        .ok_or_else(|| invalid("decorated Theme Family XML length overflows"))?;
    if output_len > super::MAX_XML_BYTES {
        return Err(limit(
            "decorated Theme Family XML bytes",
            super::MAX_XML_BYTES,
        ));
    }
    let mut declarations = String::new();
    declarations
        .try_reserve_exact(declaration_bytes)
        .map_err(|_| invalid("Theme Family namespace allocation failed"))?;
    for binding in &candidate.inherited_bindings {
        declarations.push(' ');
        declarations.push_str("xmlns");
        if !binding.prefix.is_empty() {
            declarations.push(':');
            declarations.push_str(std::str::from_utf8(&binding.prefix).map_err(xml_error)?);
        }
        declarations.push_str("=\"");
        declarations.push_str(&escape_xml(
            std::str::from_utf8(&binding.uri).map_err(xml_error)?,
        ));
        declarations.push('"');
    }
    let mut output = family_buffer(output_len)?;
    output.extend_from_slice(&fragment[..insertion]);
    output.extend_from_slice(declarations.as_bytes());
    output.extend_from_slice(&fragment[insertion..]);
    Ok(output)
}

fn family_child_bytes(family: &Family, inherited: &[Binding]) -> Result<Vec<u8>> {
    let mut bytes = family_codec::write(family)?;
    // A UTF-8 BOM frames a standalone XML document; it is not character data
    // that can be embedded between extension elements. Strip only the exact
    // leading BOM after the fragment codec has validated the source.
    if bytes.starts_with(b"\xEF\xBB\xBF") {
        bytes.drain(..3);
    }
    if has_xml_declaration(&bytes) {
        return Err(invalid("Theme Family child contains an XML declaration"));
    }
    remove_inherited_namespace_declarations(bytes, inherited)
}

fn has_xml_declaration(bytes: &[u8]) -> bool {
    let offset = usize::from(bytes.starts_with(b"\xEF\xBB\xBF")) * 3;
    bytes
        .get(offset..)
        .is_some_and(|value| value.starts_with(b"<?xml"))
}

fn inherited_namespace_removal_ranges(
    bytes: &[u8],
    inherited: &[Binding],
) -> Result<Vec<Range<usize>>> {
    let root_range = family_root_open_range(bytes)?;
    let root = bytes
        .get(root_range.clone())
        .ok_or_else(|| invalid("Theme Family root range is invalid"))?;
    let expected = inherited
        .iter()
        .map(|binding| (binding.prefix.as_slice(), binding.uri.as_slice()))
        .collect::<HashMap<_, _>>();
    let mut ranges = Vec::new();
    let mut cursor = 1usize;
    while cursor < root.len()
        && !is_xml_space(root[cursor])
        && root[cursor] != b'/'
        && root[cursor] != b'>'
    {
        cursor += 1;
    }
    while cursor < root.len() {
        while cursor < root.len() && is_xml_space(root[cursor]) {
            cursor += 1;
        }
        if cursor >= root.len() || root[cursor] == b'>' || root[cursor] == b'/' {
            break;
        }
        let name_start = cursor;
        while cursor < root.len()
            && !is_xml_space(root[cursor])
            && root[cursor] != b'='
            && root[cursor] != b'>'
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < root.len() && is_xml_space(root[cursor]) {
            cursor += 1;
        }
        if root.get(cursor) != Some(&b'=') {
            return Err(invalid("Theme Family root attribute lacks equals"));
        }
        cursor += 1;
        while cursor < root.len() && is_xml_space(root[cursor]) {
            cursor += 1;
        }
        let quote = *root
            .get(cursor)
            .ok_or_else(|| invalid("Theme Family root attribute value is missing"))?;
        if quote != b'\'' && quote != b'"' {
            return Err(invalid("Theme Family root attribute value is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < root.len() && root[cursor] != quote {
            cursor += 1;
        }
        if cursor >= root.len() {
            return Err(invalid("Theme Family root attribute value is unterminated"));
        }
        let value_end = cursor;
        if let Some(prefix) = namespace_declaration(&root[name_start..name_end])
            && let Some(uri) = expected.get(prefix)
        {
            let uri = std::str::from_utf8(uri).map_err(xml_error)?;
            let escaped = escape_xml(uri);
            if root[value_start..value_end] == *uri.as_bytes()
                || root[value_start..value_end] == *escaped.as_bytes()
            {
                let mut range_start = name_start;
                if range_start > 0 && is_xml_space(root[range_start - 1]) {
                    range_start -= 1;
                }
                ranges.push(root_range.start + range_start..root_range.start + value_end + 1);
            }
        }
        cursor += 1;
    }
    Ok(ranges)
}

fn family_child_len(family: &Family, inherited: &[Binding], max_xml_bytes: usize) -> Result<usize> {
    let mut length = family_codec::serialized_len(family)?;
    // The standalone intermediate must also fit the caller's cap.
    if length > max_xml_bytes {
        return Err(limit("patched Theme XML bytes", max_xml_bytes));
    }
    if let Some(source) = family.source_state() {
        if has_xml_declaration(source.xml.as_ref()) {
            return Err(invalid("Theme Family child contains an XML declaration"));
        }
        let bom = usize::from(source.xml.starts_with(b"\xEF\xBB\xBF")) * 3;
        length = length
            .checked_sub(bom)
            .ok_or_else(|| invalid("Theme Family child length underflows"))?;
        if !inherited.is_empty() {
            for range in inherited_namespace_removal_ranges(source.xml.as_ref(), inherited)? {
                length = length
                    .checked_sub(range.len())
                    .ok_or_else(|| invalid("Theme Family child length underflows"))?;
            }
        }
    }
    Ok(length)
}

fn remove_inherited_namespace_declarations(
    bytes: Vec<u8>,
    inherited: &[Binding],
) -> Result<Vec<u8>> {
    if inherited.is_empty() {
        return Ok(bytes);
    }
    let ranges = inherited_namespace_removal_ranges(&bytes, inherited)?;
    if ranges.is_empty() {
        return Ok(bytes);
    }
    let removed = ranges.iter().try_fold(0usize, |total, range| {
        total
            .checked_add(range.len())
            .ok_or_else(|| invalid("Theme Family namespace removal length overflows"))
    })?;
    let mut output = family_buffer(
        bytes
            .len()
            .checked_sub(removed)
            .ok_or_else(|| invalid("Theme Family namespace removal underflows"))?,
    )?;
    let mut previous = 0usize;
    for range in ranges {
        output.extend_from_slice(&bytes[previous..range.start]);
        previous = range.end;
    }
    output.extend_from_slice(&bytes[previous..]);
    Ok(output)
}

fn family_root_open_range(bytes: &[u8]) -> Result<Range<usize>> {
    let bom_len = ReaderOrigin::of(bytes).skipped();
    let parse_bytes = &bytes[bom_len..];
    let origin = ReaderOrigin::of(parse_bytes);
    let mut reader = NsReader::from_reader(parse_bytes);
    reader.config_mut().trim_text(false);
    loop {
        let start = position(&reader, origin, bom_len)?;
        let (_, event) = reader.read_resolved_event().map_err(xml_error)?;
        let end = position(&reader, origin, bom_len)?;
        match event {
            Event::Start(_) | Event::Empty(_) => return Ok(start..end),
            Event::Eof => return Err(invalid("Theme Family root element is missing")),
            _ => {},
        }
    }
}

fn ensure_splice_output_limit(
    source: &[u8],
    range: &Range<usize>,
    replacement_len: usize,
    max_xml_bytes: usize,
) -> Result<()> {
    let output_len = source
        .len()
        .checked_sub(range.len())
        .and_then(|value| value.checked_add(replacement_len))
        .ok_or_else(|| invalid("Theme XML patched length overflows"))?;
    if output_len > max_xml_bytes {
        return Err(limit("patched Theme XML bytes", max_xml_bytes));
    }
    Ok(())
}

fn empty_element_replacement_len(
    slash: usize,
    child_len: usize,
    closing_name_len: usize,
) -> Result<usize> {
    slash
        .checked_add(1)
        .and_then(|value| value.checked_add(child_len))
        .and_then(|value| value.checked_add(2))
        .and_then(|value| value.checked_add(closing_name_len))
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| invalid("Theme XML replacement length overflows"))
}

fn ensure_container_insertion_limit(
    source: &[u8],
    container: &Container,
    payload_len: usize,
    max_xml_bytes: usize,
) -> Result<()> {
    let replacement_len = if container.empty {
        let source_range = source
            .get(container.range.clone())
            .ok_or_else(|| invalid("Theme extLst range is invalid"))?;
        let slash = empty_element_slash(source_range)
            .ok_or_else(|| invalid("Theme empty extLst slash is missing"))?;
        empty_element_replacement_len(slash, payload_len, container.qualified_name.len())?
    } else {
        payload_len
    };
    let range = if container.empty {
        container.range.clone()
    } else {
        let end = container
            .end_start
            .ok_or_else(|| invalid("Theme extLst closing range is missing"))?;
        end..end
    };
    ensure_splice_output_limit(source, &range, replacement_len, max_xml_bytes)
}

fn ensure_root_insertion_limit(
    source: &[u8],
    root: &Root,
    payload_len: usize,
    max_xml_bytes: usize,
) -> Result<()> {
    let replacement_len = if root.empty {
        let source_range = source
            .get(root.range.clone())
            .ok_or_else(|| invalid("Theme root range is invalid"))?;
        let slash = empty_element_slash(source_range)
            .ok_or_else(|| invalid("Theme empty-root slash is missing"))?;
        empty_element_replacement_len(slash, payload_len, root.qualified_name.len())?
    } else {
        payload_len
    };
    let range = if root.empty {
        root.range.clone()
    } else {
        let end = root
            .end_start
            .ok_or_else(|| invalid("Theme root closing range is missing"))?;
        end..end
    };
    ensure_splice_output_limit(source, &range, replacement_len, max_xml_bytes)
}

fn serialized_extension_len(prefix: &[u8], uri: &str, family_len: usize) -> Result<usize> {
    let name_len = qualified_len(prefix, b"ext")?;
    let uri_len = escape_xml(uri).len();
    1usize
        .checked_add(name_len)
        .and_then(|value| value.checked_add(b" uri=\"".len()))
        .and_then(|value| value.checked_add(uri_len))
        .and_then(|value| value.checked_add(2))
        .and_then(|value| value.checked_add(family_len))
        .and_then(|value| value.checked_add(2))
        .and_then(|value| value.checked_add(name_len))
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| invalid("Theme extension length overflows"))
}

fn serialized_ext_list_len(prefix: &[u8], uri: &str, family_len: usize) -> Result<usize> {
    let list_len = qualified_len(prefix, b"extLst")?;
    let extension_len = serialized_extension_len(prefix, uri, family_len)?;
    1usize
        .checked_add(list_len)
        .and_then(|value| value.checked_add(1))
        .and_then(|value| value.checked_add(extension_len))
        .and_then(|value| value.checked_add(2))
        .and_then(|value| value.checked_add(list_len))
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| invalid("Theme extension list length overflows"))
}

fn qualified_len(prefix: &[u8], local: &[u8]) -> Result<usize> {
    prefix
        .len()
        .checked_add(usize::from(!prefix.is_empty()))
        .and_then(|value| value.checked_add(local.len()))
        .ok_or_else(|| invalid("Theme qualified-name length overflows"))
}

fn ensure_extension_insertion_limit(
    source: &[u8],
    extension: &Extension,
    family_len: usize,
    max_xml_bytes: usize,
) -> Result<()> {
    if extension.empty {
        let source_range = source
            .get(extension.range.clone())
            .ok_or_else(|| invalid("Theme extension range is invalid"))?;
        let slash = empty_element_slash(source_range)
            .ok_or_else(|| invalid("Theme empty extension slash is missing"))?;
        let length =
            empty_element_replacement_len(slash, family_len, extension.qualified_name.len())?;
        ensure_splice_output_limit(source, &extension.range, length, max_xml_bytes)
    } else {
        let end = extension
            .end_start
            .ok_or_else(|| invalid("Theme extension closing range is missing"))?;
        ensure_splice_output_limit(source, &(end..end), family_len, max_xml_bytes)
    }
}

fn insert_into_ext(
    source: &[u8],
    extension: &Extension,
    family: &[u8],
    max_xml_bytes: usize,
) -> Result<Vec<u8>> {
    if extension.empty {
        let source_range = source
            .get(extension.range.clone())
            .ok_or_else(|| invalid("Theme extension range is invalid"))?;
        let slash = empty_element_slash(source_range)
            .ok_or_else(|| invalid("Theme empty extension slash is missing"))?;
        let replacement_len =
            empty_element_replacement_len(slash, family.len(), extension.qualified_name.len())?;
        ensure_splice_output_limit(source, &extension.range, replacement_len, max_xml_bytes)?;
        let mut replacement = theme_buffer(replacement_len)?;
        replacement.extend_from_slice(&source_range[..slash]);
        replacement.push(b'>');
        replacement.extend_from_slice(family);
        replacement.extend_from_slice(b"</");
        replacement.extend_from_slice(&extension.qualified_name);
        replacement.push(b'>');
        splice_with_limit(
            source,
            &[Replacement {
                range: extension.range.clone(),
                bytes: replacement,
            }],
            max_xml_bytes,
        )
    } else {
        let end = extension
            .end_start
            .ok_or_else(|| invalid("Theme extension closing range is missing"))?;
        ensure_splice_output_limit(source, &(end..end), family.len(), max_xml_bytes)?;
        splice_with_limit(
            source,
            &[Replacement {
                range: end..end,
                bytes: family.to_vec(),
            }],
            max_xml_bytes,
        )
    }
}

fn insert_into_container(
    source: &[u8],
    container: &Container,
    extension: &[u8],
    max_xml_bytes: usize,
) -> Result<Vec<u8>> {
    if container.empty {
        let source_range = source
            .get(container.range.clone())
            .ok_or_else(|| invalid("Theme extLst range is invalid"))?;
        let slash = empty_element_slash(source_range)
            .ok_or_else(|| invalid("Theme empty extLst slash is missing"))?;
        let replacement_len =
            empty_element_replacement_len(slash, extension.len(), container.qualified_name.len())?;
        ensure_splice_output_limit(source, &container.range, replacement_len, max_xml_bytes)?;
        let mut replacement = theme_buffer(replacement_len)?;
        replacement.extend_from_slice(&source_range[..slash]);
        replacement.push(b'>');
        replacement.extend_from_slice(extension);
        replacement.extend_from_slice(b"</");
        replacement.extend_from_slice(&container.qualified_name);
        replacement.push(b'>');
        splice_with_limit(
            source,
            &[Replacement {
                range: container.range.clone(),
                bytes: replacement,
            }],
            max_xml_bytes,
        )
    } else {
        let end = container
            .end_start
            .ok_or_else(|| invalid("Theme extLst closing range is missing"))?;
        ensure_splice_output_limit(source, &(end..end), extension.len(), max_xml_bytes)?;
        splice_with_limit(
            source,
            &[Replacement {
                range: end..end,
                bytes: extension.to_vec(),
            }],
            max_xml_bytes,
        )
    }
}

fn insert_before_root_close(
    source: &[u8],
    root: &Root,
    extension: &[u8],
    max_xml_bytes: usize,
) -> Result<Vec<u8>> {
    if root.empty {
        let source_range = source
            .get(root.range.clone())
            .ok_or_else(|| invalid("Theme root range is invalid"))?;
        let slash = empty_element_slash(source_range)
            .ok_or_else(|| invalid("Theme empty-root slash is missing"))?;
        let replacement_len =
            empty_element_replacement_len(slash, extension.len(), root.qualified_name.len())?;
        ensure_splice_output_limit(source, &root.range, replacement_len, max_xml_bytes)?;
        let mut replacement = theme_buffer(replacement_len)?;
        replacement.extend_from_slice(&source_range[..slash]);
        replacement.push(b'>');
        replacement.extend_from_slice(extension);
        replacement.extend_from_slice(b"</");
        replacement.extend_from_slice(&root.qualified_name);
        replacement.push(b'>');
        splice_with_limit(
            source,
            &[Replacement {
                range: root.range.clone(),
                bytes: replacement,
            }],
            max_xml_bytes,
        )
    } else {
        let end = root
            .end_start
            .ok_or_else(|| invalid("Theme root closing range is missing"))?;
        ensure_splice_output_limit(source, &(end..end), extension.len(), max_xml_bytes)?;
        splice_with_limit(
            source,
            &[Replacement {
                range: end..end,
                bytes: extension.to_vec(),
            }],
            max_xml_bytes,
        )
    }
}

fn make_extension(prefix: &[u8], uri: &str, family: &[u8]) -> Result<Vec<u8>> {
    let mut output = theme_buffer(serialized_extension_len(prefix, uri, family.len())?)?;
    let name = qualified(prefix, b"ext");
    output.extend_from_slice(b"<");
    output.extend_from_slice(&name);
    output.extend_from_slice(b" uri=\"");
    output.extend_from_slice(escape_xml(uri).as_bytes());
    output.extend_from_slice(b"\">");
    output.extend_from_slice(family);
    output.extend_from_slice(b"</");
    output.extend_from_slice(&name);
    output.push(b'>');
    Ok(output)
}

fn make_ext_list(prefix: &[u8], uri: &str, family: &[u8]) -> Result<Vec<u8>> {
    let mut output = theme_buffer(serialized_ext_list_len(prefix, uri, family.len())?)?;
    let list = qualified(prefix, b"extLst");
    let extension = make_extension(prefix, uri, family)?;
    output.extend_from_slice(b"<");
    output.extend_from_slice(&list);
    output.push(b'>');
    output.extend_from_slice(&extension);
    output.extend_from_slice(b"</");
    output.extend_from_slice(&list);
    output.push(b'>');
    Ok(output)
}

fn qualified(prefix: &[u8], local: &[u8]) -> Vec<u8> {
    if prefix.is_empty() {
        local.to_vec()
    } else {
        let mut value = prefix.to_vec();
        value.push(b':');
        value.extend_from_slice(local);
        value
    }
}

#[derive(Debug)]
struct Replacement {
    range: Range<usize>,
    bytes: Vec<u8>,
}

fn splice_with_limit(
    source: &[u8],
    replacements: &[Replacement],
    max_xml_bytes: usize,
) -> Result<Vec<u8>> {
    let mut replacements = replacements.iter().collect::<Vec<_>>();
    replacements.sort_by_key(|replacement| replacement.range.start);
    let mut output_len = source.len();
    let mut previous = 0usize;
    for replacement in &replacements {
        if replacement.range.start < previous
            || replacement.range.end < replacement.range.start
            || replacement.range.end > source.len()
        {
            return Err(invalid("Theme XML replacement range is invalid"));
        }
        output_len = output_len
            .checked_sub(replacement.range.len())
            .and_then(|value| value.checked_add(replacement.bytes.len()))
            .ok_or_else(|| invalid("Theme XML patched length overflows"))?;
        previous = replacement.range.end;
    }
    if output_len > max_xml_bytes {
        return Err(limit("patched Theme XML bytes", max_xml_bytes));
    }
    let mut output = theme_buffer(output_len)?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&source[cursor..replacement.range.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.range.end;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

fn validate_result(xml: &[u8]) -> Result<Scanned> {
    let scanned = scan(xml)?;
    let _ = parse_family(xml, &scanned)?;
    Ok(scanned)
}

fn validate_output_limit(max_xml_bytes: usize) -> Result<()> {
    if max_xml_bytes == 0 || max_xml_bytes > MAX_XML_BYTES {
        return Err(invalid(format!(
            "Theme XML output limit must be between 1 and {MAX_XML_BYTES}"
        )));
    }
    Ok(())
}

fn validate_namespace_prefix(prefix: &[u8]) -> Result<()> {
    if prefix.len() > MAX_NAMESPACE_BYTES {
        return Err(limit(
            "Theme XML namespace prefix bytes",
            MAX_NAMESPACE_BYTES,
        ));
    }
    Ok(())
}

fn theme_buffer(length: usize) -> Result<Vec<u8>> {
    if length > MAX_XML_BYTES {
        return Err(limit("Theme XML bytes", MAX_XML_BYTES));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|_| invalid("Theme XML output allocation failed"))?;
    Ok(output)
}

fn family_buffer(length: usize) -> Result<Vec<u8>> {
    if length > super::MAX_XML_BYTES {
        return Err(limit("Theme Family XML bytes", super::MAX_XML_BYTES));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|_| invalid("Theme Family output allocation failed"))?;
    Ok(output)
}

fn apply_namespace_declarations(
    element: &BytesStart<'_>,
    scope: &mut Vec<Binding>,
    scope_bytes: &mut usize,
    reader: &NsReader<&[u8]>,
) -> Result<usize> {
    let mut added = 0usize;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let raw = attribute.key.as_ref();
        let Some(prefix) = namespace_declaration(raw) else {
            continue;
        };
        validate_namespace_prefix(prefix)?;
        if scope
            .iter()
            .rev()
            .take(added)
            .any(|binding| binding.prefix == prefix)
        {
            return Err(invalid("Theme XML has duplicate namespace declarations"));
        }
        let value = decode_attribute(&attribute, reader)?;
        family_codec::validate_namespace_binding(prefix, &value)?;
        if value.len() > MAX_NAMESPACE_BYTES {
            return Err(limit("Theme XML namespace bytes", MAX_NAMESPACE_BYTES));
        }
        let binding_bytes = prefix
            .len()
            .checked_add(value.len())
            .ok_or_else(|| invalid("Theme XML namespace scope length overflows"))?;
        if scope.len() >= MAX_ACTIVE_NAMESPACE_BINDINGS {
            return Err(limit(
                "Theme XML active namespace bindings",
                MAX_ACTIVE_NAMESPACE_BINDINGS,
            ));
        }
        *scope_bytes = scope_bytes
            .checked_add(binding_bytes)
            .ok_or_else(|| invalid("Theme XML namespace scope length overflows"))?;
        if *scope_bytes > MAX_XML_BYTES {
            return Err(limit("Theme XML namespace scope bytes", MAX_XML_BYTES));
        }
        scope.push(Binding {
            prefix: prefix.to_vec(),
            uri: value.into_bytes(),
        });
        added += 1;
    }
    Ok(added)
}

fn validate_element_attributes(
    element: &BytesStart<'_>,
    scope: &[Binding],
    reader: &NsReader<&[u8]>,
) -> Result<()> {
    let mut count = 0usize;
    let mut seen = Vec::<(Vec<u8>, Vec<u8>)>::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        count += 1;
        if count > MAX_ATTRIBUTES {
            return Err(limit("Theme XML attributes", MAX_ATTRIBUTES));
        }
        let raw = attribute.key.as_ref();
        validate_namespace_prefix(namespace_declaration(raw).unwrap_or_else(|| qname_prefix(raw)))?;
        if !is_qualified_name(std::str::from_utf8(raw).map_err(xml_error)?) {
            return Err(invalid("Theme XML attribute name is invalid"));
        }
        let value = decode_attribute(&attribute, reader)?;
        if value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(limit(
                "Theme XML attribute value bytes",
                MAX_ATTRIBUTE_VALUE_BYTES,
            ));
        }
        if namespace_declaration(raw).is_some() {
            continue;
        }
        let prefix = qname_prefix(raw);
        let namespace = if prefix.is_empty() {
            Vec::new()
        } else {
            lookup_binding(scope, prefix).ok_or_else(|| {
                invalid(format!(
                    "Theme XML attribute uses undeclared namespace prefix '{}'",
                    String::from_utf8_lossy(prefix)
                ))
            })?
        };
        let local = qname_local(raw).to_vec();
        if seen.iter().any(|(known_namespace, known_local)| {
            known_namespace == &namespace && known_local == &local
        }) {
            return Err(invalid("Theme XML has duplicate expanded attributes"));
        }
        seen.push((namespace, local));
    }
    Ok(())
}

fn decode_attribute(
    attribute: &quick_xml::events::attributes::Attribute<'_>,
    reader: &NsReader<&[u8]>,
) -> Result<String> {
    family_codec::validate_raw_attribute_value(attribute.value.as_ref(), "Theme XML attribute")?;
    let value = attribute
        .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
        .map(|value| value.into_owned())
        .map_err(xml_error)?;
    family_codec::validate_event_text(value.as_bytes(), "Theme XML attribute")?;
    Ok(value)
}

fn validate_raw_theme_text(value: &[u8]) -> Result<()> {
    if value.windows(3).any(|window| window == b"]]>") {
        return Err(invalid(
            "Theme XML text contains the forbidden raw ']]>' delimiter",
        ));
    }
    Ok(())
}

fn namespace_declaration(raw: &[u8]) -> Option<&[u8]> {
    if raw == b"xmlns" {
        Some(b"")
    } else {
        raw.strip_prefix(b"xmlns:")
    }
}

fn namespace_declaration_names(element: &BytesStart<'_>) -> Vec<Vec<u8>> {
    first_wins(element)
        .flatten()
        .filter_map(|attribute| {
            namespace_declaration(attribute.key.as_ref()).map(ToOwned::to_owned)
        })
        .collect()
}

fn lookup_binding(scope: &[Binding], prefix: &[u8]) -> Option<Vec<u8>> {
    scope
        .iter()
        .rev()
        .find(|binding| binding.prefix.as_slice() == prefix)
        .map(|binding| binding.uri.clone())
}

fn resolved_namespace(resolved: &ResolveResult<'_>) -> Result<Vec<u8>> {
    match resolved {
        ResolveResult::Bound(Namespace(value)) => {
            if value.len() > MAX_NAMESPACE_BYTES {
                return Err(limit("Theme XML namespace bytes", MAX_NAMESPACE_BYTES));
            }
            std::str::from_utf8(value).map_err(xml_error)?;
            Ok(value.to_vec())
        },
        ResolveResult::Unbound => Ok(Vec::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "Theme XML element uses undeclared namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn is_drawingml_namespace(namespace: &[u8]) -> bool {
    namespace == DRAWINGML_NAMESPACE.as_bytes()
        || namespace == DRAWINGML_NAMESPACE_STRICT.as_bytes()
}

fn qname_prefix(name: &[u8]) -> &[u8] {
    name.iter()
        .position(|byte| *byte == b':')
        .map_or(&[][..], |index| &name[..index])
}

fn qname_local(name: &[u8]) -> &[u8] {
    name.iter()
        .rposition(|byte| *byte == b':')
        .map_or(name, |index| &name[index + 1..])
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value.iter().copied().all(is_xml_space)
}

fn is_xml_whitespace_character(character: char) -> bool {
    matches!(character, ' ' | '\t' | '\r' | '\n')
}

fn element_only_context(stack: &[Frame], extensions: &[ExtensionRecord]) -> bool {
    match stack.last().map(|frame| frame.kind) {
        Some(Kind::Root | Kind::ExtList) => true,
        Some(Kind::Ext(index)) => extensions
            .get(index)
            .and_then(|extension| extension.uri.as_deref())
            .and_then(ExtensionProfile::from_uri)
            .is_some(),
        _ => false,
    }
}

fn is_xml_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n')
}

fn open_tag_end(xml: &[u8]) -> Result<usize> {
    let mut quote = None;
    for (index, byte) in xml.iter().copied().enumerate() {
        match quote {
            Some(current) if byte == current => quote = None,
            Some(_) => {},
            None if byte == b'\'' || byte == b'"' => quote = Some(byte),
            None if byte == b'>' => return Ok(index + 1),
            None => {},
        }
    }
    Err(invalid("Theme XML start tag is unterminated"))
}

fn empty_element_slash(xml: &[u8]) -> Option<usize> {
    let mut cursor = xml.len();
    while cursor > 0 && is_xml_space(xml[cursor - 1]) {
        cursor -= 1;
    }
    if cursor == 0 || xml[cursor - 1] != b'>' {
        return None;
    }
    cursor -= 1;
    while cursor > 0 && is_xml_space(xml[cursor - 1]) {
        cursor -= 1;
    }
    (cursor > 0 && xml[cursor - 1] == b'/').then_some(cursor - 1)
}

/// The byte offset in the complete input of the reader's position: the
/// reader reads the input after its first `offset` bytes, and `origin` is the
/// [`ReaderOrigin`] of that slice.
fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin, offset: usize) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .and_then(|position| position.checked_add(offset))
        .ok_or_else(|| invalid("Theme XML offset exceeds usize"))
}

fn bump_nodes(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| invalid("Theme XML node count overflows"))?;
    if *nodes > MAX_NODES {
        return Err(limit("Theme XML nodes", MAX_NODES));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
