//! Bounded source scanning for worksheet-owned drawing relationships.
//!
//! A worksheet's `drawing` element is a relationship owner in the direct
//! `CT_Worksheet` content model.  It is easy to accidentally discover a
//! lookalike below `extLst`, foreign markup, or an MCE branch when searching
//! the XML by local name.  This scanner keeps the ownership boundary explicit:
//! it resolves expanded names from decoded namespace declarations, admits only
//! direct SpreadsheetML children of the worksheet root, and refuses a direct
//! `mc:AlternateContent` subtree that hides a drawing owner.
//!
//! The scanner records relationship IDs only.  Worksheet and drawing parts,
//! package relationships, and payload bytes remain owned by the workbook
//! facade.  Namespace state is updated in place and restored by scope marks;
//! it never clones an in-scope map for each XML node.

use std::collections::HashMap;
use std::str;

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesDecl, BytesStart, Event};
use quick_xml::name::{PrefixDeclaration, QName};
use quick_xml::reader::Reader;

use crate::error::{Result, allocation, invalid};
use litchi_ooxml_common::xml::attributes::count_up_to;
use litchi_ooxml_common::xml::{decode_xml_reference, is_ncname};

const SPREADSHEETML: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SPREADSHEETML: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const RELATIONSHIPS: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS: &[u8] = b"http://www.w3.org/2000/xmlns/";

/// Hard upper bound for one worksheet source member admitted by this scanner.
pub const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
/// Hard upper bound for XML events inspected by this scanner.
pub const MAX_XML_NODES: usize = 1_000_000;
/// Hard upper bound for XML nesting inspected by this scanner.
pub const MAX_XML_DEPTH: usize = 256;
/// Hard upper bound for direct worksheet drawing references retained by a scan.
pub const MAX_DRAWING_REFERENCES: usize = 1;
/// Hard upper bound for attributes inspected on one worksheet element.
pub const MAX_ATTRIBUTES: usize = 256;
/// Hard upper bound for namespace declarations on one worksheet element.
pub const MAX_NAMESPACE_DECLARATIONS: usize = 256;
/// Hard upper bound for active namespace declarations.
pub const MAX_ACTIVE_NAMESPACE_BINDINGS: usize = 16 * 1024;
/// Hard upper bound for one namespace URI's lexical bytes.
pub const MAX_NAMESPACE_BYTES: usize = 16 * 1024;
/// Hard upper bound for one decoded non-namespace attribute value.
pub const MAX_ATTRIBUTE_VALUE_BYTES: usize = 1024 * 1024;
/// Hard upper bound for one relationship ID after XML decoding.
pub const MAX_RELATIONSHIP_ID_BYTES: usize = 256;
/// Hard upper bound for one XML name or namespace prefix.
pub const MAX_NAME_BYTES: usize = 4 * 1024;

/// Caller-configurable finite limits for worksheet source scanning.
///
/// Every field is clamped to the corresponding scanner hard ceiling.  The
/// caller can make an operation narrower, but cannot use a larger value to
/// disable the structural safety boundary.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[must_use]
pub struct WorksheetSourceLimits {
    /// Maximum worksheet XML bytes inspected.
    pub max_xml_bytes: usize,
    /// Maximum XML events inspected.
    pub max_nodes: usize,
    /// Maximum XML nesting depth.
    pub max_depth: usize,
    /// Maximum direct drawing references retained.
    pub max_drawing_references: usize,
    /// Maximum attributes on one element.
    pub max_attributes: usize,
    /// Maximum namespace declarations on one element.
    pub max_namespace_declarations: usize,
    /// Maximum active namespace bindings.
    pub max_active_namespace_bindings: usize,
    /// Maximum namespace URI lexical bytes.
    pub max_namespace_bytes: usize,
    /// Maximum decoded non-namespace attribute bytes.
    pub max_attribute_value_bytes: usize,
    /// Maximum decoded relationship ID bytes.
    pub max_relationship_id_bytes: usize,
}

impl Default for WorksheetSourceLimits {
    fn default() -> Self {
        Self {
            max_xml_bytes: MAX_XML_BYTES,
            max_nodes: MAX_XML_NODES,
            max_depth: MAX_XML_DEPTH,
            max_drawing_references: MAX_DRAWING_REFERENCES,
            max_attributes: MAX_ATTRIBUTES,
            max_namespace_declarations: MAX_NAMESPACE_DECLARATIONS,
            max_active_namespace_bindings: MAX_ACTIVE_NAMESPACE_BINDINGS,
            max_namespace_bytes: MAX_NAMESPACE_BYTES,
            max_attribute_value_bytes: MAX_ATTRIBUTE_VALUE_BYTES,
            max_relationship_id_bytes: MAX_RELATIONSHIP_ID_BYTES,
        }
    }
}

impl WorksheetSourceLimits {
    /// Derive the worksheet-source policy from the package read profile.
    ///
    /// The worksheet scanner has a narrower structural policy than the OPC
    /// reader.  This conversion preserves the caller's finite XML, depth,
    /// relationship, attribute, and member-byte ceilings while keeping the
    /// scanner's hard caps in force.  Callers that scan an already-owned
    /// worksheet member may construct this type directly instead.
    pub fn from_read_limits(limits: litchi_opc::ReadLimits) -> Self {
        let defaults = Self::default();
        let part_bytes = usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX);
        Self {
            max_xml_bytes: defaults.max_xml_bytes.min(part_bytes),
            max_nodes: defaults.max_nodes.min(limits.max_xml_events()),
            max_depth: defaults.max_depth.min(limits.max_xml_depth()),
            max_drawing_references: defaults
                .max_drawing_references
                .min(limits.max_relationships_per_part()),
            max_attributes: defaults.max_attributes,
            max_namespace_declarations: defaults.max_namespace_declarations,
            max_active_namespace_bindings: defaults.max_active_namespace_bindings,
            max_namespace_bytes: defaults
                .max_namespace_bytes
                .min(limits.max_xml_attribute_bytes()),
            max_attribute_value_bytes: defaults
                .max_attribute_value_bytes
                .min(limits.max_xml_attribute_bytes()),
            max_relationship_id_bytes: defaults
                .max_relationship_id_bytes
                .min(limits.max_relationship_target_bytes()),
        }
    }

    fn bounded(self) -> Self {
        Self {
            max_xml_bytes: self.max_xml_bytes.min(MAX_XML_BYTES),
            max_nodes: self.max_nodes.min(MAX_XML_NODES),
            max_depth: self.max_depth.min(MAX_XML_DEPTH),
            max_drawing_references: self.max_drawing_references.min(MAX_DRAWING_REFERENCES),
            max_attributes: self.max_attributes.min(MAX_ATTRIBUTES),
            max_namespace_declarations: self
                .max_namespace_declarations
                .min(MAX_NAMESPACE_DECLARATIONS),
            max_active_namespace_bindings: self
                .max_active_namespace_bindings
                .min(MAX_ACTIVE_NAMESPACE_BINDINGS),
            max_namespace_bytes: self.max_namespace_bytes.min(MAX_NAMESPACE_BYTES),
            max_attribute_value_bytes: self
                .max_attribute_value_bytes
                .min(MAX_ATTRIBUTE_VALUE_BYTES),
            max_relationship_id_bytes: self
                .max_relationship_id_bytes
                .min(MAX_RELATIONSHIP_ID_BYTES),
        }
    }
}

/// One direct `<worksheet>/<drawing>` relationship reference.
#[derive(Clone, Debug, PartialEq, Eq)]
#[must_use]
pub struct WorksheetDrawingReference {
    relationship_id: Box<str>,
}

impl WorksheetDrawingReference {
    /// Return the decoded `r:id` token.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }
}

/// Result of scanning one worksheet source member.
#[derive(Clone, Debug, PartialEq, Eq)]
#[must_use]
pub struct WorksheetSourceScan {
    references: Vec<WorksheetDrawingReference>,
}

impl WorksheetSourceScan {
    /// Scan one worksheet using the hard bounded default policy.
    pub fn scan(source: &[u8]) -> Result<Self> {
        Self::scan_with_limits(source, WorksheetSourceLimits::default())
    }

    /// Scan one worksheet using a caller-supplied finite policy.
    pub fn scan_with_limits(source: &[u8], limits: WorksheetSourceLimits) -> Result<Self> {
        Scanner::new(source, limits.bounded())?.run()
    }

    /// Direct worksheet drawing owners in source order.
    pub fn drawings(&self) -> &[WorksheetDrawingReference] {
        &self.references
    }

    /// Alias emphasizing that the values are relationship references.
    pub fn references(&self) -> &[WorksheetDrawingReference] {
        &self.references
    }

    /// Return the only normative direct drawing owner, when present.
    pub fn drawing(&self, ordinal: usize) -> Result<&WorksheetDrawingReference> {
        self.references.get(ordinal).ok_or_else(|| {
            invalid(format!(
                "worksheet drawing ordinal {ordinal} is outside {} direct drawings",
                self.references.len()
            ))
        })
    }

    /// Borrow all decoded relationship IDs without allocating another vector.
    pub fn relationship_ids(&self) -> impl Iterator<Item = &str> {
        self.references
            .iter()
            .map(WorksheetDrawingReference::relationship_id)
    }
}

/// Scan direct worksheet drawing references and return their decoded IDs.
///
/// This convenience function is intended for package/edit owners that only
/// need relationship lookup.  `WorksheetSourceScan` should be retained by
/// read facades that also need to inspect the bounded inventory.
pub fn scan_worksheet_drawing_references(
    source: &[u8],
    limits: WorksheetSourceLimits,
) -> Result<Vec<String>> {
    let scan = WorksheetSourceScan::scan_with_limits(source, limits)?;
    let mut references = Vec::new();
    references
        .try_reserve_exact(scan.references.len())
        .map_err(|source| allocation("worksheet drawing references", source))?;
    for reference in &scan.references {
        let mut value = String::new();
        value
            .try_reserve_exact(reference.relationship_id.len())
            .map_err(|source| allocation("worksheet relationship ID", source))?;
        value.push_str(&reference.relationship_id);
        references.push(value);
    }
    Ok(references)
}

/// Short alias for callers that prefer a source-scanner verb.
pub fn scan(source: &[u8], limits: WorksheetSourceLimits) -> Result<WorksheetSourceScan> {
    WorksheetSourceScan::scan_with_limits(source, limits)
}

struct Scanner<'a> {
    limits: WorksheetSourceLimits,
    reader: Reader<&'a [u8]>,
    namespaces: NamespaceState,
    depth: usize,
    root_seen: bool,
    declaration_seen: bool,
    direct_alternate_content_depth: Option<usize>,
    opaque_alternate_content_depth: Option<usize>,
    hidden_drawing_seen: bool,
    references: Vec<WorksheetDrawingReference>,
    nodes: usize,
}

impl<'a> Scanner<'a> {
    fn new(source: &'a [u8], limits: WorksheetSourceLimits) -> Result<Self> {
        if source.len() > limits.max_xml_bytes {
            return Err(invalid(format!(
                "worksheet XML exceeds {} bytes",
                limits.max_xml_bytes
            )));
        }
        str::from_utf8(source)
            .map_err(|error| invalid(format!("worksheet XML is not UTF-8: {error}")))?;
        let mut references = Vec::new();
        references
            .try_reserve_exact(limits.max_drawing_references)
            .map_err(|source| allocation("worksheet drawing references", source))?;
        let mut reader = Reader::from_reader(source);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        reader.config_mut().check_comments = true;
        Ok(Self {
            limits,
            reader,
            namespaces: NamespaceState::default(),
            depth: 0,
            root_seen: false,
            declaration_seen: false,
            direct_alternate_content_depth: None,
            opaque_alternate_content_depth: None,
            hidden_drawing_seen: false,
            references,
            nodes: 0,
        })
    }

    fn run(mut self) -> Result<WorksheetSourceScan> {
        let mut buffer = Vec::new();
        loop {
            self.nodes = self
                .nodes
                .checked_add(1)
                .ok_or_else(|| invalid("worksheet XML event count overflows"))?;
            if self.nodes > self.limits.max_nodes {
                return Err(invalid("worksheet XML exceeds caller event limit"));
            }
            let event = self
                .reader
                .read_event_into(&mut buffer)
                .map_err(|error| invalid(error.to_string()))?;
            let decoder = self.reader.decoder();
            let event = match event {
                Event::Start(element) => {
                    self.start(element, decoder)?;
                    EventKind::Start
                },
                Event::Empty(element) => {
                    self.empty(element, decoder)?;
                    EventKind::Empty
                },
                Event::End(element) => {
                    self.end(element)?;
                    EventKind::End
                },
                Event::Decl(declaration) => {
                    self.declaration(declaration)?;
                    EventKind::Other
                },
                Event::Text(text) => {
                    self.text(text)?;
                    EventKind::Other
                },
                Event::CData(data) => {
                    self.cdata(data)?;
                    EventKind::Other
                },
                Event::Comment(comment) => {
                    self.comment(comment)?;
                    EventKind::Other
                },
                Event::PI(instruction) => {
                    self.processing_instruction(instruction, decoder)?;
                    EventKind::Other
                },
                Event::GeneralRef(reference) => {
                    self.general_reference(reference)?;
                    EventKind::Other
                },
                Event::DocType(_) => return Err(invalid("worksheet XML DTDs are not permitted")),
                Event::Eof => break,
            };
            if matches!(event, EventKind::Other) {
                buffer.clear();
                continue;
            }
            buffer.clear();
        }
        if !self.root_seen || self.depth != 0 || !self.namespaces.is_empty() {
            return Err(invalid("worksheet XML is incomplete"));
        }
        if self.hidden_drawing_seen {
            return Err(invalid(
                "worksheet drawing reference under direct AlternateContent is refused",
            ));
        }
        Ok(WorksheetSourceScan {
            references: self.references,
        })
    }

    fn declaration(&mut self, declaration: BytesDecl<'_>) -> Result<()> {
        if self.nodes != 1 || self.declaration_seen || self.root_seen {
            return Err(invalid("worksheet XML declaration is out of place"));
        }
        validate_declaration(&declaration)?;
        self.declaration_seen = true;
        Ok(())
    }

    fn text(&self, text: quick_xml::events::BytesText<'_>) -> Result<()> {
        if text.as_ref().windows(3).any(|window| window == b"]]>") {
            return Err(invalid(
                "worksheet XML text contains the forbidden ']]>' sequence",
            ));
        }
        let value = text
            .xml_content(XmlVersion::Implicit1_0)
            .map_err(|error| invalid(format!("worksheet XML text is not decodable: {error}")))?;
        validate_xml_chars(&value)?;
        if self.depth == 0 && !is_xml_whitespace(&value) {
            return Err(invalid("worksheet XML has text outside its root"));
        }
        Ok(())
    }

    fn cdata(&self, data: quick_xml::events::BytesCData<'_>) -> Result<()> {
        if data.as_ref().windows(3).any(|window| window == b"]]>") {
            return Err(invalid(
                "worksheet XML CDATA contains the forbidden ']]>' sequence",
            ));
        }
        let value = data
            .xml_content(XmlVersion::Implicit1_0)
            .map_err(|error| invalid(format!("worksheet XML CDATA is not decodable: {error}")))?;
        validate_xml_chars(&value)?;
        if self.depth == 0 {
            return Err(invalid("worksheet XML has CDATA outside its root"));
        }
        Ok(())
    }

    fn comment(&self, comment: quick_xml::events::BytesText<'_>) -> Result<()> {
        let bytes = comment.as_ref();
        if bytes.windows(2).any(|pair| pair == b"--") || bytes.last() == Some(&b'-') {
            return Err(invalid("worksheet XML comment contains a forbidden '--'"));
        }
        let value = comment
            .decode()
            .map_err(|error| invalid(format!("worksheet XML comment is not decodable: {error}")))?;
        validate_xml_chars(&value)
    }

    fn processing_instruction(
        &self,
        instruction: quick_xml::events::BytesPI<'_>,
        decoder: Decoder,
    ) -> Result<()> {
        let target = decoder
            .decode(instruction.target())
            .map_err(|error| invalid(format!("worksheet XML PI target is invalid: {error}")))?;
        validate_qname(target.as_bytes(), "processing-instruction")?;
        if target.eq_ignore_ascii_case("xml") {
            return Err(invalid(
                "worksheet XML processing-instruction target is reserved",
            ));
        }
        let content = decoder
            .decode(instruction.content())
            .map_err(|error| invalid(format!("worksheet XML PI is not decodable: {error}")))?;
        validate_xml_chars(&content)
    }

    fn general_reference(&self, reference: quick_xml::events::BytesRef<'_>) -> Result<()> {
        if self.depth == 0 {
            return Err(invalid("worksheet XML has a reference outside its root"));
        }
        let value = decode_xml_reference(&reference)
            .map_err(|error| invalid(format!("worksheet XML reference is invalid: {error}")))?;
        validate_xml_chars(&value)
    }

    fn start(&mut self, element: BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let element_depth = self
            .depth
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet XML nesting overflows"))?;
        if element_depth > self.limits.max_depth {
            return Err(invalid("worksheet XML exceeds caller depth limit"));
        }
        let scope = self.namespaces.push(&element, decoder, self.limits)?;
        let namespace = self.namespaces.resolve_element(element.name())?;
        if self.depth == 0 {
            if self.root_seen || element.name().local_name().as_ref() != b"worksheet" {
                self.namespaces.pop(scope);
                return Err(invalid("worksheet XML has an invalid root"));
            }
            if !is_namespace(namespace.as_deref(), SPREADSHEETML, STRICT_SPREADSHEETML) {
                self.namespaces.pop(scope);
                return Err(invalid("worksheet XML root has an unsupported namespace"));
            }
            self.root_seen = true;
        } else {
            self.inspect_element(&element, namespace.as_deref(), decoder)?;
        }
        self.depth = element_depth;
        if self.depth == 2
            && self.direct_alternate_content_depth.is_none()
            && namespace.as_deref().is_some_and(|value| value == MCE)
            && element.name().local_name().as_ref() == b"AlternateContent"
        {
            // `depth == 1` above denotes the root's direct children.  The
            // frame is recorded as the depth of the AlternateContent element
            // after opening so its descendants remain marked until close.
            self.direct_alternate_content_depth = Some(self.depth - 1);
        }
        Ok(())
    }

    fn empty(&mut self, element: BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let element_depth = self
            .depth
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet XML nesting overflows"))?;
        if element_depth > self.limits.max_depth {
            return Err(invalid("worksheet XML exceeds caller depth limit"));
        }
        let scope = self.namespaces.push(&element, decoder, self.limits)?;
        let namespace = self.namespaces.resolve_element(element.name())?;
        if self.depth == 0 {
            if self.root_seen
                || element.name().local_name().as_ref() != b"worksheet"
                || !is_namespace(namespace.as_deref(), SPREADSHEETML, STRICT_SPREADSHEETML)
            {
                self.namespaces.pop(scope);
                return Err(invalid("worksheet XML has an invalid root"));
            }
            self.root_seen = true;
        } else {
            self.inspect_element(&element, namespace.as_deref(), decoder)?;
        }
        self.namespaces.pop(scope);
        if self.opaque_alternate_content_depth == Some(element_depth) {
            self.opaque_alternate_content_depth = None;
        }
        Ok(())
    }

    fn end(&mut self, element: quick_xml::events::BytesEnd<'_>) -> Result<()> {
        if self.depth == 0 {
            return Err(invalid("worksheet XML has an unexpected end element"));
        }
        // `check_end_names` validates lexical pairing.  Resolving the closing
        // QName also rejects an unbound prefix before its scope is restored.
        let _ = self.namespaces.resolve_element(element.name())?;
        self.depth -= 1;
        if self.opaque_alternate_content_depth == Some(self.depth + 1) {
            self.opaque_alternate_content_depth = None;
        }
        let scope = self.namespaces.pop_last()?;
        if self.direct_alternate_content_depth == Some(self.depth) {
            self.direct_alternate_content_depth = None;
        }
        self.namespaces.pop(scope);
        Ok(())
    }

    fn inspect_element(
        &mut self,
        element: &BytesStart<'_>,
        namespace: Option<&[u8]>,
        decoder: Decoder,
    ) -> Result<()> {
        if namespace.is_some_and(|value| value == MCE)
            && element.name().local_name().as_ref() == b"AlternateContent"
            && self.depth == 1
        {
            // The corresponding depth marker is installed by `start` after
            // this function returns.  No drawing can be direct while this
            // element itself is being inspected.
            return Ok(());
        }
        if self.direct_alternate_content_depth.is_some() {
            if self.opaque_alternate_content_depth.is_some() {
                // A foreign wrapper inside a direct MCE owner is opaque as a
                // whole.  A SpreadsheetML-looking descendant below it must
                // not turn that foreign payload into a hidden owner refusal.
                return Ok(());
            }
            if namespace.is_none_or(|value| value != MCE)
                && !is_namespace(namespace, SPREADSHEETML, STRICT_SPREADSHEETML)
            {
                // Choice/Fallback descendants in the MCE vocabulary and
                // SpreadsheetML branches remain inspectable.  Any other
                // namespace is an opaque foreign wrapper until its end tag.
                self.opaque_alternate_content_depth = Some(self.depth + 1);
                return Ok(());
            }
        }
        if !is_namespace(namespace, SPREADSHEETML, STRICT_SPREADSHEETML)
            || element.name().local_name().as_ref() != b"drawing"
        {
            return Ok(());
        }
        if self.direct_alternate_content_depth.is_some() {
            self.hidden_drawing_seen = true;
            return Ok(());
        }
        if self.depth != 1 {
            // Foreign and opaque worksheet extension descendants are not
            // relationship owners under CT_Worksheet.
            return Ok(());
        }
        if self.references.len() >= self.limits.max_drawing_references {
            return Err(invalid(
                "worksheet contains more than one direct drawing reference",
            ));
        }
        let relationship_id = relationship_id(&self.namespaces, element, decoder, self.limits)?;
        self.references.push(WorksheetDrawingReference {
            relationship_id: relationship_id.into_boxed_str(),
        });
        Ok(())
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum EventKind {
    Start,
    Empty,
    End,
    Other,
}

fn is_namespace(namespace: Option<&[u8]>, transitional: &[u8], strict: &[u8]) -> bool {
    namespace.is_some_and(|value| value == transitional || value == strict)
}

fn relationship_id(
    namespaces: &NamespaceState,
    element: &BytesStart<'_>,
    decoder: Decoder,
    limits: WorksheetSourceLimits,
) -> Result<String> {
    let mut relationship_id = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let (namespace, local) = namespaces.resolve_attribute(attribute.key)?;
        if !is_namespace(namespace.as_deref(), RELATIONSHIPS, STRICT_RELATIONSHIPS)
            || local.as_ref() != b"id"
        {
            continue;
        }
        if relationship_id.is_some() {
            return Err(invalid(
                "worksheet drawing reference has duplicate relationship IDs",
            ));
        }
        if attribute.value.len() > limits.max_relationship_id_bytes.saturating_mul(4) {
            return Err(invalid(
                "worksheet relationship ID lexical value is too large",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?;
        let value = litchi_ooxml_common::xml::xsd_token_atom(&value)
            .ok_or_else(|| invalid("worksheet relationship ID is not one token"))?;
        if value.len() > limits.max_relationship_id_bytes || !is_ncname(value) || value.is_empty() {
            return Err(invalid("worksheet relationship ID is invalid"));
        }
        let mut decoded = String::new();
        decoded
            .try_reserve_exact(value.len())
            .map_err(|source| allocation("worksheet relationship ID", source))?;
        decoded.push_str(value);
        relationship_id = Some(decoded);
    }
    relationship_id.ok_or_else(|| invalid("worksheet drawing reference has no relationship ID"))
}

#[derive(Default)]
struct NamespaceState {
    bindings: HashMap<Vec<u8>, Vec<u8>>,
    changes: Vec<BindingChange>,
    scopes: Vec<usize>,
}

struct BindingChange {
    prefix: Vec<u8>,
    previous: Option<Vec<u8>>,
}

impl NamespaceState {
    fn is_empty(&self) -> bool {
        self.scopes.is_empty()
    }

    fn push(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        limits: WorksheetSourceLimits,
    ) -> Result<usize> {
        let mut attributes = 0usize;
        let mut declarations = 0usize;
        for attribute in element.attributes().with_checks(true) {
            let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
            if attribute.key.as_ref().len() > MAX_NAME_BYTES {
                return Err(invalid("worksheet XML attribute name is too large"));
            }
            attributes = attributes
                .checked_add(1)
                .ok_or_else(|| invalid("worksheet XML attribute count overflows"))?;
            if attributes > limits.max_attributes {
                return Err(invalid("worksheet XML exceeds caller attribute limit"));
            }
            if attribute.value.len() > limits.max_attribute_value_bytes {
                return Err(invalid(
                    "worksheet XML attribute value exceeds caller limit",
                ));
            }
            if attribute.value.as_ref().contains(&b'<') {
                return Err(invalid("worksheet XML attribute contains a literal '<'"));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| invalid(format!("worksheet XML attribute is invalid: {error}")))?;
            validate_xml_chars(&value)?;
            if attribute.key.as_namespace_binding().is_some() {
                declarations = declarations
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet namespace declaration count overflows"))?;
                if declarations > limits.max_namespace_declarations {
                    return Err(invalid(
                        "worksheet XML exceeds caller namespace declaration limit",
                    ));
                }
                if attribute.value.len() > limits.max_namespace_bytes {
                    return Err(invalid("worksheet namespace URI exceeds caller limit"));
                }
            }
        }
        let active = self
            .bindings
            .len()
            .checked_add(declarations)
            .ok_or_else(|| invalid("worksheet active namespace count overflows"))?;
        if active > limits.max_active_namespace_bindings {
            return Err(invalid(
                "worksheet XML exceeds caller active namespace limit",
            ));
        }
        self.scopes
            .try_reserve(1)
            .map_err(|source| allocation("worksheet namespace scopes", source))?;
        self.changes
            .try_reserve(declarations)
            .map_err(|source| allocation("worksheet namespace changes", source))?;
        self.bindings
            .try_reserve(declarations)
            .map_err(|source| allocation("worksheet namespace bindings", source))?;
        let scope = self.changes.len();
        self.scopes.push(scope);
        for attribute in element.attributes().with_checks(true) {
            let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
            let Some(declaration) = attribute.key.as_namespace_binding() else {
                continue;
            };
            let prefix = match declaration {
                PrefixDeclaration::Default => &[][..],
                PrefixDeclaration::Named(prefix) => prefix,
            };
            validate_prefix(prefix)?;
            let uri = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| invalid(error.to_string()))?;
            if uri.len() > limits.max_namespace_bytes {
                return Err(invalid("worksheet namespace URI exceeds caller limit"));
            }
            validate_binding(prefix, uri.as_bytes())?;
            let key = clone_bytes(prefix, "worksheet namespace prefix")?;
            let map_key = clone_bytes(&key, "worksheet namespace prefix")?;
            let previous = self.bindings.insert(
                map_key,
                clone_bytes(uri.as_bytes(), "worksheet namespace URI")?,
            );
            self.changes.push(BindingChange {
                prefix: key,
                previous,
            });
        }
        self.validate_attributes(element, limits.max_attributes)?;
        Ok(scope)
    }

    fn validate_attributes(&self, element: &BytesStart<'_>, max_attributes: usize) -> Result<()> {
        let mut expanded: Vec<(Option<Vec<u8>>, Vec<u8>)> = Vec::new();
        // `push` has refused a tag with more than `max_attributes`; counting
        // no further, and without quick-xml's duplicate check, keeps this
        // reservation bounded on its own.
        expanded
            .try_reserve_exact(count_up_to(element, max_attributes))
            .map_err(|source| allocation("worksheet expanded attribute names", source))?;
        for attribute in element.attributes().with_checks(true) {
            let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
            if attribute.key.as_namespace_binding().is_some() {
                continue;
            }
            let (namespace, local) = self.resolve_attribute(attribute.key)?;
            if expanded.iter().any(|known| {
                known.0.as_deref() == namespace.as_deref() && known.1.as_slice() == local.as_ref()
            }) {
                return Err(invalid(
                    "worksheet XML contains duplicate expanded attributes",
                ));
            }
            expanded.push((
                namespace,
                clone_bytes(local.as_ref(), "worksheet attribute local name")?,
            ));
        }
        Ok(())
    }

    fn pop_last(&mut self) -> Result<usize> {
        self.scopes
            .last()
            .copied()
            .ok_or_else(|| invalid("worksheet namespace scope underflow"))
    }

    fn pop(&mut self, scope: usize) {
        let _ = self.scopes.pop();
        while self.changes.len() > scope {
            let change = self
                .changes
                .pop()
                .expect("namespace scope change exists while unwinding");
            if let Some(previous) = change.previous {
                self.bindings.insert(change.prefix, previous);
            } else {
                self.bindings.remove(&change.prefix);
            }
        }
    }

    fn resolve_element(&self, name: QName<'_>) -> Result<Option<Vec<u8>>> {
        validate_qname(name.as_ref(), "element")?;
        let (_, prefix) = name.decompose();
        self.resolve_prefix(prefix.map(|prefix| prefix.into_inner()), true)
    }

    fn resolve_attribute<'a>(
        &self,
        name: QName<'a>,
    ) -> Result<(Option<Vec<u8>>, quick_xml::name::LocalName<'a>)> {
        validate_qname(name.as_ref(), "attribute")?;
        let (local, prefix) = name.decompose();
        let namespace = match prefix {
            Some(prefix) => self.resolve_prefix(Some(prefix.as_ref()), false)?,
            None => None,
        };
        Ok((namespace, local))
    }

    fn resolve_prefix(&self, prefix: Option<&[u8]>, prefixed: bool) -> Result<Option<Vec<u8>>> {
        let Some(prefix) = prefix else {
            let Some(uri) = self.bindings.get(&[][..]).filter(|uri| !uri.is_empty()) else {
                return Ok(None);
            };
            return Ok(Some(clone_bytes(uri, "worksheet namespace URI")?));
        };
        if prefix == b"xml" {
            return Ok(Some(clone_bytes(XML, "worksheet namespace URI")?));
        }
        if prefix == b"xmlns" {
            return Err(invalid("worksheet XML uses the reserved xmlns prefix"));
        }
        let Some(uri) = self.bindings.get(prefix) else {
            return Err(invalid(
                "worksheet XML name uses an unbound namespace prefix",
            ));
        };
        if uri.is_empty() && prefixed {
            return Err(invalid(
                "worksheet XML prefixed name uses an empty namespace binding",
            ));
        }
        if uri.is_empty() {
            Ok(None)
        } else {
            Ok(Some(clone_bytes(uri, "worksheet namespace URI")?))
        }
    }
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let content = str::from_utf8(declaration.as_ref())
        .map_err(|error| invalid(format!("worksheet XML declaration is not UTF-8: {error}")))?;
    if content.len() < 3 {
        return Err(invalid("worksheet XML declaration is incomplete"));
    }
    let start = BytesStart::from_content(content, 3);
    let mut state = 0u8;
    let mut attributes = 0usize;
    for attribute in start.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("worksheet XML declaration is invalid: {error}")))?;
        attributes = attributes
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet XML declaration attribute count overflows"))?;
        if attributes > 3 {
            return Err(invalid("worksheet XML declaration has too many attributes"));
        }
        let key = attribute.key.as_ref();
        let value = attribute.value.as_ref();
        if key.len() > MAX_NAME_BYTES || value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(invalid("worksheet XML declaration attribute is too large"));
        }
        if value.contains(&b'<') || value.contains(&b'&') {
            return Err(invalid(
                "worksheet XML declaration contains a forbidden raw attribute character",
            ));
        }
        let value = str::from_utf8(value).map_err(|error| {
            invalid(format!(
                "worksheet XML declaration attribute is not UTF-8: {error}"
            ))
        })?;
        validate_xml_chars(value)?;
        let next = match (state, key, value) {
            (0, b"version", "1.0") => 1,
            (1, b"encoding", value) if value.eq_ignore_ascii_case("UTF-8") => 2,
            (1 | 2, b"standalone", "yes" | "no") => 3,
            (0..=3, _, _) => {
                return Err(invalid(
                    "worksheet XML declaration has a duplicate or out-of-order attribute",
                ));
            },
            _ => {
                return Err(invalid("worksheet XML declaration is invalid"));
            },
        };
        state = next;
    }
    if state == 0 {
        return Err(invalid("worksheet XML declaration is missing version"));
    }
    Ok(())
}

fn validate_qname(raw: &[u8], kind: &str) -> Result<()> {
    if raw.is_empty() || raw.len() > MAX_NAME_BYTES {
        return Err(invalid(format!("worksheet XML {kind} QName is invalid")));
    }
    let name = str::from_utf8(raw)
        .map_err(|error| invalid(format!("worksheet XML {kind} QName is not UTF-8: {error}")))?;
    let mut parts = name.split(':');
    let prefix_or_local = parts.next().unwrap_or_default();
    if !is_ncname(prefix_or_local) {
        return Err(invalid(format!("worksheet XML {kind} QName is invalid")));
    }
    if let Some(local) = parts.next() {
        if !is_ncname(local) || parts.next().is_some() {
            return Err(invalid(format!("worksheet XML {kind} QName is invalid")));
        }
    }
    Ok(())
}

fn validate_xml_chars(value: &str) -> Result<()> {
    if value.chars().all(is_xml_character) {
        Ok(())
    } else {
        Err(invalid(
            "worksheet XML contains a forbidden XML 1.0 character",
        ))
    }
}

fn is_xml_character(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{a}' | '\u{d}')
        || ('\u{20}'..='\u{d7ff}').contains(&value)
        || ('\u{e000}'..='\u{fffd}').contains(&value)
        || ('\u{10000}'..='\u{10ffff}').contains(&value)
}

fn is_xml_whitespace(value: &str) -> bool {
    value
        .chars()
        .all(|character| matches!(character, ' ' | '\t' | '\r' | '\n'))
}

fn clone_bytes(value: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(value.len())
        .map_err(|source| allocation(resource, source))?;
    copy.extend_from_slice(value);
    Ok(copy)
}

fn validate_prefix(prefix: &[u8]) -> Result<()> {
    if prefix.len() > MAX_NAME_BYTES {
        return Err(invalid("worksheet XML namespace prefix is too large"));
    }
    if prefix.is_empty() || prefix == b"xml" || is_ncname(str::from_utf8(prefix).unwrap_or("")) {
        return Ok(());
    }
    Err(invalid("worksheet XML namespace prefix is invalid"))
}

fn validate_binding(prefix: &[u8], uri: &[u8]) -> Result<()> {
    if prefix == b"xml" {
        if uri != XML {
            return Err(invalid(
                "worksheet XML has an invalid xml namespace binding",
            ));
        }
        return Ok(());
    }
    if prefix == b"xmlns" || uri == XMLNS || (prefix != b"xml" && uri == XML) {
        return Err(invalid(
            "worksheet XML has an invalid reserved namespace binding",
        ));
    }
    if !prefix.is_empty() && uri.is_empty() {
        return Err(invalid("worksheet XML prefixed namespace binding is empty"));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::error::Error;

    /// ` a00000="0" a00001="1" ...`: `count` distinct attribute names.
    fn distinct_attributes(count: usize) -> String {
        (0..count)
            .map(|index| format!(" a{index:05}=\"{index}\""))
            .collect()
    }

    fn worksheet(attributes: &str) -> String {
        format!(
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData{attributes}/></worksheet>"
        )
    }

    fn scan(xml: &str, max_attributes: usize) -> Result<WorksheetSourceScan> {
        WorksheetSourceScan::scan_with_limits(
            xml.as_bytes(),
            WorksheetSourceLimits {
                max_attributes,
                ..WorksheetSourceLimits::default()
            },
        )
    }

    fn refused<T>(result: Result<T>, expected: &str) -> bool {
        matches!(result, Err(Error::Invalid(message)) if message.contains(expected))
    }

    #[test]
    fn attribute_limits_and_duplicates_are_refused_as_before() {
        assert!(
            scan(
                &worksheet(&distinct_attributes(MAX_ATTRIBUTES)),
                MAX_ATTRIBUTES
            )
            .is_ok()
        );
        assert!(refused(
            scan(
                &worksheet(&distinct_attributes(MAX_ATTRIBUTES + 1)),
                MAX_ATTRIBUTES
            ),
            "exceeds caller attribute limit"
        ));
        assert!(scan(&worksheet(&distinct_attributes(4)), 4).is_ok());
        assert!(refused(
            scan(&worksheet(&distinct_attributes(5)), 4),
            "exceeds caller attribute limit"
        ));
        assert!(refused(
            scan(&worksheet(" a='1' a='2'"), MAX_ATTRIBUTES),
            "duplicated attribute"
        ));
        assert!(refused(
            scan(
                &worksheet(" xmlns:p='urn:p' xmlns:q='urn:p' p:a='1' q:a='2'"),
                MAX_ATTRIBUTES
            ),
            "duplicate expanded attributes"
        ));
        // 20,000 distinct names followed by 20,000 repeats of the last one.
        let mut attributes = distinct_attributes(20_000);
        for _ in 0..20_000 {
            attributes.push_str(" a19999=\"r\"");
        }
        assert!(refused(
            scan(&worksheet(&attributes), MAX_ATTRIBUTES),
            "exceeds caller attribute limit"
        ));
    }

    #[test]
    fn expanded_names_are_validated_past_the_bounded_reservation() {
        // The reservation counts at most `max_attributes + 1` attributes;
        // validation still reads every one of them.
        let namespaces = NamespaceState::default();
        let tag = BytesStart::from_content("e a='1' b='2' c='3' d='4' e='5'", 1);
        assert!(namespaces.validate_attributes(&tag, 2).is_ok());
        let tag = BytesStart::from_content("e a='1' b='2' c='3' d='4' a='5'", 1);
        assert!(refused(
            namespaces.validate_attributes(&tag, 2),
            "duplicated attribute"
        ));
    }
}
