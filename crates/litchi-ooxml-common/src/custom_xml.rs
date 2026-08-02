//! Host-neutral Custom XML Data Storage package grammar.
//!
//! Payload XML is validated and retained as inert bytes. This module never
//! retrieves schemas, validates against a schema, runs transforms, resolves
//! external entities, or interprets application-specific payloads.

use crate::xml::decode_xml_reference;
use crate::{
    Error, ExpandedName, MceCapabilities, MceLimits, Result, process_markup_compatibility,
};
use litchi_opc::part::XmlPart;
use litchi_opc::{ContentType, OpcPackage, PackURI, Part, TargetMode};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesDecl, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

/// Transitional Custom XML Data Storage namespace.
pub const TRANSITIONAL_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
/// Strict Custom XML Data Storage namespace.
pub const STRICT_NAMESPACE: &str = "http://purl.oclc.org/ooxml/officeDocument/customXml";
/// Transitional relationship from a host part to a Custom XML data part.
pub const TRANSITIONAL_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
/// Strict relationship from a host part to a Custom XML data part.
pub const STRICT_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";
/// Transitional relationship from a data part to its properties part.
pub const TRANSITIONAL_PROPS_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
/// Strict relationship from a data part to its properties part.
pub const STRICT_PROPS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customXmlProps";
/// Content type required for a Custom XML properties part.
pub const PROPS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";

/// Maximum bytes accepted in one Custom XML payload part.
pub const MAX_PART_BYTES: usize = 16 * 1024 * 1024;
/// Maximum bytes accepted in one Custom XML properties part.
pub const MAX_PROPS_BYTES: usize = 1024 * 1024;
/// Maximum XML element nesting in payloads and properties.
pub const MAX_DEPTH: usize = 256;
/// Maximum Custom XML relationship occurrences in a package.
pub const MAX_ITEMS: usize = 4096;
/// Maximum schema references in one properties part.
pub const MAX_SCHEMA_REFS: usize = 4096;
/// Maximum elements scanned in one payload.
pub const MAX_ELEMENTS: usize = 1_000_000;
/// Maximum aggregate bytes in property strings.
pub const MAX_STRING_BYTES: usize = 4 * 1024 * 1024;

/// OOXML namespace and relationship family used for newly-authored parts.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Conformance {
    /// ECMA-376 transitional namespace family.
    Transitional,
    /// ISO/IEC 29500 strict namespace family.
    Strict,
}

impl Conformance {
    /// Namespace used by the properties vocabulary.
    #[must_use]
    pub const fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_NAMESPACE,
            Self::Strict => STRICT_NAMESPACE,
        }
    }

    /// Relationship used from a host part to the data part.
    #[must_use]
    pub const fn relationship(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_RELATIONSHIP,
            Self::Strict => STRICT_RELATIONSHIP,
        }
    }

    /// Relationship used from the data part to its properties part.
    #[must_use]
    pub const fn props_relationship(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_PROPS_RELATIONSHIP,
            Self::Strict => STRICT_PROPS_RELATIONSHIP,
        }
    }
}

/// Contents of a Custom XML Data Storage Properties part (ISO/IEC 29500 §22.5).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Props {
    /// Globally unique `ST_Guid` associated with the data part.
    pub id: String,
    /// Target namespaces of associated schemas; these URIs are never resolved.
    pub schemas: Vec<String>,
}

/// One relationship occurrence targeting a Custom XML Data Storage part.
#[derive(Debug, PartialEq, Eq)]
pub struct Item {
    /// Part that owns the data relationship.
    source: PackURI,
    /// Relationship ID on [`Self::source`].
    rel_id: String,
    /// Canonical package name of the data part.
    part: PackURI,
    /// Declared XML-based content type.
    content_type: String,
    /// Expanded name of the payload document element.
    root: ExpandedName,
    /// Exact, uninterpreted payload bytes.
    xml: Arc<Vec<u8>>,
    /// Canonical package name of the optional properties part.
    props_part: Option<PackURI>,
    /// Parsed optional properties.
    props: Option<Props>,
}

impl Item {
    /// Part that owns this data relationship.
    #[must_use]
    pub fn source(&self) -> &PackURI {
        &self.source
    }

    /// Relationship ID on [`Self::source`].
    #[must_use]
    pub fn rel_id(&self) -> &str {
        &self.rel_id
    }

    /// Canonical package name of the data part.
    #[must_use]
    pub fn part(&self) -> &PackURI {
        &self.part
    }

    /// Declared XML-based content type.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Expanded name of the payload document element.
    #[must_use]
    pub fn root(&self) -> &ExpandedName {
        &self.root
    }

    /// Exact, uninterpreted payload bytes borrowed from shared OPC storage.
    #[must_use]
    pub fn xml(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Canonical package name of the optional properties part.
    #[must_use]
    pub fn props_part(&self) -> Option<&PackURI> {
        self.props_part.as_ref()
    }

    /// Parsed optional properties.
    #[must_use]
    pub fn props(&self) -> Option<&Props> {
        self.props.as_ref()
    }
}

/// Properties authoring request.
///
/// Grouping these fields makes an incomplete properties request
/// unrepresentable: the part name, relationship ID, and value are all present
/// or all absent through [`Option`].
#[derive(Debug, PartialEq, Eq)]
pub struct NewProps {
    /// Package name for the new properties part.
    pub part: PackURI,
    /// Relationship ID from the new data part to the properties part.
    pub rel_id: String,
    /// Typed properties value to serialize.
    pub value: Props,
}

/// Deterministic package authoring request.
#[derive(Debug, PartialEq, Eq)]
pub struct NewItem {
    /// Existing part that will own the new data relationship.
    pub source: PackURI,
    /// Relationship ID from [`Self::source`] to [`Self::part`].
    pub rel_id: String,
    /// Package name for the new data part.
    pub part: PackURI,
    /// XML-based content type for the payload.
    pub content_type: String,
    /// Exact payload bytes to store after validation.
    pub xml: Vec<u8>,
    /// Optional, type-safe properties authoring request.
    pub props: Option<NewProps>,
    /// Namespace and relationship family to author.
    pub conformance: Conformance,
}

/// Parse a Custom XML Data Storage Properties part with bounded MCE handling.
pub fn read_props(xml: &[u8]) -> Result<Props> {
    require_at_most("custom XML properties bytes", xml.len(), MAX_PROPS_BYTES)?;
    let mut capabilities = MceCapabilities::ooxml_baseline();
    capabilities
        .understand_namespace(TRANSITIONAL_NAMESPACE)
        .understand_namespace(STRICT_NAMESPACE);
    let limits = MceLimits {
        max_input_bytes: MAX_PROPS_BYTES,
        max_output_bytes: MAX_PROPS_BYTES.saturating_mul(2),
        max_depth: MAX_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, &capabilities, &limits)?;
    parse_props_xml(processed.xml.as_ref())
}

/// Serialize properties in stable schema order.
pub fn write_props(props: &Props, conformance: Conformance) -> Result<Vec<u8>> {
    validate_props(props)?;
    let output_len = props_output_len(props, conformance)?;
    let mut out = Vec::with_capacity(output_len);
    out.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.extend_from_slice(b"<ds:datastoreItem xmlns:ds=\"");
    push_escaped_attr(&mut out, conformance.namespace());
    out.extend_from_slice(b"\" ds:itemID=\"");
    push_escaped_attr(&mut out, &props.id);
    out.extend_from_slice(b"\">");
    if !props.schemas.is_empty() {
        out.extend_from_slice(b"<ds:schemaRefs>");
        for uri in &props.schemas {
            out.extend_from_slice(b"<ds:schemaRef ds:uri=\"");
            push_escaped_attr(&mut out, uri);
            out.extend_from_slice(b"\"/>");
        }
        out.extend_from_slice(b"</ds:schemaRefs>");
    }
    out.extend_from_slice(b"</ds:datastoreItem>");
    if out.len() != output_len {
        return invalid("custom XML properties size calculation disagrees with serialization");
    }
    Ok(out)
}

/// Discover and validate every explicit Custom XML Data Storage relationship.
pub fn discover(package: &OpcPackage) -> Result<Vec<Item>> {
    let mut items = Vec::new();
    scan(
        package,
        |source, rel_id, part, data, root, props_part, props| {
            items.push(Item {
                source: source.clone(),
                rel_id: rel_id.into(),
                part,
                content_type: data.content_type().into(),
                root,
                xml: data.blob_arc(),
                props_part,
                props,
            });
            Ok(())
        },
    )?;
    items.sort_unstable_by(|left, right| {
        left.source
            .as_str()
            .cmp(right.source.as_str())
            .then_with(|| left.rel_id.cmp(&right.rel_id))
    });
    Ok(items)
}

/// Atomically add a validated data part, optional properties part, and relationships.
///
/// All fallible graph and XML work happens before package mutation. Defensive
/// rollback also covers an unexpected insertion or relationship failure, so an
/// error never exposes a partially-created Custom XML item.
pub fn add(package: &mut OpcPackage, item: NewItem) -> Result<()> {
    validate_content_type(&item.content_type)?;
    validate_payload(&item.xml)?;
    require_rel_id(&item.rel_id, "custom XML relationship")?;
    if item.part.as_str() == "/" {
        return invalid("custom XML data part cannot be the package root");
    }

    let source = package.get_part(&item.source).map_err(|error| {
        Error::Missing(format!(
            "custom XML source '{}': {error}",
            item.source.as_str()
        ))
    })?;
    if source.rels().get(&item.rel_id).is_some() {
        return invalid(format!(
            "relationship '{}' already exists on '{}'",
            item.rel_id,
            item.source.as_str()
        ));
    }
    package.validate_new_part_name(&item.part)?;

    if let Some(new_props) = &item.props {
        validate_props(&new_props.value)?;
        require_rel_id(&new_props.rel_id, "custom XML properties relationship")?;
        if new_props.part.as_str() == "/" {
            return invalid("custom XML properties part cannot be the package root");
        }
        package.validate_new_part_name(&new_props.part)?;
    }
    validate_new_names(&item.part, item.props.as_ref().map(|props| &props.part))?;

    if let Some(candidate_id) = item.props.as_ref().map(|props| props.value.id.as_str()) {
        scan(package, |_, _, _, _, _, _, props| {
            if props
                .as_ref()
                .is_some_and(|existing| candidate_id.eq_ignore_ascii_case(&existing.id))
            {
                return invalid(format!("duplicate custom XML itemID '{candidate_id}'"));
            }
            Ok(())
        })?;
    }

    let NewItem {
        source,
        rel_id,
        part,
        content_type,
        xml,
        props,
        conformance,
    } = item;
    let prepared_props = if let Some(NewProps {
        part,
        rel_id,
        value,
    }) = props
    {
        let xml = write_props(&value, conformance)?;
        Some((part, rel_id, xml))
    } else {
        None
    };

    let mut data = XmlPart::new(part.clone(), content_type, xml);
    if let Some((props_part, props_rel_id, _)) = prepared_props.as_ref() {
        let target = props_part.relative_ref(part.base_uri());
        data.rels_mut().try_add_relationship(
            conformance.props_relationship().into(),
            target,
            props_rel_id.clone(),
            TargetMode::Internal,
        )?;
    }

    let inserted_props = if let Some((props_part, _, props_xml)) = prepared_props {
        package.try_add_part(Box::new(XmlPart::new(
            props_part.clone(),
            PROPS_CONTENT_TYPE.into(),
            props_xml,
        )))?;
        Some(props_part)
    } else {
        None
    };

    if let Err(error) = package.try_add_part(Box::new(data)) {
        rollback_parts(package, &part, inserted_props.as_ref());
        return Err(error.into());
    }

    let target = part.relative_ref(source.base_uri());
    let relation_result =
        package
            .get_part_mut(&source)
            .map_err(Error::from)
            .and_then(|source_part| {
                source_part
                    .rels_mut()
                    .try_add_relationship(
                        conformance.relationship().into(),
                        target,
                        rel_id,
                        TargetMode::Internal,
                    )
                    .map(|_| ())
                    .map_err(Error::from)
            });
    if let Err(error) = relation_result {
        rollback_parts(package, &part, inserted_props.as_ref());
        return Err(error);
    }
    package.unsign();
    Ok(())
}

/// Validate a typed properties value without serializing it.
pub fn validate_props(props: &Props) -> Result<()> {
    if !valid_guid(&props.id) {
        return invalid(format!("custom XML itemID '{}' is not ST_Guid", props.id));
    }
    require_at_most(
        "custom XML schema references",
        props.schemas.len(),
        MAX_SCHEMA_REFS,
    )?;
    let mut string_bytes = props.id.len();
    for schema in &props.schemas {
        if schema.is_empty() {
            return invalid("custom XML schema reference URI must not be empty");
        }
        validate_xml_chars(schema)?;
        string_bytes = string_bytes.checked_add(schema.len()).ok_or_else(|| {
            limit(
                "custom XML property string bytes",
                MAX_STRING_BYTES,
                usize::MAX,
            )
        })?;
    }
    require_at_most(
        "custom XML property string bytes",
        string_bytes,
        MAX_STRING_BYTES,
    )
}

/// Return whether `value` is the braced hexadecimal `ST_Guid` lexical form.
#[must_use]
pub fn valid_guid(value: &str) -> bool {
    let Some(inner) = value
        .strip_prefix('{')
        .and_then(|value| value.strip_suffix('}'))
    else {
        return false;
    };
    let groups = [8, 4, 4, 4, 12];
    let mut parts = inner.split('-');
    groups.into_iter().all(|length| {
        parts.next().is_some_and(|part| {
            part.len() == length && part.bytes().all(|byte| byte.is_ascii_hexdigit())
        })
    }) && parts.next().is_none()
}

/// Validate a bounded XML payload and return its expanded document-element name.
pub fn validate_payload(xml: &[u8]) -> Result<ExpandedName> {
    require_at_most("custom XML payload bytes", xml.len(), MAX_PART_BYTES)?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root = None;
    let mut closed_root = false;
    let mut elements = 0usize;
    let mut version = XmlVersion::Implicit1_0;
    let mut event_seen = false;

    loop {
        let event = reader.read_event_into(&mut buffer)?;
        match event {
            Event::Decl(declaration) => {
                if event_seen {
                    return invalid("custom XML payload has a late XML declaration");
                }
                version = validate_declaration(&declaration)?;
            },
            Event::Start(element) => {
                require_nested_depth(depth)?;
                elements = bump_elements(elements)?;
                let name = inspect_element(&reader, &element, version, depth == 0)?;
                if depth == 0 {
                    if closed_root || root.is_some() {
                        return invalid("custom XML payload has multiple roots");
                    }
                    root = name;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("custom XML depth", MAX_DEPTH, usize::MAX))?;
            },
            Event::Empty(element) => {
                require_nested_depth(depth)?;
                elements = bump_elements(elements)?;
                let name = inspect_element(&reader, &element, version, depth == 0)?;
                if depth == 0 {
                    if closed_root || root.is_some() {
                        return invalid("custom XML payload has multiple roots");
                    }
                    root = name;
                    closed_root = true;
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("custom XML payload has an unexpected end tag".into())
                })?;
                if depth == 0 {
                    closed_root = true;
                }
            },
            Event::DocType(_) => return invalid("DTD is forbidden in custom XML payloads"),
            Event::GeneralRef(reference) => {
                if depth == 0 {
                    return invalid("custom XML payload has a reference outside its root");
                }
                let value = decode_xml_reference(&reference)?;
                validate_xml_chars(&value)?;
            },
            Event::Text(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&decoded)?;
                if depth == 0 && !is_xml_whitespace(decoded.as_bytes()) {
                    return invalid("custom XML payload has text outside its root");
                }
            },
            Event::CData(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&decoded)?;
                if depth == 0 {
                    return invalid("custom XML payload has CDATA outside its root");
                }
            },
            Event::Comment(text) => {
                let text = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&text)?;
            },
            Event::PI(instruction) => validate_instruction(&reader, &instruction)?,
            Event::Eof => break,
        }
        event_seen = true;
        buffer.clear();
    }
    if depth != 0 || !closed_root {
        return invalid("custom XML payload has no complete root element");
    }
    root.ok_or_else(|| Error::Invalid("custom XML payload has no root".into()))
}

/// Validate that a Custom XML data-part content type is a well-formed XML media type.
pub fn validate_content_type(content_type: &str) -> Result<()> {
    let parsed = ContentType::new(content_type).map_err(|_| Error::ContentType {
        expected: "a well-formed XML media type".into(),
        actual: content_type.into(),
    })?;
    let media_type = parsed.as_str().split(';').next().unwrap_or_default();
    let Some((type_name, subtype)) = media_type.split_once('/') else {
        return Err(Error::ContentType {
            expected: "an XML media type".into(),
            actual: content_type.into(),
        });
    };
    let xml = (type_name.eq_ignore_ascii_case("application")
        || type_name.eq_ignore_ascii_case("text"))
        && subtype.eq_ignore_ascii_case("xml");
    let suffix = subtype.len() > 4
        && subtype
            .get(subtype.len() - 4..)
            .is_some_and(|value| value.eq_ignore_ascii_case("+xml"));
    if xml || suffix {
        Ok(())
    } else {
        Err(Error::ContentType {
            expected: "an XML media type".into(),
            actual: content_type.into(),
        })
    }
}

fn scan(
    package: &OpcPackage,
    mut visit: impl FnMut(
        &PackURI,
        &str,
        PackURI,
        &dyn Part,
        ExpandedName,
        Option<PackURI>,
        Option<Props>,
    ) -> Result<()>,
) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_data_relationship(relationship.reltype()))
    {
        return invalid("package root cannot source a Custom XML Data Storage relationship");
    }

    let mut occurrences = 0usize;
    let mut property_ids: HashMap<String, PackURI> = HashMap::new();
    let mut cached_props: HashMap<PackURI, Props> = HashMap::new();
    let mut props_owners: HashMap<PackURI, PackURI> = HashMap::new();
    let mut cached_roots: HashMap<PackURI, ExpandedName> = HashMap::new();

    for source in package.iter_parts() {
        for relationship in source
            .rels()
            .iter()
            .filter(|relationship| is_data_relationship(relationship.reltype()))
        {
            occurrences = occurrences
                .checked_add(1)
                .ok_or_else(|| limit("custom XML items", MAX_ITEMS, usize::MAX))?;
            require_at_most("custom XML items", occurrences, MAX_ITEMS)?;
            if relationship.is_external() {
                return Err(Error::Relationship(format!(
                    "custom XML relationship '{}' from '{}' must be internal",
                    relationship.r_id(),
                    source.partname().as_str()
                )));
            }
            let requested_name = relationship.target_partname().map_err(|error| {
                Error::Relationship(format!(
                    "invalid custom XML target '{}': {error}",
                    relationship.r_id()
                ))
            })?;
            let data = package.get_part(&requested_name).map_err(|error| {
                Error::Missing(format!(
                    "custom XML part '{}': {error}",
                    requested_name.as_str()
                ))
            })?;
            let part = data.partname().clone();
            validate_content_type(data.content_type())?;
            let root = if let Some(root) = cached_roots.get(&part) {
                root.clone()
            } else {
                let root = validate_payload(data.blob())?;
                cached_roots.insert(part.clone(), root.clone());
                root
            };
            let (props_part, props) = resolve_props(
                package,
                data,
                &mut property_ids,
                &mut cached_props,
                &mut props_owners,
            )?;
            visit(
                source.partname(),
                relationship.r_id(),
                part,
                data,
                root,
                props_part,
                props,
            )?;
        }
    }
    Ok(())
}

fn resolve_props(
    package: &OpcPackage,
    data: &dyn Part,
    property_ids: &mut HashMap<String, PackURI>,
    cache: &mut HashMap<PackURI, Props>,
    owners: &mut HashMap<PackURI, PackURI>,
) -> Result<(Option<PackURI>, Option<Props>)> {
    let mut relationships = data.rels().iter();
    let Some(relationship) = relationships.next() else {
        return Ok((None, None));
    };
    if relationships.next().is_some() {
        return invalid(format!(
            "custom XML part '{}' has more than one outbound relationship",
            data.partname().as_str()
        ));
    }
    if !is_props_relationship(relationship.reltype()) {
        return Err(Error::Relationship(format!(
            "custom XML part '{}' has forbidden relationship type '{}'",
            data.partname().as_str(),
            relationship.reltype()
        )));
    }
    if relationship.is_external() {
        return Err(Error::Relationship(
            "custom XML properties relationship must be internal".into(),
        ));
    }
    let requested_name = relationship.target_partname().map_err(|error| {
        Error::Relationship(format!("invalid custom XML properties target: {error}"))
    })?;
    let part = package.get_part(&requested_name).map_err(|error| {
        Error::Missing(format!(
            "custom XML properties part '{}': {error}",
            requested_name.as_str()
        ))
    })?;
    let part_name = part.partname().clone();
    if let Some(existing_owner) = owners.insert(part_name.clone(), data.partname().clone())
        && existing_owner != *data.partname()
    {
        return invalid(format!(
            "custom XML properties part '{}' is shared by '{}' and '{}'",
            part_name.as_str(),
            existing_owner.as_str(),
            data.partname().as_str()
        ));
    }
    let props = if let Some(props) = cache.get(&part_name) {
        props.clone()
    } else {
        if part.content_type() != PROPS_CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: PROPS_CONTENT_TYPE.into(),
                actual: part.content_type().into(),
            });
        }
        if !part.rels().is_empty() {
            return invalid(format!(
                "custom XML properties part '{}' must not have relationships",
                part_name.as_str()
            ));
        }
        let props = read_props(part.blob())?;
        let key = props.id.to_ascii_lowercase();
        if let Some(existing) = property_ids.insert(key, part_name.clone())
            && existing != part_name
        {
            return invalid(format!("duplicate custom XML itemID '{}'", props.id));
        }
        cache.insert(part_name.clone(), props.clone());
        props
    };
    Ok((Some(part_name), Some(props)))
}

fn validate_new_names(part: &PackURI, props_part: Option<&PackURI>) -> Result<()> {
    let mut candidates = OpcPackage::new();
    candidates.try_add_part(Box::new(XmlPart::new(
        part.clone(),
        "application/xml".into(),
        Vec::new(),
    )))?;
    if let Some(props_part) = props_part {
        candidates.try_add_part(Box::new(XmlPart::new(
            props_part.clone(),
            PROPS_CONTENT_TYPE.into(),
            Vec::new(),
        )))?;
    }
    Ok(())
}

fn rollback_parts(package: &mut OpcPackage, part: &PackURI, props_part: Option<&PackURI>) {
    package.remove_part(part);
    if let Some(props_part) = props_part {
        package.remove_part(props_part);
    }
}

#[derive(Debug)]
struct ResolvedElement {
    namespace: String,
    local_name: String,
    attributes: Vec<(String, String, String)>,
}

fn inspect_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    version: XmlVersion,
    capture: bool,
) -> Result<Option<ExpandedName>> {
    validate_qname(element.name().as_ref(), "element")?;
    let (resolved, local) = reader.resolver().resolve_element(element.name());
    let namespace = resolved_namespace(resolved, "element")?;
    let namespace =
        std::str::from_utf8(namespace).map_err(|error| Error::Xml(error.to_string()))?;
    let local =
        std::str::from_utf8(local.as_ref()).map_err(|error| Error::Xml(error.to_string()))?;
    let captured = capture.then(|| ExpandedName {
        namespace: namespace.into(),
        local_name: local.into(),
    });

    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        validate_xml_chars(&value)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            validate_namespace_declaration(attribute.key.as_ref())?;
            continue;
        }
        validate_qname(attribute.key.as_ref(), "attribute")?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(resolved, "attribute")?;
        std::str::from_utf8(namespace).map_err(|error| Error::Xml(error.to_string()))?;
        std::str::from_utf8(local.as_ref()).map_err(|error| Error::Xml(error.to_string()))?;
        if !seen.insert((namespace.to_vec(), local.as_ref().to_vec())) {
            return invalid(format!(
                "duplicate expanded XML attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            ));
        }
    }
    Ok(captured)
}

fn resolve_props_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    version: XmlVersion,
) -> Result<ResolvedElement> {
    validate_qname(element.name().as_ref(), "element")?;
    let (resolved, local) = reader.resolver().resolve_element(element.name());
    let namespace = std::str::from_utf8(resolved_namespace(resolved, "element")?)
        .map_err(|error| Error::Xml(error.to_string()))?
        .to_owned();
    let local_name = std::str::from_utf8(local.as_ref())
        .map_err(|error| Error::Xml(error.to_string()))?
        .to_owned();
    let mut attributes = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        validate_xml_chars(&value)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            validate_namespace_declaration(attribute.key.as_ref())?;
            continue;
        }
        validate_qname(attribute.key.as_ref(), "attribute")?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let attribute_namespace = std::str::from_utf8(resolved_namespace(resolved, "attribute")?)
            .map_err(|error| Error::Xml(error.to_string()))?
            .to_owned();
        let attribute_name = std::str::from_utf8(local.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?
            .to_owned();
        if !seen.insert((attribute_namespace.clone(), attribute_name.clone())) {
            return invalid(format!(
                "duplicate expanded XML attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            ));
        }
        attributes.push((attribute_namespace, attribute_name, value));
    }
    Ok(ResolvedElement {
        namespace,
        local_name,
        attributes,
    })
}

fn resolved_namespace<'a>(result: ResolveResult<'a>, kind: &str) -> Result<&'a [u8]> {
    match result {
        ResolveResult::Bound(Namespace(value)) => Ok(value),
        ResolveResult::Unbound => Ok(b""),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound XML {kind} namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

fn parse_props_xml(xml: &[u8]) -> Result<Props> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_comments = true;
    let mut buffer = Vec::new();
    let mut contexts = Vec::new();
    let mut id = None;
    let mut schemas = Vec::new();
    let mut root_namespace = None;
    let mut seen_schemas = false;
    let mut closed_root = false;
    let mut version = XmlVersion::Implicit1_0;
    let mut event_seen = false;

    loop {
        let event = reader.read_event_into(&mut buffer)?;
        match event {
            Event::Decl(declaration) => {
                if event_seen {
                    return invalid("custom XML properties have a late XML declaration");
                }
                version = validate_declaration(&declaration)?;
            },
            Event::Start(element) => {
                require_nested_depth(contexts.len())?;
                let element = resolve_props_element(&reader, &element, version)?;
                let context = props_start(
                    &element,
                    contexts.last().copied(),
                    &mut id,
                    &mut schemas,
                    &mut root_namespace,
                    &mut seen_schemas,
                )?;
                contexts.push(context);
            },
            Event::Empty(element) => {
                require_nested_depth(contexts.len())?;
                let element = resolve_props_element(&reader, &element, version)?;
                props_start(
                    &element,
                    contexts.last().copied(),
                    &mut id,
                    &mut schemas,
                    &mut root_namespace,
                    &mut seen_schemas,
                )?;
                if contexts.is_empty() {
                    closed_root = true;
                }
            },
            Event::End(_) => {
                if contexts.pop().is_none() {
                    return invalid("unexpected custom XML properties end tag");
                }
                if contexts.is_empty() {
                    closed_root = true;
                }
            },
            Event::DocType(_) => return invalid("DTD is forbidden in custom XML properties"),
            Event::Text(text) => {
                let text = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&text)?;
                if !is_xml_whitespace(text.as_bytes()) {
                    return invalid("text is not permitted in custom XML properties");
                }
            },
            Event::CData(text) => {
                let text = text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                validate_xml_chars(&text)?;
                if contexts.is_empty() {
                    return invalid("CDATA is not permitted outside custom XML properties");
                }
                if !is_xml_whitespace(text.as_bytes()) {
                    return invalid("CDATA is not permitted in custom XML properties");
                }
            },
            Event::GeneralRef(reference) => {
                if contexts.is_empty() {
                    return invalid("references are not permitted outside custom XML properties");
                }
                let value = decode_xml_reference(&reference)?;
                validate_xml_chars(&value)?;
                if !is_xml_whitespace(value.as_bytes()) {
                    return invalid("references are not permitted in custom XML properties");
                }
            },
            Event::Comment(text) => {
                text.decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
            },
            Event::PI(_) => {
                return invalid("processing instructions are forbidden in custom XML properties");
            },
            Event::Eof => break,
        }
        event_seen = true;
        buffer.clear();
    }
    if !closed_root || !contexts.is_empty() {
        return invalid("custom XML properties root is incomplete");
    }
    let props = Props {
        id: id.ok_or_else(|| Error::Invalid("datastoreItem requires ds:itemID".into()))?,
        schemas,
    };
    validate_props(&props)?;
    Ok(props)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PropsContext {
    Item,
    Schemas,
    Schema,
}

fn props_start(
    element: &ResolvedElement,
    parent: Option<PropsContext>,
    id: &mut Option<String>,
    schemas: &mut Vec<String>,
    root_namespace: &mut Option<String>,
    seen_schemas: &mut bool,
) -> Result<PropsContext> {
    let supported_namespace = is_custom_namespace(&element.namespace);
    match parent {
        None if root_namespace.is_none()
            && supported_namespace
            && element.local_name == "datastoreItem" =>
        {
            reject_attributes_except(element, &[(element.namespace.as_str(), "itemID")])?;
            *id = Some(required_attr(element, &element.namespace, "itemID")?.into());
            *root_namespace = Some(element.namespace.clone());
            Ok(PropsContext::Item)
        },
        Some(PropsContext::Item)
            if root_namespace.as_deref() == Some(element.namespace.as_str())
                && element.local_name == "schemaRefs" =>
        {
            if *seen_schemas {
                return invalid("datastoreItem has multiple schemaRefs elements");
            }
            *seen_schemas = true;
            reject_attributes_except(element, &[])?;
            Ok(PropsContext::Schemas)
        },
        Some(PropsContext::Schemas)
            if root_namespace.as_deref() == Some(element.namespace.as_str())
                && element.local_name == "schemaRef" =>
        {
            let next = schemas.len().checked_add(1).ok_or_else(|| {
                limit("custom XML schema references", MAX_SCHEMA_REFS, usize::MAX)
            })?;
            require_at_most("custom XML schema references", next, MAX_SCHEMA_REFS)?;
            reject_attributes_except(element, &[(element.namespace.as_str(), "uri")])?;
            schemas.push(required_attr(element, &element.namespace, "uri")?.into());
            Ok(PropsContext::Schema)
        },
        _ => invalid(format!(
            "unexpected custom XML properties element {{{}}}{}",
            element.namespace, element.local_name
        )),
    }
}

fn reject_attributes_except(element: &ResolvedElement, allowed: &[(&str, &str)]) -> Result<()> {
    for (namespace, local_name, _) in &element.attributes {
        if !allowed.iter().any(|(allowed_namespace, allowed_name)| {
            namespace == allowed_namespace && local_name == allowed_name
        }) {
            return invalid(format!(
                "unexpected attribute {{{namespace}}}{local_name} on {}",
                element.local_name
            ));
        }
    }
    Ok(())
}

fn required_attr<'a>(
    element: &'a ResolvedElement,
    namespace: &str,
    local_name: &str,
) -> Result<&'a str> {
    element
        .attributes
        .iter()
        .find(|(candidate_namespace, candidate_name, _)| {
            candidate_namespace == namespace && candidate_name == local_name
        })
        .map(|(_, _, value)| value.as_str())
        .ok_or_else(|| {
            Error::Invalid(format!(
                "{} requires {{{namespace}}}{local_name}",
                element.local_name
            ))
        })
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<XmlVersion> {
    let version = declaration.xml_version()?;
    let declaration_text =
        std::str::from_utf8(declaration.as_ref()).map_err(|error| Error::Xml(error.to_string()))?;
    let raw = BytesStart::from_content(declaration_text, 3);
    let mut declaration_state = 0u8;
    for attribute in raw.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.prefix().is_some() {
            return invalid(format!(
                "unexpected XML declaration attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            ));
        }
        declaration_state = match (declaration_state, attribute.key.as_ref()) {
            (0, b"version") => 1,
            (1, b"encoding") => 2,
            (1 | 2, b"standalone") => 3,
            _ => {
                return invalid(format!(
                    "unexpected or out-of-order XML declaration attribute '{}'",
                    String::from_utf8_lossy(attribute.key.as_ref())
                ));
            },
        };
        std::str::from_utf8(attribute.value.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(encoding) = declaration.encoding() {
        let encoding = encoding.map_err(|error| Error::Xml(error.to_string()))?;
        let encoding =
            std::str::from_utf8(&encoding).map_err(|error| Error::Xml(error.to_string()))?;
        if !valid_encoding_name(encoding) {
            return invalid(format!(
                "XML declaration encoding '{encoding}' is not an EncName"
            ));
        }
    }
    if let Some(standalone) = declaration.standalone() {
        let standalone = standalone.map_err(|error| Error::Xml(error.to_string()))?;
        if !matches!(standalone.as_ref(), b"yes" | b"no") {
            return invalid("XML declaration standalone must be 'yes' or 'no'");
        }
    }
    Ok(version)
}

fn validate_instruction(
    reader: &NsReader<&[u8]>,
    instruction: &quick_xml::events::BytesPI<'_>,
) -> Result<()> {
    let target = reader
        .decoder()
        .decode(instruction.target())
        .map_err(|error| Error::Xml(error.to_string()))?;
    if !valid_xml_name(&target) {
        return Err(Error::Xml(format!(
            "invalid processing-instruction target '{target}'"
        )));
    }
    if target.eq_ignore_ascii_case("xml") {
        return invalid("processing-instruction target cannot be 'xml'");
    }
    let content = reader
        .decoder()
        .decode(instruction.content())
        .map_err(|error| Error::Xml(error.to_string()))?;
    validate_xml_chars(&content)?;
    Ok(())
}

fn props_output_len(props: &Props, conformance: Conformance) -> Result<usize> {
    const DECL: &[u8] = b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    const ROOT_PREFIX: &[u8] = b"<ds:datastoreItem xmlns:ds=\"";
    const ID_PREFIX: &[u8] = b"\" ds:itemID=\"";
    const ROOT_OPEN_END: &[u8] = b"\">";
    const SCHEMAS_OPEN: &[u8] = b"<ds:schemaRefs>";
    const SCHEMA_PREFIX: &[u8] = b"<ds:schemaRef ds:uri=\"";
    const SCHEMA_END: &[u8] = b"\"/>";
    const SCHEMAS_END: &[u8] = b"</ds:schemaRefs>";
    const ROOT_END: &[u8] = b"</ds:datastoreItem>";

    let mut total = 0usize;
    add_len(&mut total, DECL.len())?;
    add_len(&mut total, ROOT_PREFIX.len())?;
    add_len(&mut total, escaped_attr_len(conformance.namespace())?)?;
    add_len(&mut total, ID_PREFIX.len())?;
    add_len(&mut total, escaped_attr_len(&props.id)?)?;
    add_len(&mut total, ROOT_OPEN_END.len())?;
    if !props.schemas.is_empty() {
        add_len(&mut total, SCHEMAS_OPEN.len())?;
        for schema in &props.schemas {
            add_len(&mut total, SCHEMA_PREFIX.len())?;
            add_len(&mut total, escaped_attr_len(schema)?)?;
            add_len(&mut total, SCHEMA_END.len())?;
        }
        add_len(&mut total, SCHEMAS_END.len())?;
    }
    add_len(&mut total, ROOT_END.len())?;
    require_at_most(
        "custom XML serialized properties bytes",
        total,
        MAX_PROPS_BYTES,
    )?;
    Ok(total)
}

fn escaped_attr_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |total, character| {
        let width = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            '\t' | '\n' | '\r' => 5,
            _ => character.len_utf8(),
        };
        total.checked_add(width).ok_or_else(|| {
            limit(
                "custom XML serialized properties bytes",
                MAX_PROPS_BYTES,
                usize::MAX,
            )
        })
    })
}

fn add_len(total: &mut usize, value: usize) -> Result<()> {
    *total = total.checked_add(value).ok_or_else(|| {
        limit(
            "custom XML serialized properties bytes",
            MAX_PROPS_BYTES,
            usize::MAX,
        )
    })?;
    Ok(())
}

fn push_escaped_attr(out: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '>' => out.extend_from_slice(b"&gt;"),
            '"' => out.extend_from_slice(b"&quot;"),
            '\'' => out.extend_from_slice(b"&apos;"),
            '\t' => out.extend_from_slice(b"&#x9;"),
            '\n' => out.extend_from_slice(b"&#xA;"),
            '\r' => out.extend_from_slice(b"&#xD;"),
            _ => {
                let mut encoded = [0; 4];
                out.extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
            },
        }
    }
}

fn validate_xml_chars(value: &str) -> Result<()> {
    if value.chars().any(|character| {
        !matches!(
            character,
            '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(Error::Xml(
            "custom XML contains a character forbidden by XML 1.0".into(),
        ));
    }
    Ok(())
}

fn require_rel_id(value: &str, label: &str) -> Result<()> {
    if valid_ncname(value) {
        Ok(())
    } else {
        Err(Error::Relationship(format!(
            "{label} ID '{value}' is not an XML NCName"
        )))
    }
}

fn validate_qname(value: &[u8], kind: &str) -> Result<()> {
    let value = std::str::from_utf8(value).map_err(|error| Error::Xml(error.to_string()))?;
    let mut parts = value.split(':');
    let first = parts.next().unwrap_or_default();
    let second = parts.next();
    let valid = valid_ncname(first) && second.is_none_or(valid_ncname) && parts.next().is_none();
    if valid {
        Ok(())
    } else {
        Err(Error::Xml(format!("invalid XML {kind} QName '{value}'")))
    }
}

fn validate_namespace_declaration(value: &[u8]) -> Result<()> {
    if value == b"xmlns" {
        return Ok(());
    }
    let Some(prefix) = value.strip_prefix(b"xmlns:") else {
        return Err(Error::Xml("invalid XML namespace declaration".into()));
    };
    let prefix = std::str::from_utf8(prefix).map_err(|error| Error::Xml(error.to_string()))?;
    if valid_ncname(prefix) {
        Ok(())
    } else {
        Err(Error::Xml(format!(
            "invalid XML namespace prefix '{prefix}'"
        )))
    }
}

fn valid_ncname(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(is_ncname_start) && characters.all(is_ncname_character)
}

fn valid_xml_name(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(is_name_start) && characters.all(is_name_character)
}

fn valid_encoding_name(value: &str) -> bool {
    let mut bytes = value.bytes();
    bytes.next().is_some_and(|byte| byte.is_ascii_alphabetic())
        && bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
}

fn is_ncname_start(character: char) -> bool {
    character != ':' && is_name_start(character)
}

fn is_ncname_character(character: char) -> bool {
    character != ':' && is_name_character(character)
}

fn is_name_start(character: char) -> bool {
    matches!(
        character,
        ':' | 'A'..='Z' | '_' | 'a'..='z'
            | '\u{C0}'..='\u{D6}'
            | '\u{D8}'..='\u{F6}'
            | '\u{F8}'..='\u{2FF}'
            | '\u{370}'..='\u{37D}'
            | '\u{37F}'..='\u{1FFF}'
            | '\u{200C}'..='\u{200D}'
            | '\u{2070}'..='\u{218F}'
            | '\u{2C00}'..='\u{2FEF}'
            | '\u{3001}'..='\u{D7FF}'
            | '\u{F900}'..='\u{FDCF}'
            | '\u{FDF0}'..='\u{FFFD}'
            | '\u{10000}'..='\u{EFFFF}'
    )
}

fn is_name_character(character: char) -> bool {
    is_name_start(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{B7}' | '\u{300}'..='\u{36F}' | '\u{203F}'..='\u{2040}'
        )
}

fn require_nested_depth(depth: usize) -> Result<()> {
    if depth >= MAX_DEPTH {
        Err(limit(
            "custom XML depth",
            MAX_DEPTH,
            depth.saturating_add(1),
        ))
    } else {
        Ok(())
    }
}

fn bump_elements(elements: usize) -> Result<usize> {
    let next = elements
        .checked_add(1)
        .ok_or_else(|| limit("custom XML elements", MAX_ELEMENTS, usize::MAX))?;
    require_at_most("custom XML elements", next, MAX_ELEMENTS)?;
    Ok(next)
}

fn require_at_most(resource: &'static str, actual: usize, max: usize) -> Result<()> {
    if actual > max {
        Err(limit(resource, max, actual))
    } else {
        Ok(())
    }
}

fn limit(resource: &'static str, max: usize, actual: usize) -> Error {
    Error::Limit {
        resource,
        max,
        actual,
    }
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Invalid(message.into()))
}

fn is_data_relationship(value: &str) -> bool {
    matches!(value, TRANSITIONAL_RELATIONSHIP | STRICT_RELATIONSHIP)
}

fn is_props_relationship(value: &str) -> bool {
    matches!(
        value,
        TRANSITIONAL_PROPS_RELATIONSHIP | STRICT_PROPS_RELATIONSHIP
    )
}

fn is_custom_namespace(value: &str) -> bool {
    matches!(value, TRANSITIONAL_NAMESPACE | STRICT_NAMESPACE)
}

fn is_namespace_declaration(value: &[u8]) -> bool {
    value == b"xmlns" || value.starts_with(b"xmlns:")
}

fn is_xml_whitespace(value: &[u8]) -> bool {
    value
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::part::BlobPart;

    const POI_XLSX: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/spreadsheet/customIndexedColors.xlsx");
    const LO_DOCX: &[u8] = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/core/objectpositioning/data/do-not-capture-draw-objs-on-page-draw-wrap-none.docx"
    );

    #[test]
    fn loads_poi_and_libreoffice_reference_fixtures() {
        let poi = OpcPackage::from_bytes(POI_XLSX).unwrap();
        let items = discover(&poi).unwrap();
        assert_eq!(items.len(), 1);
        assert_eq!(items[0].root().local_name, "easyPacket");
        assert!(items[0].props().unwrap().schemas.is_empty());

        let libreoffice = OpcPackage::from_bytes(LO_DOCX).unwrap();
        let items = discover(&libreoffice).unwrap();
        assert_eq!(items.len(), 1);
        assert_eq!(items[0].root().local_name, "Sources");
        assert_eq!(
            items[0].props().unwrap().schemas,
            ["http://schemas.openxmlformats.org/officeDocument/2006/bibliography"]
        );
    }

    #[test]
    fn strict_writer_is_deterministic_and_round_trips() {
        let props = sample_props();
        let first = write_props(&props, Conformance::Strict).unwrap();
        let second = write_props(&props, Conformance::Strict).unwrap();
        assert_eq!(first, second);
        assert!(
            std::str::from_utf8(&first)
                .unwrap()
                .contains(STRICT_NAMESPACE)
        );
        assert_eq!(read_props(&first).unwrap(), props);
    }

    #[test]
    fn attribute_whitespace_round_trips_without_normalization_loss() {
        let mut props = sample_props();
        props.schemas = vec!["urn:line\nnext\tlast\rreturn".into()];
        let xml = write_props(&props, Conformance::Transitional).unwrap();
        assert_eq!(read_props(&xml).unwrap(), props);
    }

    #[test]
    fn mce_selects_fallback_schema_reference() {
        let xml = format!(
            r#"<ds:datastoreItem xmlns:ds="{TRANSITIONAL_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported" ds:itemID="{{11111111-1111-1111-1111-111111111111}}"><ds:schemaRefs><mc:AlternateContent><mc:Choice Requires="x"><ds:schemaRef ds:uri="urn:wrong"/></mc:Choice><mc:Fallback><ds:schemaRef ds:uri="urn:right"/></mc:Fallback></mc:AlternateContent></ds:schemaRefs></ds:datastoreItem>"#
        );
        assert_eq!(read_props(xml.as_bytes()).unwrap().schemas, ["urn:right"]);
    }

    #[test]
    fn package_writer_round_trips_without_interpreting_payload() {
        let mut package = package_with_source();
        package.relate_to(
            "_xmlsignatures/origin.sigs",
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        assert!(package.is_signed());
        add(
            &mut package,
            NewItem {
                source: PackURI::new("/word/document.xml").unwrap(),
                rel_id: "rIdData".into(),
                part: PackURI::new("/customXml/item1.xml").unwrap(),
                content_type: "application/xml".into(),
                xml: b"<customer xmlns=\"urn:customer\" id=\"7\"/>".to_vec(),
                props: Some(NewProps {
                    part: PackURI::new("/customXml/itemProps1.xml").unwrap(),
                    rel_id: "rIdProps".into(),
                    value: sample_props(),
                }),
                conformance: Conformance::Transitional,
            },
        )
        .unwrap();
        assert!(!package.is_signed());
        let items = discover(&package).unwrap();
        assert_eq!(items.len(), 1);
        assert_eq!(
            items[0].xml(),
            b"<customer xmlns=\"urn:customer\" id=\"7\"/>"
        );
        assert_eq!(items[0].props().unwrap(), &sample_props());
    }

    #[test]
    fn failed_add_is_transactional() {
        let mut package = package_with_source();
        let source = PackURI::new("/word/document.xml").unwrap();
        let before_parts = package.part_count();
        let before_rels = package.get_part(&source).unwrap().rels().len();
        let same_part = PackURI::new("/customXml/item1.xml").unwrap();
        let error = add(
            &mut package,
            NewItem {
                source: source.clone(),
                rel_id: "rIdData".into(),
                part: same_part.clone(),
                content_type: "application/xml".into(),
                xml: b"<root/>".to_vec(),
                props: Some(NewProps {
                    part: same_part,
                    rel_id: "rIdProps".into(),
                    value: sample_props(),
                }),
                conformance: Conformance::Transitional,
            },
        )
        .unwrap_err();
        assert!(matches!(error, Error::Opc(_)));
        assert_eq!(package.part_count(), before_parts);
        assert_eq!(package.get_part(&source).unwrap().rels().len(), before_rels);

        let error = add(
            &mut package,
            NewItem {
                source: source.clone(),
                rel_id: "1 invalid".into(),
                part: PackURI::new("/customXml/item2.xml").unwrap(),
                content_type: "application/xml".into(),
                xml: b"<root/>".to_vec(),
                props: None,
                conformance: Conformance::Transitional,
            },
        )
        .unwrap_err();
        assert!(matches!(error, Error::Relationship(_)));
        assert_eq!(package.part_count(), before_parts);
        assert_eq!(package.get_part(&source).unwrap().rels().len(), before_rels);
    }

    #[test]
    fn rejects_malformed_properties_payloads_and_package_graphs() {
        assert!(read_props(br#"<!DOCTYPE x><x/>"#).is_err());
        let missing_id = format!(r#"<ds:datastoreItem xmlns:ds="{TRANSITIONAL_NAMESPACE}"/>"#);
        assert!(read_props(missing_id.as_bytes()).is_err());
        let duplicate_refs = format!(
            r#"<ds:datastoreItem xmlns:ds="{TRANSITIONAL_NAMESPACE}" ds:itemID="{{11111111-1111-1111-1111-111111111111}}"><ds:schemaRefs/><ds:schemaRefs/></ds:datastoreItem>"#
        );
        assert!(read_props(duplicate_refs.as_bytes()).is_err());
        assert!(validate_payload(br#"<!DOCTYPE x><x/>"#).is_err());
        assert!(validate_payload(b"<a><b></a>").is_err());
        assert!(validate_payload(b"&#32;<root/>").is_err());
        assert!(validate_payload(b"<![CDATA[ ]]><root/>").is_err());
        assert!(validate_payload(b"<root>&unknown;</root>").is_err());
        assert!(validate_payload(b"<root>&#x110000;</root>").is_err());
        assert!(validate_payload(b"<root>\0</root>").is_err());
        assert!(validate_payload(b"<1root/>").is_err());

        let mut package = OpcPackage::new();
        let mut source = BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            "application/xml".into(),
            b"<document/>".to_vec(),
        );
        source.rels_mut().add_relationship(
            TRANSITIONAL_RELATIONSHIP.into(),
            "https://example.invalid/data.xml".into(),
            "rId1".into(),
            true,
        );
        package.add_part(Box::new(source));
        assert!(discover(&package).is_err());
    }

    #[test]
    fn enforces_guid_depth_size_and_content_type_caps() {
        let mut invalid_guid = sample_props();
        invalid_guid.id = "not-a-guid".into();
        assert!(write_props(&invalid_guid, Conformance::Transitional).is_err());

        let too_deep = format!(
            "{}<leaf/>{}",
            "<x>".repeat(MAX_DEPTH),
            "</x>".repeat(MAX_DEPTH)
        );
        let error = validate_payload(too_deep.as_bytes()).unwrap_err();
        assert!(matches!(
            error,
            Error::Limit {
                resource: "custom XML depth",
                ..
            }
        ));
        assert!(validate_payload(&vec![b' '; MAX_PART_BYTES + 1]).is_err());
        assert!(validate_content_type("not-a-media-type+xml").is_err());
        assert!(validate_content_type("application/vnd.example+xml").is_ok());
    }

    #[test]
    fn rejects_invalid_declarations_attributes_and_xml_characters() {
        assert!(validate_payload(br#" <!--before--><?xml version="1.0"?><root/>"#).is_err());
        assert!(validate_payload(br#"<?xml version="1.0" bad="x"?><root/>"#).is_err());
        assert!(
            validate_payload(br#"<?xml version="1.0" standalone="yes" encoding="UTF-8"?><root/>"#)
                .is_err()
        );
        assert!(validate_payload(br#"<?xml version="1.0" standalone="maybe"?><root/>"#).is_err());
        assert!(validate_payload(br#"<p:root/>"#).is_err());
        assert!(validate_payload(br#"<root 1id="value"/>"#).is_err());
        assert!(
            validate_payload(br#"<root xmlns:a="urn:x" xmlns:b="urn:x" a:id="1" b:id="2"/>"#)
                .is_err()
        );
        let mut props = sample_props();
        props.schemas = vec!["urn:\0bad".into()];
        assert!(write_props(&props, Conformance::Transitional).is_err());
    }

    fn package_with_source() -> OpcPackage {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                .into(),
            b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>".to_vec(),
        )));
        package
    }

    fn sample_props() -> Props {
        Props {
            id: "{11111111-1111-1111-1111-111111111111}".into(),
            schemas: vec!["urn:customer".into()],
        }
    }
}
