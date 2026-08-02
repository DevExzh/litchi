//! Bounded, inert Ribbon customization storage shared by every OOXML host.
//!
//! Ribbon callback names and image payloads remain opaque. This module only
//! validates the declared OPC graph and the Custom UI document boundary; it
//! never invokes callbacks, resolves commands, decodes images, or contacts an
//! external resource. XML is consumed without transcoding and therefore must
//! be UTF-8; an encoding declaration must say `UTF-8` case-insensitively.

use crate::{Error, Result};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{OpcPackage, PackURI, Part, XmlPart};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesDecl, BytesPI, BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;
use std::borrow::Cow;
use std::cmp::Ordering;
use std::collections::{HashSet, VecDeque};

const V2007_NAMESPACE: &str = "http://schemas.microsoft.com/office/2006/01/customui";
const V2010_NAMESPACE: &str = "http://schemas.microsoft.com/office/2009/07/customui";
const UI2_NAMESPACE: &str = "http://schemas.microsoft.com/office/2007/10/customui";
const LEGACY_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility";
const MODERN_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility";
const CONTENT_TYPE: &str = "application/xml";
const PART_NAME_ATTEMPTS: usize = 10_000;
const MAX_GRAPH_LINKS: usize = 1_000_000;
const MAX_PART_NAMES: usize = 100_000;
const MAX_PART_NAME_BYTES: usize = 64 * 1024 * 1024;
const MAX_IMAGE_GC_EDGES: usize = 262_144;

/// A package-level Ribbon relationship family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Family {
    /// Office 2007 Custom UI relationship family.
    Legacy,
    /// Office 2010 and CustomUI2 relationship family.
    Modern,
}

impl Family {
    /// Package relationship type for this family.
    #[must_use]
    pub const fn relationship(self) -> &'static str {
        match self {
            Self::Legacy => LEGACY_RELATIONSHIP,
            Self::Modern => MODERN_RELATIONSHIP,
        }
    }

    fn from_relationship(value: &str) -> Option<Self> {
        if value == LEGACY_RELATIONSHIP {
            Some(Self::Legacy)
        } else if value == MODERN_RELATIONSHIP {
            Some(Self::Modern)
        } else {
            None
        }
    }
}

/// Custom UI vocabulary selected by the root namespace.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Version {
    /// Office 2007 Custom UI vocabulary.
    V2007,
    /// Office 2010 Ribbon and Backstage vocabulary.
    V2010,
    /// CustomUI2 vocabulary documented by the Office extensions.
    Ui2,
}

impl Version {
    /// Root namespace required by this vocabulary.
    #[must_use]
    pub const fn namespace(self) -> &'static str {
        match self {
            Self::V2007 => V2007_NAMESPACE,
            Self::V2010 => V2010_NAMESPACE,
            Self::Ui2 => UI2_NAMESPACE,
        }
    }

    /// Package relationship type required by this vocabulary.
    #[must_use]
    pub const fn relationship(self) -> &'static str {
        self.family().relationship()
    }

    /// Package relationship family containing this vocabulary.
    #[must_use]
    pub const fn family(self) -> Family {
        match self {
            Self::V2007 => Family::Legacy,
            Self::V2010 | Self::Ui2 => Family::Modern,
        }
    }

    const fn default_part(self) -> &'static str {
        match self {
            Self::V2007 => "/customUI/customUI.xml",
            Self::V2010 => "/customUI/customUI14.xml",
            Self::Ui2 => "/customUI/customUI2.xml",
        }
    }
}

/// A Ribbon customization borrowing its package-owned XML allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Ui<'a> {
    part: &'a PackURI,
    id: &'a str,
    version: Version,
    xml: &'a [u8],
}

impl<'a> Ui<'a> {
    /// Canonical package part containing the Custom UI XML.
    #[must_use]
    #[inline]
    pub const fn part(self) -> &'a PackURI {
        self.part
    }

    /// Low-level package relationship ID.
    ///
    /// Prefer [`Set::effective`] and [`remove`] for semantic operations.
    #[must_use]
    #[inline]
    pub const fn id(self) -> &'a str {
        self.id
    }

    /// Custom UI vocabulary identified by the relationship and root namespace.
    #[must_use]
    #[inline]
    pub const fn version(self) -> Version {
        self.version
    }

    /// Original package-owned XML bytes.
    #[must_use]
    #[inline]
    pub const fn xml(self) -> &'a [u8] {
        self.xml
    }
}

/// Fixed Ribbon family slots for one package.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Set<'a> {
    legacy: Option<Ui<'a>>,
    modern: Option<Ui<'a>>,
}

impl<'a> Set<'a> {
    /// Office 2007 relationship-family slot.
    #[must_use]
    #[inline]
    pub const fn legacy(self) -> Option<Ui<'a>> {
        self.legacy
    }

    /// Office 2010 and CustomUI2 relationship-family slot.
    #[must_use]
    #[inline]
    pub const fn modern(self) -> Option<Ui<'a>> {
        self.modern
    }

    /// Customization used by modern-first consumers.
    #[must_use]
    #[inline]
    pub const fn effective(self) -> Option<Ui<'a>> {
        match self.modern {
            Some(value) => Some(value),
            None => self.legacy,
        }
    }

    /// Present slots in stable legacy-then-modern order.
    #[inline]
    pub fn iter(self) -> impl Iterator<Item = Ui<'a>> {
        [self.legacy, self.modern].into_iter().flatten()
    }

    const fn get(self, family: Family) -> Option<Ui<'a>> {
        match family {
            Family::Legacy => self.legacy,
            Family::Modern => self.modern,
        }
    }

    fn insert(&mut self, value: Ui<'a>) {
        let slot = match value.version.family() {
            Family::Legacy => &mut self.legacy,
            Family::Modern => &mut self.modern,
        };
        *slot = Some(value);
    }

    fn require_empty(self, family: Family) -> Result<()> {
        if self.get(family).is_some() {
            return Err(Error::Relationship(format!(
                "a package may contain at most one {:?} Ribbon relationship",
                family
            )));
        }
        Ok(())
    }
}

/// Resource ceilings applied while validating Ribbon package data.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum bytes in each Custom UI XML part.
    pub xml_bytes: usize,
    /// Maximum Custom UI XML element nesting depth.
    pub depth: usize,
    /// Maximum XML events and attributes in each Custom UI part.
    pub nodes: usize,
    /// Maximum aggregate image relationships across both Ribbon parts.
    pub images: usize,
}

impl Limits {
    /// Conservative defaults for untrusted Office packages.
    #[must_use]
    pub const fn standard() -> Self {
        Self {
            xml_bytes: 4 * 1024 * 1024,
            depth: 128,
            nodes: 262_144,
            images: 4_096,
        }
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::standard()
    }
}

/// Read both Ribbon family slots with safe default limits.
pub fn load(package: &OpcPackage) -> Result<Set<'_>> {
    load_with(package, &Limits::standard())
}

/// Read both Ribbon family slots with explicit resource limits.
pub fn load_with<'a>(package: &'a OpcPackage, limits: &Limits) -> Result<Set<'a>> {
    let mut scanned = 0usize;
    reject_part_sourced_ribbons(package, &mut scanned)?;

    let mut result = Set::default();
    let mut images = 0usize;
    for relationship in package.rels().iter() {
        bump_graph_links(&mut scanned)?;
        let Some(family) = Family::from_relationship(relationship.reltype()) else {
            continue;
        };
        result.require_empty(family)?;
        require_internal_target(relationship, "Ribbon")?;
        let target = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid Ribbon relationship target: {error}"))
        })?;
        let part = package.get_part(&target).map_err(|error| {
            Error::Missing(format!(
                "Ribbon part '{}' does not exist: {error}",
                target.as_str()
            ))
        })?;
        require_content_type(part, CONTENT_TYPE)?;
        let version = validate_xml(part.blob(), family, limits)?;
        validate_images(package, part, &mut images, limits)?;
        result.insert(Ui {
            part: part.partname(),
            id: relationship.r_id(),
            version,
            xml: part.blob(),
        });
    }
    Ok(result)
}

/// Create or replace one Ribbon family using safe default limits.
///
/// The XML allocation is moved into the OPC part. A byte-identical update is
/// a true no-op and preserves package signatures. Input must be UTF-8 XML.
pub fn put(package: &mut OpcPackage, version: Version, xml: Vec<u8>) -> Result<()> {
    put_with(package, version, xml, &Limits::standard())
}

/// Create or replace one Ribbon family using explicit resource limits.
pub fn put_with(
    package: &mut OpcPackage,
    version: Version,
    xml: Vec<u8>,
    limits: &Limits,
) -> Result<()> {
    let parsed = validate_xml(&xml, version.family(), limits)?;
    if parsed != version {
        return Err(Error::Invalid(format!(
            "Ribbon root namespace does not match requested {version:?} version"
        )));
    }

    let existing = {
        let ribbons = load_with(package, limits)?;
        match ribbons.get(version.family()) {
            Some(current) if current.version == version && current.xml == xml => return Ok(()),
            Some(current) => {
                let part = current.part.clone();
                let id = current.id.to_owned();
                let shared = has_other_inbound(package, &part, Some(&id))?;
                Some((part, id, shared))
            },
            None => None,
        }
    };

    if let Some((part, _, false)) = existing.as_ref() {
        package.get_part_mut(part)?.set_blob(xml);
        package.unsign();
        return Ok(());
    }

    let part = available_part_name(package, version)?;
    let target = part.as_str().trim_start_matches('/').to_owned();
    let mut replacement = XmlPart::new(part, CONTENT_TYPE.to_owned(), xml);
    if let Some((shared, _, true)) = existing.as_ref() {
        let base = replacement.partname().base_uri().to_owned();
        for relationship in package.get_part(shared)?.rels().iter() {
            let image = relationship.target_partname()?;
            replacement.rels_mut().try_add_relationship(
                relationship.reltype().to_owned(),
                image.relative_ref(&base),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )?;
        }
    }
    let relationship_id = existing.map(|(_, id, _)| id);
    if let Some(id) = relationship_id.as_deref()
        && package.rels_mut().remove(id).is_none()
    {
        return Err(Error::Relationship(format!(
            "Ribbon relationship '{id}' disappeared before commit"
        )));
    }
    package.add_part(Box::new(replacement));
    if let Some(id) = relationship_id {
        package
            .rels_mut()
            .add_relationship(version.relationship().to_owned(), target, id, false);
    } else {
        package.relate_to(&target, version.relationship());
    }
    package.unsign();
    Ok(())
}

/// Remove one Ribbon family and collect only parts that become unreferenced.
///
/// The complete deletion plan is resolved before mutation. Shared Ribbon or
/// image parts remain in the package. An absent family returns `Ok(false)` and
/// preserves signatures.
pub fn remove(package: &mut OpcPackage, family: Family) -> Result<bool> {
    let (relationship_id, ribbon_part, images) = {
        let ribbons = load(package)?;
        let Some(selected) = ribbons.get(family) else {
            return Ok(false);
        };
        let ribbon_part = selected.part.clone();
        let relationship_id = selected.id.to_owned();
        let remove_part =
            !has_other_inbound(package, &ribbon_part, Some(relationship_id.as_str()))?;
        let images = if remove_part {
            let candidates = ribbon_images(package, &ribbon_part)?;
            removable_images(package, &ribbon_part, &candidates)?
        } else {
            Vec::new()
        };
        (relationship_id, remove_part.then_some(ribbon_part), images)
    };

    if package.rels_mut().remove(&relationship_id).is_none() {
        return Err(Error::Relationship(format!(
            "Ribbon relationship '{relationship_id}' disappeared before commit"
        )));
    }
    if let Some(part) = ribbon_part {
        let _ = package.remove_part(&part);
        for image in images {
            let _ = package.remove_part(&image);
        }
    }
    package.unsign();
    Ok(true)
}

fn reject_part_sourced_ribbons(package: &OpcPackage, scanned: &mut usize) -> Result<()> {
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            bump_graph_links(scanned)?;
            if Family::from_relationship(relationship.reltype()).is_some() {
                return Err(Error::Relationship(format!(
                    "Ribbon relationship '{}' must be sourced by the package, not part '{}'",
                    relationship.r_id(),
                    part.partname().as_str()
                )));
            }
        }
    }
    Ok(())
}

fn validate_images(
    package: &OpcPackage,
    ribbon: &dyn Part,
    count: &mut usize,
    limits: &Limits,
) -> Result<()> {
    for relationship in ribbon.rels().iter() {
        if !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE) {
            return Err(Error::Relationship(format!(
                "Ribbon part '{}' may relate only to image parts; '{}' has type '{}'",
                ribbon.partname().as_str(),
                relationship.r_id(),
                relationship.reltype()
            )));
        }
        *count = count.checked_add(1).ok_or(Error::Limit {
            resource: "Ribbon image relationships",
            max: limits.images,
            actual: usize::MAX,
        })?;
        if *count > limits.images {
            return Err(Error::Limit {
                resource: "Ribbon image relationships",
                max: limits.images,
                actual: *count,
            });
        }
        require_internal_target(relationship, "Ribbon image")?;
        let target = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid Ribbon image target: {error}"))
        })?;
        let image = package.get_part(&target).map_err(|error| {
            Error::Missing(format!(
                "Ribbon image part '{}' does not exist: {error}",
                target.as_str()
            ))
        })?;
        if !is_image_content_type(image.content_type()) {
            return Err(Error::ContentType {
                expected: "image/*".to_owned(),
                actual: image.content_type().to_owned(),
            });
        }
    }
    Ok(())
}

fn require_internal_target(relationship: &litchi_opc::Relationship, context: &str) -> Result<()> {
    if relationship.is_external() {
        return Err(Error::Relationship(format!(
            "{context} relationship '{}' must be internal",
            relationship.r_id()
        )));
    }
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        return Err(Error::Relationship(format!(
            "{context} relationship '{}' target cannot contain a query or fragment",
            relationship.r_id()
        )));
    }
    Ok(())
}

fn require_content_type(part: &dyn Part, expected: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(Error::ContentType {
            expected: expected.to_owned(),
            actual: part.content_type().to_owned(),
        })
    }
}

fn is_image_content_type(value: &str) -> bool {
    if !valid_content_type(value) {
        return false;
    }
    let essence = value.split(';').next().unwrap_or(value);
    let Some((kind, subtype)) = essence.split_once('/') else {
        return false;
    };
    kind.eq_ignore_ascii_case("image") && is_mime_token(subtype)
}

fn valid_content_type(value: &str) -> bool {
    let mut components = value.split(';');
    let Some((kind, subtype)) = components.next().unwrap_or_default().split_once('/') else {
        return false;
    };
    if !is_mime_token(kind) || !is_mime_token(subtype) {
        return false;
    }
    components.all(|parameter| {
        let Some((name, raw_value)) = parameter.split_once('=') else {
            return false;
        };
        if !is_mime_token(name) {
            return false;
        }
        let value = raw_value
            .strip_prefix('"')
            .and_then(|value| value.strip_suffix('"'))
            .unwrap_or(raw_value);
        is_mime_token(value)
            && (!raw_value.contains('"')
                || (raw_value.starts_with('"') && raw_value.ends_with('"')))
    })
}

fn is_mime_token(value: &str) -> bool {
    !value.is_empty()
        && value.bytes().all(|byte| {
            byte.is_ascii_alphanumeric()
                || matches!(
                    byte,
                    b'!' | b'#'
                        | b'$'
                        | b'%'
                        | b'&'
                        | b'\''
                        | b'*'
                        | b'+'
                        | b'-'
                        | b'.'
                        | b'^'
                        | b'_'
                        | b'`'
                        | b'|'
                        | b'~'
                )
        })
}

fn validate_xml(xml: &[u8], family: Family, limits: &Limits) -> Result<Version> {
    if xml.len() > limits.xml_bytes {
        return Err(Error::Limit {
            resource: "Ribbon XML bytes",
            max: limits.xml_bytes,
            actual: xml.len(),
        });
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_comments = true;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut first_event = true;
    let mut version = None;

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event()?;
        let was_first = first_event;
        first_event = false;
        match event {
            Event::Decl(declaration) => {
                count_node(&mut nodes, limits)?;
                if !was_first || declaration_seen || root_seen {
                    return Err(Error::Invalid(
                        "Ribbon XML declaration must be the first event and occur once".into(),
                    ));
                }
                validate_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::DocType(_) => {
                return Err(Error::Invalid("DTD is forbidden in Ribbon XML".into()));
            },
            Event::PI(instruction) => {
                count_node(&mut nodes, limits)?;
                validate_instruction(&reader, &instruction)?;
            },
            Event::Start(element) => {
                count_node(&mut nodes, limits)?;
                validate_qname(element.name().as_ref(), "element")?;
                let root_version = if depth == 0 {
                    if root_seen || root_closed {
                        return Err(Error::Invalid(
                            "Ribbon XML must contain exactly one root".into(),
                        ));
                    }
                    Some(validate_root(&namespace, &element, decoder, family)?)
                } else {
                    validate_element_namespace(&namespace)?;
                    None
                };
                validate_attributes(&reader, &element, decoder, &mut nodes, limits)?;
                if let Some(root_version) = root_version {
                    version = Some(root_version);
                    root_seen = true;
                }
                depth = depth.checked_add(1).ok_or(Error::Limit {
                    resource: "Ribbon XML depth",
                    max: limits.depth,
                    actual: usize::MAX,
                })?;
                if depth > limits.depth {
                    return Err(Error::Limit {
                        resource: "Ribbon XML depth",
                        max: limits.depth,
                        actual: depth,
                    });
                }
            },
            Event::Empty(element) => {
                count_node(&mut nodes, limits)?;
                validate_qname(element.name().as_ref(), "element")?;
                let child_depth = depth.checked_add(1).ok_or(Error::Limit {
                    resource: "Ribbon XML depth",
                    max: limits.depth,
                    actual: usize::MAX,
                })?;
                if child_depth > limits.depth {
                    return Err(Error::Limit {
                        resource: "Ribbon XML depth",
                        max: limits.depth,
                        actual: child_depth,
                    });
                }
                let root_version = if depth == 0 {
                    if root_seen || root_closed {
                        return Err(Error::Invalid(
                            "Ribbon XML must contain exactly one root".into(),
                        ));
                    }
                    Some(validate_root(&namespace, &element, decoder, family)?)
                } else {
                    validate_element_namespace(&namespace)?;
                    None
                };
                validate_attributes(&reader, &element, decoder, &mut nodes, limits)?;
                if let Some(root_version) = root_version {
                    version = Some(root_version);
                    root_seen = true;
                    root_closed = true;
                }
            },
            Event::End(_) => {
                count_node(&mut nodes, limits)?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("Ribbon XML has an unexpected end element".into())
                })?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                count_node(&mut nodes, limits)?;
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(format!("invalid Ribbon XML text: {error}")))?;
                validate_xml_chars(&value)?;
                if depth == 0 && !value.trim().is_empty() {
                    return Err(Error::Invalid(
                        "Ribbon XML contains text outside its root".into(),
                    ));
                }
            },
            Event::CData(text) => {
                count_node(&mut nodes, limits)?;
                if depth == 0 {
                    return Err(Error::Invalid(
                        "Ribbon XML contains CDATA outside its root".into(),
                    ));
                }
                let value = text
                    .decode()
                    .map_err(|error| Error::Xml(format!("invalid Ribbon CDATA: {error}")))?;
                validate_xml_chars(&value)?;
            },
            Event::GeneralRef(reference) => {
                count_node(&mut nodes, limits)?;
                if depth == 0 {
                    return Err(Error::Invalid(
                        "Ribbon XML contains an entity reference outside its root".into(),
                    ));
                }
                validate_reference(&reference)?;
            },
            Event::Comment(comment) => {
                count_node(&mut nodes, limits)?;
                let value = comment
                    .decode()
                    .map_err(|error| Error::Xml(format!("invalid Ribbon comment: {error}")))?;
                validate_xml_chars(&value)?;
            },
            Event::Eof => break,
        }
    }

    if !root_seen || !root_closed || depth != 0 {
        return Err(Error::Invalid(
            "Ribbon XML must contain one complete customUI root".into(),
        ));
    }
    version.ok_or_else(|| Error::Invalid("Ribbon XML has no customUI root".into()))
}

fn validate_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    nodes: &mut usize,
    limits: &Limits,
) -> Result<()> {
    let mut expanded = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        count_node(nodes, limits)?;
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(format!("invalid Ribbon XML attribute: {error}")))?;
        validate_xml_chars(&value)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            validate_namespace_declaration(attribute.key.as_ref(), &value)?;
            continue;
        }
        validate_qname(attribute.key.as_ref(), "attribute")?;
        let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = normalized_namespace(&namespace, decoder, "attribute")?;
        let QName(raw_name) = attribute.key;
        let local_name = raw_name
            .rsplit(|byte| *byte == b':')
            .next()
            .unwrap_or(raw_name);
        if !expanded.insert((namespace, local_name)) {
            return Err(Error::Invalid(format!(
                "duplicate expanded Ribbon attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
    }
    Ok(())
}

fn validate_element_namespace(namespace: &ResolveResult<'_>) -> Result<()> {
    if let ResolveResult::Unknown(prefix) = namespace {
        return Err(Error::Invalid(format!(
            "unbound Ribbon element namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )));
    }
    Ok(())
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn validate_namespace_declaration(name: &[u8], value: &str) -> Result<()> {
    let prefix = name.strip_prefix(b"xmlns:");
    if let Some(prefix) = prefix {
        let prefix = std::str::from_utf8(prefix)
            .map_err(|error| Error::Xml(format!("invalid namespace prefix: {error}")))?;
        if !valid_ncname(prefix) || prefix == "xmlns" {
            return Err(Error::Invalid(format!(
                "invalid Ribbon namespace prefix '{prefix}'"
            )));
        }
        if value.is_empty() {
            return Err(Error::Invalid(format!(
                "Ribbon namespace prefix '{prefix}' cannot be undeclared in XML 1.0"
            )));
        }
        if (prefix == "xml") != (value == "http://www.w3.org/XML/1998/namespace") {
            return Err(Error::Invalid(
                "the XML namespace URI may be bound only to the 'xml' prefix".into(),
            ));
        }
    } else if value == "http://www.w3.org/XML/1998/namespace" {
        return Err(Error::Invalid(
            "the XML namespace URI may be bound only to the 'xml' prefix".into(),
        ));
    }
    if value == "http://www.w3.org/2000/xmlns/" {
        return Err(Error::Invalid(
            "the xmlns namespace URI cannot be rebound".into(),
        ));
    }
    if value
        .bytes()
        .any(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
    {
        return Err(Error::Invalid(
            "Ribbon namespace URI cannot contain XML whitespace".into(),
        ));
    }
    Ok(())
}

fn normalized_namespace<'a>(
    namespace: &ResolveResult<'a>,
    decoder: quick_xml::encoding::Decoder,
    kind: &str,
) -> Result<Cow<'a, str>> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => normalize_namespace_value(value, decoder),
        ResolveResult::Unbound => Ok(Cow::Borrowed("")),
        ResolveResult::Unknown(prefix) => Err(Error::Invalid(format!(
            "unbound Ribbon {kind} namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn normalize_namespace_value<'a>(
    value: &'a [u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Cow<'a, str>> {
    let decoded = decoder
        .decode(value)
        .map_err(|error| Error::Xml(format!("invalid Ribbon namespace URI: {error}")))?;
    let normalized = match decoded {
        Cow::Borrowed(value) => quick_xml::escape::unescape(value)
            .map_err(|error| Error::Xml(format!("invalid Ribbon namespace URI: {error}")))?,
        Cow::Owned(value) => Cow::Owned(
            quick_xml::escape::unescape(&value)
                .map_err(|error| Error::Xml(format!("invalid Ribbon namespace URI: {error}")))?
                .into_owned(),
        ),
    };
    validate_xml_chars(&normalized)?;
    Ok(normalized)
}

fn validate_reference(reference: &quick_xml::events::BytesRef<'_>) -> Result<()> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| Error::Xml(format!("invalid Ribbon character reference: {error}")))?
    {
        return if is_xml_char(character) {
            Ok(())
        } else {
            Err(Error::Xml(format!(
                "Ribbon character reference U+{:04X} is forbidden by XML 1.0",
                u32::from(character)
            )))
        };
    }
    let name = reference
        .decode()
        .map_err(|error| Error::Xml(format!("invalid Ribbon entity reference: {error}")))?;
    if matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot") {
        Ok(())
    } else {
        Err(Error::Invalid(format!(
            "unsupported Ribbon entity reference '&{name};'"
        )))
    }
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let version = declaration.xml_version()?;
    if version != XmlVersion::Explicit1_0 {
        return Err(Error::Invalid(
            "Ribbon XML declaration must use version 1.0".into(),
        ));
    }
    let declaration_text = std::str::from_utf8(declaration.as_ref())
        .map_err(|error| Error::Xml(format!("invalid Ribbon XML declaration: {error}")))?;
    let raw = BytesStart::from_content(declaration_text, 3);
    let mut state = 0u8;
    for attribute in raw.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.prefix().is_some() {
            return Err(Error::Invalid(format!(
                "unexpected Ribbon XML declaration attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
        state = match (state, attribute.key.as_ref()) {
            (0, b"version") => 1,
            (1, b"encoding") => 2,
            (1 | 2, b"standalone") => 3,
            _ => {
                return Err(Error::Invalid(format!(
                    "unexpected or out-of-order Ribbon XML declaration attribute '{}'",
                    String::from_utf8_lossy(attribute.key.as_ref())
                )));
            },
        };
        std::str::from_utf8(attribute.value.as_ref())
            .map_err(|error| Error::Xml(format!("invalid Ribbon XML declaration: {error}")))?;
    }
    if let Some(encoding) = declaration.encoding() {
        let encoding = encoding.map_err(|error| Error::Xml(error.to_string()))?;
        let encoding = std::str::from_utf8(&encoding)
            .map_err(|error| Error::Xml(format!("invalid Ribbon XML encoding: {error}")))?;
        if !valid_encoding_name(encoding) {
            return Err(Error::Invalid(format!(
                "Ribbon XML encoding '{encoding}' is not an EncName"
            )));
        }
        if !encoding.eq_ignore_ascii_case("UTF-8") {
            return Err(Error::Invalid(format!(
                "Ribbon XML encoding '{encoding}' is unsupported; Ribbon XML must be UTF-8"
            )));
        }
    }
    if let Some(standalone) = declaration.standalone() {
        let standalone = standalone.map_err(|error| Error::Xml(error.to_string()))?;
        if !matches!(standalone.as_ref(), b"yes" | b"no") {
            return Err(Error::Invalid(
                "Ribbon XML standalone must be 'yes' or 'no'".into(),
            ));
        }
    }
    Ok(())
}

fn validate_instruction(reader: &NsReader<&[u8]>, instruction: &BytesPI<'_>) -> Result<()> {
    let target = reader
        .decoder()
        .decode(instruction.target())
        .map_err(|error| Error::Xml(format!("invalid Ribbon instruction target: {error}")))?;
    if !valid_xml_name(&target) || target.eq_ignore_ascii_case("xml") {
        return Err(Error::Invalid(format!(
            "invalid Ribbon processing-instruction target '{target}'"
        )));
    }
    let content = reader
        .decoder()
        .decode(instruction.content())
        .map_err(|error| Error::Xml(format!("invalid Ribbon instruction content: {error}")))?;
    validate_xml_chars(&content)
}

fn validate_qname(value: &[u8], kind: &str) -> Result<()> {
    let value = std::str::from_utf8(value)
        .map_err(|error| Error::Xml(format!("invalid Ribbon {kind} name: {error}")))?;
    let mut components = value.split(':');
    let first = components.next().unwrap_or_default();
    let second = components.next();
    let valid = match second {
        Some(local) => valid_ncname(first) && valid_ncname(local) && components.next().is_none(),
        None => valid_ncname(first),
    };
    if valid {
        Ok(())
    } else {
        Err(Error::Invalid(format!(
            "invalid Ribbon {kind} QName '{value}'"
        )))
    }
}

fn validate_xml_chars(value: &str) -> Result<()> {
    if value.chars().all(is_xml_char) {
        Ok(())
    } else {
        Err(Error::Xml(
            "Ribbon XML contains a character forbidden by XML 1.0".into(),
        ))
    }
}

fn is_xml_char(character: char) -> bool {
    matches!(
        character,
        '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
    )
}

fn valid_encoding_name(value: &str) -> bool {
    let mut bytes = value.bytes();
    bytes.next().is_some_and(|byte| byte.is_ascii_alphabetic())
        && bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
}

fn valid_xml_name(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(is_name_start) && characters.all(is_name_character)
}

fn valid_ncname(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(is_ncname_start) && characters.all(is_ncname_character)
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

fn validate_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    family: Family,
) -> Result<Version> {
    if element.local_name().as_ref() != b"customUI" {
        return Err(Error::Invalid(
            "Ribbon XML root element must be customUI".into(),
        ));
    }
    let namespace = normalized_namespace(namespace, decoder, "root element")?;
    match (family, namespace.as_ref()) {
        (Family::Legacy, V2007_NAMESPACE) => Ok(Version::V2007),
        (Family::Modern, V2010_NAMESPACE) => Ok(Version::V2010),
        (Family::Modern, UI2_NAMESPACE) => Ok(Version::Ui2),
        _ => Err(Error::Invalid(
            "Ribbon XML root namespace does not match its package relationship".into(),
        )),
    }
}

fn count_node(nodes: &mut usize, limits: &Limits) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or(Error::Limit {
        resource: "Ribbon XML nodes",
        max: limits.nodes,
        actual: usize::MAX,
    })?;
    if *nodes > limits.nodes {
        return Err(Error::Limit {
            resource: "Ribbon XML nodes",
            max: limits.nodes,
            actual: *nodes,
        });
    }
    Ok(())
}

fn available_part_name(package: &OpcPackage, version: Version) -> Result<PackURI> {
    let existing = sorted_part_names(package)?;
    for suffix in 0..PART_NAME_ATTEMPTS {
        let path = if suffix == 0 {
            version.default_part().to_owned()
        } else {
            format!("/customUI/customUI{suffix}.xml")
        };
        let candidate = PackURI::new(&path)
            .map_err(|error| Error::Uri(format!("Ribbon part URI '{path}': {error}")))?;
        let folded = path.to_ascii_lowercase();
        if !part_name_conflicts(&existing, &folded) {
            package.validate_new_part_name(&candidate)?;
            return Ok(candidate);
        }
    }
    Err(Error::Invalid(
        "unable to allocate a unique Ribbon part name".into(),
    ))
}

fn sorted_part_names(package: &OpcPackage) -> Result<Vec<String>> {
    let mut names = Vec::with_capacity(package.part_count().min(MAX_PART_NAMES));
    let mut bytes = 0usize;
    for part in package.iter_parts() {
        let count = names.len().checked_add(1).ok_or(Error::Limit {
            resource: "Ribbon part-name allocation scan",
            max: MAX_PART_NAMES,
            actual: usize::MAX,
        })?;
        if count > MAX_PART_NAMES {
            return Err(Error::Limit {
                resource: "Ribbon part-name allocation scan",
                max: MAX_PART_NAMES,
                actual: count,
            });
        }
        bytes = bytes
            .checked_add(part.partname().as_str().len())
            .ok_or(Error::Limit {
                resource: "Ribbon part-name allocation bytes",
                max: MAX_PART_NAME_BYTES,
                actual: usize::MAX,
            })?;
        if bytes > MAX_PART_NAME_BYTES {
            return Err(Error::Limit {
                resource: "Ribbon part-name allocation bytes",
                max: MAX_PART_NAME_BYTES,
                actual: bytes,
            });
        }
        names.push(part.partname().as_str().to_ascii_lowercase());
    }
    names.sort_unstable();
    names.dedup();
    Ok(names)
}

fn part_name_conflicts(existing: &[String], candidate: &str) -> bool {
    if sorted_name_exists(existing, candidate) {
        return true;
    }
    for (index, _) in candidate.match_indices('/').skip(1) {
        if sorted_name_exists(existing, &candidate[..index]) {
            return true;
        }
    }
    let descendant_prefix = format!("{candidate}/");
    let position = existing.partition_point(|name| name.as_str() < descendant_prefix.as_str());
    existing
        .get(position)
        .is_some_and(|name| name.starts_with(&descendant_prefix))
}

fn sorted_name_exists(existing: &[String], wanted: &str) -> bool {
    existing
        .binary_search_by(|name| name.as_str().cmp(wanted))
        .is_ok()
}

fn ribbon_images(package: &OpcPackage, ribbon: &PackURI) -> Result<Vec<PackURI>> {
    let part = package.get_part(ribbon)?;
    let mut images = Vec::new();
    for relationship in part.rels().iter() {
        let target = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid Ribbon image target: {error}"))
        })?;
        let target = package.get_part(&target)?.partname();
        images.push(target.clone());
    }
    images.sort_unstable_by(compare_names);
    images.dedup_by(|left, right| same_name(left, right));
    Ok(images)
}

fn has_other_inbound(
    package: &OpcPackage,
    target: &PackURI,
    skipped_package_relationship: Option<&str>,
) -> Result<bool> {
    let mut scanned = 0usize;
    for relationship in package.rels().iter() {
        bump_graph_links(&mut scanned)?;
        if skipped_package_relationship == Some(relationship.r_id()) || relationship.is_external() {
            continue;
        }
        let related = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid package relationship target: {error}"))
        })?;
        if same_name(&related, target) {
            return Ok(true);
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            bump_graph_links(&mut scanned)?;
            if relationship.is_external() {
                continue;
            }
            let related = relationship.target_partname().map_err(|error| {
                Error::Relationship(format!(
                    "invalid relationship target from '{}': {error}",
                    source.partname().as_str()
                ))
            })?;
            if same_name(&related, target) {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn removable_images(
    package: &OpcPackage,
    removed_ribbon: &PackURI,
    candidates: &[PackURI],
) -> Result<Vec<PackURI>> {
    let names = NameIndex::new(candidates);
    let mut keep = vec![false; candidates.len()];
    let mut outgoing = vec![Vec::new(); candidates.len()];
    let mut scanned = 0usize;
    let mut edges = 0usize;

    for relationship in package.rels().iter() {
        bump_graph_links(&mut scanned)?;
        if relationship.is_external() {
            continue;
        }
        let target = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid package relationship target: {error}"))
        })?;
        if let Some(index) = names.get(&target) {
            keep[index] = true;
        }
    }

    for source in package.iter_parts() {
        if same_name(source.partname(), removed_ribbon) {
            continue;
        }
        let source_index = names.get(source.partname());
        for relationship in source.rels().iter() {
            bump_graph_links(&mut scanned)?;
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname().map_err(|error| {
                Error::Relationship(format!(
                    "invalid relationship target from '{}': {error}",
                    source.partname().as_str()
                ))
            })?;
            let Some(target_index) = names.get(&target) else {
                continue;
            };
            match source_index {
                Some(source_index) => {
                    edges = edges.checked_add(1).ok_or(Error::Limit {
                        resource: "Ribbon image garbage-collection edges",
                        max: MAX_IMAGE_GC_EDGES,
                        actual: usize::MAX,
                    })?;
                    if edges > MAX_IMAGE_GC_EDGES {
                        return Err(Error::Limit {
                            resource: "Ribbon image garbage-collection edges",
                            max: MAX_IMAGE_GC_EDGES,
                            actual: edges,
                        });
                    }
                    outgoing[source_index].push(target_index);
                },
                None => keep[target_index] = true,
            }
        }
    }

    for targets in &mut outgoing {
        targets.sort_unstable();
        targets.dedup();
    }

    let mut pending: VecDeque<_> = keep
        .iter()
        .enumerate()
        .filter_map(|(index, keep)| keep.then_some(index))
        .collect();
    while let Some(source) = pending.pop_front() {
        for &target in &outgoing[source] {
            if !keep[target] {
                keep[target] = true;
                pending.push_back(target);
            }
        }
    }

    Ok(candidates
        .iter()
        .zip(keep)
        .filter(|(_, keep)| !keep)
        .map(|(candidate, _)| candidate.clone())
        .collect())
}

fn bump_graph_links(scanned: &mut usize) -> Result<()> {
    *scanned = scanned.checked_add(1).ok_or(Error::Limit {
        resource: "Ribbon package graph relationships",
        max: MAX_GRAPH_LINKS,
        actual: usize::MAX,
    })?;
    if *scanned > MAX_GRAPH_LINKS {
        return Err(Error::Limit {
            resource: "Ribbon package graph relationships",
            max: MAX_GRAPH_LINKS,
            actual: *scanned,
        });
    }
    Ok(())
}

struct NameIndex<'a> {
    values: &'a [PackURI],
    order: Vec<usize>,
}

impl<'a> NameIndex<'a> {
    fn new(values: &'a [PackURI]) -> Self {
        let mut order: Vec<_> = (0..values.len()).collect();
        order.sort_unstable_by(|left, right| compare_names(&values[*left], &values[*right]));
        Self { values, order }
    }

    fn get(&self, wanted: &PackURI) -> Option<usize> {
        self.order
            .binary_search_by(|index| compare_names(&self.values[*index], wanted))
            .ok()
            .map(|position| self.order[position])
    }
}

fn compare_names(left: &PackURI, right: &PackURI) -> Ordering {
    left.as_str()
        .bytes()
        .map(|byte| byte.to_ascii_lowercase())
        .cmp(right.as_str().bytes().map(|byte| byte.to_ascii_lowercase()))
}

fn same_name(left: &PackURI, right: &PackURI) -> bool {
    left.is_equivalent_to(right)
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::BlobPart;
    use litchi_opc::constants::relationship_type as rt;
    use std::sync::Arc;

    const XML_2007: &[u8] =
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#;
    const XML_2010: &[u8] =
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"/>"#;
    const XML_UI2: &[u8] =
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2007/10/customui"/>"#;

    #[test]
    fn fixed_slots_borrow_package_bytes_and_prefer_modern() {
        let mut package = OpcPackage::new();
        put(&mut package, Version::V2007, XML_2007.to_vec()).unwrap();
        put(&mut package, Version::V2010, XML_2010.to_vec()).unwrap();

        let ribbons = load(&package).unwrap();
        let legacy = ribbons.legacy().unwrap();
        let modern = ribbons.modern().unwrap();
        assert_eq!(ribbons.iter().collect::<Vec<_>>(), [legacy, modern]);
        assert_eq!(ribbons.effective(), Some(modern));
        assert_eq!(legacy.version(), Version::V2007);
        assert_eq!(modern.version(), Version::V2010);
        assert_eq!(
            modern.xml().as_ptr(),
            package.get_part(modern.part()).unwrap().blob().as_ptr()
        );
    }

    #[test]
    fn modern_vocabulary_updates_in_place() {
        let mut package = OpcPackage::new();
        put(&mut package, Version::Ui2, XML_UI2.to_vec()).unwrap();
        let part = load(&package).unwrap().modern().unwrap().part().clone();

        put(&mut package, Version::V2010, XML_2010.to_vec()).unwrap();
        let modern = load(&package).unwrap().modern().unwrap();
        assert_eq!(modern.part(), &part);
        assert_eq!(modern.version(), Version::V2010);
        assert_eq!(package.part_count(), 1);
    }

    #[test]
    fn shared_ribbon_is_forked_before_update() {
        let mut package = OpcPackage::new();
        let original = raw_ribbon(&mut package, Version::V2007, "/addons/ui.xml", XML_2007);
        let image = add_image(&mut package, "/addons/images/icon.png", "image/png");
        package
            .get_part_mut(&original)
            .unwrap()
            .relate_to("images/icon.png", rt::IMAGE);
        let original_bytes = package.get_part(&original).unwrap().blob_arc();
        let source = add_source(&mut package, "/word/document.xml");
        package
            .get_part_mut(&source)
            .unwrap()
            .relate_to("../addons/ui.xml", "urn:shared-ribbon");
        let relationship_id = load(&package).unwrap().legacy().unwrap().id().to_owned();
        let replacement =
            br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon/></customUI>"#
                .to_vec();
        let allocation = replacement.as_ptr();

        put(&mut package, Version::V2007, replacement).unwrap();

        let updated = load(&package).unwrap().legacy().unwrap();
        assert_ne!(updated.part(), &original);
        assert_eq!(updated.part().as_str(), "/customUI/customUI.xml");
        assert_eq!(updated.id(), relationship_id);
        assert_eq!(updated.xml().as_ptr(), allocation);
        assert!(Arc::ptr_eq(
            &original_bytes,
            &package.get_part(&original).unwrap().blob_arc()
        ));
        assert_eq!(package.get_part(&original).unwrap().blob(), XML_2007);
        let updated_part = package.get_part(updated.part()).unwrap();
        assert_eq!(updated_part.rels().len(), 1);
        let image_relationship = updated_part.rels().iter().next().unwrap();
        assert_eq!(image_relationship.target_ref(), "../addons/images/icon.png");
        assert_eq!(image_relationship.target_partname().unwrap(), image);
    }

    #[test]
    fn identical_put_is_a_signature_and_allocation_preserving_noop() {
        let mut package = OpcPackage::new();
        put(&mut package, Version::V2007, XML_2007.to_vec()).unwrap();
        let part = load(&package).unwrap().legacy().unwrap().part().clone();
        let before = package.get_part(&part).unwrap().blob_arc();
        sign_marker(&mut package);

        put(&mut package, Version::V2007, XML_2007.to_vec()).unwrap();

        assert!(package.is_signed());
        let after = package.get_part(&part).unwrap().blob_arc();
        assert!(Arc::ptr_eq(&before, &after));

        put(
            &mut package,
            Version::V2007,
            br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon/></customUI>"#
                .to_vec(),
        )
        .unwrap();
        assert!(!package.is_signed());
    }

    #[test]
    fn name_collisions_do_not_copy_the_moved_payload() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/CUSTOMui/CUSTOMui.XML").unwrap(),
            "application/octet-stream".into(),
            vec![7],
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/customUI/customUI1.xml/child").unwrap(),
            "application/octet-stream".into(),
            vec![8],
        )));
        let xml = XML_2007.to_vec();
        let allocation = xml.as_ptr();

        put(&mut package, Version::V2007, xml).unwrap();

        let ui = load(&package).unwrap().legacy().unwrap();
        assert_eq!(ui.part().as_str(), "/customUI/customUI2.xml");
        assert_eq!(ui.xml().as_ptr(), allocation);
    }

    #[test]
    fn xml_bytes_depth_nodes_entities_and_namespaces_are_bounded() {
        let mut package = OpcPackage::new();
        let tiny_bytes = Limits {
            xml_bytes: XML_2007.len() - 1,
            ..Limits::standard()
        };
        assert!(matches!(
            put_with(&mut package, Version::V2007, XML_2007.to_vec(), &tiny_bytes),
            Err(Error::Limit { .. })
        ));
        assert_eq!(package.part_count(), 0);

        let shallow = Limits {
            depth: 1,
            ..Limits::standard()
        };
        let nested = br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon/></customUI>"#;
        assert!(matches!(
            put_with(&mut package, Version::V2007, nested.to_vec(), &shallow),
            Err(Error::Limit { .. })
        ));

        let few_nodes = Limits {
            nodes: 1,
            ..Limits::standard()
        };
        assert!(matches!(
            put_with(&mut package, Version::V2007, XML_2007.to_vec(), &few_nodes),
            Err(Error::Limit { .. })
        ));
        assert!(put(
            &mut package,
            Version::V2007,
            br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">&madeUp;</customUI>"#
                .to_vec()
        )
        .is_err());
        assert!(put(
            &mut package,
            Version::V2007,
            br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><bad:node/></customUI>"#
                .to_vec()
        )
        .is_err());
        assert!(put(
            &mut package,
            Version::V2007,
            br#"<!DOCTYPE customUI><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#
                .to_vec()
        )
        .is_err());
        assert_eq!(package.part_count(), 0);
    }

    #[test]
    fn xml_declarations_characters_namespaces_and_inert_instructions_are_strict() {
        let mut valid = OpcPackage::new();
        let normalized_namespace = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><?safe inert?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customu&#x69;"><?inside retained?></customUI>"#;
        put(&mut valid, Version::V2007, normalized_namespace.to_vec()).unwrap();
        assert_eq!(
            load(&valid).unwrap().legacy().unwrap().xml(),
            normalized_namespace
        );

        for invalid in [
            br#"<?xml?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#
                .as_slice(),
            br#"<?xml version="1.1"?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#
                .as_slice(),
            br#"<?xml version="1.0" encoding="UTF-16"?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#
                .as_slice(),
            b"<?xml version=\"1.0\" encoding=\"US-ASCII\"?><customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\" label=\"caf\xC3\xA9\"/>"
                .as_slice(),
            br#"<!--before--><?xml version="1.0"?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#
                .as_slice(),
            br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">&#1;</customUI>"#
                .as_slice(),
            br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" label="&#xB;"/>"#
                .as_slice(),
            br#"<?xml illegal?><customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#
                .as_slice(),
        ] {
            let mut package = OpcPackage::new();
            assert!(put(&mut package, Version::V2007, invalid.to_vec()).is_err());
            assert_eq!(package.part_count(), 0);
        }

        let mut duplicate_expanded = OpcPackage::new();
        assert!(
            put(
                &mut duplicate_expanded,
                Version::V2007,
                br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" xmlns:a="urn:x" xmlns:b="urn:&#x78;" a:value="1" b:value="2"/>"#
                    .to_vec(),
            )
            .is_err()
        );

        let mut raw_control =
            br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">"#.to_vec();
        raw_control.push(1);
        raw_control.extend_from_slice(b"</customUI>");
        let mut package = OpcPackage::new();
        assert!(put(&mut package, Version::V2007, raw_control).is_err());
    }

    #[test]
    fn relationship_location_cardinality_and_root_family_are_strict() {
        let mut external = OpcPackage::new();
        external.relate_to_external("https://example.invalid/ui.xml", LEGACY_RELATIONSHIP);
        assert!(matches!(load(&external), Err(Error::Relationship(_))));

        let mut duplicate = OpcPackage::new();
        raw_ribbon(
            &mut duplicate,
            Version::V2010,
            "/customUI/one.xml",
            XML_2010,
        );
        raw_ribbon(&mut duplicate, Version::Ui2, "/customUI/two.xml", XML_UI2);
        assert!(matches!(load(&duplicate), Err(Error::Relationship(_))));

        let mut mismatched = OpcPackage::new();
        raw_ribbon(
            &mut mismatched,
            Version::V2007,
            "/customUI/customUI.xml",
            XML_2010,
        );
        assert!(matches!(load(&mismatched), Err(Error::Invalid(_))));

        let mut part_sourced = OpcPackage::new();
        let source = PackURI::new("/word/document.xml").unwrap();
        part_sourced.add_part(Box::new(XmlPart::new(
            source.clone(),
            "application/xml".into(),
            b"<document/>".to_vec(),
        )));
        part_sourced
            .get_part_mut(&source)
            .unwrap()
            .relate_to("../customUI/customUI.xml", LEGACY_RELATIONSHIP);
        assert!(matches!(load(&part_sourced), Err(Error::Relationship(_))));
    }

    #[test]
    fn ribbon_outbound_relationships_are_internal_images_only() {
        let mut wrong_type = package_with_legacy();
        let ribbon = legacy_part(&wrong_type);
        let data = PackURI::new("/customUI/data.bin").unwrap();
        wrong_type.add_part(Box::new(BlobPart::new(
            data,
            "application/octet-stream".into(),
            vec![1],
        )));
        wrong_type
            .get_part_mut(&ribbon)
            .unwrap()
            .relate_to("data.bin", "urn:not-an-image");
        assert!(matches!(load(&wrong_type), Err(Error::Relationship(_))));

        let mut wrong_media = package_with_legacy();
        let ribbon = legacy_part(&wrong_media);
        let data = PackURI::new("/customUI/image.bin").unwrap();
        wrong_media.add_part(Box::new(BlobPart::new(
            data,
            "application/octet-stream".into(),
            vec![1],
        )));
        wrong_media
            .get_part_mut(&ribbon)
            .unwrap()
            .relate_to("image.bin", rt::IMAGE);
        assert!(matches!(load(&wrong_media), Err(Error::ContentType { .. })));

        let mut external = package_with_legacy();
        let ribbon = legacy_part(&external);
        external
            .get_part_mut(&ribbon)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::IMAGE.into(),
                "https://example.invalid/image.png".into(),
                "rIdImage".into(),
                true,
            );
        assert!(matches!(load(&external), Err(Error::Relationship(_))));

        let mut queried = package_with_legacy();
        let ribbon = legacy_part(&queried);
        add_image(&mut queried, "/customUI/image.png", "image/png");
        queried
            .get_part_mut(&ribbon)
            .unwrap()
            .relate_to("image.png?variant=2", rt::IMAGE);
        assert!(matches!(load(&queried), Err(Error::Relationship(_))));

        for content_type in ["image/ png", "image/png;garbage"] {
            let mut malformed = package_with_legacy();
            let ribbon = legacy_part(&malformed);
            add_image(&mut malformed, "/customUI/image.png", content_type);
            malformed
                .get_part_mut(&ribbon)
                .unwrap()
                .relate_to("image.png", rt::IMAGE);
            assert!(matches!(load(&malformed), Err(Error::ContentType { .. })));
        }
    }

    #[test]
    fn aggregate_image_relationships_are_bounded() {
        let mut package = package_with_legacy();
        let ribbon = legacy_part(&package);
        add_image(&mut package, "/customUI/image.png", "IMAGE/PNG");
        package
            .get_part_mut(&ribbon)
            .unwrap()
            .relate_to("image.png", rt::STRICT_IMAGE);
        let limits = Limits {
            images: 0,
            ..Limits::standard()
        };
        assert!(matches!(
            load_with(&package, &limits),
            Err(Error::Limit { .. })
        ));
    }

    #[test]
    fn remove_collects_unreferenced_ribbon_and_image_parts() {
        let mut package = package_with_legacy();
        let ribbon = legacy_part(&package);
        let image = add_image(&mut package, "/customUI/image.png", "image/png");
        package
            .get_part_mut(&ribbon)
            .unwrap()
            .relate_to("image.png", rt::IMAGE);
        sign_marker(&mut package);

        assert!(remove(&mut package, Family::Legacy).unwrap());
        assert!(package.get_part(&ribbon).is_err());
        assert!(package.get_part(&image).is_err());
        assert!(!package.is_signed());
        assert!(!remove(&mut package, Family::Legacy).unwrap());
    }

    #[test]
    fn remove_preserves_shared_ribbon_and_image_parts() {
        let mut shared_image = package_with_legacy();
        let ribbon = legacy_part(&shared_image);
        let image = add_image(&mut shared_image, "/customUI/image.png", "image/png");
        shared_image
            .get_part_mut(&ribbon)
            .unwrap()
            .relate_to("image.png", rt::IMAGE);
        let source = add_source(&mut shared_image, "/word/document.xml");
        shared_image
            .get_part_mut(&source)
            .unwrap()
            .relate_to("../customUI/image.png", rt::IMAGE);
        assert!(remove(&mut shared_image, Family::Legacy).unwrap());
        assert!(shared_image.get_part(&ribbon).is_err());
        assert!(shared_image.get_part(&image).is_ok());

        let mut shared_ribbon = package_with_legacy();
        let ribbon = legacy_part(&shared_ribbon);
        let image = add_image(&mut shared_ribbon, "/customUI/image.png", "image/png");
        shared_ribbon
            .get_part_mut(&ribbon)
            .unwrap()
            .relate_to("image.png", rt::IMAGE);
        let source = add_source(&mut shared_ribbon, "/word/document.xml");
        shared_ribbon
            .get_part_mut(&source)
            .unwrap()
            .relate_to("../customUI/customUI.xml", "urn:shared-ribbon");
        assert!(remove(&mut shared_ribbon, Family::Legacy).unwrap());
        assert!(shared_ribbon.get_part(&ribbon).is_ok());
        assert!(shared_ribbon.get_part(&image).is_ok());
    }

    #[test]
    fn image_cycles_are_collected_unless_reachable_from_a_kept_part() {
        let mut unanchored = package_with_image_cycle();
        let first = PackURI::new("/customUI/one.png").unwrap();
        let second = PackURI::new("/customUI/two.png").unwrap();
        remove(&mut unanchored, Family::Legacy).unwrap();
        assert!(unanchored.get_part(&first).is_err());
        assert!(unanchored.get_part(&second).is_err());

        let mut anchored = package_with_image_cycle();
        anchored.relate_to("customUI/one.png", "urn:keep-image");
        remove(&mut anchored, Family::Legacy).unwrap();
        assert!(anchored.get_part(&first).is_ok());
        assert!(anchored.get_part(&second).is_ok());
    }

    #[test]
    fn failed_mutations_leave_graph_bytes_and_signatures_untouched() {
        let mut package = package_with_legacy();
        let ribbon = legacy_part(&package);
        let before = package.get_part(&ribbon).unwrap().blob_arc();
        sign_marker(&mut package);
        assert!(put(&mut package, Version::V2010, b"<broken".to_vec()).is_err());
        assert!(package.is_signed());
        assert!(Arc::ptr_eq(
            &before,
            &package.get_part(&ribbon).unwrap().blob_arc()
        ));

        let data = PackURI::new("/customUI/data.bin").unwrap();
        package.add_part(Box::new(BlobPart::new(
            data,
            "application/octet-stream".into(),
            vec![1],
        )));
        package
            .get_part_mut(&ribbon)
            .unwrap()
            .relate_to("data.bin", "urn:not-an-image");
        assert!(remove(&mut package, Family::Legacy).is_err());
        assert!(package.is_signed());
        assert!(package.get_part(&ribbon).is_ok());
        assert!(package.rels().iter().any(|relationship| {
            Family::from_relationship(relationship.reltype()) == Some(Family::Legacy)
        }));

        let mut absent = OpcPackage::new();
        sign_marker(&mut absent);
        assert!(!remove(&mut absent, Family::Modern).unwrap());
        assert!(absent.is_signed());
    }

    fn package_with_legacy() -> OpcPackage {
        let mut package = OpcPackage::new();
        put(&mut package, Version::V2007, XML_2007.to_vec()).unwrap();
        package
    }

    fn package_with_image_cycle() -> OpcPackage {
        let mut package = package_with_legacy();
        let ribbon = legacy_part(&package);
        let first = add_image(&mut package, "/customUI/one.png", "image/png");
        let second = add_image(&mut package, "/customUI/two.png", "image/png");
        package
            .get_part_mut(&ribbon)
            .unwrap()
            .relate_to("one.png", rt::IMAGE);
        package
            .get_part_mut(&ribbon)
            .unwrap()
            .relate_to("two.png", rt::IMAGE);
        package
            .get_part_mut(&first)
            .unwrap()
            .relate_to("two.png", rt::IMAGE);
        package
            .get_part_mut(&second)
            .unwrap()
            .relate_to("one.png", rt::IMAGE);
        package
    }

    fn raw_ribbon(package: &mut OpcPackage, version: Version, name: &str, xml: &[u8]) -> PackURI {
        let part = PackURI::new(name).unwrap();
        package.add_part(Box::new(XmlPart::new(
            part.clone(),
            CONTENT_TYPE.into(),
            xml.to_vec(),
        )));
        package.relate_to(name.trim_start_matches('/'), version.relationship());
        part
    }

    fn legacy_part(package: &OpcPackage) -> PackURI {
        load(package).unwrap().legacy().unwrap().part().clone()
    }

    fn add_image(package: &mut OpcPackage, name: &str, content_type: &str) -> PackURI {
        let part = PackURI::new(name).unwrap();
        package.add_part(Box::new(BlobPart::new(
            part.clone(),
            content_type.into(),
            vec![1, 2, 3],
        )));
        part
    }

    fn add_source(package: &mut OpcPackage, name: &str) -> PackURI {
        let part = PackURI::new(name).unwrap();
        package.add_part(Box::new(XmlPart::new(
            part.clone(),
            "application/xml".into(),
            b"<source/>".to_vec(),
        )));
        part
    }

    fn sign_marker(package: &mut OpcPackage) {
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        assert!(package.is_signed());
    }
}
