//! Typed, bounded DrawingML table-style catalogs and transactional package CRUD.

use crate::{Error, Result};
use bitflags::bitflags;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI, TargetMode};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;
use std::fmt::{self, Write as _};
use std::ops::Range;
use std::str::FromStr;

const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const AS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/tableStyles";
const DEFAULT_XML: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
    r#"def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>"#,
);

const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_PRESENTATION_BYTES: usize = 32 * 1024 * 1024;
const MAX_NODES: usize = 250_000;
const MAX_DEPTH: usize = 128;
const MAX_ATTRIBUTES: usize = 64;
const MAX_ATTRIBUTE_BYTES: usize = 4_096;
const MAX_STYLES: usize = 4_096;
const MAX_GRAPH_PARTS: usize = 65_536;
const MAX_GRAPH_RELATIONSHIPS: usize = 262_144;
const PART_NAME_ATTEMPTS: usize = 4_096;
const RELATIONSHIP_ID_ATTEMPTS: usize = 65_536;

/// Namespace and relationship profile used by a table-style catalog.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    fn drawing(self) -> &'static str {
        match self {
            Self::Transitional => A,
            Self::Strict => AS,
        }
    }

    fn relationship(self) -> &'static str {
        match self {
            Self::Transitional => rt::TABLE_STYLES,
            Self::Strict => STRICT_REL,
        }
    }

    fn office_document(self) -> &'static str {
        match self {
            Self::Transitional => rt::OFFICE_DOCUMENT,
            Self::Strict => rt::STRICT_OFFICE_DOCUMENT,
        }
    }
}

/// A validated DrawingML table-style GUID stored without heap allocation.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Id([u8; 16]);

impl Id {
    /// Parse the required braced GUID wire form.
    pub fn parse(value: &str) -> Result<Self> {
        value.parse()
    }

    fn write_to(self, output: &mut String) -> Result<()> {
        output
            .try_reserve(38)
            .map_err(|source| allocation("table-style GUID encoding", source))?;
        write!(
            output,
            "{{{:02X}{:02X}{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
            self.0[0],
            self.0[1],
            self.0[2],
            self.0[3],
            self.0[4],
            self.0[5],
            self.0[6],
            self.0[7],
            self.0[8],
            self.0[9],
            self.0[10],
            self.0[11],
            self.0[12],
            self.0[13],
            self.0[14],
            self.0[15],
        )
        .map_err(|_| Error::Write)
    }
}

impl FromStr for Id {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        let bytes = value.as_bytes();
        if bytes.len() != 38 || bytes.first() != Some(&b'{') || bytes.last() != Some(&b'}') {
            return Err(invalid("table-style ID must be a braced GUID"));
        }
        for position in [9usize, 14, 19, 24] {
            if bytes.get(position) != Some(&b'-') {
                return Err(invalid("table-style ID has invalid GUID separators"));
            }
        }
        let mut decoded = [0u8; 16];
        let mut source = 1usize;
        for byte in &mut decoded {
            while matches!(source, 9 | 14 | 19 | 24) {
                source += 1;
            }
            let high = hex(*bytes
                .get(source)
                .ok_or_else(|| invalid("short table-style GUID"))?)
            .ok_or_else(|| invalid("table-style ID contains a non-hex digit"))?;
            let low = hex(*bytes
                .get(source + 1)
                .ok_or_else(|| invalid("short table-style GUID"))?)
            .ok_or_else(|| invalid("table-style ID contains a non-hex digit"))?;
            *byte = (high << 4) | low;
            source += 2;
        }
        Ok(Self(decoded))
    }
}

impl fmt::Display for Id {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "{{{:02X}{:02X}{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
            self.0[0],
            self.0[1],
            self.0[2],
            self.0[3],
            self.0[4],
            self.0[5],
            self.0[6],
            self.0[7],
            self.0[8],
            self.0[9],
            self.0[10],
            self.0[11],
            self.0[12],
            self.0[13],
            self.0[14],
            self.0[15],
        )
    }
}

impl fmt::Debug for Id {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::Display::fmt(self, formatter)
    }
}

bitflags! {
    /// Conditional table regions defined by one style, packed into two bytes.
    #[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
    pub struct Parts: u16 {
        const WHOLE = 1 << 0;
        const ODD_ROW = 1 << 1;
        const EVEN_ROW = 1 << 2;
        const ODD_COLUMN = 1 << 3;
        const EVEN_COLUMN = 1 << 4;
        const FIRST_COLUMN = 1 << 5;
        const LAST_COLUMN = 1 << 6;
        const FIRST_ROW = 1 << 7;
        const LAST_ROW = 1 << 8;
        const SOUTH_EAST = 1 << 9;
        const SOUTH_WEST = 1 << 10;
        const NORTH_EAST = 1 << 11;
        const NORTH_WEST = 1 << 12;
        const BACKGROUND = 1 << 13;
    }
}

impl Parts {
    /// Return the DrawingML element name for one single-region flag.
    pub fn xml_name(self) -> Option<&'static str> {
        PARTS
            .iter()
            .find_map(|(part, name)| (*part == self).then_some(*name))
    }

    fn from_xml_name(name: &[u8]) -> Option<Self> {
        PARTS
            .iter()
            .find_map(|(part, candidate)| (candidate.as_bytes() == name).then_some(*part))
    }
}

const PARTS: [(Parts, &str); 14] = [
    (Parts::BACKGROUND, "tblBg"),
    (Parts::WHOLE, "wholeTbl"),
    (Parts::ODD_ROW, "band1H"),
    (Parts::EVEN_ROW, "band2H"),
    (Parts::ODD_COLUMN, "band1V"),
    (Parts::EVEN_COLUMN, "band2V"),
    (Parts::LAST_COLUMN, "lastCol"),
    (Parts::FIRST_COLUMN, "firstCol"),
    (Parts::LAST_ROW, "lastRow"),
    (Parts::SOUTH_EAST, "seCell"),
    (Parts::SOUTH_WEST, "swCell"),
    (Parts::FIRST_ROW, "firstRow"),
    (Parts::NORTH_EAST, "neCell"),
    (Parts::NORTH_WEST, "nwCell"),
];

#[derive(Debug)]
struct Attr {
    name: String,
    value: String,
}

#[derive(Debug)]
enum Payload {
    Shared {
        raw: Range<usize>,
        body: Range<usize>,
        exact: bool,
    },
    Owned {
        xml: Vec<u8>,
        body: Range<usize>,
        exact: bool,
    },
}

/// One typed `a:tblStyle` definition.
///
/// Loaded definitions retain their complete inert XML payload. Renaming a
/// definition preserves that payload; [`Self::reset_parts`] deliberately
/// replaces detailed cell formatting with empty region declarations.
pub struct Def {
    id: Id,
    name: String,
    parts: Parts,
    attrs: Vec<Attr>,
    payload: Payload,
}

impl fmt::Debug for Def {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Def")
            .field("id", &self.id)
            .field("name", &self.name)
            .field("parts", &self.parts)
            .finish_non_exhaustive()
    }
}

impl Def {
    /// Create a style with no conditional region payloads.
    pub fn new(id: Id, name: impl Into<String>) -> Result<Self> {
        let name = name.into();
        validate_name(&name)?;
        Ok(Self {
            id,
            name,
            parts: Parts::empty(),
            attrs: Vec::new(),
            payload: Payload::Owned {
                xml: Vec::new(),
                body: 0..0,
                exact: false,
            },
        })
    }

    pub fn id(&self) -> Id {
        self.id
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn parts(&self) -> Parts {
        self.parts
    }

    pub fn has(&self, parts: Parts) -> bool {
        self.parts.contains(parts)
    }

    /// Rename this detached definition while preserving its cell-style body.
    pub fn rename(&mut self, name: impl Into<String>) -> Result<String> {
        let name = name.into();
        validate_name(&name)?;
        if self.name == name {
            return Ok(name);
        }
        invalidate_payload(&mut self.payload);
        Ok(std::mem::replace(&mut self.name, name))
    }

    /// Replace detailed cell formatting with the selected empty regions.
    ///
    /// The explicit `reset` name makes the destructive payload change visible
    /// at the call site. Existing opaque formatting is otherwise preserved.
    pub fn reset_parts(&mut self, parts: Parts) -> Parts {
        let previous = self.parts;
        self.parts = parts;
        self.payload = Payload::Owned {
            xml: Vec::new(),
            body: 0..0,
            exact: false,
        };
        previous
    }

    fn materialize(&mut self, source: &[u8]) -> Result<()> {
        let Payload::Shared { raw, body, exact } = &self.payload else {
            return Ok(());
        };
        let raw_bytes = source
            .get(raw.clone())
            .ok_or_else(|| invalid("table-style source range is invalid"))?;
        let body_start = body
            .start
            .checked_sub(raw.start)
            .ok_or_else(|| invalid("table-style body precedes its element"))?;
        let body_end = body
            .end
            .checked_sub(raw.start)
            .ok_or_else(|| invalid("table-style body precedes its element"))?;
        let mut xml = Vec::new();
        xml.try_reserve_exact(raw_bytes.len())
            .map_err(|source| allocation("detached table-style XML", source))?;
        xml.extend_from_slice(raw_bytes);
        let exact = *exact;
        self.payload = Payload::Owned {
            xml,
            body: body_start..body_end,
            exact,
        };
        Ok(())
    }
}

/// Ordered table-style catalog (`a:tblStyleLst`).
///
/// The facade keeps GUIDs typed, region settings compact, and loaded XML
/// source-backed. An unchanged load→put moves the original producer bytes
/// back into OPC without normalization.
pub struct List {
    conformance: Conformance,
    default: Id,
    defs: Vec<Def>,
    root_attrs: Vec<Attr>,
    source: Vec<u8>,
    dirty: bool,
}

impl fmt::Debug for List {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("List")
            .field("conformance", &self.conformance)
            .field("default", &self.default)
            .field("defs", &self.defs)
            .field("source_bytes", &self.source.len())
            .field("dirty", &self.dirty)
            .finish()
    }
}

impl List {
    /// Create an empty catalog with an explicitly selected default style.
    pub fn new(conformance: Conformance, default: Id) -> Self {
        Self {
            conformance,
            default,
            defs: Vec::new(),
            root_attrs: Vec::new(),
            source: Vec::new(),
            dirty: true,
        }
    }

    /// Parse and take ownership of bounded table-style XML.
    pub fn parse(xml: impl Into<Vec<u8>>) -> Result<Self> {
        parse_owned(xml.into())
    }

    pub fn conformance(&self) -> Conformance {
        self.conformance
    }

    pub fn default(&self) -> Id {
        self.default
    }

    pub fn set_default(&mut self, id: Id) -> Id {
        if self.default == id {
            return id;
        }
        self.dirty = true;
        std::mem::replace(&mut self.default, id)
    }

    pub fn styles(&self) -> &[Def] {
        &self.defs
    }

    pub fn len(&self) -> usize {
        self.defs.len()
    }

    pub fn is_empty(&self) -> bool {
        self.defs.is_empty()
    }

    /// Checked raw-position lookup for ordered inspection.
    pub fn at(&self, index: usize) -> Option<&Def> {
        self.defs.get(index)
    }

    /// Preferred stable-identity lookup.
    pub fn get(&self, id: Id) -> Option<&Def> {
        self.defs.iter().find(|style| style.id == id)
    }

    /// Return every definition with this non-identity display name.
    ///
    /// DrawingML permits duplicate and empty `styleName` values, so this
    /// method deliberately returns all matches rather than selecting one.
    pub fn named<'a>(&'a self, name: &'a str) -> impl Iterator<Item = &'a Def> + 'a {
        self.defs.iter().filter(move |style| style.name == name)
    }

    pub fn add(&mut self, style: Def) -> Result<()> {
        validate_def(&style)?;
        self.ensure_unique_id(style.id, None)?;
        if self.defs.len() >= MAX_STYLES {
            return Err(limit("table-style count", MAX_STYLES));
        }
        self.defs
            .try_reserve(1)
            .map_err(|source| allocation("table-style insertion", source))?;
        self.defs.push(style);
        self.dirty = true;
        Ok(())
    }

    /// Rename one style by stable ID while retaining its opaque formatting.
    pub fn rename(&mut self, id: Id, name: impl Into<String>) -> Result<String> {
        let name = name.into();
        validate_name(&name)?;
        let style = self
            .defs
            .iter_mut()
            .find(|style| style.id == id)
            .ok_or_else(|| invalid(format!("table style {id} was not found")))?;
        if style.name == name {
            return Ok(name);
        }
        invalidate_payload(&mut style.payload);
        let previous = std::mem::replace(&mut style.name, name);
        self.dirty = true;
        Ok(previous)
    }

    /// Replace one style while retaining the selected stable GUID.
    pub fn replace(&mut self, id: Id, mut replacement: Def) -> Result<Def> {
        replacement.id = id;
        validate_def(&replacement)?;
        self.ensure_unique_id(id, Some(id))?;
        let style = self
            .defs
            .iter_mut()
            .find(|style| style.id == id)
            .ok_or_else(|| invalid(format!("table style {id} was not found")))?;
        style.materialize(&self.source)?;
        let previous = std::mem::replace(style, replacement);
        self.dirty = true;
        Ok(previous)
    }

    /// Remove one non-default style by stable GUID.
    pub fn remove(&mut self, id: Id) -> Result<Option<Def>> {
        if id == self.default {
            return Err(invalid(
                "cannot remove the selected default table style; select another default first",
            ));
        }
        let Some(position) = self.defs.iter().position(|style| style.id == id) else {
            return Ok(None);
        };
        self.defs[position].materialize(&self.source)?;
        let removed = self.defs.remove(position);
        self.dirty = true;
        Ok(Some(removed))
    }

    /// Return original producer XML when the list has not been edited.
    pub fn source_xml(&self) -> Option<&[u8]> {
        (!self.dirty).then_some(self.source.as_slice())
    }

    /// Consume and encode the catalog, moving unchanged source bytes directly.
    pub fn into_xml(self) -> Result<Vec<u8>> {
        validate_list(&self)?;
        if !self.dirty {
            return Ok(self.source);
        }
        let xml = encode(&self)?;
        let parsed = scan(&xml)?;
        if parsed.conformance != self.conformance
            || parsed.default != self.default
            || parsed.defs.len() != self.defs.len()
            || parsed.defs.iter().zip(&self.defs).any(|(left, right)| {
                left.id != right.id || left.name != right.name || left.parts != right.parts
            })
        {
            return Err(invalid("encoded table-style catalog did not round-trip"));
        }
        Ok(xml)
    }

    fn ensure_unique_id(&self, id: Id, except: Option<Id>) -> Result<()> {
        if self
            .defs
            .iter()
            .any(|style| style.id == id && Some(style.id) != except)
        {
            return Err(invalid(format!("duplicate table-style ID {id}")));
        }
        Ok(())
    }
}

/// Return the deterministic Transitional default catalog used by new decks.
pub fn default_xml() -> &'static str {
    DEFAULT_XML
}

/// A borrowed, fully validated presentation relationship to a table-style
/// catalog.
#[derive(Clone, Copy, Debug)]
pub struct Link<'a> {
    id: &'a str,
    kind: &'a str,
    target: &'a str,
}

impl<'a> Link<'a> {
    /// Return the exact relationship ID stored by the producer.
    pub fn id(self) -> &'a str {
        self.id
    }

    /// Return the exact Strict or Transitional relationship type.
    pub fn kind(self) -> &'a str {
        self.kind
    }

    /// Return the producer's unmodified relative target reference.
    pub fn target(self) -> &'a str {
        self.target
    }
}

/// Borrow the optional, fully validated table-style relationship.
///
/// This is the low-level package-writer seam. It validates the complete graph
/// and returns exact producer relationship fields without copying catalog XML.
pub fn link(package: &OpcPackage) -> Result<Option<Link<'_>>> {
    let graph = inspect_graph(package)?;
    let Some(attachment) = graph.attachment else {
        return Ok(None);
    };
    let relationship = package
        .get_part(&graph.presentation)?
        .rels()
        .get(&attachment.relationship_id)
        .ok_or_else(|| invalid("validated table-style relationship disappeared"))?;
    Ok(Some(Link {
        id: relationship.r_id(),
        kind: relationship.reltype(),
        target: relationship.target_ref(),
    }))
}

/// Validate the package graph and return its presentation conformance.
pub fn conformance(package: &OpcPackage) -> Result<Conformance> {
    Ok(inspect_graph(package)?.conformance)
}

/// Validate the presentation and table-style attachment graph without
/// allocating an owned copy of the catalog XML.
///
/// This is intended for package writers that need to preserve the optional
/// attachment while rebuilding the presentation part.
pub fn present(package: &OpcPackage) -> Result<bool> {
    Ok(link(package)?.is_some())
}

/// Load the presentation's optional validated table-style catalog.
///
/// The returned list copies the bounded part payload exactly once so it can
/// outlive the package borrow and later move unchanged bytes through [`put`].
pub fn load(package: &OpcPackage) -> Result<Option<List>> {
    let graph = inspect_graph(package)?;
    let Some(attachment) = graph.attachment else {
        return Ok(None);
    };
    let xml = own_blob(
        package.get_part(&attachment.part)?.blob(),
        "table-style XML ownership",
    )?;
    Ok(Some(parse_owned(xml)?))
}

/// Create or replace the presentation's table-style catalog atomically.
///
/// Loaded, unedited catalogs retain exact producer bytes. A byte-identical
/// load→put is a signature-preserving no-op; changed XML moves into a staged
/// part only after conformance, topology, and relationship checks succeed.
pub fn put(package: &mut OpcPackage, list: List) -> Result<bool> {
    let graph = inspect_graph(package)?;
    if list.conformance != graph.conformance {
        return Err(invalid(
            "table-style conformance differs from the presentation",
        ));
    }
    let xml = list.into_xml()?;
    if let Some(attachment) = graph.attachment {
        let stored = package.get_part(&attachment.part)?.blob();
        if stored == xml || semantic_xml_eq(stored, &xml)? {
            return Ok(false);
        }
        let staged = BlobPart::new(attachment.part, ct::PML_TABLE_STYLES.into(), xml);
        package.add_part(Box::new(staged));
        package.unsign();
        return Ok(true);
    }

    let part_name = available_part_name(package)?;
    let relationship_id = {
        let presentation = package.get_part(&graph.presentation)?;
        available_relationship_id(presentation)?
    };
    let target = part_name.relative_ref(graph.presentation.base_uri());
    let mut staged_presentation = package.get_part(&graph.presentation)?.clone_part();
    staged_presentation.rels_mut().try_add_relationship(
        graph.conformance.relationship().into(),
        target,
        relationship_id,
        TargetMode::Internal,
    )?;
    package.validate_new_part_name(&part_name)?;
    let staged = BlobPart::new(part_name, ct::PML_TABLE_STYLES.into(), xml);

    package.add_part(Box::new(staged));
    package.add_part(staged_presentation);
    package.unsign();
    Ok(true)
}

/// Remove and return the optional table-style catalog atomically.
///
/// Absence is an idempotent, signature-preserving `Ok(None)`. A catalog with
/// any unexpected inbound edge is rejected before either relationship or part
/// is changed.
pub fn remove(package: &mut OpcPackage) -> Result<Option<List>> {
    let graph = inspect_graph(package)?;
    let Some(attachment) = graph.attachment else {
        return Ok(None);
    };
    let xml = own_blob(
        package.get_part(&attachment.part)?.blob(),
        "removed table-style XML",
    )?;
    let list = parse_owned(xml)?;
    let mut staged_presentation = package.get_part(&graph.presentation)?.clone_part();
    if staged_presentation
        .rels_mut()
        .remove(&attachment.relationship_id)
        .is_none()
    {
        return Err(invalid(
            "validated table-style relationship disappeared before commit",
        ));
    }

    package.add_part(staged_presentation);
    let _ = package.remove_part(&attachment.part);
    package.unsign();
    Ok(Some(list))
}

struct Graph {
    presentation: PackURI,
    conformance: Conformance,
    attachment: Option<Attachment>,
}

struct Attachment {
    relationship_id: String,
    part: PackURI,
}

fn inspect_graph(package: &OpcPackage) -> Result<Graph> {
    if package.part_count() > MAX_GRAPH_PARTS {
        return Err(limit("table-style package parts", MAX_GRAPH_PARTS));
    }
    let mut main_relationship = None;
    for relationship in package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        )
    }) {
        if main_relationship.replace(relationship).is_some() {
            return Err(invalid("package has multiple main-document relationships"));
        }
    }
    let main_relationship = main_relationship
        .ok_or_else(|| invalid("package main-document relationship is missing"))?;
    if main_relationship.is_external() {
        return Err(invalid(
            "package main-document relationship cannot be external",
        ));
    }
    let requested_presentation = main_relationship.target_partname()?;
    let presentation = package.get_part(&requested_presentation)?;
    require_presentation_content_type(presentation.content_type())?;
    let presentation_name = presentation.partname().clone();
    let conformance = presentation_conformance(presentation.blob())?;
    if main_relationship.reltype() != conformance.office_document() {
        return Err(invalid(
            "package main-document relationship conformance differs from the presentation",
        ));
    }
    let mut selected = None;
    for relationship in presentation
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::TABLE_STYLES | STRICT_REL))
    {
        if selected.is_some() {
            return Err(invalid(
                "presentation has multiple table-style relationships",
            ));
        }
        if relationship.reltype() != conformance.relationship() {
            return Err(invalid(
                "table-style relationship conformance differs from the presentation",
            ));
        }
        if relationship.is_external() {
            return Err(invalid("table-style relationship cannot be external"));
        }
        let requested = relationship.target_partname()?;
        let part = package.get_part(&requested)?;
        let part_name = part.partname().clone();
        validate_part_name(&part_name)?;
        if part.content_type() != ct::PML_TABLE_STYLES {
            return Err(Error::ContentType {
                expected: ct::PML_TABLE_STYLES.into(),
                actual: part.content_type().into(),
            });
        }
        if part.rels().iter().next().is_some() {
            return Err(invalid(
                "table-style part must not own package relationships",
            ));
        }
        let parsed = scan(part.blob())?;
        if parsed.conformance != conformance {
            return Err(invalid(
                "table-style XML conformance differs from the presentation",
            ));
        }
        selected = Some(Attachment {
            relationship_id: relationship.r_id().to_owned(),
            part: part_name,
        });
    }
    if let Some(attachment) = &selected {
        validate_inbound(package, &presentation_name, attachment)?;
    } else if package
        .iter_parts()
        .any(|part| part.content_type() == ct::PML_TABLE_STYLES)
    {
        return Err(invalid(
            "orphan table-style part exists without a presentation relationship",
        ));
    }
    Ok(Graph {
        presentation: presentation_name,
        conformance,
        attachment: selected,
    })
}

fn validate_inbound(
    package: &OpcPackage,
    presentation: &PackURI,
    attachment: &Attachment,
) -> Result<()> {
    let mut links = 0usize;
    let mut expected = 0usize;
    for relationship in package.rels().iter() {
        inspect_inbound(
            package,
            None,
            relationship,
            presentation,
            attachment,
            &mut expected,
        )?;
        links = checked_link_count(links)?;
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            inspect_inbound(
                package,
                Some(source.partname()),
                relationship,
                presentation,
                attachment,
                &mut expected,
            )?;
            links = checked_link_count(links)?;
        }
    }
    if expected != 1 {
        return Err(invalid(
            "table-style part must have exactly one presentation owner",
        ));
    }
    Ok(())
}

fn inspect_inbound(
    package: &OpcPackage,
    source: Option<&PackURI>,
    relationship: &litchi_opc::Relationship,
    presentation: &PackURI,
    attachment: &Attachment,
    expected: &mut usize,
) -> Result<()> {
    if relationship.is_external() {
        return Ok(());
    }
    let requested = relationship.target_partname()?;
    let targets_style = requested == attachment.part
        || package
            .get_part(&requested)
            .is_ok_and(|part| part.partname() == &attachment.part);
    if !targets_style {
        return Ok(());
    }
    if source == Some(presentation) && relationship.r_id() == attachment.relationship_id {
        *expected = expected
            .checked_add(1)
            .ok_or_else(|| invalid("table-style inbound count overflow"))?;
        return Ok(());
    }
    Err(invalid(format!(
        "table-style part '{}' has an unexpected inbound relationship '{}' from '{}'",
        attachment.part.as_str(),
        relationship.r_id(),
        source.map_or("package root", PackURI::as_str),
    )))
}

fn checked_link_count(value: usize) -> Result<usize> {
    let value = value
        .checked_add(1)
        .ok_or_else(|| limit("table-style package relationships", MAX_GRAPH_RELATIONSHIPS))?;
    if value > MAX_GRAPH_RELATIONSHIPS {
        Err(limit(
            "table-style package relationships",
            MAX_GRAPH_RELATIONSHIPS,
        ))
    } else {
        Ok(value)
    }
}

fn require_presentation_content_type(value: &str) -> Result<()> {
    if matches!(
        value,
        ct::PML_PRESENTATION_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    ) {
        Ok(())
    } else {
        Err(Error::ContentType {
            expected: "a PresentationML presentation, slideshow, or template main part".into(),
            actual: value.into(),
        })
    }
}

fn presentation_conformance(xml: &[u8]) -> Result<Conformance> {
    if xml.len() > MAX_PRESENTATION_BYTES {
        return Err(limit("presentation XML bytes", MAX_PRESENTATION_BYTES));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut profile = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut closed = false;
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                bump_presentation_node(&mut nodes)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("presentation XML depth", MAX_DEPTH))?;
                if depth > MAX_DEPTH {
                    return Err(limit("presentation XML depth", MAX_DEPTH));
                }
                if depth == 1 {
                    if profile.is_some() || element.local_name().as_ref() != b"presentation" {
                        return Err(invalid("invalid presentation root"));
                    }
                    profile = presentation_namespace(&namespace);
                    if profile.is_none() {
                        return Err(invalid("presentation root has the wrong namespace"));
                    }
                }
            },
            Event::Empty(element) => {
                bump_presentation_node(&mut nodes)?;
                if depth == 0 {
                    if profile.is_some() || element.local_name().as_ref() != b"presentation" {
                        return Err(invalid("invalid presentation root"));
                    }
                    profile = presentation_namespace(&namespace);
                    if profile.is_none() {
                        return Err(invalid("presentation root has the wrong namespace"));
                    }
                    closed = true;
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("presentation XML nesting underflow"));
                }
                if depth == 1 {
                    let selected = profile.ok_or_else(|| invalid("missing presentation root"))?;
                    if element.local_name().as_ref() != b"presentation"
                        || presentation_namespace(&namespace) != Some(selected)
                    {
                        return Err(invalid("presentation root closes with a wrong element"));
                    }
                    closed = true;
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "presentation XML must not contain a DTD or processing instruction",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || !closed {
        return Err(invalid("presentation XML root is unterminated"));
    }
    profile.ok_or_else(|| invalid("missing presentation root"))
}

fn presentation_namespace(namespace: &ResolveResult<'_>) -> Option<Conformance> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == P.as_bytes() => {
            Some(Conformance::Transitional)
        },
        ResolveResult::Bound(Namespace(value)) if *value == PS.as_bytes() => {
            Some(Conformance::Strict)
        },
        _ => None,
    }
}

fn bump_presentation_node(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("presentation XML nodes", MAX_NODES))?;
    if *nodes > MAX_NODES {
        Err(limit("presentation XML nodes", MAX_NODES))
    } else {
        Ok(())
    }
}

fn validate_part_name(part: &PackURI) -> Result<()> {
    if part.as_str().starts_with("/ppt/") && part.as_str().ends_with(".xml") {
        Ok(())
    } else {
        Err(invalid("table-style part must be an XML part below /ppt/"))
    }
}

fn available_part_name(package: &OpcPackage) -> Result<PackURI> {
    for number in 1..=PART_NAME_ATTEMPTS {
        let path = if number == 1 {
            "/ppt/tableStyles.xml".to_owned()
        } else {
            format!("/ppt/tableStyles{number}.xml")
        };
        let candidate = PackURI::new(&path).map_err(Error::Invalid)?;
        if package.get_part(&candidate).is_err() {
            package.validate_new_part_name(&candidate)?;
            return Ok(candidate);
        }
    }
    Err(limit(
        "table-style part-name allocation attempts",
        PART_NAME_ATTEMPTS,
    ))
}

fn available_relationship_id(owner: &dyn Part) -> Result<String> {
    for number in 1..=RELATIONSHIP_ID_ATTEMPTS {
        let candidate = format!("rId{number}");
        if owner.rels().get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(limit(
        "table-style relationship-ID allocation attempts",
        RELATIONSHIP_ID_ATTEMPTS,
    ))
}

fn own_blob(blob: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(blob.len())
        .map_err(|source| allocation(resource, source))?;
    owned.extend_from_slice(blob);
    Ok(owned)
}

#[derive(Debug, PartialEq, Eq)]
enum SemanticToken {
    Start {
        namespace: String,
        local: String,
        attributes: Vec<(String, String, String)>,
    },
    End {
        namespace: String,
        local: String,
    },
    Text(String),
    Comment(String),
}

struct SemanticCursor<'a> {
    reader: NsReader<&'a [u8]>,
    pending_end: Option<(String, String)>,
    spaces: Vec<Space>,
    nodes: usize,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Space {
    Default,
    Preserve,
}

impl<'a> SemanticCursor<'a> {
    fn new(xml: &'a [u8]) -> Self {
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().check_end_names = true;
        Self {
            reader,
            pending_end: None,
            spaces: Vec::new(),
            nodes: 0,
        }
    }

    fn next(&mut self) -> Result<Option<SemanticToken>> {
        if let Some((namespace, local)) = self.pending_end.take() {
            let _ = self.spaces.pop();
            return Ok(Some(SemanticToken::End { namespace, local }));
        }
        loop {
            let decoder = self.reader.decoder();
            let event = self.reader.read_event().map_err(xml_error)?.into_owned();
            let resolver = self.reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) => {
                    bump_node(&mut self.nodes)?;
                    let (namespace, local, attributes) =
                        semantic_element(&resolver, &namespace, &element, decoder)?;
                    self.push_space(&attributes);
                    return Ok(Some(SemanticToken::Start {
                        namespace,
                        local,
                        attributes,
                    }));
                },
                Event::Empty(element) => {
                    bump_node(&mut self.nodes)?;
                    let (namespace, local, attributes) =
                        semantic_element(&resolver, &namespace, &element, decoder)?;
                    self.push_space(&attributes);
                    self.pending_end = Some((namespace.clone(), local.clone()));
                    return Ok(Some(SemanticToken::Start {
                        namespace,
                        local,
                        attributes,
                    }));
                },
                Event::End(element) => {
                    let _ = self.spaces.pop();
                    return Ok(Some(SemanticToken::End {
                        namespace: resolved_namespace(&namespace)?,
                        local: std::str::from_utf8(element.local_name().as_ref())
                            .map_err(xml_error)?
                            .to_owned(),
                    }));
                },
                Event::Text(text) => {
                    let decoded = text.decode().map_err(xml_error)?;
                    let text = quick_xml::escape::unescape(&decoded)
                        .map_err(xml_error)?
                        .into_owned();
                    if !text.chars().all(char::is_whitespace)
                        || self.spaces.len() >= 3
                        || self.spaces.last() == Some(&Space::Preserve)
                    {
                        return Ok(Some(SemanticToken::Text(text)));
                    }
                },
                Event::GeneralRef(reference) => {
                    return Ok(Some(SemanticToken::Text(
                        litchi_ooxml_common::xml::decode_xml_reference(&reference)?,
                    )));
                },
                Event::CData(text) => {
                    return Ok(Some(SemanticToken::Text(
                        text.decode().map_err(xml_error)?.into_owned(),
                    )));
                },
                Event::Comment(comment) => {
                    return Ok(Some(SemanticToken::Comment(
                        comment.decode().map_err(xml_error)?.into_owned(),
                    )));
                },
                Event::Decl(_) => {},
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("forbidden markup in table-style XML"));
                },
                Event::Eof => return Ok(None),
            }
        }
    }

    fn push_space(&mut self, attributes: &[(String, String, String)]) {
        const XML: &str = "http://www.w3.org/XML/1998/namespace";
        let inherited = self.spaces.last().copied().unwrap_or(Space::Default);
        let selected = attributes
            .iter()
            .find(|(namespace, local, _)| namespace == XML && local == "space")
            .map_or(inherited, |(_, _, value)| match value.as_str() {
                "default" => Space::Default,
                "preserve" => Space::Preserve,
                // The catalog scanner deliberately retains opaque attributes.
                // Unknown xml:space values must therefore disable whitespace
                // normalization rather than risk discarding caller data.
                _ => Space::Preserve,
            });
        self.spaces.push(selected);
    }
}

fn semantic_xml_eq(left: &[u8], right: &[u8]) -> Result<bool> {
    let mut left = SemanticCursor::new(left);
    let mut right = SemanticCursor::new(right);
    loop {
        let left = left.next()?;
        let right = right.next()?;
        if left != right {
            return Ok(false);
        }
        if left.is_none() {
            return Ok(true);
        }
    }
}

fn semantic_element(
    resolver: &NamespaceResolver,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<(String, String, Vec<(String, String, String)>)> {
    let namespace = resolved_namespace(namespace)?;
    let local = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    let mut attributes = Vec::new();
    let mut bytes = 0usize;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        if attributes.len() >= MAX_ATTRIBUTES {
            return Err(limit("table-style attribute count", MAX_ATTRIBUTES));
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        let namespace = resolved_namespace(&namespace)?;
        let local = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bytes = bytes
            .checked_add(namespace.len())
            .and_then(|total| total.checked_add(local.len()))
            .and_then(|total| total.checked_add(value.len()))
            .ok_or_else(|| limit("table-style attribute bytes", MAX_ATTRIBUTE_BYTES))?;
        if bytes > MAX_ATTRIBUTE_BYTES {
            return Err(limit("table-style attribute bytes", MAX_ATTRIBUTE_BYTES));
        }
        attributes
            .try_reserve(1)
            .map_err(|source| allocation("table-style semantic attributes", source))?;
        attributes.push((namespace, local, value));
    }
    attributes.sort_unstable();
    Ok((namespace, local, attributes))
}

fn resolved_namespace(namespace: &ResolveResult<'_>) -> Result<String> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound table-style XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::constants::relationship_type;

    const DEFAULT: &str = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
    const FIRST: &str = "{5940675A-B579-460E-94D1-54222C63F5DA}";
    const SECOND: &str = "{11111111-2222-3333-4444-555555555555}";

    #[test]
    fn ids_are_typed_case_insensitive_and_canonical() {
        let upper = Id::parse(FIRST).unwrap();
        let lower = Id::parse("{5940675a-b579-460e-94d1-54222c63f5da}").unwrap();
        assert_eq!(upper, lower);
        assert_eq!(upper.to_string(), FIRST);
        for invalid in [
            "5940675A-B579-460E-94D1-54222C63F5DA",
            "{5940675A-B579-460E-94D1-54222C63F5D}",
            "{5940675A_B579-460E-94D1-54222C63F5DA}",
            "{5940675Z-B579-460E-94D1-54222C63F5DA}",
        ] {
            assert!(Id::parse(invalid).is_err(), "accepted {invalid}");
        }
    }

    #[test]
    fn source_bytes_duplicate_names_background_and_extensions_are_preserved() {
        let xml = format!(
            r#"<?xml version="1.0"?><d:tblStyleLst xmlns:d="{A}" xmlns:x="urn:test" def="{DEFAULT}">
  <d:tblStyle styleId="{FIRST}" styleName=""><d:tblBg><d:fill/></d:tblBg><d:extLst><d:ext uri="x"><x:data/></d:ext></d:extLst></d:tblStyle>
  <d:tblStyle styleId="{SECOND}" styleName=""><d:wholeTbl/></d:tblStyle>
</d:tblStyleLst>"#,
        );
        let mut list = List::parse(xml.as_bytes().to_vec()).unwrap();
        assert_eq!(list.conformance(), Conformance::Transitional);
        assert_eq!(list.named("").count(), 2);
        assert!(
            list.get(Id::parse(FIRST).unwrap())
                .unwrap()
                .has(Parts::BACKGROUND)
        );
        assert_eq!(list.source_xml(), Some(xml.as_bytes()));

        let unchanged = list.into_xml().unwrap();
        assert_eq!(unchanged, xml.as_bytes());

        list = List::parse(unchanged).unwrap();
        list.rename(Id::parse(FIRST).unwrap(), "Renamed").unwrap();
        let changed = list.into_xml().unwrap();
        let parsed = List::parse(changed).unwrap();
        let renamed = parsed.get(Id::parse(FIRST).unwrap()).unwrap();
        assert_eq!(renamed.name(), "Renamed");
        assert!(renamed.has(Parts::BACKGROUND));
        assert!(
            parsed
                .source_xml()
                .unwrap()
                .windows(6)
                .any(|value| value == b"extLst")
        );
    }

    #[test]
    fn semantic_noop_keeps_preserved_and_opaque_whitespace_distinct() {
        let preserved = format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x" xml:space="preserve"> </a:tblStyle></a:tblStyleLst>"#,
        );
        let changed_preserved = format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x" xml:space="preserve">  </a:tblStyle></a:tblStyleLst>"#,
        );
        List::parse(preserved.as_bytes().to_vec()).unwrap();
        List::parse(changed_preserved.as_bytes().to_vec()).unwrap();
        assert!(!semantic_xml_eq(preserved.as_bytes(), changed_preserved.as_bytes()).unwrap());

        let opaque = format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:extLst><a:ext uri="x"><x:data xmlns:x="urn:test"> </x:data></a:ext></a:extLst></a:tblStyle></a:tblStyleLst>"#,
        );
        let changed_opaque = format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:extLst><a:ext uri="x"><x:data xmlns:x="urn:test">  </x:data></a:ext></a:extLst></a:tblStyle></a:tblStyleLst>"#,
        );
        let mut package = synthetic(
            ct::PML_PRESENTATION_MAIN,
            Conformance::Transitional,
            Some(opaque.as_bytes()),
        );
        package.relate_to(
            "_xmlsignatures/origin.sigs",
            relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        assert!(package.is_signed());

        let candidate = List::parse(changed_opaque.as_bytes().to_vec()).unwrap();
        assert!(put(&mut package, candidate).unwrap());
        assert!(!package.is_signed());
        assert_eq!(
            load(&package).unwrap().unwrap().source_xml(),
            Some(changed_opaque.as_bytes())
        );
    }

    #[test]
    fn semantic_comparison_publishes_changed_opaque_comments() {
        let original = format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:extLst><!-- producer note --><a:ext uri="x"/></a:extLst></a:tblStyle></a:tblStyleLst>"#,
        );
        let changed = format!(
            r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:extLst><!-- updated producer note --><a:ext uri="x"/></a:extLst></a:tblStyle></a:tblStyleLst>"#,
        );
        let mut package = synthetic(
            ct::PML_PRESENTATION_MAIN,
            Conformance::Transitional,
            Some(original.as_bytes()),
        );

        assert!(
            put(
                &mut package,
                List::parse(changed.as_bytes().to_vec()).unwrap()
            )
            .unwrap()
        );
        assert_eq!(
            load(&package).unwrap().unwrap().source_xml(),
            Some(changed.as_bytes())
        );
    }

    #[test]
    fn semantic_crud_is_checked_and_failure_atomic() {
        let default = Id::parse(DEFAULT).unwrap();
        let first = Id::parse(FIRST).unwrap();
        let second = Id::parse(SECOND).unwrap();
        let mut list = List::new(Conformance::Transitional, default);
        let mut a = Def::new(first, "Shared name").unwrap();
        a.reset_parts(Parts::WHOLE | Parts::FIRST_ROW);
        list.add(a).unwrap();
        list.add(Def::new(second, "Shared name").unwrap()).unwrap();
        assert_eq!(list.named("Shared name").count(), 2);

        let before = list.len();
        assert!(list.add(Def::new(first, "Duplicate ID").unwrap()).is_err());
        assert_eq!(list.len(), before);
        assert!(list.remove(default).is_err());
        assert_eq!(list.len(), before);

        let old = list.rename(first, "Renamed").unwrap();
        assert_eq!(old, "Shared name");
        let removed = list.remove(second).unwrap().unwrap();
        assert_eq!(removed.id(), second);
        assert_eq!(list.len(), 1);
        let round_trip = List::parse(list.into_xml().unwrap()).unwrap();
        assert_eq!(round_trip.get(first).unwrap().name(), "Renamed");
        assert!(round_trip.get(first).unwrap().has(Parts::FIRST_ROW));
    }

    #[test]
    fn parser_rejects_requiredness_duplicates_mixed_dialects_and_forbidden_markup() {
        let cases = [
            format!(r#"<a:tblStyleLst xmlns:a="{A}"/>"#),
            format!(
                r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleName="x"/></a:tblStyleLst>"#
            ),
            format!(
                r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}"/></a:tblStyleLst>"#
            ),
            format!(
                r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="a"/><a:tblStyle styleId="{FIRST}" styleName="b"/></a:tblStyleLst>"#
            ),
            format!(
                r#"<a:tblStyleLst xmlns:a="{A}" xmlns:s="{AS}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><s:firstRow/></a:tblStyle></a:tblStyleLst>"#
            ),
            format!(
                r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:firstRow/><a:firstRow/></a:tblStyle></a:tblStyleLst>"#
            ),
            format!(
                r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:firstCol/><a:lastCol/></a:tblStyle></a:tblStyleLst>"#
            ),
            format!(
                r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x"><a:extLst/><a:firstRow/></a:tblStyle></a:tblStyleLst>"#
            ),
            format!(
                r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x">text</a:tblStyle></a:tblStyleLst>"#
            ),
            format!(
                r#"<a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"><a:tblStyle styleId="{FIRST}" styleName="x">&amp;</a:tblStyle></a:tblStyleLst>"#
            ),
            format!(r#"<!DOCTYPE x><a:tblStyleLst xmlns:a="{A}" def="{DEFAULT}"/>"#),
        ];
        for xml in cases {
            assert!(List::parse(xml.into_bytes()).is_err());
        }
    }

    #[test]
    fn package_crud_round_trips_all_main_profiles_and_both_conformances() {
        for content_type in [
            ct::PML_PRESENTATION_MAIN,
            ct::PML_SLIDESHOW_MAIN,
            ct::PML_TEMPLATE_MAIN,
            ct::PML_PRES_MACRO_MAIN,
            ct::PML_SLIDESHOW_MACRO_MAIN,
            ct::PML_TEMPLATE_MACRO_MAIN,
        ] {
            for conformance in [Conformance::Transitional, Conformance::Strict] {
                let mut package = synthetic(content_type, conformance, None);
                let mut list = List::new(conformance, Id::parse(DEFAULT).unwrap());
                let mut style = Def::new(Id::parse(FIRST).unwrap(), "Created").unwrap();
                style.reset_parts(Parts::BACKGROUND | Parts::WHOLE | Parts::FIRST_ROW);
                list.add(style).unwrap();
                assert!(put(&mut package, list).unwrap());

                let loaded = load(&package).unwrap().unwrap();
                assert_eq!(loaded.conformance(), conformance);
                assert!(
                    loaded
                        .get(Id::parse(FIRST).unwrap())
                        .unwrap()
                        .has(Parts::BACKGROUND)
                );
                package.relate_to(
                    "_xmlsignatures/origin.sigs",
                    relationship_type::DIGITAL_SIGNATURE_ORIGIN,
                );
                assert!(package.is_signed());
                assert!(!put(&mut package, loaded).unwrap());
                assert!(package.is_signed());
                let semantic = List::parse(
                    format!(
                        r#"<d:tblStyleLst def="{DEFAULT}" xmlns:d="{}">
  <d:tblStyle styleName="Created" styleId="{FIRST}"><d:tblBg></d:tblBg><d:wholeTbl></d:wholeTbl><d:firstRow></d:firstRow></d:tblStyle>
</d:tblStyleLst>"#,
                        conformance.drawing(),
                    )
                    .into_bytes(),
                )
                .unwrap();
                assert!(!put(&mut package, semantic).unwrap());
                assert!(package.is_signed());

                let mut changed = load(&package).unwrap().unwrap();
                changed
                    .rename(Id::parse(FIRST).unwrap(), "Changed")
                    .unwrap();
                assert!(put(&mut package, changed).unwrap());
                assert!(!package.is_signed());
                assert_eq!(
                    load(&package)
                        .unwrap()
                        .unwrap()
                        .get(Id::parse(FIRST).unwrap())
                        .unwrap()
                        .name(),
                    "Changed"
                );

                let removed = remove(&mut package).unwrap().unwrap();
                assert_eq!(removed.conformance(), conformance);
                assert!(load(&package).unwrap().is_none());
                assert!(remove(&mut package).unwrap().is_none());
            }
        }
    }

    #[test]
    fn shared_inbound_topology_is_rejected_before_mutation() {
        let mut package = synthetic(
            ct::PML_PRESENTATION_MAIN,
            Conformance::Transitional,
            Some(default_xml().as_bytes()),
        );
        let observer_name = PackURI::new("/ppt/observer.xml").unwrap();
        let mut observer = BlobPart::new(
            observer_name,
            "application/xml".into(),
            b"<observer/>".to_vec(),
        );
        observer.rels_mut().add_relationship(
            "urn:test:shared".into(),
            "tableStyles.xml".into(),
            "rIdShared".into(),
            false,
        );
        package.add_part(Box::new(observer));
        let part_count = package.part_count();
        let presentation = package.main_document_part().unwrap().partname().clone();
        let rel_count = package.get_part(&presentation).unwrap().rels().len();

        assert!(load(&package).is_err());
        assert!(remove(&mut package).is_err());
        assert_eq!(package.part_count(), part_count);
        assert_eq!(
            package.get_part(&presentation).unwrap().rels().len(),
            rel_count
        );
    }

    #[test]
    fn root_relationship_must_match_the_presentation_dialect() {
        let mut package = synthetic(
            ct::PML_PRESENTATION_MAIN,
            Conformance::Strict,
            Some(format!(r#"<a:tblStyleLst xmlns:a="{AS}" def="{DEFAULT}"/>"#).as_bytes()),
        );
        let relationship_id = package.rels().iter().next().unwrap().r_id().to_owned();
        assert!(package.rels_mut().remove(&relationship_id).is_some());
        package.rels_mut().add_relationship(
            rt::OFFICE_DOCUMENT.into(),
            "ppt/presentation.xml".into(),
            relationship_id,
            false,
        );

        assert!(present(&package).is_err());
        assert!(load(&package).is_err());
    }

    fn synthetic(
        content_type: &str,
        conformance: Conformance,
        catalog: Option<&[u8]>,
    ) -> OpcPackage {
        let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
        let mut presentation = BlobPart::new(
            presentation_name,
            content_type.into(),
            format!(
                "<p:presentation xmlns:p=\"{}\"/>",
                match conformance {
                    Conformance::Transitional => P,
                    Conformance::Strict => PS,
                }
            )
            .into_bytes(),
        );
        let mut package = OpcPackage::new();
        if let Some(catalog) = catalog {
            presentation.rels_mut().add_relationship(
                conformance.relationship().into(),
                "tableStyles.xml".into(),
                "rIdStyles".into(),
                false,
            );
            package.add_part(Box::new(BlobPart::new(
                PackURI::new("/ppt/tableStyles.xml").unwrap(),
                ct::PML_TABLE_STYLES.into(),
                catalog.to_vec(),
            )));
        }
        package.add_part(Box::new(presentation));
        package.relate_to("ppt/presentation.xml", conformance.office_document());
        package
    }
}

fn hex(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}

fn invalidate_payload(payload: &mut Payload) {
    match payload {
        Payload::Shared { exact, .. } | Payload::Owned { exact, .. } => *exact = false,
    }
}

#[derive(Debug)]
struct Parsed {
    conformance: Conformance,
    default: Id,
    defs: Vec<ParsedDef>,
    root_attrs: Vec<Attr>,
}

#[derive(Debug)]
struct ParsedDef {
    id: Id,
    name: String,
    parts: Parts,
    attrs: Vec<Attr>,
    raw: Range<usize>,
    body: Range<usize>,
}

struct OpenDef {
    id: Id,
    name: String,
    parts: Parts,
    attrs: Vec<Attr>,
    raw_start: usize,
    body_start: usize,
    extension: bool,
    last_child: Option<usize>,
}

fn parse_owned(source: Vec<u8>) -> Result<List> {
    let parsed = scan(&source)?;
    let mut defs = Vec::new();
    defs.try_reserve_exact(parsed.defs.len())
        .map_err(|source| allocation("table-style index", source))?;
    for style in parsed.defs {
        defs.push(Def {
            id: style.id,
            name: style.name,
            parts: style.parts,
            attrs: style.attrs,
            payload: Payload::Shared {
                raw: style.raw,
                body: style.body,
                exact: true,
            },
        });
    }
    Ok(List {
        conformance: parsed.conformance,
        default: parsed.default,
        defs,
        root_attrs: parsed.root_attrs,
        source,
        dirty: false,
    })
}

fn scan(xml: &[u8]) -> Result<Parsed> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("table-style XML bytes", MAX_XML_BYTES));
    }
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut saw_root = false;
    let mut closed_root = false;
    let mut conformance = None;
    let mut default = None;
    let mut root_attrs = Vec::new();
    let mut defs = Vec::new();
    let mut open = None;

    loop {
        let start = xml_position(&reader)?;
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = xml_position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                bump_node(&mut nodes)?;
                depth = checked_depth(depth)?;
                if depth == 1 {
                    if saw_root || element.local_name().as_ref() != b"tblStyleLst" {
                        return Err(invalid(
                            "table-style part must contain one tblStyleLst root",
                        ));
                    }
                    let profile = drawing_conformance(&namespace)
                        .ok_or_else(|| invalid("table-style root has the wrong namespace"))?;
                    let (id, attrs) = parse_root_attrs(&element, decoder)?;
                    saw_root = true;
                    conformance = Some(profile);
                    default = Some(id);
                    root_attrs = attrs;
                } else if depth == 2 {
                    let profile = conformance
                        .ok_or_else(|| invalid("table-style root profile is missing"))?;
                    require_drawing(&namespace, profile, element.name(), b"tblStyle")?;
                    if defs.len() >= MAX_STYLES {
                        return Err(limit("table-style count", MAX_STYLES));
                    }
                    let (id, name, attrs) = parse_def_attrs(&element, decoder)?;
                    open = Some(OpenDef {
                        id,
                        name,
                        parts: Parts::empty(),
                        attrs,
                        raw_start: start,
                        body_start: end,
                        extension: false,
                        last_child: None,
                    });
                } else if depth == 3 {
                    record_part(&namespace, conformance, element.name(), &mut open)?;
                }
            },
            Event::Empty(element) => {
                bump_node(&mut nodes)?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("table-style XML depth", MAX_DEPTH))?;
                if child_depth > MAX_DEPTH {
                    return Err(limit("table-style XML depth", MAX_DEPTH));
                }
                if child_depth == 1 {
                    if saw_root || element.local_name().as_ref() != b"tblStyleLst" {
                        return Err(invalid(
                            "table-style part must contain one tblStyleLst root",
                        ));
                    }
                    let profile = drawing_conformance(&namespace)
                        .ok_or_else(|| invalid("table-style root has the wrong namespace"))?;
                    let (id, attrs) = parse_root_attrs(&element, decoder)?;
                    saw_root = true;
                    closed_root = true;
                    conformance = Some(profile);
                    default = Some(id);
                    root_attrs = attrs;
                } else if child_depth == 2 {
                    let profile = conformance
                        .ok_or_else(|| invalid("table-style root profile is missing"))?;
                    require_drawing(&namespace, profile, element.name(), b"tblStyle")?;
                    if defs.len() >= MAX_STYLES {
                        return Err(limit("table-style count", MAX_STYLES));
                    }
                    let (id, name, attrs) = parse_def_attrs(&element, decoder)?;
                    defs.try_reserve(1)
                        .map_err(|source| allocation("table-style parse index", source))?;
                    defs.push(ParsedDef {
                        id,
                        name,
                        parts: Parts::empty(),
                        attrs,
                        raw: start..end,
                        body: end..end,
                    });
                } else if child_depth == 3 {
                    record_part(&namespace, conformance, element.name(), &mut open)?;
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("table-style XML nesting underflow"));
                }
                if depth == 2 {
                    let profile = conformance
                        .ok_or_else(|| invalid("table-style root profile is missing"))?;
                    require_drawing(&namespace, profile, element.name(), b"tblStyle")?;
                    let style = open
                        .take()
                        .ok_or_else(|| invalid("table-style closing element has no start"))?;
                    if style.body_start > start || style.raw_start > style.body_start {
                        return Err(invalid("table-style source ranges are invalid"));
                    }
                    defs.try_reserve(1)
                        .map_err(|source| allocation("table-style parse index", source))?;
                    defs.push(ParsedDef {
                        id: style.id,
                        name: style.name,
                        parts: style.parts,
                        attrs: style.attrs,
                        raw: style.raw_start..end,
                        body: style.body_start..start,
                    });
                } else if depth == 1 {
                    let profile = conformance
                        .ok_or_else(|| invalid("table-style root profile is missing"))?;
                    require_drawing(&namespace, profile, element.name(), b"tblStyleLst")?;
                    closed_root = true;
                }
                depth -= 1;
            },
            Event::Text(text) => {
                if depth <= 2
                    && text
                        .decode()
                        .map_err(xml_error)?
                        .chars()
                        .any(|value| !value.is_whitespace())
                {
                    return Err(invalid(
                        "table-style root or definition contains text content",
                    ));
                }
            },
            Event::GeneralRef(_) if depth <= 2 => {
                return Err(invalid(
                    "table-style root or definition contains a character reference",
                ));
            },
            Event::CData(_) => return Err(invalid("table-style XML must not contain CDATA")),
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "table-style XML must not contain a DTD or processing instruction",
                ));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
            _ => {},
        }
    }
    if !saw_root || !closed_root || depth != 0 || open.is_some() {
        return Err(invalid("table-style XML root is missing or unterminated"));
    }
    let parsed = Parsed {
        conformance: conformance.ok_or_else(|| invalid("missing table-style conformance"))?,
        default: default.ok_or_else(|| invalid("table-style default GUID is required"))?,
        defs,
        root_attrs,
    };
    validate_parsed(&parsed)?;
    Ok(parsed)
}

fn parse_root_attrs(element: &BytesStart<'_>, decoder: Decoder) -> Result<(Id, Vec<Attr>)> {
    let mut default = None;
    let mut extras = Vec::new();
    for (name, value) in attributes(element, decoder)? {
        if name == "def" {
            if default.replace(Id::parse(&value)?).is_some() {
                return Err(invalid("table-style root declares def twice"));
            }
        } else {
            extras
                .try_reserve(1)
                .map_err(|source| allocation("table-style root attributes", source))?;
            extras.push(Attr { name, value });
        }
    }
    Ok((
        default.ok_or_else(|| invalid("table-style default GUID is required"))?,
        extras,
    ))
}

fn parse_def_attrs(element: &BytesStart<'_>, decoder: Decoder) -> Result<(Id, String, Vec<Attr>)> {
    let mut id = None;
    let mut name = None;
    let mut extras = Vec::new();
    for (attribute, value) in attributes(element, decoder)? {
        match attribute.as_str() {
            "styleId" => {
                if id.replace(Id::parse(&value)?).is_some() {
                    return Err(invalid("table style declares styleId twice"));
                }
            },
            "styleName" => {
                validate_name(&value)?;
                if name.replace(value).is_some() {
                    return Err(invalid("table style declares styleName twice"));
                }
            },
            _ => {
                extras
                    .try_reserve(1)
                    .map_err(|source| allocation("table-style attributes", source))?;
                extras.push(Attr {
                    name: attribute,
                    value,
                });
            },
        }
    }
    Ok((
        id.ok_or_else(|| invalid("table style requires a styleId GUID"))?,
        name.ok_or_else(|| invalid("table style requires a styleName attribute"))?,
        extras,
    ))
}

fn attributes(element: &BytesStart<'_>, decoder: Decoder) -> Result<Vec<(String, String)>> {
    let mut output = Vec::new();
    let mut bytes = 0usize;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if output.len() >= MAX_ATTRIBUTES {
            return Err(limit("table-style attribute count", MAX_ATTRIBUTES));
        }
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bytes = bytes
            .checked_add(name.len())
            .and_then(|value_len| value_len.checked_add(value.len()))
            .ok_or_else(|| limit("table-style attribute bytes", MAX_ATTRIBUTE_BYTES))?;
        if bytes > MAX_ATTRIBUTE_BYTES {
            return Err(limit("table-style attribute bytes", MAX_ATTRIBUTE_BYTES));
        }
        output
            .try_reserve(1)
            .map_err(|source| allocation("table-style attribute decoding", source))?;
        output.push((name, value));
    }
    Ok(output)
}

fn record_part(
    namespace: &ResolveResult<'_>,
    conformance: Option<Conformance>,
    name: QName<'_>,
    open: &mut Option<OpenDef>,
) -> Result<()> {
    let profile = conformance.ok_or_else(|| invalid("table-style root profile is missing"))?;
    if !is_drawing(namespace, profile) {
        return Err(invalid(
            "table-style region uses the wrong DrawingML namespace",
        ));
    }
    let style = open
        .as_mut()
        .ok_or_else(|| invalid("table-style region has no owning definition"))?;
    if name.local_name().as_ref() == b"extLst" {
        if style.extension {
            return Err(invalid("table style declares extLst twice"));
        }
        style.extension = true;
        return Ok(());
    }
    if style.extension {
        return Err(invalid("table-style region appears after extLst"));
    }
    let part = Parts::from_xml_name(name.local_name().as_ref())
        .ok_or_else(|| invalid("unexpected direct child in table-style definition"))?;
    if style.parts.intersects(part) {
        return Err(invalid("table style declares a conditional region twice"));
    }
    let order = PARTS
        .iter()
        .position(|(candidate, _)| *candidate == part)
        .ok_or_else(|| invalid("table-style region has no schema order"))?;
    if style.last_child.is_some_and(|previous| previous >= order) {
        return Err(invalid(
            "table-style regions do not follow the schema sequence",
        ));
    }
    style.last_child = Some(order);
    style.parts.insert(part);
    Ok(())
}

fn drawing_conformance(namespace: &ResolveResult<'_>) -> Option<Conformance> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == A.as_bytes() => {
            Some(Conformance::Transitional)
        },
        ResolveResult::Bound(Namespace(value)) if *value == AS.as_bytes() => {
            Some(Conformance::Strict)
        },
        _ => None,
    }
}

fn is_drawing(namespace: &ResolveResult<'_>, conformance: Conformance) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == conformance.drawing().as_bytes())
}

fn require_drawing(
    namespace: &ResolveResult<'_>,
    conformance: Conformance,
    actual: QName<'_>,
    expected: &[u8],
) -> Result<()> {
    if is_drawing(namespace, conformance) && actual.local_name().as_ref() == expected {
        Ok(())
    } else {
        Err(invalid("unexpected table-style element or namespace"))
    }
}

fn validate_parsed(parsed: &Parsed) -> Result<()> {
    if parsed.defs.len() > MAX_STYLES {
        return Err(limit("table-style count", MAX_STYLES));
    }
    let mut ids = HashSet::new();
    ids.try_reserve(parsed.defs.len())
        .map_err(|source| allocation("table-style ID validation", source))?;
    for style in &parsed.defs {
        validate_name(&style.name)?;
        if !ids.insert(style.id) {
            return Err(invalid(format!("duplicate table-style ID {}", style.id)));
        }
    }
    Ok(())
}

fn validate_list(list: &List) -> Result<()> {
    if list.defs.len() > MAX_STYLES {
        return Err(limit("table-style count", MAX_STYLES));
    }
    let mut ids = HashSet::new();
    ids.try_reserve(list.defs.len())
        .map_err(|source| allocation("table-style ID validation", source))?;
    for style in &list.defs {
        validate_def(style)?;
        if !ids.insert(style.id) {
            return Err(invalid(format!("duplicate table-style ID {}", style.id)));
        }
    }
    Ok(())
}

fn validate_def(style: &Def) -> Result<()> {
    validate_name(&style.name)?;
    if style.parts.bits() & !Parts::all().bits() != 0 {
        return Err(invalid("table style contains unknown region flags"));
    }
    Ok(())
}

fn validate_name(name: &str) -> Result<()> {
    if name.len() > MAX_ATTRIBUTE_BYTES {
        return Err(limit("table-style name bytes", MAX_ATTRIBUTE_BYTES));
    }
    if !name.chars().all(xml_char) {
        return Err(invalid(
            "table-style name contains an invalid XML character",
        ));
    }
    Ok(())
}

fn encode(list: &List) -> Result<Vec<u8>> {
    let mut output = String::new();
    output
        .try_reserve(list.source.len().clamp(256, MAX_XML_BYTES))
        .map_err(|source| allocation("table-style XML encoding", source))?;
    append(
        &mut output,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a=""#,
    )?;
    append(&mut output, list.conformance.drawing())?;
    append(&mut output, "\"")?;
    for attribute in &list.root_attrs {
        if attribute.name == "xmlns:a" || attribute.name == "def" {
            continue;
        }
        append_attribute(&mut output, &attribute.name, &attribute.value)?;
    }
    append(&mut output, " def=\"")?;
    list.default.write_to(&mut output)?;
    if list.defs.is_empty() {
        append(&mut output, "\"/>")?;
    } else {
        append(&mut output, "\">")?;
        for style in &list.defs {
            encode_def(&mut output, style, &list.source)?;
        }
        append(&mut output, "</a:tblStyleLst>")?;
    }
    Ok(output.into_bytes())
}

fn encode_def(output: &mut String, style: &Def, source: &[u8]) -> Result<()> {
    let (raw, body, exact) = payload(style, source)?;
    if exact {
        return append_bytes(output, raw);
    }
    append(output, "<a:tblStyle styleId=\"")?;
    style.id.write_to(output)?;
    append(output, "\" styleName=\"")?;
    append_escaped(output, &style.name)?;
    append(output, "\"")?;
    for attribute in &style.attrs {
        if matches!(attribute.name.as_str(), "styleId" | "styleName") {
            continue;
        }
        append_attribute(output, &attribute.name, &attribute.value)?;
    }
    if body.is_empty() && style.parts.is_empty() {
        return append(output, "/>");
    }
    append(output, ">")?;
    if body.is_empty() {
        for (part, name) in PARTS {
            if style.parts.contains(part) {
                append(output, "<a:")?;
                append(output, name)?;
                append(output, "/>")?;
            }
        }
    } else {
        append_bytes(output, body)?;
    }
    append(output, "</a:tblStyle>")
}

fn payload<'a>(style: &'a Def, source: &'a [u8]) -> Result<(&'a [u8], &'a [u8], bool)> {
    match &style.payload {
        Payload::Shared { raw, body, exact } => Ok((
            source
                .get(raw.clone())
                .ok_or_else(|| invalid("table-style raw source range is invalid"))?,
            source
                .get(body.clone())
                .ok_or_else(|| invalid("table-style body source range is invalid"))?,
            *exact,
        )),
        Payload::Owned { xml, body, exact } => Ok((
            xml,
            xml.get(body.clone())
                .ok_or_else(|| invalid("detached table-style body range is invalid"))?,
            *exact,
        )),
    }
}

fn append_attribute(output: &mut String, name: &str, value: &str) -> Result<()> {
    append(output, " ")?;
    append(output, name)?;
    append(output, "=\"")?;
    append_escaped(output, value)?;
    append(output, "\"")
}

fn append_escaped(output: &mut String, value: &str) -> Result<()> {
    let escaped = quick_xml::escape::escape(value);
    append(output, escaped.as_ref())
}

fn append_bytes(output: &mut String, value: &[u8]) -> Result<()> {
    append(output, std::str::from_utf8(value).map_err(xml_error)?)
}

fn append(output: &mut String, value: &str) -> Result<()> {
    let length = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| limit("encoded table-style XML bytes", MAX_XML_BYTES))?;
    if length > MAX_XML_BYTES {
        return Err(limit("encoded table-style XML bytes", MAX_XML_BYTES));
    }
    output
        .try_reserve(value.len())
        .map_err(|source| allocation("table-style XML encoding", source))?;
    output.push_str(value);
    Ok(())
}

fn xml_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(value as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

fn checked_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| limit("table-style XML depth", MAX_DEPTH))?;
    if depth > MAX_DEPTH {
        Err(limit("table-style XML depth", MAX_DEPTH))
    } else {
        Ok(depth)
    }
}

fn bump_node(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("table-style XML nodes", MAX_NODES))?;
    if *nodes > MAX_NODES {
        Err(limit("table-style XML nodes", MAX_NODES))
    } else {
        Ok(())
    }
}

fn xml_position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("table-style XML offset exceeds usize"))
}
