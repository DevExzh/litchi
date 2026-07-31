//! WordprocessingML glossary documents (building blocks / AutoText).

use litchi_core::sheet::Result;
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::{
    Reader, XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
};
use std::collections::{HashMap, HashSet, VecDeque};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WS: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(crate) const REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";
pub(crate) const STRICT_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/glossaryDocument";
pub(crate) const CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml";
const MAX: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_NODES: usize = 1_000_000;
const MAX_PARTS: usize = 100_000;
const MAX_VALUES: usize = 4096;
const MAX_STRING: usize = 1024 * 1024;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum DocPartType {
    None,
    Normal,
    AutoExpand,
    Toolbar,
    Speller,
    FormField,
    SdtPlaceholder,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum InsertionBehavior {
    Content,
    Paragraph,
    Page,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocPartGallery(String);
impl DocPartGallery {
    pub fn as_str(&self) -> &str {
        &self.0
    }
    pub fn parse(v: &str) -> Result<Self> {
        if GALLERIES.contains(&v) {
            Ok(Self(v.into()))
        } else {
            Err(invalid(format!("invalid document-part gallery '{v}'")))
        }
    }
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocPartName {
    pub value: String,
    pub decorated: Option<bool>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocPartCategory {
    pub name: String,
    pub gallery: DocPartGallery,
}
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct DocPartProperties {
    pub name: Option<DocPartName>,
    pub style: Option<String>,
    pub category: Option<DocPartCategory>,
    pub all_types: Option<bool>,
    pub types: Vec<DocPartType>,
    pub behaviors: Vec<InsertionBehavior>,
    pub description: Option<String>,
    pub guid: Option<String>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct GlossaryEntry {
    pub properties: Option<DocPartProperties>,
    /// Full inert `w:docPartBody` subtree.
    pub body_xml: Option<Vec<u8>>,
}
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct GlossaryDocument {
    /// Full inert `w:background` subtree.
    pub background_xml: Option<Vec<u8>>,
    pub entries: Vec<GlossaryEntry>,
}

/// A relationship owned by the glossary document or one of its auxiliary parts.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct GlossaryRelationship {
    pub id: String,
    pub relationship_type: String,
    pub target: String,
    pub external: bool,
}

/// An auxiliary glossary-owned OPC part, such as styles, numbering, theme, or media.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct GlossaryAuxiliaryPart {
    pub part_name: String,
    pub content_type: String,
    pub data: Vec<u8>,
    pub relationships: Vec<GlossaryRelationship>,
}

/// A complete glossary graph staged for atomic package installation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct GlossaryPackage {
    pub document: GlossaryDocument,
    pub strict: bool,
    pub relationships: Vec<GlossaryRelationship>,
    pub auxiliary_parts: Vec<GlossaryAuxiliaryPart>,
}

impl GlossaryPackage {
    pub fn new(document: GlossaryDocument, strict: bool) -> Self {
        Self {
            document,
            strict,
            relationships: Vec::new(),
            auxiliary_parts: Vec::new(),
        }
    }
}

impl GlossaryDocument {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX {
            return Err(invalid("glossary document exceeds 32 MiB"));
        }
        let x = litchi_ooxml_common::mce::process_ooxml(xml)?;
        if x.len() > MAX {
            return Err(invalid("processed glossary document exceeds 32 MiB"));
        }
        project(&parse_dom(x.as_ref())?)
    }
    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        validate(self)?;
        let w = if strict { WS } else { W };
        let r = if strict { RS } else { R };
        let mut x = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument xmlns:w="{w}" xmlns:r="{r}">"#
        );
        if let Some(v) = &self.background_xml {
            opaque(&mut x, v, strict)?;
        }
        if !self.entries.is_empty() {
            x.push_str("<w:docParts>");
            for entry in &self.entries {
                write_entry(&mut x, entry, strict)?;
            }
            x.push_str("</w:docParts>");
        }
        x.push_str("</w:glossaryDocument>");
        if x.len() > MAX {
            return Err(invalid("serialized glossary document exceeds 32 MiB"));
        }
        Ok(x.into_bytes())
    }

    pub fn entries(&self) -> &[GlossaryEntry] {
        &self.entries
    }

    pub fn find_entry(&self, name: &str) -> Option<(usize, &GlossaryEntry)> {
        self.entries.iter().enumerate().find(|(_, entry)| {
            entry
                .properties
                .as_ref()
                .and_then(|properties| properties.name.as_ref())
                .is_some_and(|entry_name| entry_name.value.eq_ignore_ascii_case(name))
        })
    }

    pub fn add_entry(&mut self, entry: GlossaryEntry) -> Result<usize> {
        let mut staged = self.clone();
        staged.entries.push(entry);
        validate(&staged)?;
        let index = self.entries.len();
        *self = staged;
        Ok(index)
    }

    pub fn replace_entry(&mut self, index: usize, entry: GlossaryEntry) -> Result<GlossaryEntry> {
        if index >= self.entries.len() {
            return Err(invalid(format!(
                "glossary entry index {index} is out of bounds"
            )));
        }
        let mut staged = self.clone();
        let previous = std::mem::replace(&mut staged.entries[index], entry);
        validate(&staged)?;
        *self = staged;
        Ok(previous)
    }

    pub fn update_entry<F>(&mut self, index: usize, update: F) -> Result<()>
    where
        F: FnOnce(&mut GlossaryEntry) -> Result<()>,
    {
        if index >= self.entries.len() {
            return Err(invalid(format!(
                "glossary entry index {index} is out of bounds"
            )));
        }
        let mut staged = self.clone();
        update(&mut staged.entries[index])?;
        validate(&staged)?;
        *self = staged;
        Ok(())
    }

    pub fn remove_entry(&mut self, index: usize) -> Result<GlossaryEntry> {
        if index >= self.entries.len() {
            return Err(invalid(format!(
                "glossary entry index {index} is out of bounds"
            )));
        }
        Ok(self.entries.remove(index))
    }

    pub fn remove_entry_named(&mut self, name: &str) -> Option<GlossaryEntry> {
        let index = self.find_entry(name)?.0;
        Some(self.entries.remove(index))
    }

    pub fn move_entry(&mut self, from: usize, to: usize) -> Result<()> {
        if from >= self.entries.len() || to >= self.entries.len() {
            return Err(invalid("glossary reorder index is out of bounds"));
        }
        if from != to {
            let entry = self.entries.remove(from);
            self.entries.insert(to, entry);
        }
        Ok(())
    }

    pub fn clear_entries(&mut self) -> usize {
        let count = self.entries.len();
        self.entries.clear();
        count
    }
}

impl GlossaryEntry {
    pub fn new(name: impl Into<String>, body_xml: Vec<u8>) -> Result<Self> {
        let entry = Self {
            properties: Some(DocPartProperties {
                name: Some(DocPartName {
                    value: name.into(),
                    decorated: None,
                }),
                ..DocPartProperties::default()
            }),
            body_xml: Some(body_xml),
        };
        validate(&GlossaryDocument {
            background_xml: None,
            entries: vec![entry.clone()],
        })?;
        Ok(entry)
    }
}

pub fn load_from_package(package: &OpcPackage) -> Result<Option<GlossaryDocument>> {
    let main = package.main_document_part()?;
    let mut found = main
        .rels()
        .iter()
        .filter(|x| matches!(x.reltype(), REL | STRICT_REL));
    let Some(rel) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid("main document has multiple glossary relationships"));
    }
    if rel.is_external() {
        return Err(invalid("glossary relationship cannot be external"));
    }
    let uri: PackURI = rel.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != CT {
        return Err(invalid(format!(
            "glossary part '{uri}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    let glossary = GlossaryDocument::parse(part.blob())?;
    validate_relationship_integrity(&glossary, part)?;
    Ok(Some(glossary))
}

/// Load the typed document together with its owned auxiliary OPC graph.
pub fn load_package_from_package(package: &OpcPackage) -> Result<Option<GlossaryPackage>> {
    let Some((_, root_uri)) = existing_glossary(package)? else {
        return Ok(None);
    };
    let document = load_from_package(package)?.ok_or_else(|| invalid("missing glossary"))?;
    let root = package.get_part(&root_uri)?;
    validate_existing_targets(package, root)?;
    let strict = package
        .main_document_part()?
        .rels()
        .iter()
        .find(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
        .is_some_and(|relationship| relationship.reltype() == STRICT_REL);
    let relationships = copy_relationships(root);
    let owned = glossary_owned_parts(package, &root_uri)?;
    let mut names: Vec<_> = owned.into_iter().filter(|uri| uri != &root_uri).collect();
    names.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    let mut auxiliary_parts = Vec::with_capacity(names.len());
    for uri in names {
        let part = package.get_part(&uri)?;
        validate_existing_targets(package, part)?;
        auxiliary_parts.push(GlossaryAuxiliaryPart {
            part_name: uri.as_str().to_string(),
            content_type: part.content_type().to_string(),
            data: part.blob().to_vec(),
            relationships: copy_relationships(part),
        });
    }
    Ok(Some(GlossaryPackage {
        document,
        strict,
        relationships,
        auxiliary_parts,
    }))
}

fn copy_relationships(part: &dyn Part) -> Vec<GlossaryRelationship> {
    part.rels()
        .iter()
        .map(|relationship| GlossaryRelationship {
            id: relationship.r_id().to_string(),
            relationship_type: relationship.reltype().to_string(),
            target: relationship.target_ref().to_string(),
            external: relationship.is_external(),
        })
        .collect()
}

fn validate_existing_targets(package: &OpcPackage, part: &dyn Part) -> Result<()> {
    for relationship in part.rels().iter() {
        if !allowed_relationship_type(relationship.reltype()) {
            return Err(invalid(format!(
                "unsupported glossary relationship type '{}'",
                relationship.reltype()
            )));
        }
        if relationship.is_external() {
            if !relationship.reltype().ends_with("/hyperlink") {
                return Err(invalid("only glossary hyperlinks may be external"));
            }
        } else {
            let target = relationship.target_partname()?;
            if package.get_part(&target).is_err() {
                return Err(invalid(format!(
                    "glossary relationship '{}' has dangling target '{}'",
                    relationship.r_id(),
                    target.as_str()
                )));
            }
        }
    }
    Ok(())
}

/// Install or replace a complete glossary graph after validating it in memory.
pub fn store_in_package(package: &mut OpcPackage, value: GlossaryPackage) -> Result<()> {
    let root_uri = PackURI::new("/word/glossary/document.xml")?;
    let xml = value.document.to_xml(value.strict)?;
    let mut root = BlobPart::new(root_uri.clone(), CT.to_string(), xml);
    add_relationships(&mut root, &value.relationships)?;
    validate_relationship_integrity(&value.document, &root)?;

    if value.auxiliary_parts.len() > MAX_VALUES {
        return Err(invalid("glossary auxiliary part limit exceeded"));
    }
    let mut staged: HashMap<PackURI, BlobPart> = HashMap::new();
    for auxiliary in value.auxiliary_parts {
        if auxiliary.data.len() > MAX {
            return Err(invalid("glossary auxiliary part exceeds 32 MiB"));
        }
        if auxiliary.content_type.trim().is_empty() {
            return Err(invalid("glossary auxiliary content type is empty"));
        }
        let uri = PackURI::new(&auxiliary.part_name)?;
        if !uri.as_str().starts_with("/word/glossary/") || uri == root_uri {
            return Err(invalid(format!(
                "glossary auxiliary part '{}' is outside its owned namespace",
                uri.as_str()
            )));
        }
        let mut part = BlobPart::new(uri.clone(), auxiliary.content_type, auxiliary.data);
        add_relationships(&mut part, &auxiliary.relationships)?;
        if staged.insert(uri.clone(), part).is_some() {
            return Err(invalid(format!(
                "duplicate glossary auxiliary part '{uri}'"
            )));
        }
    }

    let old = existing_glossary(package)?;
    let old_owned = if let Some((_, uri)) = &old {
        glossary_owned_parts(package, uri)?
    } else {
        HashSet::new()
    };
    for uri in staged.keys().chain(std::iter::once(&root_uri)) {
        if package.get_part(uri).is_ok() && !old_owned.contains(uri) {
            return Err(invalid(format!(
                "glossary part '{uri}' collides with an existing part"
            )));
        }
    }

    validate_staged_targets(package, &root, &staged, &old_owned)?;

    let main_uri = package.main_document_part()?.partname().clone();
    if let Some((relationship_id, _)) = &old {
        package
            .get_part_mut(&main_uri)?
            .rels_mut()
            .remove(relationship_id);
    }
    for uri in &old_owned {
        package.remove_part(uri);
    }
    for (_, part) in staged {
        package.add_part(Box::new(part));
    }
    package.add_part(Box::new(root));
    let relationship_type = if value.strict { STRICT_REL } else { REL };
    package
        .get_part_mut(&main_uri)?
        .relate_to("glossary/document.xml", relationship_type);
    Ok(())
}

/// Remove the glossary relationship and only its reachable owned parts.
pub fn remove_from_package(package: &mut OpcPackage) -> Result<Option<GlossaryDocument>> {
    let Some((relationship_id, root_uri)) = existing_glossary(package)? else {
        return Ok(None);
    };
    let glossary = load_from_package(package)?.ok_or_else(|| invalid("missing glossary"))?;
    let owned = glossary_owned_parts(package, &root_uri)?;
    let main_uri = package.main_document_part()?.partname().clone();
    package
        .get_part_mut(&main_uri)?
        .rels_mut()
        .remove(&relationship_id);
    for uri in owned {
        package.remove_part(&uri);
    }
    Ok(Some(glossary))
}

fn existing_glossary(package: &OpcPackage) -> Result<Option<(String, PackURI)>> {
    let main = package.main_document_part()?;
    let mut relationships = main
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL));
    let Some(relationship) = relationships.next() else {
        return Ok(None);
    };
    if relationships.next().is_some() {
        return Err(invalid("main document has multiple glossary relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("glossary relationship cannot be external"));
    }
    Ok(Some((
        relationship.r_id().to_string(),
        relationship.target_partname()?,
    )))
}

fn add_relationships(part: &mut BlobPart, relationships: &[GlossaryRelationship]) -> Result<()> {
    if relationships.len() > MAX_VALUES {
        return Err(invalid("glossary relationship limit exceeded"));
    }
    let mut ids = HashSet::new();
    for relationship in relationships {
        bounded(&relationship.id)?;
        bounded(&relationship.relationship_type)?;
        bounded(&relationship.target)?;
        if relationship.id.is_empty() || !ids.insert(relationship.id.clone()) {
            return Err(invalid("duplicate or empty glossary relationship ID"));
        }
        if !allowed_relationship_type(&relationship.relationship_type) {
            return Err(invalid(format!(
                "unsupported glossary relationship type '{}'",
                relationship.relationship_type
            )));
        }
        if relationship.external && !relationship.relationship_type.ends_with("/hyperlink") {
            return Err(invalid("only glossary hyperlinks may be external"));
        }
        part.rels_mut().try_add_relationship(
            relationship.relationship_type.clone(),
            relationship.target.clone(),
            relationship.id.clone(),
            if relationship.external {
                litchi_opc::TargetMode::External
            } else {
                litchi_opc::TargetMode::Internal
            },
        )?;
    }
    Ok(())
}

fn allowed_relationship_type(value: &str) -> bool {
    if value == "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects" {
        return true;
    }
    let Some(kind) = value
        .strip_prefix("http://schemas.openxmlformats.org/officeDocument/2006/relationships/")
        .or_else(|| value.strip_prefix("http://purl.oclc.org/ooxml/officeDocument/relationships/"))
    else {
        return false;
    };
    matches!(
        kind,
        "styles"
            | "numbering"
            | "settings"
            | "fontTable"
            | "webSettings"
            | "theme"
            | "comments"
            | "image"
            | "hyperlink"
    )
}

fn validate_relationship_integrity(document: &GlossaryDocument, part: &dyn Part) -> Result<()> {
    let mut referenced = HashSet::new();
    if let Some(background) = &document.background_xml {
        collect_relationship_references(&parse_dom(background)?, &mut referenced);
    }
    for entry in &document.entries {
        if let Some(body) = &entry.body_xml {
            collect_relationship_references(&parse_dom(body)?, &mut referenced);
        }
    }
    for id in referenced {
        if part.rels().get(&id).is_none() {
            return Err(invalid(format!(
                "glossary XML references missing relationship '{id}'"
            )));
        }
    }
    for relationship in part.rels().iter() {
        if !allowed_relationship_type(relationship.reltype()) {
            return Err(invalid(format!(
                "unsupported glossary relationship type '{}'",
                relationship.reltype()
            )));
        }
        if relationship.is_external() && !relationship.reltype().ends_with("/hyperlink") {
            return Err(invalid("only glossary hyperlinks may be external"));
        }
    }
    Ok(())
}

fn collect_relationship_references(node: &Node, output: &mut HashSet<String>) {
    for attribute in &node.attrs {
        if matches!(attribute.ns.as_str(), R | RS)
            && matches!(attribute.l.as_str(), "id" | "embed" | "link")
        {
            output.insert(attribute.v.clone());
        }
    }
    for content in &node.content {
        if let Content::Node(child) = content {
            collect_relationship_references(child, output);
        }
    }
}

fn validate_staged_targets(
    package: &OpcPackage,
    root: &BlobPart,
    staged: &HashMap<PackURI, BlobPart>,
    old_owned: &HashSet<PackURI>,
) -> Result<()> {
    for part in
        std::iter::once(root as &dyn Part).chain(staged.values().map(|part| part as &dyn Part))
    {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            let target = relationship.target_partname()?;
            if target == *root.partname() || staged.contains_key(&target) {
                continue;
            }
            if old_owned.contains(&target) || package.get_part(&target).is_err() {
                return Err(invalid(format!(
                    "glossary relationship '{}' has dangling target '{}'",
                    relationship.r_id(),
                    target.as_str()
                )));
            }
        }
    }
    Ok(())
}

fn glossary_owned_parts(package: &OpcPackage, root: &PackURI) -> Result<HashSet<PackURI>> {
    let mut owned = HashSet::new();
    let mut queue = VecDeque::from([root.clone()]);
    while let Some(uri) = queue.pop_front() {
        if !uri.as_str().starts_with("/word/glossary/") || !owned.insert(uri.clone()) {
            continue;
        }
        let part = package.get_part(&uri)?;
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            let target = relationship.target_partname()?;
            if target.as_str().starts_with("/word/glossary/") {
                queue.push_back(target);
            }
        }
    }
    Ok(owned)
}

#[derive(Clone)]
struct Attr {
    q: String,
    ns: String,
    l: String,
    v: String,
}
#[derive(Clone)]
enum Content {
    Node(Node),
    Text(String),
    CData(String),
    Comment(String),
}
#[derive(Clone)]
struct Node {
    q: String,
    ns: String,
    l: String,
    attrs: Vec<Attr>,
    bindings: Vec<(String, String)>,
    content: Vec<Content>,
}
fn parse_dom(xml: &[u8]) -> Result<Node> {
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut rd = Reader::from_reader(xml);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut count = 0;
    loop {
        let d = rd.decoder();
        match rd.read_event() {
            Ok(Event::Start(e)) => {
                count += 1;
                if count > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(invalid("glossary XML resource limit exceeded"));
                }
                stack.push(make(&e, d, &stack)?);
            },
            Ok(Event::Empty(e)) => {
                count += 1;
                if count > MAX_NODES {
                    return Err(invalid("glossary node limit exceeded"));
                }
                let n = make(&e, d, &stack)?;
                attach(&mut stack, &mut root, n)?;
            },
            Ok(Event::End(_)) => {
                let n = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected glossary closing element"))?;
                attach(&mut stack, &mut root, n)?;
            },
            Ok(Event::Text(t)) => {
                let v = t.decode().map_err(xml_error)?.into_owned();
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Text(v))
                } else if !v.trim().is_empty() {
                    return Err(invalid("text outside glossary root"));
                }
            },
            Ok(Event::CData(t)) => {
                if let Some(n) = stack.last_mut() {
                    n.content
                        .push(Content::CData(t.decode().map_err(xml_error)?.into_owned()))
                } else {
                    return Err(invalid("CDATA outside glossary root"));
                }
            },
            Ok(Event::Comment(t)) => {
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Comment(
                        t.decode().map_err(xml_error)?.into_owned(),
                    ))
                }
            },
            Ok(Event::GeneralRef(t)) => {
                if let Some(n) = stack.last_mut() {
                    n.content
                        .push(Content::Text(litchi_ooxml_common::xml::decode_xml_reference(&t)?))
                } else {
                    return Err(invalid("entity outside glossary root"));
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Ok(Event::Decl(_)) => {},
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(e)),
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated glossary XML"));
    }
    root.ok_or_else(|| invalid("missing glossary root"))
}
fn make(e: &BytesStart<'_>, d: Decoder, stack: &[Node]) -> Result<Node> {
    let q = std::str::from_utf8(e.name().as_ref())
        .map_err(xml_error)?
        .to_string();
    let mut bindings = stack.last().map(|x| x.bindings.clone()).unwrap_or_default();
    let mut raw = Vec::new();
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        raw.push((
            std::str::from_utf8(a.key.as_ref())
                .map_err(xml_error)?
                .to_string(),
            a.decoded_and_normalized_value(XmlVersion::Implicit1_0, d)
                .map_err(xml_error)?
                .into_owned(),
        ));
    }
    for (k, v) in &raw {
        if k == "xmlns" || k.starts_with("xmlns:") {
            let key = k.strip_prefix("xmlns:").unwrap_or("").to_string();
            if let Some(x) = bindings.iter_mut().find(|x| x.0 == key) {
                x.1 = v.clone()
            } else {
                bindings.push((key, v.clone()))
            }
        }
    }
    let (pr, lo) = split(&q)?;
    let local = lo.to_string();
    let ns = resolve(&bindings, pr)?;
    let mut attrs = Vec::new();
    for (q, v) in raw {
        if q == "xmlns" || q.starts_with("xmlns:") {
            continue;
        }
        let (pr, lo) = split(&q)?;
        let ans = if pr.is_empty() {
            String::new()
        } else {
            resolve(&bindings, pr)?
        };
        let local = lo.to_string();
        attrs.push(Attr {
            q,
            ns: ans,
            l: local,
            v,
        });
    }
    Ok(Node {
        q,
        ns,
        l: local,
        attrs,
        bindings,
        content: Vec::new(),
    })
}
fn attach(stack: &mut [Node], root: &mut Option<Node>, n: Node) -> Result<()> {
    if let Some(p) = stack.last_mut() {
        p.content.push(Content::Node(n))
    } else if root.replace(n).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn project(n: &Node) -> Result<GlossaryDocument> {
    expect(n, "glossaryDocument")?;
    noattrs(n)?;
    let c = kids(n)?;
    let mut out = GlossaryDocument::default();
    let mut index = 0;
    if c.first().is_some_and(|x| x.l == "background") {
        out.background_xml = Some(node_xml(c[0], false)?);
        index = 1;
    }
    if index < c.len() {
        let parts = c[index];
        expect(parts, "docParts")?;
        noattrs(parts)?;
        for p in kids(parts)? {
            if out.entries.len() >= MAX_PARTS {
                return Err(invalid("glossary entry limit exceeded"));
            }
            out.entries.push(parse_entry(p)?);
        }
        index += 1;
    }
    if index != c.len() {
        return Err(invalid("unexpected glossary root child"));
    }
    validate(&out)?;
    Ok(out)
}
fn parse_entry(n: &Node) -> Result<GlossaryEntry> {
    expect(n, "docPart")?;
    noattrs(n)?;
    let c = kids(n)?;
    if c.len() > 2 {
        return Err(invalid("docPart has too many children"));
    }
    let mut properties = None;
    let mut body_xml = None;
    let mut i = 0;
    if c.first().is_some_and(|x| x.l == "docPartPr") {
        properties = Some(parse_properties(c[0])?);
        i = 1;
    }
    if i < c.len() {
        expect(c[i], "docPartBody")?;
        body_xml = Some(node_xml(c[i], false)?);
        i += 1;
    }
    if i != c.len() {
        return Err(invalid("invalid docPart child order"));
    }
    Ok(GlossaryEntry {
        properties,
        body_xml,
    })
}
fn parse_properties(n: &Node) -> Result<DocPartProperties> {
    noattrs(n)?;
    let mut out = DocPartProperties::default();
    let mut order = 0;
    for c in kids(n)? {
        expect_w(c)?;
        let i = match c.l.as_str() {
            "name" => 0,
            "style" => 1,
            "category" => 2,
            "types" => 3,
            "behaviors" => 4,
            "description" => 5,
            "guid" => 6,
            _ => return Err(invalid("unexpected docPartPr child")),
        };
        if i < order {
            return Err(invalid("docPartPr children out of order"));
        }
        order = i;
        match i {
            0 => {
                if out.name.is_some() {
                    return Err(invalid("duplicate building-block name"));
                }
                let value = wattr_get(c, "val")?.ok_or_else(|| invalid("missing w:val"))?;
                bounded(&value)?;
                out.name = Some(DocPartName {
                    value,
                    decorated: onoff(c, "decorated")?,
                });
                only_w(c, &["val", "decorated"])?;
                leaf(c)?;
            },
            1 => set(&mut out.style, wval(c)?)?,
            2 => {
                if out.category.is_some() {
                    return Err(invalid("duplicate category"));
                }
                let k = kids(c)?;
                if k.len() != 2 || k[0].l != "name" || k[1].l != "gallery" {
                    return Err(invalid("category requires name then gallery"));
                }
                out.category = Some(DocPartCategory {
                    name: wval(k[0])?,
                    gallery: DocPartGallery::parse(&wval(k[1])?)?,
                });
            },
            3 => {
                if !out.types.is_empty() || out.all_types.is_some() {
                    return Err(invalid("duplicate types"));
                }
                out.all_types = onoff(c, "all")?;
                only_w(c, &["all"])?;
                for t in kids(c)? {
                    if out.types.len() >= MAX_VALUES {
                        return Err(invalid("type limit exceeded"));
                    }
                    expect(t, "type")?;
                    out.types.push(parse_type(&wval(t)?)?);
                }
            },
            4 => {
                if !out.behaviors.is_empty() {
                    return Err(invalid("duplicate behaviors"));
                }
                noattrs(c)?;
                for b in kids(c)? {
                    if out.behaviors.len() >= MAX_VALUES {
                        return Err(invalid("behavior limit exceeded"));
                    }
                    expect(b, "behavior")?;
                    out.behaviors.push(parse_behavior(&wval(b)?)?);
                }
            },
            5 => set(&mut out.description, wval(c)?)?,
            6 => set(&mut out.guid, wval(c)?)?,
            _ => unreachable!(),
        }
    }
    Ok(out)
}

fn write_entry(x: &mut String, e: &GlossaryEntry, s: bool) -> Result<()> {
    x.push_str("<w:docPart>");
    if let Some(p) = &e.properties {
        write_properties(x, p);
    }
    if let Some(b) = &e.body_xml {
        opaque(x, b, s)?;
    }
    x.push_str("</w:docPart>");
    Ok(())
}
fn write_properties(x: &mut String, p: &DocPartProperties) {
    x.push_str("<w:docPartPr>");
    if let Some(n) = &p.name {
        x.push_str("<w:name");
        wattr(x, "val", &n.value);
        wonoff(x, "decorated", n.decorated);
        x.push_str("/>");
    }
    if let Some(v) = &p.style {
        leafval(x, "style", v);
    }
    if let Some(c) = &p.category {
        x.push_str("<w:category>");
        leafval(x, "name", &c.name);
        leafval(x, "gallery", c.gallery.as_str());
        x.push_str("</w:category>");
    }
    if p.all_types.is_some() || !p.types.is_empty() {
        x.push_str("<w:types");
        wonoff(x, "all", p.all_types);
        x.push('>');
        for v in &p.types {
            leafval(x, "type", type_str(*v));
        }
        x.push_str("</w:types>");
    }
    if !p.behaviors.is_empty() {
        x.push_str("<w:behaviors>");
        for v in &p.behaviors {
            leafval(x, "behavior", behavior_str(*v));
        }
        x.push_str("</w:behaviors>");
    }
    if let Some(v) = &p.description {
        leafval(x, "description", v);
    }
    if let Some(v) = &p.guid {
        leafval(x, "guid", v);
    }
    x.push_str("</w:docPartPr>");
}
fn leafval(x: &mut String, n: &str, v: &str) {
    x.push_str(&format!("<w:{n}"));
    wattr(x, "val", v);
    x.push_str("/>")
}
fn wattr(x: &mut String, n: &str, v: &str) {
    x.push_str(&format!(" w:{n}=\""));
    esc(x, v);
    x.push('"')
}
fn wonoff(x: &mut String, n: &str, v: Option<bool>) {
    if let Some(v) = v {
        wattr(x, n, if v { "1" } else { "0" })
    }
}

fn validate(v: &GlossaryDocument) -> Result<()> {
    if v.entries.len() > MAX_PARTS {
        return Err(invalid("glossary entry limit exceeded"));
    }
    if let Some(x) = &v.background_xml {
        let root = parse_dom(x)?;
        expect(&root, "background")?;
    }
    let mut names = HashSet::new();
    for e in &v.entries {
        if let Some(x) = &e.body_xml {
            let root = parse_dom(x)?;
            expect(&root, "docPartBody")?;
        }
        if let Some(p) = &e.properties {
            for s in [
                p.name.as_ref().map(|x| x.value.as_str()),
                p.style.as_deref(),
                p.category.as_ref().map(|x| x.name.as_str()),
                p.description.as_deref(),
                p.guid.as_deref(),
            ]
            .into_iter()
            .flatten()
            {
                bounded(s)?;
            }
            if p.types.len() > MAX_VALUES || p.behaviors.len() > MAX_VALUES {
                return Err(invalid("glossary metadata value limit exceeded"));
            }
            if let Some(name) = &p.name {
                if name.value.trim().is_empty() {
                    return Err(invalid("building-block name cannot be empty"));
                }
                if !names.insert(name.value.to_lowercase()) {
                    return Err(invalid(format!(
                        "duplicate case-insensitive building-block name '{}'",
                        name.value
                    )));
                }
            }
            if let Some(category) = &p.category
                && category.name.trim().is_empty()
            {
                return Err(invalid("building-block category name cannot be empty"));
            }
            let unique_types: HashSet<_> = p.types.iter().copied().collect();
            let unique_behaviors: HashSet<_> = p.behaviors.iter().copied().collect();
            if unique_types.len() != p.types.len() || unique_behaviors.len() != p.behaviors.len() {
                return Err(invalid("duplicate glossary type or insertion behavior"));
            }
            if let Some(guid) = &p.guid
                && !valid_guid(guid)
            {
                return Err(invalid(format!("invalid building-block GUID '{guid}'")));
            }
        }
    }
    Ok(())
}

fn valid_guid(value: &str) -> bool {
    let value = value
        .strip_prefix('{')
        .and_then(|value| value.strip_suffix('}'))
        .unwrap_or(value);
    value.len() == 36
        && value.bytes().enumerate().all(|(index, byte)| match index {
            8 | 13 | 18 | 23 => byte == b'-',
            _ => byte.is_ascii_hexdigit(),
        })
}
fn opaque(x: &mut String, b: &[u8], strict: bool) -> Result<()> {
    parse_dom(b)?;
    let mut s = std::str::from_utf8(b).map_err(xml_error)?.to_string();
    if strict {
        s = s.replace(W, WS).replace(R, RS)
    } else {
        s = s.replace(WS, W).replace(RS, R)
    }
    x.push_str(&s);
    Ok(())
}
fn node_xml(n: &Node, s: bool) -> Result<Vec<u8>> {
    let mut x = String::new();
    node_write(&mut x, n, s)?;
    Ok(x.into_bytes())
}
fn node_write(x: &mut String, n: &Node, s: bool) -> Result<()> {
    x.push('<');
    x.push_str(&n.q);
    for (p, u) in &n.bindings {
        if p.is_empty() {
            x.push_str(" xmlns=\"")
        } else {
            x.push_str(" xmlns:");
            x.push_str(p);
            x.push_str("=\"")
        }
        esc(x, mapns(u, s));
        x.push('"');
    }
    for a in &n.attrs {
        x.push(' ');
        x.push_str(&a.q);
        x.push_str("=\"");
        esc(x, &a.v);
        x.push('"');
    }
    if n.content.is_empty() {
        x.push_str("/>");
        return Ok(());
    }
    x.push('>');
    for c in &n.content {
        match c {
            Content::Node(n) => node_write(x, n, s)?,
            Content::Text(v) => text(x, v),
            Content::CData(v) => {
                x.push_str("<![CDATA[");
                x.push_str(v);
                x.push_str("]]>");
            },
            Content::Comment(v) => {
                x.push_str("<!--");
                x.push_str(v);
                x.push_str("-->");
            },
        }
    }
    x.push_str("</");
    x.push_str(&n.q);
    x.push('>');
    Ok(())
}
fn mapns(v: &str, s: bool) -> &str {
    if s {
        match v {
            W => WS,
            R => RS,
            _ => v,
        }
    } else {
        match v {
            WS => W,
            RS => R,
            _ => v,
        }
    }
}

fn kids(n: &Node) -> Result<Vec<&Node>> {
    let mut v = Vec::new();
    for c in &n.content {
        match c {
            Content::Node(x) => v.push(x),
            Content::Text(x) if x.trim().is_empty() => {},
            Content::Comment(_) => {},
            _ => return Err(invalid("unexpected text in typed glossary metadata")),
        }
    }
    Ok(v)
}
fn leaf(n: &Node) -> Result<()> {
    if kids(n)?.is_empty() {
        Ok(())
    } else {
        Err(invalid("glossary metadata leaf has children"))
    }
}
fn expect(n: &Node, l: &str) -> Result<()> {
    if (n.ns == W || n.ns == WS) && n.l == l {
        Ok(())
    } else {
        Err(invalid(format!("expected WordprocessingML {l}")))
    }
}
fn expect_w(n: &Node) -> Result<()> {
    if n.ns == W || n.ns == WS {
        Ok(())
    } else {
        Err(invalid("expected WordprocessingML metadata"))
    }
}
fn wval(n: &Node) -> Result<String> {
    let v = wattr_get(n, "val")?.ok_or_else(|| invalid("missing w:val"))?;
    bounded(&v)?;
    only_w(n, &["val"])?;
    leaf(n)?;
    Ok(v)
}
fn wattr_get(n: &Node, l: &str) -> Result<Option<String>> {
    let mut v = None;
    for a in &n.attrs {
        if (a.ns == W || a.ns == WS) && a.l == l {
            if v.is_some() {
                return Err(invalid("duplicate WordprocessingML attribute"));
            }
            v = Some(a.v.clone());
        }
    }
    Ok(v)
}
fn onoff(n: &Node, l: &str) -> Result<Option<bool>> {
    match wattr_get(n, l)?.as_deref() {
        None => Ok(None),
        Some("1" | "true" | "on") => Ok(Some(true)),
        Some("0" | "false" | "off") => Ok(Some(false)),
        _ => Err(invalid(format!("invalid on/off attribute '{l}'"))),
    }
}
fn only_w(n: &Node, allowed: &[&str]) -> Result<()> {
    for a in &n.attrs {
        if !((a.ns == W || a.ns == WS) && allowed.contains(&a.l.as_str())) {
            return Err(invalid(format!("unexpected glossary attribute '{}'", a.q)));
        }
    }
    Ok(())
}
fn noattrs(n: &Node) -> Result<()> {
    if n.attrs.is_empty() {
        Ok(())
    } else {
        Err(invalid("unexpected glossary attributes"))
    }
}
fn set<T>(slot: &mut Option<T>, v: T) -> Result<()> {
    if slot.replace(v).is_some() {
        Err(invalid("duplicate glossary metadata"))
    } else {
        Ok(())
    }
}
fn parse_type(v: &str) -> Result<DocPartType> {
    match v {
        "none" => Ok(DocPartType::None),
        "normal" => Ok(DocPartType::Normal),
        "autoExp" => Ok(DocPartType::AutoExpand),
        "toolbar" => Ok(DocPartType::Toolbar),
        "speller" => Ok(DocPartType::Speller),
        "formFld" => Ok(DocPartType::FormField),
        "bbPlcHdr" => Ok(DocPartType::SdtPlaceholder),
        _ => Err(invalid(format!("invalid document-part type '{v}'"))),
    }
}
fn type_str(v: DocPartType) -> &'static str {
    match v {
        DocPartType::None => "none",
        DocPartType::Normal => "normal",
        DocPartType::AutoExpand => "autoExp",
        DocPartType::Toolbar => "toolbar",
        DocPartType::Speller => "speller",
        DocPartType::FormField => "formFld",
        DocPartType::SdtPlaceholder => "bbPlcHdr",
    }
}
fn parse_behavior(v: &str) -> Result<InsertionBehavior> {
    match v {
        "content" => Ok(InsertionBehavior::Content),
        "p" => Ok(InsertionBehavior::Paragraph),
        "pg" => Ok(InsertionBehavior::Page),
        _ => Err(invalid(format!("invalid insertion behavior '{v}'"))),
    }
}
fn behavior_str(v: InsertionBehavior) -> &'static str {
    match v {
        InsertionBehavior::Content => "content",
        InsertionBehavior::Paragraph => "p",
        InsertionBehavior::Page => "pg",
    }
}
const GALLERIES: &[&str] = &[
    "placeholder",
    "any",
    "default",
    "docParts",
    "coverPg",
    "eq",
    "ftrs",
    "hdrs",
    "pgNum",
    "tbls",
    "watermarks",
    "autoTxt",
    "txtBox",
    "pgNumT",
    "pgNumB",
    "pgNumMargins",
    "tblOfContents",
    "bib",
    "custQuickParts",
    "custCoverPg",
    "custEq",
    "custFtrs",
    "custHdrs",
    "custPgNum",
    "custTbls",
    "custWatermarks",
    "custAutoTxt",
    "custTxtBox",
    "custPgNumT",
    "custPgNumB",
    "custPgNumMargins",
    "custTblOfContents",
    "custBib",
    "custom1",
    "custom2",
    "custom3",
    "custom4",
    "custom5",
];
fn split(q: &str) -> Result<(&str, &str)> {
    if let Some((p, l)) = q.split_once(':') {
        if l.is_empty() || l.contains(':') {
            return Err(invalid("invalid QName"));
        }
        Ok((p, l))
    } else {
        Ok(("", q))
    }
}
fn resolve(b: &[(String, String)], p: &str) -> Result<String> {
    if p == "xml" {
        return Ok("http://www.w3.org/XML/1998/namespace".into());
    }
    b.iter()
        .rev()
        .find(|x| x.0 == p)
        .map(|x| x.1.clone())
        .ok_or_else(|| invalid(format!("unbound prefix '{p}'")))
}
fn bounded(v: &str) -> Result<()> {
    if v.len() > MAX_STRING {
        Err(invalid("glossary metadata string exceeds 1 MiB"))
    } else {
        Ok(())
    }
}
fn esc(x: &mut String, v: &str) {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;"),
            '<' => x.push_str("&lt;"),
            '"' => x.push_str("&quot;"),
            '\r' => x.push_str("&#xD;"),
            '\n' => x.push_str("&#xA;"),
            '\t' => x.push_str("&#x9;"),
            _ => x.push(c),
        }
    }
}
fn text(x: &mut String, v: &str) {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;"),
            '<' => x.push_str("&lt;"),
            '>' => x.push_str("&gt;"),
            _ => x.push(c),
        }
    }
}
fn invalid(v: impl Into<String>) -> Box<dyn std::error::Error + Send + Sync> {
    std::io::Error::new(std::io::ErrorKind::InvalidData, v.into()).into()
}
fn xml_error(e: impl std::fmt::Display) -> Box<dyn std::error::Error + Send + Sync> {
    invalid(e.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    fn f(b: &[u8]) -> GlossaryDocument {
        let p = OpcPackage::from_bytes(b).unwrap();
        load_from_package(&p).unwrap().unwrap()
    }
    #[test]
    fn poi_placeholders_and_strict_roundtrip() {
        let v = f(include_bytes!(
            "../../../../test-data/poi/test-data/document/Bug54849.docx"
        ));
        assert_eq!(v.entries.len(), 3);
        let p = v.entries[0].properties.as_ref().unwrap();
        assert_eq!(p.types, [DocPartType::SdtPlaceholder]);
        assert_eq!(p.category.as_ref().unwrap().gallery.as_str(), "placeholder");
        let x = v.to_xml(true).unwrap();
        assert_eq!(GlossaryDocument::parse(&x).unwrap().entries.len(), 3);
    }
    #[test]
    fn libreoffice_multiple_autotext_and_tables_stay_inert() {
        let v = f(include_bytes!(
            "../../../../test-data/libreoffice-core/sw/qa/extras/uiwriter/data/autotext-multiple.dotx"
        ));
        assert_eq!(v.entries.len(), 3);
        assert_eq!(
            v.entries[0]
                .properties
                .as_ref()
                .unwrap()
                .name
                .as_ref()
                .unwrap()
                .value,
            "Multiple"
        );
        let body = std::str::from_utf8(v.entries[2].body_xml.as_deref().unwrap()).unwrap();
        assert!(body.contains("w:tbl"));
        assert!(body.contains("jksdjkfdskjfds"));
    }
    #[test]
    fn libreoffice_empty_glossary_is_valid() {
        let v = f(include_bytes!(
            "../../../../test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/testGlossary.docx"
        ));
        assert!(v.entries.is_empty());
        assert!(v.to_xml(false).is_ok());
    }
    #[test]
    fn strict_mce_and_malformed() {
        let x = format!(
            r#"<w:glossaryDocument xmlns:w="{WS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:u" mc:Ignorable="u"><w:docParts><mc:AlternateContent><mc:Choice Requires="u"><u:x/></mc:Choice><mc:Fallback><w:docPart><w:docPartPr><w:name w:val="MCE"/><w:behaviors><w:behavior w:val="p"/></w:behaviors></w:docPartPr><w:docPartBody><w:p/></w:docPartBody></w:docPart></mc:Fallback></mc:AlternateContent></w:docParts></w:glossaryDocument>"#
        );
        assert_eq!(
            GlossaryDocument::parse(x.as_bytes()).unwrap().entries.len(),
            1
        );
        for bad in [
            format!(
                r#"<w:glossaryDocument xmlns:w="{W}"><w:docParts><w:docPart><w:docPartPr><w:behaviors><w:behavior w:val="run"/></w:behaviors></w:docPartPr></w:docPart></w:docParts></w:glossaryDocument>"#
            ),
            format!(r#"<!DOCTYPE x><w:glossaryDocument xmlns:w="{W}"/>"#),
        ] {
            assert!(
                GlossaryDocument::parse(bad.as_bytes()).is_err(),
                "accepted {bad}"
            );
        }
        let v = GlossaryDocument {
            background_xml: Some(b"<?bad?><w:background/>".to_vec()),
            ..GlossaryDocument::default()
        };
        assert!(v.to_xml(false).is_err());
    }
}
