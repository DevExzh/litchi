//! Bounded, inert PresentationML notes-slide and notes-master package graphs.

use crate::{Error, Result};
use litchi_ooxml_common::mce::process_ooxml;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI, TargetMode};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const AS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const SLIDE_CT: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const THEME_CT: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
const MAX_PRESENTATION_XML: usize = 32 * 1024 * 1024;
const MAX_SLIDE_XML: usize = 16 * 1024 * 1024;
const MAX_NOTES_XML: usize = 8 * 1024 * 1024;
const MAX_MASTER_XML: usize = 16 * 1024 * 1024;
const MAX_THEME_XML: usize = 16 * 1024 * 1024;
const MAX_TOTAL_BYTES: usize = 64 * 1024 * 1024;
const MAX_NOTES_SLIDES: usize = 4096;
const MAX_NODES: usize = 100_000;
const MAX_DEPTH: usize = 128;
const MAX_ATTRIBUTES: usize = 500_000;
const MAX_ATTRIBUTE_BYTES: usize = 8 * 1024 * 1024;

const NOTES_XML_DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const NOTES_XML_BODY_PREFIX: &str = concat!(
    "<p:cSld><p:spTree>",
    "<p:nvGrpSpPr>",
    r#"<p:cNvPr id="1" name=""/>"#,
    "<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>",
    "<p:grpSpPr>",
    r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>"#,
    r#"<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>"#,
    "</p:grpSpPr><p:sp><p:nvSpPr>",
    r#"<p:cNvPr id="2" name="Notes Placeholder"/>"#,
    r#"<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>"#,
    r#"<p:nvPr><p:ph type="body" idx="1"/></p:nvPr>"#,
    "</p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>",
    r#"<a:rPr lang="en-US" dirty="0"/><a:t>"#,
);
const NOTES_XML_SUFFIX: &str = concat!(
    "</a:t></a:r></a:p></p:txBody></p:sp>",
    "</p:spTree></p:cSld>",
    r#"<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"#,
    "</p:notes>",
);

/// Return the deterministic Transitional notes-master producer template.
///
/// This static asset is intentionally used only by the Transitional legacy
/// writer. Strict graph CRUD preserves caller-owned Strict master and theme
/// payloads; it does not currently synthesize either resource.
pub fn master_xml() -> &'static str {
    include_str!("notes/resources/generated/notesMaster.xml")
}

/// Encode one bounded Transitional plain-text speaker-notes slide.
///
/// Package relationship allocation remains the responsibility of the
/// concrete package host; this function owns only PresentationML grammar.
pub fn write_text(text: &str) -> Result<Vec<u8>> {
    write_text_with(Conformance::Transitional, text)
}

/// Encode one bounded plain-text speaker-notes slide in the chosen dialect.
///
/// Use the graph's [`Graph::conformance`] when replacing an existing notes
/// slide. The output allocation is returned by value so [`Slide::replace_xml`]
/// and [`put`] can move it directly into the package.
pub fn write_text_with(conformance: Conformance, text: &str) -> Result<Vec<u8>> {
    if text.len() > MAX_NOTES_XML {
        return Err(Error::Limit {
            resource: "speaker-notes text bytes",
            limit: MAX_NOTES_XML,
        });
    }
    if !text.chars().all(is_xml_char) {
        return Err(invalid("speaker notes contain an invalid XML character"));
    }
    let escaped = quick_xml::escape::escape(text);
    let prefix = [
        NOTES_XML_DECLARATION,
        r#"<p:notes xmlns:p=""#,
        conformance.p(),
        r#"" xmlns:a=""#,
        conformance.a(),
        r#"" xmlns:r=""#,
        conformance.r(),
        r#"">"#,
        NOTES_XML_BODY_PREFIX,
    ];
    let prefix_len = prefix
        .iter()
        .try_fold(0usize, |len, part| len.checked_add(part.len()))
        .ok_or_else(|| invalid("speaker-notes XML length overflow"))?;
    let capacity = prefix_len
        .checked_add(escaped.len())
        .and_then(|len| len.checked_add(NOTES_XML_SUFFIX.len()))
        .ok_or_else(|| invalid("speaker-notes XML length overflow"))?;
    if capacity > MAX_NOTES_XML {
        return Err(Error::Limit {
            resource: "speaker-notes XML bytes",
            limit: MAX_NOTES_XML,
        });
    }
    let mut xml = String::new();
    xml.try_reserve_exact(capacity)
        .map_err(|source| allocation("speaker-notes XML", source))?;
    for part in prefix {
        xml.push_str(part);
    }
    xml.push_str(&escaped);
    xml.push_str(NOTES_XML_SUFFIX);
    Ok(xml.into_bytes())
}

fn is_xml_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(value as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

/// PresentationML namespace and relationship conformance used by a notes graph.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    fn p(self) -> &'static str {
        if self == Self::Strict { PS } else { P }
    }
    fn a(self) -> &'static str {
        if self == Self::Strict { AS } else { A }
    }
    fn r(self) -> &'static str {
        if self == Self::Strict { RS } else { R }
    }
    fn notes_slide_rel(self) -> &'static str {
        if self == Self::Strict {
            rt::STRICT_NOTES_SLIDE
        } else {
            rt::NOTES_SLIDE
        }
    }
    fn notes_master_rel(self) -> &'static str {
        if self == Self::Strict {
            rt::STRICT_NOTES_MASTER
        } else {
            rt::NOTES_MASTER
        }
    }
    fn slide_rel(self) -> &'static str {
        if self == Self::Strict {
            "http://purl.oclc.org/ooxml/officeDocument/relationships/slide"
        } else {
            rt::SLIDE
        }
    }
    fn theme_rel(self) -> &'static str {
        if self == Self::Strict {
            "http://purl.oclc.org/ooxml/officeDocument/relationships/theme"
        } else {
            rt::THEME
        }
    }
}

/// Owned notes-master theme resource.
#[derive(Debug, PartialEq, Eq)]
pub struct Theme {
    relationship_id: String,
    part_name: String,
    content_type: String,
    data: Vec<u8>,
}

impl Theme {
    /// Return the validated package part name for diagnostics.
    pub fn part(&self) -> &str {
        &self.part_name
    }

    /// Return the validated resource content type.
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Lend the inert theme XML payload.
    pub fn xml(&self) -> &[u8] {
        &self.data
    }

    /// Replace the inert theme XML, returning the previous allocation.
    pub fn replace_xml(&mut self, xml: Vec<u8>) -> Vec<u8> {
        std::mem::replace(&mut self.data, xml)
    }
}

/// Owned notes-master resource and its theme.
#[derive(Debug, PartialEq, Eq)]
pub struct Master {
    presentation_relationship_id: String,
    part_name: String,
    content_type: String,
    data: Vec<u8>,
    theme: Theme,
}

impl Master {
    /// Return the validated package part name for diagnostics.
    pub fn part(&self) -> &str {
        &self.part_name
    }

    /// Return the validated resource content type.
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Lend the inert notes-master XML payload.
    pub fn xml(&self) -> &[u8] {
        &self.data
    }

    /// Replace the inert notes-master XML, returning the previous allocation.
    pub fn replace_xml(&mut self, xml: Vec<u8>) -> Vec<u8> {
        std::mem::replace(&mut self.data, xml)
    }

    /// Lend the owned notes-master theme resource.
    pub fn theme(&self) -> &Theme {
        &self.theme
    }

    /// Mutably lend the owned notes-master theme resource.
    pub fn theme_mut(&mut self) -> &mut Theme {
        &mut self.theme
    }
}

/// Owned speaker-notes resource for one slide.
#[derive(Debug, PartialEq, Eq)]
pub struct Slide {
    slide_part_name: String,
    slide_relationship_id: String,
    part_name: String,
    content_type: String,
    data: Vec<u8>,
    backlink_relationship_id: String,
    notes_master_relationship_id: String,
}

impl Slide {
    /// Return the validated owning slide part name for diagnostics.
    pub fn owner(&self) -> &str {
        &self.slide_part_name
    }

    /// Return the validated notes-slide part name for diagnostics.
    pub fn part(&self) -> &str {
        &self.part_name
    }

    /// Return the validated resource content type.
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Lend the inert notes-slide XML payload.
    pub fn xml(&self) -> &[u8] {
        &self.data
    }

    /// Replace the inert notes-slide XML, returning the previous allocation.
    pub fn replace_xml(&mut self, xml: Vec<u8>) -> Vec<u8> {
        std::mem::replace(&mut self.data, xml)
    }

    /// Flatten the inert notes XML to its DrawingML text runs.
    pub fn text(&self) -> Result<Option<String>> {
        let processed = process_ooxml(&self.data)?;
        let mut reader = Reader::from_reader(processed.as_ref());
        reader.config_mut().trim_text(true);
        let mut in_text = false;
        let mut value = String::new();
        loop {
            match reader.read_event() {
                Ok(Event::Start(element)) if element.local_name().as_ref() == b"t" => {
                    in_text = true
                },
                Ok(Event::Text(text)) if in_text => {
                    let decoded = text.decode().map_err(xml_error)?;
                    let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                    if !value.is_empty() {
                        value.push('\n');
                    }
                    value.push_str(&decoded);
                },
                Ok(Event::End(element)) if element.local_name().as_ref() == b"t" => in_text = false,
                Ok(Event::Eof) => break,
                Err(error) => return Err(xml_error(error)),
                _ => {},
            }
        }
        Ok((!value.is_empty()).then_some(value))
    }
}

/// Complete owned notes graph for one presentation.
///
/// Topology identities remain private. XML payload replacement is explicit,
/// and [`put`] consumes the graph so successful storage moves those buffers.
#[derive(Debug, PartialEq, Eq)]
pub struct Graph {
    conformance: Conformance,
    master: Master,
    slides: Vec<Slide>,
}

impl Graph {
    /// Return the graph's Strict or Transitional namespace profile.
    pub fn conformance(&self) -> Conformance {
        self.conformance
    }

    /// Lend the shared notes-master resource.
    pub fn master(&self) -> &Master {
        &self.master
    }

    /// Mutably lend the shared notes-master resource.
    pub fn master_mut(&mut self) -> &mut Master {
        &mut self.master
    }

    /// Lend notes slides in presentation order.
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Mutably lend notes slides in presentation order.
    pub fn slides_mut(&mut self) -> &mut [Slide] {
        &mut self.slides
    }
}

#[derive(Debug)]
struct ThemeIndex {
    relationship_id: String,
    part_name: PackURI,
    content_type: String,
}

#[derive(Debug)]
struct MasterIndex {
    presentation_relationship_id: String,
    part_name: PackURI,
    content_type: String,
    theme: ThemeIndex,
}

#[derive(Debug)]
struct SlideIndex {
    slide_part_name: PackURI,
    slide_relationship_id: String,
    part_name: PackURI,
    content_type: String,
    backlink_relationship_id: String,
    notes_master_relationship_id: String,
}

#[derive(Debug)]
struct GraphIndex {
    conformance: Conformance,
    master: MasterIndex,
    slides: Vec<SlideIndex>,
}

#[derive(Default)]
struct XmlScan {
    relationship_attributes: Vec<String>,
    notes_master_ids: Vec<String>,
    slide_ids: Vec<String>,
}

/// Load and validate the complete bounded notes graph for a presentation part.
///
/// The returned graph is lifetime-free and independently editable, so each
/// validated notes, master, and theme payload is copied exactly once. Package
/// deletion uses the metadata-only index and does not perform these copies.
pub fn load(package: &OpcPackage, presentation_name: &PackURI) -> Result<Option<Graph>> {
    let Some(index) = load_index(package, presentation_name)? else {
        return Ok(None);
    };
    Ok(Some(materialize(package, index)?))
}

/// Validate and index the complete notes graph without copying resource payloads.
fn load_index(package: &OpcPackage, presentation_name: &PackURI) -> Result<Option<GraphIndex>> {
    let presentation = package.get_part(presentation_name)?;
    if !is_presentation_main_content_type(presentation.content_type()) {
        return Err(invalid("notes graph requires a PresentationML main part"));
    }
    let conformance = root_conformance(presentation.blob(), MAX_PRESENTATION_XML, "presentation")?;
    let presentation_scan = scan_xml(
        presentation.blob(),
        MAX_PRESENTATION_XML,
        conformance,
        "presentation",
    )?;
    if presentation_scan.notes_master_ids.len() > 1 {
        return Err(invalid(
            "presentation has multiple notesMasterId references",
        ));
    }
    if presentation_scan.slide_ids.len() > MAX_NOTES_SLIDES {
        return Err(limit("presentation slide count", MAX_NOTES_SLIDES));
    }
    let mut slide_sources = Vec::with_capacity(presentation_scan.slide_ids.len());
    for id in &presentation_scan.slide_ids {
        validate_id(id)?;
        let relationship = presentation
            .rels()
            .get(id)
            .ok_or_else(|| invalid("presentation slide reference is missing its relationship"))?;
        if relationship.reltype() != conformance.slide_rel() || relationship.is_external() {
            return Err(invalid(
                "presentation slide relationship has wrong type or target mode",
            ));
        }
        let target = relationship_target(presentation, relationship)?;
        let slide = package.get_part(&target)?;
        if slide.content_type() != SLIDE_CT {
            return Err(invalid("slide has invalid content type"));
        }
        root_conformance(slide.blob(), MAX_SLIDE_XML, "sld").and_then(|actual| {
            if actual == conformance {
                Ok(actual)
            } else {
                Err(invalid("slide conformance differs from presentation"))
            }
        })?;
        slide_sources.push((id.clone(), slide.partname().clone()));
    }
    let master_relationships: Vec<_> = presentation
        .rels()
        .iter()
        .filter(|relationship| is_notes_master_rel(relationship.reltype()))
        .collect();
    let Some(master_id) = presentation_scan.notes_master_ids.first() else {
        if !master_relationships.is_empty()
            || package.iter_parts().any(|part| {
                matches!(
                    part.content_type(),
                    ct::PML_NOTES_MASTER | ct::PML_NOTES_SLIDE
                )
            })
            || slide_sources.iter().any(|(_, name)| {
                package.get_part(name).is_ok_and(|part| {
                    part.rels()
                        .iter()
                        .any(|relationship| is_notes_slide_rel(relationship.reltype()))
                })
            })
        {
            return Err(invalid("orphan notes graph exists without notesMasterId"));
        }
        return Ok(None);
    };
    validate_id(master_id)?;
    if master_relationships.len() != 1 {
        return Err(invalid(
            "notesMasterId and presentation notes-master relationships differ",
        ));
    }
    let master_relationship = presentation
        .rels()
        .get(master_id)
        .ok_or_else(|| invalid("notesMasterId relationship is missing"))?;
    if master_relationship.reltype() != conformance.notes_master_rel()
        || master_relationship.is_external()
    {
        return Err(invalid(
            "notes-master relationship has wrong type or target mode",
        ));
    }
    let master_name = relationship_target(presentation, master_relationship)?;
    validate_leaf_path(&master_name, "/ppt/notesMasters/", "notes master")?;
    let master_part = package.get_part(&master_name)?;
    let master_part_name = master_part.partname().clone();
    require_content_type(master_part, ct::PML_NOTES_MASTER, "notes master")?;
    validate_resource_xml(
        master_part.blob(),
        MAX_MASTER_XML,
        conformance,
        "notesMaster",
        "notes master",
    )?;
    if master_part.rels().iter().count() != 1 {
        return Err(invalid(
            "notes master must have exactly one bounded theme relationship",
        ));
    }
    let theme_relationship = master_part
        .rels()
        .iter()
        .next()
        .ok_or_else(|| invalid("notes master lacks theme relationship"))?;
    validate_id(theme_relationship.r_id())?;
    if theme_relationship.reltype() != conformance.theme_rel() || theme_relationship.is_external() {
        return Err(invalid(
            "notes-master theme relationship has wrong type or target mode",
        ));
    }
    let theme_name = relationship_target(master_part, theme_relationship)?;
    validate_leaf_path(&theme_name, "/ppt/theme/", "notes-master theme")?;
    let theme_part = package.get_part(&theme_name)?;
    let theme_part_name = theme_part.partname().clone();
    require_content_type(theme_part, THEME_CT, "notes-master theme")?;
    validate_resource_xml(
        theme_part.blob(),
        MAX_THEME_XML,
        conformance,
        "theme",
        "notes-master theme",
    )?;
    if theme_part.rels().iter().next().is_some() {
        return Err(invalid(
            "notes-master theme has unsupported outbound relationships",
        ));
    }
    let mut total = checked_add(
        master_part.blob().len(),
        theme_part.blob().len(),
        "aggregate bytes",
    )?;
    let mut slides = Vec::new();
    let mut discovered = BTreeSet::new();
    for (_, slide_name) in &slide_sources {
        let slide_part = package.get_part(slide_name)?;
        let relationships: Vec<_> = slide_part
            .rels()
            .iter()
            .filter(|relationship| is_notes_slide_rel(relationship.reltype()))
            .collect();
        if relationships.len() > 1 {
            return Err(invalid("slide has multiple notes-slide relationships"));
        }
        let Some(relationship) = relationships.first() else {
            continue;
        };
        validate_id(relationship.r_id())?;
        if relationship.reltype() != conformance.notes_slide_rel() || relationship.is_external() {
            return Err(invalid(
                "notes-slide relationship has wrong type or target mode",
            ));
        }
        let notes_name = relationship_target(slide_part, relationship)?;
        validate_leaf_path(&notes_name, "/ppt/notesSlides/", "notes slide")?;
        let notes_part = package.get_part(&notes_name)?;
        let notes_part_name = notes_part.partname().clone();
        if !discovered.insert(notes_part_name.as_str().to_owned()) {
            return Err(invalid("multiple slides reference the same notes slide"));
        }
        require_content_type(notes_part, ct::PML_NOTES_SLIDE, "notes slide")?;
        validate_resource_xml(
            notes_part.blob(),
            MAX_NOTES_XML,
            conformance,
            "notes",
            "notes slide",
        )?;
        if notes_part.rels().iter().count() != 2 {
            return Err(invalid(
                "notes slide must have exactly slide and notes-master relationships",
            ));
        }
        let mut backlink = None;
        let mut notes_master = None;
        for child in notes_part.rels().iter() {
            validate_id(child.r_id())?;
            if child.is_external() {
                return Err(invalid("notes slide has an external relationship"));
            }
            if child.reltype() == conformance.slide_rel() {
                if backlink.replace(child).is_some() {
                    return Err(invalid("notes slide has multiple slide backlinks"));
                }
            } else if child.reltype() == conformance.notes_master_rel() {
                if notes_master.replace(child).is_some() {
                    return Err(invalid(
                        "notes slide has multiple notes-master relationships",
                    ));
                }
            } else {
                return Err(invalid(
                    "notes slide has an unsupported outbound relationship",
                ));
            }
        }
        let backlink = backlink.ok_or_else(|| invalid("notes slide lacks slide backlink"))?;
        let notes_master =
            notes_master.ok_or_else(|| invalid("notes slide lacks notes-master relationship"))?;
        let backlink_target = relationship_target(notes_part, backlink)?;
        if package.get_part(&backlink_target)?.partname() != slide_name {
            return Err(invalid("notes slide backlink targets the wrong slide"));
        }
        let notes_master_target = relationship_target(notes_part, notes_master)?;
        if package.get_part(&notes_master_target)?.partname() != &master_part_name {
            return Err(invalid("notes slide targets the wrong notes master"));
        }
        total = checked_add(total, notes_part.blob().len(), "aggregate bytes")?;
        if total > MAX_TOTAL_BYTES {
            return Err(limit("notes aggregate bytes", MAX_TOTAL_BYTES));
        }
        slides.push(SlideIndex {
            slide_part_name: slide_part.partname().clone(),
            slide_relationship_id: relationship.r_id().to_owned(),
            part_name: notes_part_name,
            content_type: notes_part.content_type().to_owned(),
            backlink_relationship_id: backlink.r_id().to_owned(),
            notes_master_relationship_id: notes_master.r_id().to_owned(),
        });
    }
    if package
        .iter_parts()
        .filter(|part| part.content_type() == ct::PML_NOTES_MASTER)
        .count()
        != 1
        || package
            .iter_parts()
            .filter(|part| part.content_type() == ct::PML_NOTES_SLIDE)
            .any(|part| !discovered.contains(part.partname().as_str()))
    {
        return Err(invalid("package contains orphan notes parts"));
    }
    Ok(Some(GraphIndex {
        conformance,
        master: MasterIndex {
            presentation_relationship_id: master_id.clone(),
            part_name: master_part_name,
            content_type: master_part.content_type().to_owned(),
            theme: ThemeIndex {
                relationship_id: theme_relationship.r_id().to_owned(),
                part_name: theme_part_name,
                content_type: theme_part.content_type().to_owned(),
            },
        },
        slides,
    }))
}

fn materialize(package: &OpcPackage, index: GraphIndex) -> Result<Graph> {
    let master_data = own_blob(
        package.get_part(&index.master.part_name)?.blob(),
        "notes-master payload",
    )?;
    let theme_data = own_blob(
        package.get_part(&index.master.theme.part_name)?.blob(),
        "notes-master theme payload",
    )?;
    let mut slides = Vec::new();
    slides
        .try_reserve(index.slides.len())
        .map_err(|source| allocation("notes-slide graph", source))?;
    for slide in index.slides {
        let data = own_blob(
            package.get_part(&slide.part_name)?.blob(),
            "notes-slide payload",
        )?;
        slides.push(Slide {
            slide_part_name: slide.slide_part_name.as_str().to_owned(),
            slide_relationship_id: slide.slide_relationship_id,
            part_name: slide.part_name.as_str().to_owned(),
            content_type: slide.content_type,
            data,
            backlink_relationship_id: slide.backlink_relationship_id,
            notes_master_relationship_id: slide.notes_master_relationship_id,
        });
    }
    Ok(Graph {
        conformance: index.conformance,
        master: Master {
            presentation_relationship_id: index.master.presentation_relationship_id,
            part_name: index.master.part_name.as_str().to_owned(),
            content_type: index.master.content_type,
            data: master_data,
            theme: Theme {
                relationship_id: index.master.theme.relationship_id,
                part_name: index.master.theme.part_name.as_str().to_owned(),
                content_type: index.master.theme.content_type,
                data: theme_data,
            },
        },
        slides,
    })
}

#[derive(Debug)]
struct Removal {
    slide_part_name: PackURI,
    relationship_id: String,
    notes_part_name: PackURI,
}

impl From<&SlideIndex> for Removal {
    fn from(slide: &SlideIndex) -> Self {
        Self {
            slide_part_name: slide.slide_part_name.clone(),
            relationship_id: slide.slide_relationship_id.clone(),
            notes_part_name: slide.part_name.clone(),
        }
    }
}

/// Remove the speaker-notes resource owned by one presentation slide.
///
/// The complete notes graph and every inbound edge to the selected resource
/// are validated before mutation. Missing notes are an idempotent `Ok(false)`.
/// Shared notes-master and theme resources are retained.
pub fn remove(
    package: &mut OpcPackage,
    presentation_name: &PackURI,
    slide_name: &PackURI,
) -> Result<bool> {
    let Some(index) = load_index(package, presentation_name)? else {
        return Ok(false);
    };
    let slide_name = package.get_part(slide_name)?.partname().clone();
    let Some(slide) = index
        .slides
        .iter()
        .find(|slide| slide.slide_part_name == slide_name)
    else {
        return Ok(false);
    };
    let removals = [Removal::from(slide)];
    validate_notes_removals(package, &removals)?;
    apply_notes_removals(package, &removals)?;
    Ok(true)
}

/// Remove every speaker-notes resource from a presentation.
///
/// Returns the number of removed notes slides. The operation is idempotent,
/// validates the complete graph before mutation, and retains the shared notes
/// master and its theme so ordinary presentation layout remains unchanged.
pub fn clear(package: &mut OpcPackage, presentation_name: &PackURI) -> Result<usize> {
    let Some(index) = load_index(package, presentation_name)? else {
        return Ok(0);
    };
    let mut removals = Vec::new();
    removals
        .try_reserve(index.slides.len())
        .map_err(|source| allocation("notes-removal plan", source))?;
    removals.extend(index.slides.iter().map(Removal::from));
    if removals.is_empty() {
        return Ok(0);
    }
    validate_notes_removals(package, &removals)?;
    apply_notes_removals(package, &removals)
}

fn validate_notes_removals(package: &OpcPackage, removals: &[Removal]) -> Result<()> {
    let mut by_target = HashMap::new();
    by_target
        .try_reserve(removals.len())
        .map_err(|source| allocation("notes-removal index", source))?;
    for (index, removal) in removals.iter().enumerate() {
        if by_target
            .insert(removal.notes_part_name.clone(), index)
            .is_some()
        {
            return Err(invalid("notes-removal plan contains a duplicate target"));
        }
    }
    let mut inbound_counts = Vec::new();
    inbound_counts
        .try_reserve(removals.len())
        .map_err(|source| allocation("notes-removal counters", source))?;
    inbound_counts.resize(removals.len(), 0usize);

    for relationship in package.rels().iter() {
        validate_notes_inbound(
            package,
            None,
            relationship,
            removals,
            &by_target,
            &mut inbound_counts,
        )?;
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            validate_notes_inbound(
                package,
                Some(source.partname()),
                relationship,
                removals,
                &by_target,
                &mut inbound_counts,
            )?;
        }
    }
    if inbound_counts.iter().any(|count| *count != 1) {
        return Err(invalid(
            "notes-removal target does not have exactly one owning slide relationship",
        ));
    }
    Ok(())
}

fn validate_notes_inbound(
    package: &OpcPackage,
    source: Option<&PackURI>,
    relationship: &litchi_opc::Relationship,
    removals: &[Removal],
    by_target: &HashMap<PackURI, usize>,
    inbound_counts: &mut [usize],
) -> Result<()> {
    if relationship.is_external() {
        return Ok(());
    }
    let Ok(target) = relationship.target_partname() else {
        return Ok(());
    };
    let index = by_target.get(&target).copied().or_else(|| {
        package
            .get_part(&target)
            .ok()
            .and_then(|part| by_target.get(part.partname()).copied())
    });
    let Some(index) = index else {
        return Ok(());
    };
    let removal = &removals[index];
    if source != Some(&removal.slide_part_name) || relationship.r_id() != removal.relationship_id {
        let source = source.map_or("package root", PackURI::as_str);
        return Err(invalid(format!(
            "notes slide '{}' has an unexpected inbound relationship '{}' from '{}'",
            removal.notes_part_name.as_str(),
            relationship.r_id(),
            source,
        )));
    }
    inbound_counts[index] = inbound_counts[index]
        .checked_add(1)
        .ok_or_else(|| invalid("notes inbound relationship count overflow"))?;
    Ok(())
}

fn apply_notes_removals(package: &mut OpcPackage, removals: &[Removal]) -> Result<usize> {
    // Stage cloned slide owners before the first package mutation. Built-in
    // parts retain their shared payload allocation while relationships are
    // detached on the staged clone.
    let mut staged_slides = Vec::new();
    staged_slides
        .try_reserve(removals.len())
        .map_err(|source| allocation("notes-removal staging", source))?;
    for removal in removals {
        let slide = package.get_part(&removal.slide_part_name)?;
        let relationship = slide
            .rels()
            .get(&removal.relationship_id)
            .ok_or_else(|| invalid("validated notes relationship disappeared before commit"))?;
        if relationship.is_external() {
            return Err(invalid(
                "validated notes relationship changed before commit",
            ));
        }
        let target = relationship.target_partname()?;
        if package.get_part(&target)?.partname() != &removal.notes_part_name {
            return Err(invalid(
                "validated notes relationship changed before commit",
            ));
        }
        package.get_part(&removal.notes_part_name)?;
        let mut staged = slide.clone_part();
        if staged.rels_mut().remove(&removal.relationship_id).is_none() {
            return Err(invalid("validated notes relationship was not removed"));
        }
        staged_slides.push(staged);
    }

    // Every operation below is infallible after validation and staging. Exact
    // stored part names avoid the case-insensitive lookup/exact-removal trap.
    for staged in staged_slides {
        package.add_part(staged);
    }
    for removal in removals {
        package.remove_part(&removal.notes_part_name);
    }
    package.unsign();
    Ok(removals.len())
}

/// Deterministically replace the resources of an already coherent notes graph.
/// Validation completes before the first package mutation.
pub fn put(package: &mut OpcPackage, presentation_name: &PackURI, graph: Graph) -> Result<()> {
    let current = load_index(package, presentation_name)?
        .ok_or_else(|| invalid("store requires an existing coherent notes graph"))?;
    validate_graph(&graph)?;
    if indexed_ownership(&current) != ownership(&graph) {
        return Err(invalid(
            "store cannot retarget or orphan existing notes parts",
        ));
    }
    if graph_matches(package, &current, &graph)? {
        return Ok(());
    }
    let presentation = package.get_part(presentation_name)?;
    let presentation_scan = scan_xml(
        presentation.blob(),
        MAX_PRESENTATION_XML,
        graph.conformance,
        "presentation",
    )?;
    if presentation_scan.notes_master_ids.as_slice()
        != [graph.master.presentation_relationship_id.as_str()]
    {
        return Err(invalid(
            "presentation notesMasterId does not match graph metadata",
        ));
    }
    let mut slide_names = BTreeSet::new();
    for id in presentation_scan.slide_ids {
        let relationship = presentation
            .rels()
            .get(&id)
            .ok_or_else(|| invalid("presentation slide reference is missing"))?;
        let target = relationship_target(presentation, relationship)?;
        slide_names.insert(package.get_part(&target)?.partname().as_str().to_owned());
    }
    if graph
        .slides
        .iter()
        .any(|slide| !slide_names.contains(&slide.slide_part_name))
    {
        return Err(invalid(
            "notes graph references a slide outside the presentation",
        ));
    }

    let Graph {
        conformance,
        master,
        slides,
    } = graph;
    let Master {
        presentation_relationship_id,
        part_name: master_name,
        content_type: master_content_type,
        data: master_data,
        theme,
    } = master;
    let Theme {
        relationship_id: theme_relationship_id,
        part_name: theme_name,
        content_type: theme_content_type,
        data: theme_data,
    } = theme;
    let theme_uri = PackURI::new(theme_name).map_err(Error::Invalid)?;
    let master_uri = PackURI::new(master_name).map_err(Error::Invalid)?;

    let theme_part = BlobPart::new(theme_uri.clone(), theme_content_type, theme_data);
    let mut master_part = BlobPart::new(master_uri.clone(), master_content_type, master_data);
    master_part.rels_mut().try_add_relationship(
        conformance.theme_rel().into(),
        theme_uri.relative_ref(master_uri.base_uri()),
        theme_relationship_id,
        TargetMode::Internal,
    )?;

    let mut note_parts = Vec::new();
    note_parts
        .try_reserve(slides.len())
        .map_err(|source| allocation("notes-part staging", source))?;
    let mut by_slide = BTreeMap::new();
    for slide in slides {
        let Slide {
            slide_part_name,
            slide_relationship_id,
            part_name,
            content_type,
            data,
            backlink_relationship_id,
            notes_master_relationship_id,
        } = slide;
        let notes_uri = PackURI::new(part_name).map_err(Error::Invalid)?;
        let slide_uri = PackURI::new(&slide_part_name).map_err(Error::Invalid)?;
        let mut notes_part = BlobPart::new(notes_uri.clone(), content_type, data);
        notes_part.rels_mut().try_add_relationship(
            conformance.slide_rel().into(),
            slide_uri.relative_ref(notes_uri.base_uri()),
            backlink_relationship_id,
            TargetMode::Internal,
        )?;
        notes_part.rels_mut().try_add_relationship(
            conformance.notes_master_rel().into(),
            master_uri.relative_ref(notes_uri.base_uri()),
            notes_master_relationship_id,
            TargetMode::Internal,
        )?;
        if by_slide
            .insert(slide_part_name, (notes_uri, slide_relationship_id))
            .is_some()
        {
            return Err(invalid("notes graph contains duplicate slide owners"));
        }
        note_parts.push(notes_part);
    }

    let mut staged_presentation = package.get_part(presentation_name)?.clone_part();
    let presentation_ids: Vec<_> = staged_presentation
        .rels()
        .iter()
        .filter(|relationship| is_notes_master_rel(relationship.reltype()))
        .map(|relationship| relationship.r_id().to_owned())
        .collect();
    for id in presentation_ids {
        staged_presentation.rels_mut().remove(&id);
    }
    staged_presentation.rels_mut().try_add_relationship(
        conformance.notes_master_rel().into(),
        master_uri.relative_ref(presentation_name.base_uri()),
        presentation_relationship_id,
        TargetMode::Internal,
    )?;

    let mut staged_slides = Vec::new();
    staged_slides
        .try_reserve(slide_names.len())
        .map_err(|source| allocation("notes slide-owner staging", source))?;
    for slide_name in slide_names {
        let uri = PackURI::new(&slide_name).map_err(Error::Invalid)?;
        let mut part = package.get_part(&uri)?.clone_part();
        let ids: Vec<_> = part
            .rels()
            .iter()
            .filter(|relationship| is_notes_slide_rel(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_owned())
            .collect();
        for id in ids {
            part.rels_mut().remove(&id);
        }
        if let Some((notes_uri, relationship_id)) = by_slide.remove(&slide_name) {
            part.rels_mut().try_add_relationship(
                conformance.notes_slide_rel().into(),
                notes_uri.relative_ref(uri.base_uri()),
                relationship_id,
                TargetMode::Internal,
            )?;
        }
        staged_slides.push(part);
    }
    if !by_slide.is_empty() {
        return Err(invalid("notes graph contains an unknown slide owner"));
    }

    // Commit is infallible after all URI, graph, allocation, and relationship
    // checks succeed. Owned payload buffers move into their canonical parts.
    package.add_part(Box::new(theme_part));
    package.add_part(Box::new(master_part));
    for part in note_parts {
        package.add_part(Box::new(part));
    }
    package.add_part(staged_presentation);
    for slide in staged_slides {
        package.add_part(slide);
    }
    package.unsign();
    Ok(())
}

/// Load the validated notes resource owned by one physical slide part.
///
/// This is a focused package-layer operation. Semantic slide selection belongs
/// to the higher-level PPTX facade.
pub fn slide(package: &OpcPackage, slide_name: &PackURI) -> Result<Option<Slide>> {
    let presentation_name = package.main_document_part()?.partname().clone();
    let Some(index) = load_index(package, &presentation_name)? else {
        return Ok(None);
    };
    let slide_name = package.get_part(slide_name)?.partname();
    let Some(slide) = index
        .slides
        .into_iter()
        .find(|slide| &slide.slide_part_name == slide_name)
    else {
        return Ok(None);
    };
    let data = own_blob(
        package.get_part(&slide.part_name)?.blob(),
        "notes-slide payload",
    )?;
    Ok(Some(Slide {
        slide_part_name: slide.slide_part_name.as_str().to_owned(),
        slide_relationship_id: slide.slide_relationship_id,
        part_name: slide.part_name.as_str().to_owned(),
        content_type: slide.content_type,
        data,
        backlink_relationship_id: slide.backlink_relationship_id,
        notes_master_relationship_id: slide.notes_master_relationship_id,
    }))
}

fn validate_graph(graph: &Graph) -> Result<()> {
    if graph.slides.len() > MAX_NOTES_SLIDES {
        return Err(limit("notes-slide count", MAX_NOTES_SLIDES));
    }
    validate_id(&graph.master.presentation_relationship_id)?;
    validate_id(&graph.master.theme.relationship_id)?;
    if graph.master.content_type != ct::PML_NOTES_MASTER
        || graph.master.theme.content_type != THEME_CT
    {
        return Err(invalid("notes master or theme has invalid content type"));
    }
    let master_uri = PackURI::new(&graph.master.part_name).map_err(Error::Invalid)?;
    validate_leaf_path(&master_uri, "/ppt/notesMasters/", "notes master")?;
    let theme_uri = PackURI::new(&graph.master.theme.part_name).map_err(Error::Invalid)?;
    validate_leaf_path(&theme_uri, "/ppt/theme/", "notes-master theme")?;
    validate_resource_xml(
        &graph.master.data,
        MAX_MASTER_XML,
        graph.conformance,
        "notesMaster",
        "notes master",
    )?;
    validate_resource_xml(
        &graph.master.theme.data,
        MAX_THEME_XML,
        graph.conformance,
        "theme",
        "notes-master theme",
    )?;
    let mut total = checked_add(
        graph.master.data.len(),
        graph.master.theme.data.len(),
        "aggregate bytes",
    )?;
    let mut sources = HashSet::new();
    let mut parts = HashSet::new();
    parts.insert(graph.master.part_name.as_str());
    parts.insert(graph.master.theme.part_name.as_str());
    for slide in &graph.slides {
        validate_id(&slide.slide_relationship_id)?;
        validate_id(&slide.backlink_relationship_id)?;
        validate_id(&slide.notes_master_relationship_id)?;
        if slide.backlink_relationship_id == slide.notes_master_relationship_id {
            return Err(invalid("notes-slide relationship IDs collide"));
        }
        if slide.content_type != ct::PML_NOTES_SLIDE {
            return Err(invalid("notes slide has invalid content type"));
        }
        let source = PackURI::new(&slide.slide_part_name).map_err(Error::Invalid)?;
        validate_leaf_path(&source, "/ppt/slides/", "slide")?;
        let uri = PackURI::new(&slide.part_name).map_err(Error::Invalid)?;
        validate_leaf_path(&uri, "/ppt/notesSlides/", "notes slide")?;
        if !sources.insert(slide.slide_part_name.as_str())
            || !parts.insert(slide.part_name.as_str())
        {
            return Err(invalid(
                "notes graph has duplicate source or resource part names",
            ));
        }
        validate_resource_xml(
            &slide.data,
            MAX_NOTES_XML,
            graph.conformance,
            "notes",
            "notes slide",
        )?;
        total = checked_add(total, slide.data.len(), "aggregate bytes")?;
        if total > MAX_TOTAL_BYTES {
            return Err(limit("notes aggregate bytes", MAX_TOTAL_BYTES));
        }
    }
    Ok(())
}

fn validate_resource_xml(
    xml: &[u8],
    max: usize,
    conformance: Conformance,
    root: &str,
    label: &str,
) -> Result<()> {
    let scan = scan_xml(xml, max, conformance, root)?;
    if !scan.relationship_attributes.is_empty() {
        return Err(invalid(format!(
            "{label} contains unsupported outbound relationship references"
        )));
    }
    Ok(())
}

fn root_conformance(xml: &[u8], max: usize, root: &str) -> Result<Conformance> {
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        if scan_xml(xml, max, conformance, root).is_ok() {
            return Ok(conformance);
        }
    }
    Err(invalid(format!("invalid {root} root or namespace")))
}

fn scan_xml(
    xml: &[u8],
    max: usize,
    conformance: Conformance,
    expected_root: &str,
) -> Result<XmlScan> {
    if xml.len() > max {
        return Err(limit("notes XML bytes", max));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > max {
        return Err(limit("processed notes XML bytes", max));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut attributes = 0usize;
    let mut attribute_bytes = 0usize;
    let mut root_seen = false;
    let mut scan = XmlScan::default();
    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(element) => {
                nodes += 1;
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(limit("notes XML depth", MAX_DEPTH));
                }
                if nodes > MAX_NODES {
                    return Err(limit("notes XML nodes", MAX_NODES));
                }
                inspect_element(
                    &reader,
                    &element,
                    conformance,
                    expected_root,
                    !root_seen,
                    &mut attributes,
                    &mut attribute_bytes,
                    &mut scan,
                )?;
                root_seen = true;
            },
            Event::Empty(element) => {
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(limit("notes XML nodes", MAX_NODES));
                }
                if depth >= MAX_DEPTH {
                    return Err(limit("notes XML depth", MAX_DEPTH));
                }
                inspect_element(
                    &reader,
                    &element,
                    conformance,
                    expected_root,
                    !root_seen,
                    &mut attributes,
                    &mut attribute_bytes,
                    &mut scan,
                )?;
                root_seen = true;
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("unexpected XML closing element"));
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected")),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root_seen || depth != 0 {
        return Err(invalid("missing or unterminated XML root"));
    }
    Ok(scan)
}

#[allow(clippy::too_many_arguments)]
fn inspect_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    conformance: Conformance,
    expected_root: &str,
    is_root: bool,
    attributes: &mut usize,
    attribute_bytes: &mut usize,
    scan: &mut XmlScan,
) -> Result<()> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let local = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    if is_root
        && (namespace
            != if expected_root == "theme" {
                conformance.a()
            } else {
                conformance.p()
            }
            || local != expected_root)
    {
        return Err(invalid(format!(
            "invalid {expected_root} root or namespace"
        )));
    }
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw = item.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        *attributes += 1;
        if *attributes > MAX_ATTRIBUTES {
            return Err(limit("notes XML attributes", MAX_ATTRIBUTES));
        }
        let (namespace, attr_local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let attr_local = std::str::from_utf8(attr_local.as_ref()).map_err(xml_error)?;
        let raw_value = std::str::from_utf8(item.value.as_ref()).map_err(xml_error)?;
        let value = quick_xml::escape::unescape(raw_value)
            .map_err(xml_error)?
            .into_owned();
        *attribute_bytes = attribute_bytes
            .checked_add(namespace.len() + attr_local.len() + value.len())
            .ok_or_else(|| invalid("notes XML attribute byte count overflow"))?;
        if *attribute_bytes > MAX_ATTRIBUTE_BYTES {
            return Err(limit("notes XML attribute bytes", MAX_ATTRIBUTE_BYTES));
        }
        if namespace == conformance.r() {
            scan.relationship_attributes.push(value.clone());
            if attr_local == "id" && namespace == conformance.r() {
                if local == "notesMasterId" {
                    scan.notes_master_ids.push(value.clone());
                } else if local == "sldId" {
                    scan.slide_ids.push(value.clone());
                }
            }
        }
    }
    Ok(())
}

fn ownership(graph: &Graph) -> BTreeSet<&str> {
    std::iter::once(graph.master.part_name.as_str())
        .chain(std::iter::once(graph.master.theme.part_name.as_str()))
        .chain(graph.slides.iter().map(|slide| slide.part_name.as_str()))
        .collect()
}

fn indexed_ownership(graph: &GraphIndex) -> BTreeSet<&str> {
    std::iter::once(graph.master.part_name.as_str())
        .chain(std::iter::once(graph.master.theme.part_name.as_str()))
        .chain(graph.slides.iter().map(|slide| slide.part_name.as_str()))
        .collect()
}

fn graph_matches(package: &OpcPackage, index: &GraphIndex, graph: &Graph) -> Result<bool> {
    if index.conformance != graph.conformance
        || index.master.presentation_relationship_id != graph.master.presentation_relationship_id
        || index.master.part_name.as_str() != graph.master.part_name
        || index.master.content_type != graph.master.content_type
        || index.master.theme.relationship_id != graph.master.theme.relationship_id
        || index.master.theme.part_name.as_str() != graph.master.theme.part_name
        || index.master.theme.content_type != graph.master.theme.content_type
        || index.slides.len() != graph.slides.len()
        || package.get_part(&index.master.part_name)?.blob() != graph.master.data
        || package.get_part(&index.master.theme.part_name)?.blob() != graph.master.theme.data
    {
        return Ok(false);
    }
    for (stored, candidate) in index.slides.iter().zip(&graph.slides) {
        if stored.slide_part_name.as_str() != candidate.slide_part_name
            || stored.slide_relationship_id != candidate.slide_relationship_id
            || stored.part_name.as_str() != candidate.part_name
            || stored.content_type != candidate.content_type
            || stored.backlink_relationship_id != candidate.backlink_relationship_id
            || stored.notes_master_relationship_id != candidate.notes_master_relationship_id
            || package.get_part(&stored.part_name)?.blob() != candidate.data
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn relationship_target(
    part: &dyn Part,
    relationship: &litchi_opc::Relationship,
) -> Result<PackURI> {
    if relationship.is_external() {
        return Err(invalid("external relationship is rejected"));
    }
    PackURI::from_rel_ref(part.partname().base_uri(), relationship.target_ref())
        .map_err(Error::Invalid)
}
fn validate_leaf_path(uri: &PackURI, prefix: &str, label: &str) -> Result<()> {
    let Some(rest) = uri.as_str().strip_prefix(prefix) else {
        return Err(invalid(format!("{label} is outside {prefix}")));
    };
    if rest.is_empty() || rest.contains('/') || !rest.ends_with(".xml") {
        return Err(invalid(format!("invalid {label} part path")));
    }
    Ok(())
}
fn require_content_type(part: &dyn Part, expected: &str, label: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(invalid(format!("{label} has invalid content type")))
    }
}
fn is_notes_slide_rel(value: &str) -> bool {
    matches!(value, rt::NOTES_SLIDE | rt::STRICT_NOTES_SLIDE)
}
fn is_notes_master_rel(value: &str) -> bool {
    matches!(value, rt::NOTES_MASTER | rt::STRICT_NOTES_MASTER)
}
fn is_presentation_main_content_type(value: &str) -> bool {
    matches!(
        value,
        ct::PML_PRESENTATION_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    )
}
fn validate_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID is empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        Err(invalid("invalid relationship ID"))
    } else {
        Ok(())
    }
}
fn checked_add(left: usize, right: usize, label: &str) -> Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| invalid(format!("PPTX notes {label} overflow")))
}
fn own_blob(blob: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(blob.len())
        .map_err(|source| allocation(resource, source))?;
    owned.extend_from_slice(blob);
    Ok(owned)
}
fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}
fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}
fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

#[cfg(test)]
mod tests {
    use super::*;
    const POI: &[u8] = include_bytes!("../../../test-data/poi/test-data/slideshow/prProps.pptx");
    const LO: &[u8] =
        include_bytes!("../../../test-data/libreoffice-core/oox/qa/unit/data/tdf131082.pptx");
    fn presentation() -> PackURI {
        PackURI::new("/ppt/presentation.xml").unwrap()
    }

    #[test]
    fn plain_text_writer_escapes_and_rejects_invalid_xml() {
        let xml = write_text("A < B & C").unwrap();
        assert_eq!(
            xml,
            write_text_with(Conformance::Transitional, "A < B & C").unwrap()
        );
        let xml = std::str::from_utf8(&xml).unwrap();
        assert!(xml.starts_with("<?xml version="));
        assert!(xml.contains("<a:t>A &lt; B &amp; C</a:t>"));
        assert!(xml.ends_with("</p:notes>"));
        assert!(write_text("bad\u{0}text").is_err());
    }

    #[test]
    fn notes_master_template_is_canonical_and_deterministic() {
        assert!(master_xml().contains("<p:notesMaster"));
        assert_eq!(master_xml().as_ptr(), master_xml().as_ptr());
    }

    #[test]
    fn consuming_put_moves_changed_xml_and_preserves_signed_no_ops() {
        let (mut package, name) = synthetic(Conformance::Transitional);
        let graph = load(&package, &name).unwrap().unwrap();
        package.relate_to(
            "_xmlsignatures/origin.sigs",
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        assert!(package.is_signed());

        put(&mut package, &name, graph).unwrap();
        assert!(package.is_signed());

        let mut graph = load(&package, &name).unwrap().unwrap();
        graph.slides_mut()[0].replace_xml(write_text("Updated note").unwrap());
        put(&mut package, &name, graph).unwrap();
        assert!(!package.is_signed());
        assert_eq!(
            load(&package, &name).unwrap().unwrap().slides()[0]
                .text()
                .unwrap()
                .as_deref(),
            Some("Updated note")
        );
    }
    #[test]
    fn poi_and_libreoffice_notes_graphs_load_and_store_deterministically() {
        for bytes in [POI, LO] {
            let mut package = OpcPackage::from_bytes(bytes).unwrap();
            let name = presentation();
            let graph = load(&package, &name).unwrap().unwrap();
            assert_eq!(graph.slides.len(), 1);
            assert_eq!(graph.master.content_type, ct::PML_NOTES_MASTER);
            put(&mut package, &name, graph).unwrap();
            let graph = load(&package, &name).unwrap().unwrap();
            assert_eq!(graph.slides.len(), 1);
            put(&mut package, &name, graph).unwrap();
            assert_eq!(load(&package, &name).unwrap().unwrap().slides.len(), 1);
        }
    }
    fn synthetic(conformance: Conformance) -> (OpcPackage, PackURI) {
        let p = conformance.p();
        let a = conformance.a();
        let r = conformance.r();
        let mut package = OpcPackage::new();
        let presentation = presentation();
        let mut pres=BlobPart::new(presentation.clone(),ct::PML_PRESENTATION_MAIN.into(),format!("<p:presentation xmlns:p=\"{p}\" xmlns:r=\"{r}\"><p:notesMasterIdLst><p:notesMasterId r:id=\"rIdMaster\"/></p:notesMasterIdLst><p:sldIdLst><p:sldId id=\"256\" r:id=\"rIdSlide\"/></p:sldIdLst><p:notesSz cx=\"1\" cy=\"1\"/></p:presentation>").into_bytes());
        pres.rels_mut().add_relationship(
            conformance.notes_master_rel().into(),
            "notesMasters/notesMaster1.xml".into(),
            "rIdMaster".into(),
            false,
        );
        pres.rels_mut().add_relationship(
            conformance.slide_rel().into(),
            "slides/slide1.xml".into(),
            "rIdSlide".into(),
            false,
        );
        package.add_part(Box::new(pres));
        let slide_uri = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let mut slide = BlobPart::new(
            slide_uri,
            SLIDE_CT.into(),
            format!("<p:sld xmlns:p=\"{p}\"><p:cSld/></p:sld>").into_bytes(),
        );
        slide.rels_mut().add_relationship(
            conformance.notes_slide_rel().into(),
            "../notesSlides/notesSlide1.xml".into(),
            "rIdNotes".into(),
            false,
        );
        package.add_part(Box::new(slide));
        let master_uri = PackURI::new("/ppt/notesMasters/notesMaster1.xml").unwrap();
        let mut master=BlobPart::new(master_uri,ct::PML_NOTES_MASTER.into(),format!("<p:notesMaster xmlns:p=\"{p}\" xmlns:a=\"{a}\"><p:cSld/><p:clrMap/></p:notesMaster>").into_bytes());
        master.rels_mut().add_relationship(
            conformance.theme_rel().into(),
            "../theme/theme2.xml".into(),
            "rIdTheme".into(),
            false,
        );
        package.add_part(Box::new(master));
        let notes_uri = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
        let mut notes=BlobPart::new(notes_uri,ct::PML_NOTES_SLIDE.into(),format!("<p:notes xmlns:p=\"{p}\" xmlns:a=\"{a}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><mc:AlternateContent><mc:Choice Requires=\"u\"><u:active/></mc:Choice><mc:Fallback><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Strict note</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></mc:Fallback></mc:AlternateContent><p:clrMapOvr/></p:notes>").into_bytes());
        notes.rels_mut().add_relationship(
            conformance.slide_rel().into(),
            "../slides/slide1.xml".into(),
            "rIdBack".into(),
            false,
        );
        notes.rels_mut().add_relationship(
            conformance.notes_master_rel().into(),
            "../notesMasters/notesMaster1.xml".into(),
            "rIdMaster".into(),
            false,
        );
        package.add_part(Box::new(notes));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/theme/theme2.xml").unwrap(),
            THEME_CT.into(),
            format!("<a:theme xmlns:a=\"{a}\" name=\"Notes\"/>").into_bytes(),
        )));
        (package, presentation)
    }
    #[test]
    fn strict_mce_graph_round_trips_and_projects_text() {
        let (mut package, name) = synthetic(Conformance::Strict);
        let graph = load(&package, &name).unwrap().unwrap();
        assert_eq!(
            graph.slides[0].text().unwrap().as_deref(),
            Some("Strict note")
        );
        put(&mut package, &name, graph).unwrap();
        assert_eq!(
            load(&package, &name).unwrap().unwrap().slides[0]
                .text()
                .unwrap()
                .as_deref(),
            Some("Strict note")
        );
    }

    #[test]
    fn every_presentation_main_profile_accepts_the_same_notes_graph() {
        for content_type in [
            ct::PML_PRESENTATION_MAIN,
            ct::PML_SLIDESHOW_MAIN,
            ct::PML_TEMPLATE_MAIN,
            ct::PML_PRES_MACRO_MAIN,
            ct::PML_SLIDESHOW_MACRO_MAIN,
            ct::PML_TEMPLATE_MACRO_MAIN,
        ] {
            let (mut package, name) = synthetic(Conformance::Transitional);
            package
                .get_part_mut(&name)
                .and_then(|part| part.set_content_type(content_type.to_owned()))
                .unwrap();
            assert!(load(&package, &name).unwrap().is_some());
        }
    }

    #[test]
    fn strict_text_writer_validates_and_round_trips_replacement() {
        let (mut package, name) = synthetic(Conformance::Strict);
        let mut graph = load(&package, &name).unwrap().unwrap();
        let xml = write_text_with(graph.conformance(), "Updated strict note").unwrap();
        let encoded = std::str::from_utf8(&xml).unwrap();
        assert!(encoded.contains(PS));
        assert!(encoded.contains(AS));
        assert!(encoded.contains(RS));

        graph.slides_mut()[0].replace_xml(xml);
        put(&mut package, &name, graph).unwrap();

        let graph = load(&package, &name).unwrap().unwrap();
        assert_eq!(graph.conformance(), Conformance::Strict);
        assert_eq!(
            graph.slides()[0].text().unwrap().as_deref(),
            Some("Updated strict note")
        );
    }

    #[test]
    fn strict_and_transitional_notes_removal_is_idempotent() {
        for conformance in [Conformance::Transitional, Conformance::Strict] {
            let (mut package, name) = synthetic(conformance);
            let slide = PackURI::new("/ppt/slides/slide1.xml").unwrap();
            let notes = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
            let master = PackURI::new("/ppt/notesMasters/notesMaster1.xml").unwrap();
            let theme = PackURI::new("/ppt/theme/theme2.xml").unwrap();

            assert!(remove(&mut package, &name, &slide).unwrap());
            assert!(!package.contains_part(&notes));
            assert!(package.contains_part(&master));
            assert!(package.contains_part(&theme));
            assert!(
                package
                    .get_part(&slide)
                    .unwrap()
                    .rels()
                    .iter()
                    .all(|relationship| !is_notes_slide_rel(relationship.reltype()))
            );
            assert!(load(&package, &name).unwrap().unwrap().slides.is_empty());
            assert!(!remove(&mut package, &name, &slide).unwrap());
            assert_eq!(clear(&mut package, &name).unwrap(), 0);
        }
    }

    #[test]
    fn removal_uses_the_actual_stored_part_name_after_case_folded_lookup() {
        let (mut package, name) = synthetic(Conformance::Transitional);
        let canonical = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
        let mixed_case = PackURI::new("/PPT/NOTESSLIDES/NOTESSLIDE1.XML").unwrap();
        let data = package.get_part(&canonical).unwrap().blob().to_vec();
        assert!(package.remove_part(&canonical));
        let mut notes = BlobPart::new(mixed_case.clone(), ct::PML_NOTES_SLIDE.into(), data);
        notes.rels_mut().add_relationship(
            rt::SLIDE.into(),
            "../slides/slide1.xml".into(),
            "rIdBack".into(),
            false,
        );
        notes.rels_mut().add_relationship(
            rt::NOTES_MASTER.into(),
            "../notesMasters/notesMaster1.xml".into(),
            "rIdMaster".into(),
            false,
        );
        package.add_part(Box::new(notes));

        let slide = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        assert!(remove(&mut package, &name, &slide).unwrap());
        assert!(!package.contains_part(&mixed_case));
        assert!(load(&package, &name).unwrap().unwrap().slides.is_empty());
    }

    #[test]
    fn unexpected_inbound_edge_rejects_removal_before_mutation() {
        let (mut package, name) = synthetic(Conformance::Transitional);
        let slide = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let notes = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
        let observer_name = PackURI::new("/ppt/custom/observer.xml").unwrap();
        let mut observer = BlobPart::new(
            observer_name,
            "application/xml".into(),
            b"<observer/>".to_vec(),
        );
        observer.rels_mut().add_relationship(
            "urn:test:observes-notes".into(),
            "../notesSlides/notesSlide1.xml".into(),
            "rIdObserver".into(),
            false,
        );
        package.add_part(Box::new(observer));

        let before_parts = package.part_count();
        let before_relationships = package.get_part(&slide).unwrap().rels().len();
        let error = remove(&mut package, &name, &slide).unwrap_err();
        assert!(
            error
                .to_string()
                .contains("unexpected inbound relationship")
        );
        assert_eq!(package.part_count(), before_parts);
        assert_eq!(
            package.get_part(&slide).unwrap().rels().len(),
            before_relationships
        );
        assert!(package.contains_part(&notes));
    }

    #[test]
    fn malformed_graph_rejects_clear_before_mutation() {
        let (mut package, name) = synthetic(Conformance::Transitional);
        let slide = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let notes = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
        package
            .get_part_mut(&notes)
            .unwrap()
            .set_blob(format!("<p:wrong xmlns:p=\"{P}\"/>").into_bytes());

        let before_parts = package.part_count();
        let before_relationships = package.get_part(&slide).unwrap().rels().len();
        assert!(clear(&mut package, &name).is_err());
        assert_eq!(package.part_count(), before_parts);
        assert_eq!(
            package.get_part(&slide).unwrap().rels().len(),
            before_relationships
        );
        assert!(package.contains_part(&notes));
    }

    #[test]
    fn rejects_external_wrong_root_outbound_orphan_and_caps_before_mutation() {
        let (mut package, name) = synthetic(Conformance::Transitional);
        let notes = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
        {
            let part = package.get_part_mut(&notes).unwrap();
            part.rels_mut().remove("rIdBack");
            part.rels_mut().add_relationship(
                rt::SLIDE.into(),
                "https://example.invalid/slide".into(),
                "rIdBack".into(),
                true,
            );
        }
        assert!(load(&package, &name).is_err());
        let (mut package, name) = synthetic(Conformance::Transitional);
        package
            .get_part_mut(&PackURI::new("/ppt/notesMasters/notesMaster1.xml").unwrap())
            .unwrap()
            .set_blob(format!("<p:wrong xmlns:p=\"{P}\"/>").into_bytes());
        assert!(load(&package, &name).is_err());
        let (mut package, name) = synthetic(Conformance::Transitional);
        package
            .get_part_mut(&notes)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::IMAGE.into(),
                "../media/image1.png".into(),
                "rIdImage".into(),
                false,
            );
        assert!(load(&package, &name).is_err());
        let (mut package, name) = synthetic(Conformance::Transitional);
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/notesSlides/orphan.xml").unwrap(),
            ct::PML_NOTES_SLIDE.into(),
            format!("<p:notes xmlns:p=\"{P}\"><p:cSld/></p:notes>").into_bytes(),
        )));
        assert!(load(&package, &name).is_err());
        let (mut package, name) = synthetic(Conformance::Transitional);
        let mut graph = load(&package, &name).unwrap().unwrap();
        graph.slides[0].data = vec![b' '; MAX_NOTES_XML + 1];
        let before = package.get_part(&name).unwrap().blob().to_vec();
        assert!(put(&mut package, &name, graph).is_err());
        assert_eq!(package.get_part(&name).unwrap().blob(), before);
    }
}
