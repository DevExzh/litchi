//! Slide master and slide layout authoring for PowerPoint packages.
//!
//! This module closes the authoring gap for `p:sldMaster` and `p:sldLayout`
//! parts. Every operation keeps the presentation relationship graph
//! consistent with the XML reference lists:
//!
//! - `p:sldMasterIdLst` in the presentation part stays in sync with
//!   slide-master relationships and parts.
//! - `p:sldLayoutIdLst` in each master stays in sync with slide-layout
//!   relationships and parts.
//! - Each layout keeps exactly one internal relationship back to its owning
//!   master, as the read side requires.
//!
//! After every mutation the whole master/layout graph is re-validated with
//! the same rules the read side applies, so a package that passes
//! [`validate_master_layout_graph`] resolves cleanly through
//! `Presentation::slide_masters`, `SlideMaster::slide_layouts`, and
//! `SlideLayout::master`.

use crate::error::{OoxmlError, Result};
use crate::pptx::parts::{SlideLayoutPart, SlideMasterPart};
use litchi_core::xml::escape_xml;
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::part::{BlobPart, Part};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashSet;
use std::fmt::Write as FmtWrite;

const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_SLIDE_MASTER_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideMaster";
const STRICT_SLIDE_LAYOUT_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout";

/// ECMA-376 Part 1: slide master and slide layout IDs start at 2^31.
pub const MIN_MASTER_OR_LAYOUT_ID: u32 = 2_147_483_648;
/// Shape ID 1 is reserved for the group-shape root of every shape tree.
const FIRST_SHAPE_ID: u32 = 2;
/// Bounded-input ceiling for every part this module parses or patches.
const MAX_PART_XML_BYTES: usize = 8 * 1024 * 1024;
/// Bounded-input ceiling for XML node counts while scanning.
const MAX_SCAN_NODES: usize = 100_000;
/// Bounded-input ceiling for XML nesting depth while scanning.
const MAX_SCAN_DEPTH: usize = 128;
/// Bounded ceiling for authored layout names.
const MAX_NAME_CHARS: usize = 256;
/// Bounded ceiling for placeholder shapes authored in a single operation.
const MAX_PLACEHOLDERS_PER_OPERATION: usize = 64;
/// Indentation step between the nine paragraph levels, in EMUs.
const LEVEL_MARGIN_STEP_EMU: u32 = 457200;
/// Default body font size for generated text-style levels, in hundredths of a point.
const LEVEL_FONT_SIZE_HUNDREDTHS: u32 = 1800;

const XML_DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
const SP_TREE_HEADER: &str = "<p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>";
const COLOR_MAP: &str = "<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>";

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

// ============================================================================
// Typed enums
// ============================================================================

/// Slide layout type (`ST_SlideLayoutType`, ECMA-376 Part 1).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SlideLayoutKind {
    Title,
    Text,
    TwoColumnText,
    Table,
    TextAndChart,
    ChartAndText,
    Diagram,
    Chart,
    TextAndClipArt,
    ClipArtAndText,
    TitleOnly,
    Blank,
    TextAndObject,
    ObjectAndText,
    ObjectOnly,
    Object,
    TextAndMedia,
    MediaAndText,
    ObjectOverText,
    TextOverObject,
    TextAndTwoObjects,
    TwoObjectsAndText,
    TwoObjectsOverText,
    FourObjects,
    VerticalText,
    ClipArtAndVerticalText,
    VerticalTitleAndText,
    VerticalTitleAndTextOverChart,
    TwoObjects,
    ObjectAndTwoObjects,
    TwoObjectsAndObject,
    Custom,
    SectionHeader,
    TwoTextAndTwoObjects,
    ObjectText,
    PictureWithText,
}

impl SlideLayoutKind {
    /// The spec token written to the `type` attribute of `p:sldLayout`.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Title => "title",
            Self::Text => "tx",
            Self::TwoColumnText => "twoColTx",
            Self::Table => "tbl",
            Self::TextAndChart => "txAndChart",
            Self::ChartAndText => "chartAndTx",
            Self::Diagram => "dgm",
            Self::Chart => "chart",
            Self::TextAndClipArt => "txAndClipArt",
            Self::ClipArtAndText => "clipArtAndTx",
            Self::TitleOnly => "titleOnly",
            Self::Blank => "blank",
            Self::TextAndObject => "txAndObj",
            Self::ObjectAndText => "objAndTx",
            Self::ObjectOnly => "objOnly",
            Self::Object => "obj",
            Self::TextAndMedia => "txAndMedia",
            Self::MediaAndText => "mediaAndTx",
            Self::ObjectOverText => "objOverTx",
            Self::TextOverObject => "txOverObj",
            Self::TextAndTwoObjects => "txAndTwoObj",
            Self::TwoObjectsAndText => "twoObjAndTx",
            Self::TwoObjectsOverText => "twoObjOverTx",
            Self::FourObjects => "fourObj",
            Self::VerticalText => "vertTx",
            Self::ClipArtAndVerticalText => "clipArtAndVertTx",
            Self::VerticalTitleAndText => "vertTitleAndTx",
            Self::VerticalTitleAndTextOverChart => "vertTitleAndTxOverChart",
            Self::TwoObjects => "twoObj",
            Self::ObjectAndTwoObjects => "objAndTwoObj",
            Self::TwoObjectsAndObject => "twoObjAndObj",
            Self::Custom => "cust",
            Self::SectionHeader => "secHead",
            Self::TwoTextAndTwoObjects => "twoTxTwoObj",
            Self::ObjectText => "objTx",
            Self::PictureWithText => "picTx",
        }
    }
}

/// Placeholder type (`ST_PlaceholderType`, ECMA-376 Part 1).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PlaceholderKind {
    Title,
    Body,
    CenteredTitle,
    Subtitle,
    DateTime,
    SlideNumber,
    Footer,
    Header,
    Object,
    Chart,
    Table,
    ClipArt,
    Diagram,
    Media,
    SlideImage,
    Picture,
}

impl PlaceholderKind {
    /// The spec token written to the `type` attribute of `p:ph`.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Title => "title",
            Self::Body => "body",
            Self::CenteredTitle => "ctrTitle",
            Self::Subtitle => "subTitle",
            Self::DateTime => "dt",
            Self::SlideNumber => "sldNum",
            Self::Footer => "ftr",
            Self::Header => "hdr",
            Self::Object => "obj",
            Self::Chart => "chart",
            Self::Table => "tbl",
            Self::ClipArt => "clipArt",
            Self::Diagram => "dgm",
            Self::Media => "media",
            Self::SlideImage => "sldImg",
            Self::Picture => "pic",
        }
    }

    /// Human-readable label used to build default shape names.
    const fn label(self) -> &'static str {
        match self {
            Self::Title => "Title",
            Self::Body => "Body",
            Self::CenteredTitle => "Centered Title",
            Self::Subtitle => "Subtitle",
            Self::DateTime => "Date",
            Self::SlideNumber => "Slide Number",
            Self::Footer => "Footer",
            Self::Header => "Header",
            Self::Object => "Object",
            Self::Chart => "Chart",
            Self::Table => "Table",
            Self::ClipArt => "Clip Art",
            Self::Diagram => "Diagram",
            Self::Media => "Media",
            Self::SlideImage => "Slide Image",
            Self::Picture => "Picture",
        }
    }
}

/// A placeholder shape to author on a slide master or slide layout.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PlaceholderSpec {
    /// Placeholder type written to `p:ph/@type`.
    pub kind: PlaceholderKind,
    /// Placeholder index written to `p:ph/@idx`; omitted when `None`.
    pub index: Option<u32>,
    /// Shape name written to `p:cNvPr/@name`; a default is generated when `None`.
    pub name: Option<String>,
    /// Prompt text written as a single run into the placeholder text body.
    pub text: Option<String>,
}

impl PlaceholderSpec {
    /// Create a placeholder of the given kind with no index, name, or text.
    pub const fn new(kind: PlaceholderKind) -> Self {
        Self {
            kind,
            index: None,
            name: None,
            text: None,
        }
    }

    /// Set the placeholder index (`p:ph/@idx`).
    pub fn with_index(mut self, index: u32) -> Self {
        self.index = Some(index);
        self
    }

    /// Set the shape name.
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Set the prompt text of the placeholder.
    pub fn with_text(mut self, text: impl Into<String>) -> Self {
        self.text = Some(text.into());
        self
    }

    /// The index used for identity matching; ECMA defaults `idx` to zero.
    fn effective_index(&self) -> u32 {
        self.index.unwrap_or(0)
    }
}

/// Identity of a slide master created by [`add_slide_master`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuthoredSlideMaster {
    /// The `p:sldMasterId/@id` value (always ≥ [`MIN_MASTER_OR_LAYOUT_ID`]).
    pub master_id: u32,
    /// Relationship ID from the presentation part to the master part.
    pub relationship_id: String,
    /// Part name of the new master, e.g. `/ppt/slideMasters/slideMaster2.xml`.
    pub part_name: String,
}

/// Identity of a slide layout created by [`add_slide_layout`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuthoredSlideLayout {
    /// The `p:sldLayoutId/@id` value (always ≥ [`MIN_MASTER_OR_LAYOUT_ID`]).
    pub layout_id: u32,
    /// Relationship ID from the owning master part to the layout part.
    pub relationship_id: String,
    /// Part name of the new layout, e.g. `/ppt/slideLayouts/slideLayout12.xml`.
    pub part_name: String,
    /// Part name of the owning slide master.
    pub master_part_name: String,
}

// ============================================================================
// Authoring operations
// ============================================================================

/// Create a new slide master and reference it from the presentation part.
///
/// The master is written with a color map, an empty `p:sldLayoutIdLst`, and
/// default `p:txStyles` (title, body, and other styles with nine paragraph
/// levels each). It is related to an existing theme part when one exists,
/// otherwise a new theme part is generated. The presentation part gains a
/// slide-master relationship plus a `p:sldMasterId` entry whose ID is one
/// above the current maximum (starting at [`MIN_MASTER_OR_LAYOUT_ID`]).
pub fn add_slide_master(package: &mut OpcPackage) -> Result<AuthoredSlideMaster> {
    let presentation_name = package.main_document_part()?.partname().clone();
    require_presentation_part(package.get_part(&presentation_name)?.content_type())?;
    let presentation_xml = package.get_part(&presentation_name)?.blob().to_vec();
    let entries = parse_master_id_list(&presentation_xml)?;
    let master_id = allocate_id(entries.iter().map(|entry| entry.0))?;

    let master_index = next_part_index(package, "/ppt/slideMasters/slideMaster", ".xml")?;
    let master_uri = PackURI::new(format!("/ppt/slideMasters/slideMaster{master_index}.xml"))
        .map_err(|error| OoxmlError::InvalidUri(format!("slide master partname: {error}")))?;
    let master_dir = "/ppt/slideMasters/";

    let (theme_target, new_theme) = theme_target_for_new_master(package, master_dir)?;

    let presentation = package.get_part_mut(&presentation_name)?;
    let relationship_id = presentation.relate_to(
        &format!("slideMasters/slideMaster{master_index}.xml"),
        rt::SLIDE_MASTER,
    );
    let entry = format!(
        "<p:sldMasterId xmlns:p=\"{P_NS}\" xmlns:r=\"{R_NS}\" id=\"{master_id}\" r:id=\"{}\"/>",
        escape_xml(&relationship_id)
    );
    let patched = match insert_id_list_entry(
        &presentation_xml,
        "sldMasterIdLst",
        &entry,
        IdListAnchor::AfterRootStart,
    ) {
        Ok(patched) => patched,
        Err(error) => {
            presentation.rels_mut().remove(&relationship_id);
            return Err(error);
        },
    };
    presentation.set_blob(patched);

    if let Some((theme_uri, theme_xml)) = new_theme {
        package.add_part(Box::new(BlobPart::new(
            theme_uri,
            ct::OFC_THEME.to_string(),
            theme_xml.into_bytes(),
        )));
    }
    let mut master_part = BlobPart::new(
        master_uri.clone(),
        ct::PML_SLIDE_MASTER.to_string(),
        master_xml().into_bytes(),
    );
    master_part.relate_to(&theme_target, rt::THEME);
    package.add_part(Box::new(master_part));

    invalidate_signatures(package);
    validate_master_layout_graph(package)?;
    Ok(AuthoredSlideMaster {
        master_id,
        relationship_id,
        part_name: master_uri.to_string(),
    })
}

/// Create a new slide layout attached to an existing slide master.
///
/// The master gains a slide-layout relationship plus a `p:sldLayoutId` entry
/// whose ID is one above the current maximum within that master. The layout
/// part is written with the given `ST_SlideLayoutType` kind, name, and
/// optional placeholder shapes, and carries the required relationship back to
/// its owning master.
pub fn add_slide_layout(
    package: &mut OpcPackage,
    master_part_name: &str,
    kind: SlideLayoutKind,
    name: &str,
    placeholders: &[PlaceholderSpec],
) -> Result<AuthoredSlideLayout> {
    require_name(name)?;
    require_placeholders(placeholders)?;
    let master_uri = PackURI::new(master_part_name)
        .map_err(|error| OoxmlError::InvalidUri(format!("slide master partname: {error}")))?;
    let master_part = package.get_part(&master_uri)?;
    if master_part.content_type() != ct::PML_SLIDE_MASTER {
        return Err(OoxmlError::InvalidContentType {
            expected: ct::PML_SLIDE_MASTER.to_string(),
            got: master_part.content_type().to_string(),
        });
    }
    let references = SlideMasterPart::from_part(master_part)?.slide_layout_references()?;
    let layout_id = allocate_id(
        references
            .iter()
            .filter_map(|reference| reference.layout_id()),
    )?;
    let master_xml = master_part.blob().to_vec();

    let layout_index = next_part_index(package, "/ppt/slideLayouts/slideLayout", ".xml")?;
    let layout_uri = PackURI::new(format!("/ppt/slideLayouts/slideLayout{layout_index}.xml"))
        .map_err(|error| OoxmlError::InvalidUri(format!("slide layout partname: {error}")))?;
    let layout_xml = layout_xml(kind, name, placeholders)?;

    let master_part = package.get_part_mut(&master_uri)?;
    let relationship_id = master_part.relate_to(
        &format!("../slideLayouts/slideLayout{layout_index}.xml"),
        rt::SLIDE_LAYOUT,
    );
    let entry = format!(
        "<p:sldLayoutId xmlns:p=\"{P_NS}\" xmlns:r=\"{R_NS}\" id=\"{layout_id}\" r:id=\"{}\"/>",
        escape_xml(&relationship_id)
    );
    let patched = match insert_id_list_entry(
        &master_xml,
        "sldLayoutIdLst",
        &entry,
        IdListAnchor::AfterElement("clrMap"),
    ) {
        Ok(patched) => patched,
        Err(error) => {
            master_part.rels_mut().remove(&relationship_id);
            return Err(error);
        },
    };
    master_part.set_blob(patched);

    let mut layout_part = BlobPart::new(
        layout_uri.clone(),
        ct::PML_SLIDE_LAYOUT.to_string(),
        layout_xml.into_bytes(),
    );
    layout_part.relate_to(
        &relative_target("/ppt/slideLayouts/", master_part_name)?,
        rt::SLIDE_MASTER,
    );
    package.add_part(Box::new(layout_part));

    invalidate_signatures(package);
    validate_master_layout_graph(package)?;
    Ok(AuthoredSlideLayout {
        layout_id,
        relationship_id,
        part_name: layout_uri.to_string(),
        master_part_name: master_uri.to_string(),
    })
}

/// Add or replace a placeholder shape on a slide master or slide layout.
///
/// A placeholder is identified by its `p:ph` type and index (an absent index
/// matches `idx` zero, per the ECMA default). When a matching placeholder
/// already exists its shape is replaced in place, keeping its shape ID;
/// otherwise a new shape is appended to the shape tree with the next free
/// shape ID. The optional text replaces the placeholder's prompt text.
pub fn store_placeholder_shape(
    package: &mut OpcPackage,
    part_name: &str,
    spec: &PlaceholderSpec,
) -> Result<()> {
    let uri = PackURI::new(part_name)
        .map_err(|error| OoxmlError::InvalidUri(format!("placeholder owner partname: {error}")))?;
    let part = package.get_part(&uri)?;
    let content_type = part.content_type();
    if content_type != ct::PML_SLIDE_MASTER && content_type != ct::PML_SLIDE_LAYOUT {
        return Err(OoxmlError::InvalidContentType {
            expected: format!("{} or {}", ct::PML_SLIDE_MASTER, ct::PML_SLIDE_LAYOUT),
            got: content_type.to_string(),
        });
    }
    if let Some(name) = &spec.name {
        require_name(name)?;
    }
    let xml = part.blob().to_vec();
    let existing = find_placeholder_span(&xml, spec.kind.as_str(), spec.effective_index())?;
    let shape_id = match &existing {
        Some(span) => shape_id_within(&xml[span.start..span.end])?,
        None => next_shape_id(&xml)?,
    };
    let shape = placeholder_shape_xml(shape_id, spec, true);
    let patched = match existing {
        Some(span) => replace_span(&xml, &span, shape.as_bytes())?,
        None => {
            let tree = scan_element_span(&xml, "spTree", SPTREE_DEPTH)?
                .ok_or_else(|| invalid("slide master or layout has no shape tree"))?;
            if tree.empty {
                return Err(invalid("slide master or layout has an empty shape tree"));
            }
            insert_bytes(&xml, tree.close_start, shape.as_bytes())?
        },
    };
    // The patched part must inventory the placeholder back through the same
    // scan the read side's shape parser performs.
    if find_placeholder_span(&patched, spec.kind.as_str(), spec.effective_index())?.is_none() {
        return Err(invalid("patched placeholder shape did not round-trip"));
    }
    package.get_part_mut(&uri)?.set_blob(patched);

    // Run the read-side placeholder inventory over the patched part.
    let part = package.get_part(&uri)?;
    let placeholders = if part.content_type() == ct::PML_SLIDE_MASTER {
        SlideMasterPart::from_part(part)?.placeholders()?
    } else {
        SlideLayoutPart::from_part(part)?.placeholders()?
    };
    let found = placeholders.iter().any(|shape| {
        shape
            .placeholder_type()
            .is_ok_and(|kind| kind == spec.kind.as_str())
            && shape
                .placeholder_index()
                .is_ok_and(|index| index.unwrap_or(0) == spec.effective_index())
    });
    if !found {
        return Err(invalid(
            "read-side placeholder inventory lost the authored shape",
        ));
    }
    invalidate_signatures(package);
    Ok(())
}

/// Delete a slide layout that is not referenced by any slide.
///
/// The owning master's `p:sldLayoutIdLst` entry and relationship are removed
/// together with the layout part itself. Layouts still referenced by a slide,
/// or not owned by any master, are rejected.
pub fn remove_slide_layout(package: &mut OpcPackage, layout_part_name: &str) -> Result<()> {
    let layout_uri = PackURI::new(layout_part_name)
        .map_err(|error| OoxmlError::InvalidUri(format!("slide layout partname: {error}")))?;
    let layout_part = package.get_part(&layout_uri)?;
    if layout_part.content_type() != ct::PML_SLIDE_LAYOUT {
        return Err(OoxmlError::InvalidContentType {
            expected: ct::PML_SLIDE_LAYOUT.to_string(),
            got: layout_part.content_type().to_string(),
        });
    }

    // Reject layouts still referenced by a slide.
    for part in package.iter_parts() {
        if part.content_type() != ct::PML_SLIDE {
            continue;
        }
        for relationship in part.rels().iter() {
            if matches!(
                relationship.reltype(),
                rt::SLIDE_LAYOUT | STRICT_SLIDE_LAYOUT_REL
            ) && !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|target| target == layout_uri)
            {
                return Err(invalid(format!(
                    "slide layout '{layout_part_name}' is still referenced by slide '{}'",
                    part.partname()
                )));
            }
        }
    }

    // Locate the owning master entry.
    let mut owner = None;
    for part in package.iter_parts() {
        if part.content_type() != ct::PML_SLIDE_MASTER {
            continue;
        }
        for reference in SlideMasterPart::from_part(part)?.slide_layout_references()? {
            let Some(relationship) = part.rels().get(reference.relationship_id()) else {
                continue;
            };
            if !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|target| target == layout_uri)
            {
                if owner.is_some() {
                    return Err(invalid(format!(
                        "slide layout '{layout_part_name}' is owned by more than one master"
                    )));
                }
                owner = Some((
                    part.partname().clone(),
                    reference.relationship_id().to_string(),
                ));
            }
        }
    }
    let (master_uri, relationship_id) = owner.ok_or_else(|| {
        invalid(format!(
            "slide layout '{layout_part_name}' is not referenced by any slide master"
        ))
    })?;

    let master_xml = package.get_part(&master_uri)?.blob().to_vec();
    let patched = remove_id_list_entry(&master_xml, "sldLayoutId", &relationship_id)?;
    let master_part = package.get_part_mut(&master_uri)?;
    master_part.set_blob(patched);
    if master_part.rels_mut().remove(&relationship_id).is_none() {
        return Err(OoxmlError::InvalidRelationship(format!(
            "slide master lost slide-layout relationship '{relationship_id}'"
        )));
    }
    package.remove_part(&layout_uri);

    invalidate_signatures(package);
    validate_master_layout_graph(package)?;
    Ok(())
}

// ============================================================================
// Graph validation
// ============================================================================

/// Validate the slide master and slide layout graph of a package.
///
/// This mirrors the rules the read side applies when resolving
/// `Presentation::slide_masters`, `SlideMaster::slide_layouts`, and
/// `SlideLayout::master`:
///
/// - every `p:sldMasterId` entry has a unique ID ≥ [`MIN_MASTER_OR_LAYOUT_ID`]
///   and resolves through an internal slide-master relationship to a part
///   with the slide-master content type;
/// - every `p:sldLayoutId` entry of each master resolves through an internal
///   slide-layout relationship to a part with the slide-layout content type;
/// - every referenced layout has exactly one internal slide-master
///   relationship, pointing back to the master that references it.
pub fn validate_master_layout_graph(package: &OpcPackage) -> Result<()> {
    let presentation = package.main_document_part()?;
    require_presentation_part(presentation.content_type())?;
    let entries = parse_master_id_list(presentation.blob())?;
    let mut master_parts = HashSet::new();
    for (master_id, relationship_id) in &entries {
        let relationship = presentation.rels().get(relationship_id).ok_or_else(|| {
            OoxmlError::InvalidRelationship(format!(
                "slide master ID {master_id} references missing relationship '{relationship_id}'"
            ))
        })?;
        if relationship.is_external()
            || !matches!(
                relationship.reltype(),
                rt::SLIDE_MASTER | STRICT_SLIDE_MASTER_REL
            )
        {
            return Err(OoxmlError::InvalidRelationship(format!(
                "relationship '{relationship_id}' is not an internal slide-master relationship"
            )));
        }
        let target = relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidRelationship(format!(
                "invalid slide-master relationship '{relationship_id}': {error}"
            ))
        })?;
        let master_part = package.get_part(&target)?;
        if master_part.content_type() != ct::PML_SLIDE_MASTER {
            return Err(OoxmlError::InvalidContentType {
                expected: ct::PML_SLIDE_MASTER.to_string(),
                got: master_part.content_type().to_string(),
            });
        }
        if !master_parts.insert(target.to_string()) {
            return Err(invalid(format!(
                "slide master part '{target}' is referenced more than once"
            )));
        }
        validate_master_layouts(package, master_part)?;
    }
    Ok(())
}

fn validate_master_layouts(
    package: &OpcPackage,
    master_part: &dyn litchi_opc::part::Part,
) -> Result<()> {
    let references = SlideMasterPart::from_part(master_part)?.slide_layout_references()?;
    for reference in &references {
        let relationship_id = reference.relationship_id();
        let relationship = master_part.rels().get(relationship_id).ok_or_else(|| {
            OoxmlError::InvalidRelationship(format!(
                "slide master references missing slide-layout relationship '{relationship_id}'"
            ))
        })?;
        if relationship.is_external()
            || !matches!(
                relationship.reltype(),
                rt::SLIDE_LAYOUT | STRICT_SLIDE_LAYOUT_REL
            )
        {
            return Err(OoxmlError::InvalidRelationship(format!(
                "relationship '{relationship_id}' is not an internal slide-layout relationship"
            )));
        }
        let layout_name = relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidRelationship(format!(
                "invalid slide-layout relationship '{relationship_id}': {error}"
            ))
        })?;
        let layout_part = package.get_part(&layout_name)?;
        if layout_part.content_type() != ct::PML_SLIDE_LAYOUT {
            return Err(OoxmlError::InvalidContentType {
                expected: ct::PML_SLIDE_LAYOUT.to_string(),
                got: layout_part.content_type().to_string(),
            });
        }
        // The layout must keep exactly one internal relationship back to the
        // master that references it.
        let mut back_reference = None;
        for candidate in layout_part.rels().iter() {
            if matches!(
                candidate.reltype(),
                rt::SLIDE_MASTER | STRICT_SLIDE_MASTER_REL
            ) {
                if back_reference.is_some() {
                    return Err(OoxmlError::InvalidRelationship(format!(
                        "slide layout '{layout_name}' has multiple slide-master relationships"
                    )));
                }
                back_reference = Some(candidate);
            }
        }
        let back_reference = back_reference.ok_or_else(|| {
            OoxmlError::InvalidRelationship(format!(
                "slide layout '{layout_name}' has no slide-master relationship"
            ))
        })?;
        if back_reference.is_external()
            || back_reference
                .target_partname()
                .is_ok_and(|target| target != *master_part.partname())
        {
            return Err(OoxmlError::InvalidRelationship(format!(
                "slide layout '{layout_name}' does not reference its owning master '{}'",
                master_part.partname()
            )));
        }
    }
    Ok(())
}

// ============================================================================
// XML generation
// ============================================================================

/// Serialize a new slide master part with default text styles.
fn master_xml() -> String {
    let mut xml = String::with_capacity(8192);
    xml.push_str(XML_DECL);
    xml.push_str("<p:sldMaster xmlns:a=\"");
    xml.push_str(A_NS);
    xml.push_str("\" xmlns:r=\"");
    xml.push_str(R_NS);
    xml.push_str("\" xmlns:p=\"");
    xml.push_str(P_NS);
    xml.push_str("\"><p:cSld>");
    xml.push_str(SP_TREE_HEADER);
    xml.push_str("</p:spTree></p:cSld>");
    xml.push_str(COLOR_MAP);
    xml.push_str("<p:sldLayoutIdLst/>");
    xml.push_str("<p:txStyles><p:titleStyle>");
    push_text_style_levels(&mut xml);
    xml.push_str("</p:titleStyle><p:bodyStyle>");
    push_text_style_levels(&mut xml);
    xml.push_str("</p:bodyStyle><p:otherStyle>");
    push_text_style_levels(&mut xml);
    xml.push_str("</p:otherStyle></p:txStyles></p:sldMaster>");
    xml
}

/// Write the nine paragraph levels shared by all generated text styles.
fn push_text_style_levels(xml: &mut String) {
    for level in 1..=9u32 {
        let margin = (level - 1) * LEVEL_MARGIN_STEP_EMU;
        let _ = write!(
            xml,
            "<a:lvl{level}pPr marL=\"{margin}\" algn=\"l\" defTabSz=\"457200\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"{LEVEL_FONT_SIZE_HUNDREDTHS}\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl{level}pPr>"
        );
    }
}

/// Serialize a new slide layout part.
fn layout_xml(
    kind: SlideLayoutKind,
    name: &str,
    placeholders: &[PlaceholderSpec],
) -> Result<String> {
    let mut xml = String::with_capacity(2048);
    xml.push_str(XML_DECL);
    let _ = write!(
        xml,
        "<p:sldLayout xmlns:a=\"{A_NS}\" xmlns:r=\"{R_NS}\" xmlns:p=\"{P_NS}\" type=\"{}\" matchingName=\"{}\"><p:cSld name=\"{}\">",
        kind.as_str(),
        escape_xml(name),
        escape_xml(name)
    );
    xml.push_str(SP_TREE_HEADER);
    for (offset, spec) in placeholders.iter().enumerate() {
        let shape_id = FIRST_SHAPE_ID + offset as u32;
        xml.push_str(&placeholder_shape_xml(shape_id, spec, false));
    }
    xml.push_str(
        "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>",
    );
    if xml.len() > MAX_PART_XML_BYTES {
        return Err(invalid(
            "generated slide layout exceeds the part size limit",
        ));
    }
    Ok(xml)
}

/// Serialize one placeholder shape.
///
/// When `declare_namespaces` is set the shape carries its own `xmlns`
/// declarations so it can be patched into a part with unknown prefix
/// bindings.
fn placeholder_shape_xml(
    shape_id: u32,
    spec: &PlaceholderSpec,
    declare_namespaces: bool,
) -> String {
    let name = spec
        .name
        .clone()
        .unwrap_or_else(|| format!("{} Placeholder {shape_id}", spec.kind.label()));
    let mut xml = String::with_capacity(512);
    xml.push_str("<p:sp");
    if declare_namespaces {
        let _ = write!(xml, " xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\"");
    }
    let _ = write!(
        xml,
        "><p:nvSpPr><p:cNvPr id=\"{shape_id}\" name=\"{}\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"{}\"",
        escape_xml(&name),
        spec.kind.as_str()
    );
    if let Some(index) = spec.index {
        let _ = write!(xml, " idx=\"{index}\"");
    }
    xml.push_str("/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>");
    if let Some(text) = &spec.text {
        let _ = write!(xml, "<a:r><a:t>{}</a:t></a:r>", escape_xml(text));
    }
    xml.push_str("<a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp>");
    xml
}

// ============================================================================
// Bounded XML scanning and patching
// ============================================================================

/// Depth of `p:spTree` inside master and layout parts (root → cSld → spTree).
const SPTREE_DEPTH: usize = 3;
/// Depth of `p:sp` shapes inside the shape tree.
const SHAPE_DEPTH: usize = 4;

/// Byte span of an XML element.
#[derive(Debug, Clone, Copy)]
struct ElementSpan {
    /// Offset of the `<` that opens the element.
    start: usize,
    /// Offset one past the `>` that closes the element.
    end: usize,
    /// Offset of the `</` that opens the closing tag (equals `start` for empty elements).
    close_start: usize,
    /// Whether the element uses the self-closing form.
    empty: bool,
}

/// Where a missing ID list should be created.
enum IdListAnchor {
    /// `p:sldMasterIdLst` heads the `CT_Presentation` sequence.
    AfterRootStart,
    /// `p:sldLayoutIdLst` follows `p:clrMap` in the `CT_SlideMaster` sequence.
    AfterElement(&'static str),
}

fn check_size(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_PART_XML_BYTES {
        return Err(invalid("part XML exceeds 8 MiB"));
    }
    Ok(())
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

/// Find the first element with `target` as local name at exactly `depth`.
fn scan_element_span(xml: &[u8], target: &str, depth: usize) -> Result<Option<ElementSpan>> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut stack: Vec<(usize, String)> = Vec::new();
    let mut nodes = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES || stack.len() >= MAX_SCAN_DEPTH {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                let local =
                    String::from_utf8_lossy(local_name(element.name().as_ref())).into_owned();
                stack.push((before, local));
            },
            Ok(Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                if stack.len() + 1 == depth
                    && local_name(element.name().as_ref()) == target.as_bytes()
                {
                    return Ok(Some(ElementSpan {
                        start: before,
                        end: reader.buffer_position() as usize,
                        close_start: before,
                        empty: true,
                    }));
                }
            },
            Ok(Event::End(element)) => {
                let (start, local) = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element in part XML"))?;
                if stack.len() + 1 == depth && local == target {
                    return Ok(Some(ElementSpan {
                        start,
                        end: reader.buffer_position() as usize,
                        close_start: before,
                        empty: false,
                    }));
                }
                if local_name(element.name().as_ref()) != local.as_bytes() {
                    return Err(invalid("mismatched closing element in part XML"));
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated part XML"));
    }
    Ok(None)
}

/// Insert `entry` into the ID list element `list_local`, creating the list at
/// the schema-correct position when it is missing.
fn insert_id_list_entry(
    xml: &[u8],
    list_local: &str,
    entry: &str,
    anchor: IdListAnchor,
) -> Result<Vec<u8>> {
    if let Some(span) = scan_element_span(xml, list_local, 2)? {
        if span.empty {
            let wrapped = format!(
                "<p:{list_local} xmlns:p=\"{P_NS}\" xmlns:r=\"{R_NS}\">{entry}</p:{list_local}>"
            );
            return replace_span(xml, &span, wrapped.as_bytes());
        }
        return insert_bytes(xml, span.close_start, entry.as_bytes());
    }
    let wrapped =
        format!("<p:{list_local} xmlns:p=\"{P_NS}\" xmlns:r=\"{R_NS}\">{entry}</p:{list_local}>");
    let offset = match anchor {
        IdListAnchor::AfterRootStart => root_start_end(xml)?,
        IdListAnchor::AfterElement(anchor_local) => {
            let span = scan_element_span(xml, anchor_local, 2)?.ok_or_else(|| {
                invalid(format!("part XML is missing its '{anchor_local}' anchor"))
            })?;
            span.end
        },
    };
    insert_bytes(xml, offset, wrapped.as_bytes())
}

/// Offset one past the root element's start tag.
fn root_start_end(xml: &[u8]) -> Result<usize> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(_) | Event::Empty(_)) => {
                return Ok(reader.buffer_position() as usize);
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => return Err(invalid("part XML has no root element")),
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }
}

/// Remove the ID-list entry whose `r:id` matches `relationship_id`.
fn remove_id_list_entry(xml: &[u8], entry_local: &str, relationship_id: &str) -> Result<Vec<u8>> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut nodes = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                if local_name(element.name().as_ref()) == entry_local.as_bytes()
                    && element_relationship_id(&element)?.as_deref() == Some(relationship_id)
                {
                    let span = ElementSpan {
                        start: before,
                        end: reader.buffer_position() as usize,
                        close_start: before,
                        empty: true,
                    };
                    return replace_span(xml, &span, b"");
                }
            },
            Ok(Event::Start(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                if local_name(element.name().as_ref()) == entry_local.as_bytes()
                    && element_relationship_id(&element)?.as_deref() == Some(relationship_id)
                {
                    // Consume events up to the matching closing tag so entries
                    // with extension children are removed whole.
                    let mut depth = 1usize;
                    loop {
                        match reader.read_event() {
                            Ok(Event::Start(_)) => depth += 1,
                            Ok(Event::End(_)) => {
                                depth -= 1;
                                if depth == 0 {
                                    let span = ElementSpan {
                                        start: before,
                                        end: reader.buffer_position() as usize,
                                        close_start: before,
                                        empty: false,
                                    };
                                    return replace_span(xml, &span, b"");
                                }
                            },
                            Ok(Event::Eof) => {
                                return Err(invalid("unterminated ID-list entry"));
                            },
                            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
                            _ => {},
                        }
                    }
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }
    Err(invalid(format!(
        "ID list has no entry for relationship '{relationship_id}'"
    )))
}

/// Read the relationship-namespace `id` attribute of an element.
fn element_relationship_id(element: &quick_xml::events::BytesStart<'_>) -> Result<Option<String>> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if name.rsplit_once(':').map(|(_, local)| local) == Some("id") && name.contains(':') {
            let value = std::str::from_utf8(attribute.value.as_ref())
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            return Ok(Some(value.to_owned()));
        }
    }
    Ok(None)
}

/// Find the direct `p:sp` child of the shape tree whose `p:ph` matches
/// `kind` and `index`.
fn find_placeholder_span(xml: &[u8], kind: &str, index: u32) -> Result<Option<ElementSpan>> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut shape_start = None;
    let mut nodes = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes += 1;
                depth += 1;
                if nodes > MAX_SCAN_NODES || depth > MAX_SCAN_DEPTH {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                let local = local_name(element.name().as_ref()).to_vec();
                if depth == SHAPE_DEPTH && local == b"sp" {
                    shape_start = Some(before);
                } else if local == b"ph"
                    && shape_start.is_some()
                    && placeholder_matches(&element, kind, index)?
                {
                    let start = shape_start.ok_or_else(|| invalid("missing placeholder shape"))?;
                    return Ok(Some(ElementSpan {
                        start,
                        end: shape_end(xml, start)?,
                        close_start: start,
                        empty: false,
                    }));
                }
            },
            Ok(Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                if local_name(element.name().as_ref()) == b"ph"
                    && shape_start.is_some()
                    && placeholder_matches(&element, kind, index)?
                {
                    let start = shape_start.ok_or_else(|| invalid("missing placeholder shape"))?;
                    return Ok(Some(ElementSpan {
                        start,
                        end: shape_end(xml, start)?,
                        close_start: start,
                        empty: false,
                    }));
                }
            },
            Ok(Event::End(_)) => {
                if depth == SHAPE_DEPTH {
                    shape_start = None;
                }
                if depth == 0 {
                    return Err(invalid("unexpected closing element in part XML"));
                }
                depth -= 1;
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }
    if depth != 0 {
        return Err(invalid("unterminated part XML"));
    }
    Ok(None)
}

/// Whether a `p:ph` element matches the requested type and index.
fn placeholder_matches(
    element: &quick_xml::events::BytesStart<'_>,
    kind: &str,
    index: u32,
) -> Result<bool> {
    let mut ph_type = None;
    let mut ph_index = 0u32;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        match attribute.key.as_ref() {
            b"type" => {
                ph_type = Some(
                    std::str::from_utf8(attribute.value.as_ref())
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?
                        .to_owned(),
                );
            },
            b"idx" => {
                let value = std::str::from_utf8(attribute.value.as_ref())
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                ph_index = value
                    .parse::<u32>()
                    .map_err(|_| invalid(format!("invalid placeholder index '{value}'")))?;
            },
            _ => {},
        }
    }
    // ECMA defaults: type "obj", idx 0.
    Ok(ph_type.as_deref().unwrap_or("obj") == kind && ph_index == index)
}

/// Compute the end offset of the `p:sp` element starting at `start`.
fn shape_end(xml: &[u8], start: usize) -> Result<usize> {
    let mut reader = Reader::from_reader(&xml[start..]);
    let mut depth = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(_)) => depth += 1,
            Ok(Event::Empty(_)) if depth == 0 => {
                return Ok(start + reader.buffer_position() as usize);
            },
            Ok(Event::End(_)) => {
                depth -= 1;
                if depth == 0 {
                    return Ok(start + reader.buffer_position() as usize);
                }
            },
            Ok(Event::Eof) => return Err(invalid("unterminated placeholder shape")),
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }
}

/// Extract the `p:cNvPr/@id` shape ID from a shape byte range.
fn shape_id_within(bytes: &[u8]) -> Result<u32> {
    let mut reader = Reader::from_reader(bytes);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if local_name(element.name().as_ref()) == b"cNvPr" =>
            {
                for attribute in element.attributes().with_checks(true) {
                    let attribute =
                        attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
                    if attribute.key.as_ref() == b"id" {
                        let value = std::str::from_utf8(attribute.value.as_ref())
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        return value
                            .parse::<u32>()
                            .map_err(|_| invalid("invalid shape ID in placeholder"));
                    }
                }
                return Err(invalid("placeholder shape has no shape ID"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }
    Err(invalid("placeholder shape has no non-visual properties"))
}

/// Allocate the next free shape ID for a part (max existing + 1, starting at 2).
fn next_shape_id(xml: &[u8]) -> Result<u32> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut max_id = FIRST_SHAPE_ID - 1;
    let mut nodes = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                if local_name(element.name().as_ref()) == b"cNvPr" {
                    for attribute in element.attributes().with_checks(true) {
                        let attribute =
                            attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        if attribute.key.as_ref() == b"id" {
                            let value = std::str::from_utf8(attribute.value.as_ref())
                                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                            let id = value
                                .parse::<u32>()
                                .map_err(|_| invalid(format!("invalid shape ID '{value}'")))?;
                            max_id = max_id.max(id);
                        }
                    }
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }
    max_id
        .checked_add(1)
        .ok_or_else(|| invalid("shape ID overflow"))
}

fn replace_span(xml: &[u8], span: &ElementSpan, replacement: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(xml.len() + replacement.len());
    output.extend_from_slice(&xml[..span.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[span.end..]);
    check_size(&output)?;
    Ok(output)
}

fn insert_bytes(xml: &[u8], offset: usize, value: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(xml.len() + value.len());
    output.extend_from_slice(&xml[..offset]);
    output.extend_from_slice(value);
    output.extend_from_slice(&xml[offset..]);
    check_size(&output)?;
    Ok(output)
}

// ============================================================================
// Presentation master-ID parsing and allocation
// ============================================================================

/// Parse `p:sldMasterIdLst` entries as `(id, relationship_id)` pairs.
///
/// Mirrors the read-side rules: IDs are unsigned 32-bit values at or above
/// [`MIN_MASTER_OR_LAYOUT_ID`], and both IDs and relationship IDs are unique.
fn parse_master_id_list(xml: &[u8]) -> Result<Vec<(u32, String)>> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut in_list = false;
    let mut entries = Vec::new();
    let mut nodes = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes += 1;
                depth += 1;
                if nodes > MAX_SCAN_NODES || depth > MAX_SCAN_DEPTH {
                    return Err(invalid("presentation XML resource limit exceeded"));
                }
                let local = local_name(element.name().as_ref()).to_vec();
                if depth == 2 && local == b"sldMasterIdLst" {
                    if in_list {
                        return Err(invalid("duplicate slide-master ID list"));
                    }
                    in_list = true;
                } else if depth == 3 && in_list && local == b"sldMasterId" {
                    push_master_id_entry(&mut entries, &element)?;
                }
            },
            Ok(Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("presentation XML resource limit exceeded"));
                }
                let local = local_name(element.name().as_ref()).to_vec();
                if depth == 1 && local == b"sldMasterIdLst" {
                    if in_list {
                        return Err(invalid("duplicate slide-master ID list"));
                    }
                    in_list = true;
                } else if depth == 2 && in_list && local == b"sldMasterId" {
                    push_master_id_entry(&mut entries, &element)?;
                }
            },
            Ok(Event::End(element)) => {
                if depth == 2 && local_name(element.name().as_ref()) == b"sldMasterIdLst" {
                    in_list = false;
                }
                if depth == 0 {
                    return Err(invalid("unexpected closing element in presentation XML"));
                }
                depth -= 1;
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }
    if depth != 0 {
        return Err(invalid("unterminated presentation XML"));
    }
    Ok(entries)
}

fn push_master_id_entry(
    entries: &mut Vec<(u32, String)>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<()> {
    let mut id = None;
    let mut relationship_id = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let value = std::str::from_utf8(attribute.value.as_ref())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if name == "id" {
            let parsed = value
                .parse::<u32>()
                .map_err(|_| invalid(format!("invalid slide-master ID '{value}'")))?;
            if parsed < MIN_MASTER_OR_LAYOUT_ID {
                return Err(invalid(format!(
                    "slide-master ID {parsed} is below {MIN_MASTER_OR_LAYOUT_ID}"
                )));
            }
            id = Some(parsed);
        } else if name.rsplit_once(':').map(|(_, local)| local) == Some("id") {
            relationship_id = Some(value.to_owned());
        }
    }
    let id = id.ok_or_else(|| invalid("slide-master entry is missing its ID"))?;
    let relationship_id = relationship_id
        .ok_or_else(|| invalid("slide-master entry is missing its relationship ID"))?;
    if relationship_id.is_empty() {
        return Err(invalid("empty slide-master relationship ID"));
    }
    if entries.iter().any(|(existing, _)| *existing == id) {
        return Err(invalid(format!("duplicate slide-master ID {id}")));
    }
    if entries
        .iter()
        .any(|(_, existing)| *existing == relationship_id)
    {
        return Err(invalid(format!(
            "duplicate slide-master relationship ID '{relationship_id}'"
        )));
    }
    entries.push((id, relationship_id));
    Ok(())
}

/// Allocate the next ID above the current maximum.
fn allocate_id(used: impl Iterator<Item = u32>) -> Result<u32> {
    used.max()
        .unwrap_or(MIN_MASTER_OR_LAYOUT_ID - 1)
        .checked_add(1)
        .ok_or_else(|| invalid("slide master or layout ID overflow"))
}

/// Find the lowest free numeric suffix for a part-name pattern.
fn next_part_index(package: &OpcPackage, prefix: &str, suffix: &str) -> Result<u32> {
    let mut index = 1u32;
    loop {
        let candidate = PackURI::new(format!("{prefix}{index}{suffix}"))
            .map_err(|error| OoxmlError::InvalidUri(format!("partname allocation: {error}")))?;
        if package.get_part(&candidate).is_err() {
            return Ok(index);
        }
        index = index
            .checked_add(1)
            .ok_or_else(|| invalid("part-name index overflow"))?;
    }
}

/// Resolve the theme relationship target for a newly created master.
///
/// Returns the relationship target (relative to `/ppt/slideMasters/`) and,
/// when no theme exists yet, a new theme part to add.
fn theme_target_for_new_master(
    package: &OpcPackage,
    master_dir: &str,
) -> Result<(String, Option<(PackURI, String)>)> {
    // Prefer the theme used by an existing slide master.
    for part in package.iter_parts() {
        if part.content_type() != ct::PML_SLIDE_MASTER {
            continue;
        }
        for relationship in part.rels().iter() {
            if relationship.reltype() == rt::THEME
                && !relationship.is_external()
                && let Ok(target) = relationship.target_partname()
                && package.get_part(&target).is_ok()
            {
                return Ok((relative_target(master_dir, target.as_str())?, None));
            }
        }
    }
    // Otherwise reuse any existing theme part.
    for part in package.iter_parts() {
        if part.content_type() == ct::OFC_THEME {
            return Ok((relative_target(master_dir, part.partname().as_str())?, None));
        }
    }
    // Otherwise author a fresh theme part from the default template.
    let index = next_part_index(package, "/ppt/theme/theme", ".xml")?;
    let uri = PackURI::new(format!("/ppt/theme/theme{index}.xml"))
        .map_err(|error| OoxmlError::InvalidUri(format!("theme partname: {error}")))?;
    Ok((
        format!("../theme/theme{index}.xml"),
        Some((uri, crate::pptx::template::default_theme_xml().to_string())),
    ))
}

/// Compute the relationship target for `target` relative to `source_dir`.
///
/// Both names must be absolute pack URIs; the result uses `..` segments to
/// climb out of the source directory.
fn relative_target(source_dir: &str, target: &str) -> Result<String> {
    let source = source_dir.trim_matches('/');
    let target = target.trim_start_matches('/');
    let source_segments: Vec<&str> = source.split('/').filter(|item| !item.is_empty()).collect();
    let target_segments: Vec<&str> = target.split('/').filter(|item| !item.is_empty()).collect();
    let common = source_segments
        .iter()
        .zip(target_segments.iter())
        .take_while(|(left, right)| left == right)
        .count();
    if common == 0 && !source_segments.is_empty() {
        return Err(OoxmlError::InvalidUri(format!(
            "cannot relativize '{target}' against '/{source}/'"
        )));
    }
    let mut result = String::new();
    for _ in common..source_segments.len() {
        result.push_str("../");
    }
    result.push_str(&target_segments[common..].join("/"));
    Ok(result)
}

// ============================================================================
// Misc validators
// ============================================================================

fn require_presentation_part(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
            | "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
    ) {
        Ok(())
    } else {
        Err(invalid("main document is not a PowerPoint presentation"))
    }
}

fn require_name(name: &str) -> Result<()> {
    if name.is_empty() {
        return Err(invalid("slide layout name cannot be empty"));
    }
    if name.chars().count() > MAX_NAME_CHARS {
        return Err(invalid("slide layout name exceeds 256 characters"));
    }
    Ok(())
}

fn require_placeholders(placeholders: &[PlaceholderSpec]) -> Result<()> {
    if placeholders.len() > MAX_PLACEHOLDERS_PER_OPERATION {
        return Err(invalid("too many placeholder shapes in one operation"));
    }
    let mut identities = HashSet::new();
    for spec in placeholders {
        if let Some(name) = &spec.name {
            require_name(name)?;
        }
        if !identities.insert((spec.kind, spec.effective_index())) {
            return Err(invalid(format!(
                "duplicate placeholder type '{}' with index {}",
                spec.kind.as_str(),
                spec.effective_index()
            )));
        }
    }
    Ok(())
}

fn invalidate_signatures(package: &mut OpcPackage) {
    package.unsign();
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pptx::Package;
    use litchi_opc::PackageWriter;
    use std::io::Cursor;

    fn roundtrip(package: &Package) -> Package {
        let bytes = PackageWriter::to_bytes(package.opc_package()).unwrap();
        Package::from_reader(Cursor::new(bytes)).unwrap()
    }

    #[test]
    fn authored_master_and_layouts_roundtrip_through_read_side() {
        let mut package = Package::new().unwrap();
        let master = package.add_slide_master().unwrap();
        assert_eq!(master.master_id, MIN_MASTER_OR_LAYOUT_ID + 1);

        let title_layout = package
            .add_slide_layout(
                &master.part_name,
                SlideLayoutKind::Title,
                "Custom Title",
                &[
                    PlaceholderSpec::new(PlaceholderKind::CenteredTitle)
                        .with_text("Click to edit the custom title"),
                    PlaceholderSpec::new(PlaceholderKind::Subtitle)
                        .with_index(1)
                        .with_text("Custom subtitle"),
                ],
            )
            .unwrap();
        let blank_layout = package
            .add_slide_layout(
                &master.part_name,
                SlideLayoutKind::Blank,
                "Custom Blank",
                &[],
            )
            .unwrap();
        assert!(title_layout.layout_id >= MIN_MASTER_OR_LAYOUT_ID);
        assert!(blank_layout.layout_id >= MIN_MASTER_OR_LAYOUT_ID);
        assert_ne!(title_layout.layout_id, blank_layout.layout_id);

        // Author placeholders on the master itself, then replace one.
        package
            .store_placeholder_shape(
                &master.part_name,
                &PlaceholderSpec::new(PlaceholderKind::Title).with_text("Master title"),
            )
            .unwrap();
        package
            .store_placeholder_shape(
                &master.part_name,
                &PlaceholderSpec::new(PlaceholderKind::DateTime).with_index(10),
            )
            .unwrap();
        package
            .store_placeholder_shape(
                &master.part_name,
                &PlaceholderSpec::new(PlaceholderKind::Title).with_text("Master title v2"),
            )
            .unwrap();
        package.validate_master_layout_graph().unwrap();

        let reopened = roundtrip(&package);
        reopened.validate_master_layout_graph().unwrap();
        let presentation = reopened.presentation().unwrap();
        let masters = presentation.slide_masters().unwrap();
        assert_eq!(masters.len(), 2, "default master plus authored master");

        let authored = masters
            .iter()
            .find(|candidate| candidate.part().part().partname().as_str() == master.part_name)
            .expect("authored master must resolve through the presentation");

        // Default text styles: title/body/other with nine levels each.
        let styles = authored.text_styles().unwrap().expect("authored txStyles");
        for style in [
            styles.title_style().expect("title style"),
            styles.body_style().expect("body style"),
            styles.other_style().expect("other style"),
        ] {
            assert_eq!(style.levels(), &[1, 2, 3, 4, 5, 6, 7, 8, 9]);
        }

        // Master placeholder inventory, including the replaced title text.
        let master_placeholders = authored.placeholders().unwrap();
        let titles = master_placeholders
            .iter()
            .filter(|shape| shape.placeholder_type().unwrap() == "title")
            .count();
        assert_eq!(titles, 1, "replaced title placeholder must not duplicate");
        let title = master_placeholders
            .iter()
            .find(|shape| shape.placeholder_type().unwrap() == "title")
            .unwrap();
        assert_eq!(title.text().unwrap().as_deref(), Some("Master title v2"));
        assert!(master_placeholders.iter().any(|shape| {
            shape.placeholder_type().unwrap() == "dt"
                && shape.placeholder_index().unwrap() == Some(10)
        }));

        // Layout inventory: kinds, names, placeholders, and back-references.
        let layouts = authored.slide_layouts().unwrap();
        assert_eq!(layouts.len(), 2);
        let title_layout_read = &layouts[0];
        assert_eq!(title_layout_read.metadata().unwrap().layout_type(), "title");
        assert_eq!(title_layout_read.name().unwrap(), "Custom Title");
        assert_eq!(
            title_layout_read
                .master()
                .unwrap()
                .part()
                .part()
                .partname()
                .as_str(),
            master.part_name
        );
        let layout_placeholders = title_layout_read.placeholders().unwrap();
        assert_eq!(layout_placeholders.len(), 2);
        let centered = layout_placeholders
            .iter()
            .find(|shape| shape.placeholder_type().unwrap() == "ctrTitle")
            .unwrap();
        assert_eq!(
            centered.text().unwrap().as_deref(),
            Some("Click to edit the custom title")
        );
        assert!(layout_placeholders.iter().any(|shape| {
            shape.placeholder_type().unwrap() == "subTitle"
                && shape.placeholder_index().unwrap() == Some(1)
        }));
        assert_eq!(layouts[1].metadata().unwrap().layout_type(), "blank");
        assert!(layouts[1].placeholders().unwrap().is_empty());

        // The authored master inherits a working theme relationship.
        authored.theme().unwrap();

        // The default master and its eleven layouts are untouched.
        let default_master = masters
            .iter()
            .find(|candidate| {
                candidate.part().part().partname().as_str() == "/ppt/slideMasters/slideMaster1.xml"
            })
            .unwrap();
        assert_eq!(default_master.slide_layouts().unwrap().len(), 11);
    }

    #[test]
    fn master_ids_are_unique_across_multiple_adds() {
        let mut package = Package::new().unwrap();
        let first = package.add_slide_master().unwrap();
        let second = package.add_slide_master().unwrap();
        let third = package.add_slide_master().unwrap();
        assert_eq!(first.master_id, MIN_MASTER_OR_LAYOUT_ID + 1);
        assert_eq!(second.master_id, MIN_MASTER_OR_LAYOUT_ID + 2);
        assert_eq!(third.master_id, MIN_MASTER_OR_LAYOUT_ID + 3);
        package.validate_master_layout_graph().unwrap();

        let reopened = roundtrip(&package);
        assert_eq!(
            reopened
                .presentation()
                .unwrap()
                .slide_masters()
                .unwrap()
                .len(),
            4
        );
        reopened.validate_master_layout_graph().unwrap();
    }

    #[test]
    fn authored_layout_attaches_to_default_master() {
        let mut package = Package::new().unwrap();
        let layout = package
            .add_slide_layout(
                "/ppt/slideMasters/slideMaster1.xml",
                SlideLayoutKind::TwoObjects,
                "Two Objects Extra",
                &[PlaceholderSpec::new(PlaceholderKind::Object).with_index(7)],
            )
            .unwrap();
        assert!(layout.layout_id > MIN_MASTER_OR_LAYOUT_ID + 11);

        let reopened = roundtrip(&package);
        let presentation = reopened.presentation().unwrap();
        let default_master = &presentation.slide_masters().unwrap()[0];
        let layouts = default_master.slide_layouts().unwrap();
        assert_eq!(layouts.len(), 12);
        let added = layouts
            .iter()
            .find(|candidate| candidate.name().unwrap() == "Two Objects Extra")
            .unwrap();
        assert_eq!(added.metadata().unwrap().layout_type(), "twoObj");
        let placeholders = added.placeholders().unwrap();
        assert_eq!(placeholders.len(), 1);
        assert_eq!(placeholders[0].placeholder_index().unwrap(), Some(7));
    }

    #[test]
    fn invalid_references_are_rejected() {
        let mut package = Package::new().unwrap();

        // Unknown master part.
        assert!(
            package
                .add_slide_layout(
                    "/ppt/slideMasters/slideMaster99.xml",
                    SlideLayoutKind::Blank,
                    "Nope",
                    &[],
                )
                .is_err()
        );
        // Master part name pointing at a non-master part.
        assert!(
            package
                .add_slide_layout("/ppt/presentation.xml", SlideLayoutKind::Blank, "Nope", &[],)
                .is_err()
        );
        // Placeholder authoring on a part that is not a master or layout.
        assert!(
            package
                .store_placeholder_shape(
                    "/ppt/presentation.xml",
                    &PlaceholderSpec::new(PlaceholderKind::Title),
                )
                .is_err()
        );
        // Empty layout names are rejected.
        let master = package.add_slide_master().unwrap();
        assert!(
            package
                .add_slide_layout(&master.part_name, SlideLayoutKind::Blank, "", &[])
                .is_err()
        );
        // Duplicate placeholder identities are rejected.
        assert!(
            package
                .add_slide_layout(
                    &master.part_name,
                    SlideLayoutKind::Blank,
                    "Dup",
                    &[
                        PlaceholderSpec::new(PlaceholderKind::Body).with_index(1),
                        PlaceholderSpec::new(PlaceholderKind::Body).with_index(1),
                    ],
                )
                .is_err()
        );
        // Removing unknown layouts is rejected.
        assert!(
            package
                .remove_slide_layout("/ppt/slideLayouts/slideLayout99.xml")
                .is_err()
        );
        package.validate_master_layout_graph().unwrap();
    }

    #[test]
    fn remove_layout_rejects_slide_references() {
        let mut package = Package::new().unwrap();
        let master = package.add_slide_master().unwrap();
        let layout = package
            .add_slide_layout(&master.part_name, SlideLayoutKind::Blank, "In Use", &[])
            .unwrap();

        // Attach a slide part that references the layout.
        {
            let opc = package.opc_package_mut();
            let slide_uri = PackURI::new("/ppt/slides/slide1.xml").unwrap();
            let mut slide = BlobPart::new(
                slide_uri,
                ct::PML_SLIDE.to_string(),
                b"<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>".to_vec(),
            );
            slide.relate_to(
                &format!("../{}", layout.part_name.trim_start_matches("/ppt/")),
                rt::SLIDE_LAYOUT,
            );
            opc.add_part(Box::new(slide));
        }

        assert!(package.remove_slide_layout(&layout.part_name).is_err());
        package.validate_master_layout_graph().unwrap();
    }

    #[test]
    fn remove_empty_layout_keeps_graph_consistent() {
        let mut package = Package::new().unwrap();
        let master = package.add_slide_master().unwrap();
        let layout = package
            .add_slide_layout(&master.part_name, SlideLayoutKind::Blank, "Temporary", &[])
            .unwrap();
        package
            .add_slide_layout(&master.part_name, SlideLayoutKind::TitleOnly, "Kept", &[])
            .unwrap();

        package.remove_slide_layout(&layout.part_name).unwrap();
        package.validate_master_layout_graph().unwrap();
        assert!(
            package
                .opc_package()
                .get_part(&PackURI::new(&layout.part_name).unwrap())
                .is_err(),
            "layout part must be gone"
        );

        let reopened = roundtrip(&package);
        reopened.validate_master_layout_graph().unwrap();
        let presentation = reopened.presentation().unwrap();
        let masters = presentation.slide_masters().unwrap();
        let authored = masters
            .iter()
            .find(|candidate| candidate.part().part().partname().as_str() == master.part_name)
            .unwrap();
        let layouts = authored.slide_layouts().unwrap();
        assert_eq!(layouts.len(), 1);
        assert_eq!(layouts[0].name().unwrap(), "Kept");

        // Deleting it a second time is an error.
        assert!(package.remove_slide_layout(&layout.part_name).is_err());
    }

    #[test]
    fn authored_parts_serialize_deterministically() {
        let build = || {
            let mut package = Package::new().unwrap();
            let master = package.add_slide_master().unwrap();
            package
                .add_slide_layout(
                    &master.part_name,
                    SlideLayoutKind::SectionHeader,
                    "Deterministic",
                    &[PlaceholderSpec::new(PlaceholderKind::Title).with_text("Same")],
                )
                .unwrap();
            package
        };
        let first = build();
        let second = build();
        for part_name in [
            "/ppt/slideMasters/slideMaster2.xml",
            "/ppt/slideLayouts/slideLayout12.xml",
        ] {
            let uri = PackURI::new(part_name).unwrap();
            assert_eq!(
                first.opc_package().get_part(&uri).unwrap().blob(),
                second.opc_package().get_part(&uri).unwrap().blob(),
                "part {part_name} must serialize deterministically"
            );
        }
    }
}
