//! Typed SmartArt authoring and part replacement (record 0284).
//!
//! [`Edit::add_smartart`] stages the five diagram parts of an authored
//! [`Spec`], their slide relationships and the graphic frame;
//! [`Edit::smartart`] reads the parts of an existing SmartArt frame and
//! [`Edit::set_smartart`] replaces them, creating, keeping or removing the
//! cached drawing so the data model always names exactly the drawing the
//! slide relates.

use litchi_drawingml::diagram::authoring::{Spec, colors_xml, quick_style_xml};
use litchi_drawingml::diagram::{
    DIAGRAM_COLORS_REL, DIAGRAM_DATA_REL, DIAGRAM_LAYOUT_REL, DIAGRAM_QUICK_STYLE_REL, Limits,
    MS_DIAGRAM_DRAWING_REL, edit as diagram_edit,
};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::{Edit, invalid, next_shape_id};
use crate::Result;

const DGM: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/diagram";
const DGM_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/diagram";
const R: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const R_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";

fn drawing_error(error: litchi_drawingml::Error, context: &str) -> crate::Error {
    match error {
        litchi_drawingml::Error::Invalid(_) | litchi_drawingml::Error::Xml(_) => {
            invalid(format!("{context} is invalid"))
        },
        error => crate::Error::Drawing(error),
    }
}

/// The parts of one SmartArt graphic.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SmartArtParts {
    /// The data model (`dgm:dataModel`).
    pub data: Vec<u8>,
    /// The layout definition (`dgm:layoutDef`).
    pub layout: Vec<u8>,
    /// The quick style (`dgm:styleDef`).
    pub quick_style: Vec<u8>,
    /// The colors (`dgm:colorsDef`).
    pub colors: Vec<u8>,
    /// The cached drawing (`dsp:drawing`), if the graphic has one.
    pub drawing: Option<Vec<u8>>,
}

/// The parts a SmartArt frame names, by role.
struct Targets {
    data: PackURI,
    layout: PackURI,
    quick_style: PackURI,
    colors: PackURI,
    /// The slide relationship and part of the cached drawing.
    drawing: Option<(String, PackURI)>,
}

impl Edit {
    /// Add an authored SmartArt graphic: its data model, layout, quick style,
    /// colors and Litchi's cached drawing (laid out for `bounds`), their
    /// slide relationships and a visible graphic frame.
    ///
    /// `bounds` is `(x, y, width, height)` in EMU. The parts are named
    /// `/ppt/diagrams/{data,layout,quickStyle,colors,drawing}N.xml` with the
    /// first `N` free for all five, and are generated deterministically.
    /// Returns the new frame's shape ID.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006).
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide, a spec the authoring module
    /// refuses, non-positive extents, invalid name text, a malformed shape
    /// tree, or exhausted shape IDs.
    pub fn add_smartart<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        spec: &Spec,
        name: &str,
        bounds: (i64, i64, i64, i64),
    ) -> Result<u32> {
        if bounds.2 <= 0 || bounds.3 <= 0 {
            return Err(invalid("SmartArt frame extents must be positive"));
        }
        if name.contains('\0') || name.len() > 32_767 {
            return Err(invalid("SmartArt frame name is not valid text"));
        }
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        crate::package::signature::refuse_signed_change(&candidate)?;
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(
            candidate.get_part(&selected.part_name)?.blob(),
            shape_limits,
        )?;
        let shape_id = next_shape_id(&scene, "SmartArt")?;
        let number = free_diagram_number(&candidate)?;
        let part_name = |stem: &str| {
            PackURI::new(format!("/ppt/diagrams/{stem}{number}.xml")).map_err(crate::Error::Uri)
        };
        let names = [
            (part_name("data")?, ct::DML_DIAGRAM_DATA, DIAGRAM_DATA_REL),
            (
                part_name("layout")?,
                ct::DML_DIAGRAM_LAYOUT,
                DIAGRAM_LAYOUT_REL,
            ),
            (
                part_name("quickStyle")?,
                ct::DML_DIAGRAM_STYLE,
                DIAGRAM_QUICK_STYLE_REL,
            ),
            (
                part_name("colors")?,
                ct::DML_DIAGRAM_COLORS,
                DIAGRAM_COLORS_REL,
            ),
            (
                part_name("drawing")?,
                ct::DML_DIAGRAM_DRAWING,
                MS_DIAGRAM_DRAWING_REL,
            ),
        ];
        let mut ids = Vec::with_capacity(names.len());
        for (name, _, relationship) in &names {
            let reference = name.relative_ref(selected.part_name.base_uri());
            ids.push(
                candidate
                    .get_part_mut(&selected.part_name)?
                    .relate_to(&reference, relationship),
            );
        }
        let authored = |error: litchi_drawingml::Error| drawing_error(error, "authored SmartArt");
        let limits = Limits::default();
        let blobs = [
            spec.data_xml(Some(&ids[4]), &limits)
                .map_err(authored)?
                .into_bytes(),
            spec.layout_xml(&limits).map_err(authored)?.into_bytes(),
            quick_style_xml(&limits).map_err(authored)?.into_bytes(),
            colors_xml(&limits).map_err(authored)?.into_bytes(),
            spec.drawing_xml(bounds.2, bounds.3, &limits)
                .map_err(authored)?
                .into_bytes(),
        ];
        for ((name, content_type, _), blob) in names.into_iter().zip(blobs) {
            candidate.add_part(Box::new(BlobPart::new(name, content_type.to_owned(), blob)));
        }
        let fragment = frame(shape_id, name, bounds, [&ids[0], &ids[1], &ids[2], &ids[3]]);
        let owner = candidate.get_part(&selected.part_name)?;
        let xml = super::super::xml::append_shape_with_limit(
            owner.blob(),
            fragment.as_bytes(),
            shape_limits.output_bytes(),
        )?;
        let scene = crate::shape::Scene::read(&xml, shape_limits)?;
        if !scene.iter().any(|shape| {
            shape.id() == Some(shape_id) && matches!(shape, crate::shape::Shape::Diagram(_))
        }) {
            return Err(invalid(
                "opened-presentation SmartArt frame did not round-trip semantically",
            ));
        }
        candidate.get_part_mut(&selected.part_name)?.set_blob(xml);
        crate::package::signature::settle_unsigned_change(&mut candidate)?;
        self.working = candidate;
        Ok(shape_id)
    }

    /// The parts of the SmartArt frame at a depth-first shape position of a
    /// slide.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide or shape, a shape that is not a
    /// SmartArt frame, or broken diagram relationships.
    pub fn smartart<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<SmartArtParts> {
        let selected = self.resolve_slide(slide.into())?;
        let targets = targets(
            &self.working,
            &selected.part_name,
            shape.into(),
            self.semantic_limits().shape(),
        )?;
        let read = |name: &PackURI| -> Result<Vec<u8>> {
            Ok(self.working.get_part(name)?.blob().to_vec())
        };
        Ok(SmartArtParts {
            data: read(&targets.data)?,
            layout: read(&targets.layout)?,
            quick_style: read(&targets.quick_style)?,
            colors: read(&targets.colors)?,
            drawing: match &targets.drawing {
                Some((_, name)) => Some(read(name)?),
                None => None,
            },
        })
    }

    /// Replace the parts of the SmartArt frame at a depth-first shape
    /// position of a slide.
    ///
    /// The data model must read as one. A `drawing` replaces the cached
    /// drawing, or creates one (`/ppt/diagrams/drawingN.xml` and its slide
    /// relationship) when the graphic had none; `None` removes the cached
    /// drawing and its relationship, so consumers lay the graphic out again
    /// from the data model. Either way the data model's drawing reference
    /// (`dsp:dataModelExt/@relId`) is rewritten to name exactly the drawing
    /// that remains. Returns whether anything changed; identical parts are an
    /// exact no-op.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006).
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide or shape, a shape that is not a
    /// SmartArt frame, a data model that does not read, or broken diagram
    /// relationships.
    pub fn set_smartart<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        parts: SmartArtParts,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let targets = targets(
            &self.working,
            &selected.part_name,
            shape.into(),
            self.semantic_limits().shape(),
        )?;
        let read_error =
            |error: litchi_drawingml::Error| drawing_error(error, "SmartArt data model");
        diagram_edit::Outline::read(&parts.data, &Limits::default()).map_err(read_error)?;
        let mut candidate = self.working.clone();
        let SmartArtParts {
            data,
            layout,
            quick_style,
            colors,
            drawing,
        } = parts;
        // Settle the drawing first: its relationship ID goes into the data.
        let relationship = match (drawing, &targets.drawing) {
            (Some(bytes), Some((id, name))) => {
                if candidate.get_part(name)?.blob() != bytes.as_slice() {
                    candidate.get_part_mut(name)?.set_blob(bytes);
                }
                Some(id.clone())
            },
            (Some(bytes), None) => {
                let number = free_diagram_number(&candidate)?;
                let name = PackURI::new(format!("/ppt/diagrams/drawing{number}.xml"))
                    .map_err(crate::Error::Uri)?;
                let reference = name.relative_ref(selected.part_name.base_uri());
                let id = candidate
                    .get_part_mut(&selected.part_name)?
                    .relate_to(&reference, MS_DIAGRAM_DRAWING_REL);
                candidate.add_part(Box::new(BlobPart::new(
                    name,
                    ct::DML_DIAGRAM_DRAWING.to_owned(),
                    bytes,
                )));
                Some(id)
            },
            (None, Some((id, name))) => {
                candidate
                    .get_part_mut(&selected.part_name)?
                    .rels_mut()
                    .remove(id);
                super::remove_unreferenced_dependencies(&mut candidate, vec![name.clone()])?;
                None
            },
            (None, None) => None,
        };
        let data =
            diagram_edit::set_drawing_reference(&data, relationship.as_deref(), &Limits::default())
                .map_err(read_error)?;
        for (name, bytes) in [
            (&targets.data, data),
            (&targets.layout, layout),
            (&targets.quick_style, quick_style),
            (&targets.colors, colors),
        ] {
            if candidate.get_part(name)?.blob() != bytes.as_slice() {
                candidate.get_part_mut(name)?.set_blob(bytes);
            }
        }
        if same_package(&self.working, &candidate, &selected.part_name, &targets)? {
            return Ok(false);
        }
        crate::package::signature::refuse_signed_change(&self.working)?;
        crate::package::signature::settle_unsigned_change(&mut candidate)?;
        self.working = candidate;
        Ok(true)
    }
}

/// Whether the diagram parts and the slide relationships are unchanged.
fn same_package(
    before: &OpcPackage,
    after: &OpcPackage,
    slide: &PackURI,
    targets: &Targets,
) -> Result<bool> {
    let names = [
        &targets.data,
        &targets.layout,
        &targets.quick_style,
        &targets.colors,
    ];
    for name in names {
        if before.get_part(name)?.blob() != after.get_part(name)?.blob() {
            return Ok(false);
        }
    }
    if let Some((_, name)) = &targets.drawing
        && (!after.contains_part(name)
            || before.get_part(name)?.blob() != after.get_part(name)?.blob())
    {
        return Ok(false);
    }
    let rels = |package: &OpcPackage| -> Result<Vec<(String, String, String)>> {
        Ok(package
            .get_part(slide)?
            .rels()
            .iter()
            .map(|relationship| {
                (
                    relationship.r_id().to_owned(),
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                )
            })
            .collect())
    };
    Ok(rels(before)? == rels(after)? && before.part_count() == after.part_count())
}

/// The first `N` for which no `/ppt/diagrams/*N.xml` part exists.
fn free_diagram_number(package: &OpcPackage) -> Result<u32> {
    for number in 1_u32..=1_000_000 {
        let taken = ["data", "layout", "quickStyle", "colors", "drawing"]
            .iter()
            .any(|stem| {
                PackURI::new(format!("/ppt/diagrams/{stem}{number}.xml"))
                    .is_ok_and(|name| package.contains_part(&name))
            });
        if !taken {
            return Ok(number);
        }
    }
    Err(invalid("no free diagram part name is left"))
}

/// The graphic frame of an authored SmartArt.
fn frame(shape_id: u32, name: &str, bounds: (i64, i64, i64, i64), ids: [&str; 4]) -> String {
    use litchi_core::xml::escape_xml;
    let (x, y, width, height) = bounds;
    format!(
        concat!(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{id}" name="{name}"/>"#,
            r#"<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"#,
            r#"<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/></p:xfrm>"#,
            r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">"#,
            r#"<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="{dm}" r:lo="{lo}" r:qs="{qs}" r:cs="{cs}"/>"#,
            r#"</a:graphicData></a:graphic></p:graphicFrame>"#
        ),
        id = shape_id,
        name = escape_xml(name),
        x = x,
        y = y,
        width = width,
        height = height,
        dm = escape_xml(ids[0]),
        lo = escape_xml(ids[1]),
        qs = escape_xml(ids[2]),
        cs = escape_xml(ids[3]),
    )
}

/// The relationship IDs a SmartArt frame's `dgm:relIds` names, in the order
/// data, layout, quick style, colors.
pub(crate) fn frame_relationships(fragment: &[u8]) -> Result<Option<[String; 4]>> {
    let mut reader = NsReader::from_reader(fragment);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        buffer.clear();
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|_| invalid("invalid SmartArt frame XML"))?;
        let diagram = matches!(resolved, ResolveResult::Bound(Namespace(value)) if value == DGM || value == DGM_STRICT);
        match event {
            Event::Start(element) | Event::Empty(element)
                if diagram && element.local_name().as_ref() == b"relIds" =>
            {
                let mut ids: [Option<String>; 4] = Default::default();
                for attribute in
                    litchi_ooxml_common::xml::attributes::BytesStartExt::checked_attributes(
                        &element,
                    )
                {
                    let attribute = attribute.map_err(|_| invalid("invalid SmartArt frame XML"))?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    // The fragment is cut from its slide, so the `r:` prefix
                    // may be declared only on the slide root.
                    let relationship = match namespace {
                        ResolveResult::Bound(Namespace(value)) => value == R || value == R_STRICT,
                        ResolveResult::Unknown(_) => true,
                        ResolveResult::Unbound => false,
                    };
                    if !relationship {
                        continue;
                    }
                    let slot = match local.as_ref() {
                        b"dm" => 0,
                        b"lo" => 1,
                        b"qs" => 2,
                        b"cs" => 3,
                        _ => continue,
                    };
                    ids[slot] =
                        Some(String::from_utf8_lossy(attribute.value.as_ref()).into_owned());
                }
                let [Some(dm), Some(lo), Some(qs), Some(cs)] = ids else {
                    return Ok(None);
                };
                return Ok(Some([dm, lo, qs, cs]));
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
    }
}

/// The diagram parts the SmartArt frame at `shape` on slide part `slide`
/// names.
fn targets(
    package: &OpcPackage,
    slide: &PackURI,
    shape: crate::shape::Key<'_>,
    shape_limits: crate::shape::Limits,
) -> Result<Targets> {
    let owner = package.get_part(slide)?;
    let blob = owner.blob();
    {
        let scene = crate::shape::Scene::read(blob, shape_limits)?;
        let found = scene
            .get(shape)
            .map_err(|error| invalid(format!("SmartArt frame lookup: {error}")))?
            .ok_or_else(|| invalid("the SmartArt frame was not found"))?;
        if !matches!(found, crate::shape::Shape::Diagram(_)) {
            return Err(invalid("the shape is not a SmartArt frame"));
        }
    }
    let span = super::super::xml::shape_removal_span_with_limits(blob, shape, shape_limits)?;
    let fragment = super::super::xml::span_bytes(blob, span)?;
    let ids = frame_relationships(fragment)?
        .ok_or_else(|| invalid("the SmartArt frame names no diagram parts"))?;
    let resolve =
        |id: &str, relationship: &str, local: &str, content_type: &str| -> Result<PackURI> {
            let found = owner
                .rels()
                .get(id)
                .ok_or_else(|| invalid("a SmartArt relationship is missing"))?;
            if found.is_external()
                || !crate::parts::is_relationship_type(found.reltype(), relationship, local)
            {
                return Err(invalid("a SmartArt relationship has the wrong type"));
            }
            let name = found.target_partname()?;
            if package.get_part(&name)?.content_type() != content_type {
                return Err(invalid("a SmartArt part has the wrong content type"));
            }
            Ok(name)
        };
    let data = resolve(
        &ids[0],
        DIAGRAM_DATA_REL,
        "diagramData",
        ct::DML_DIAGRAM_DATA,
    )?;
    let layout = resolve(
        &ids[1],
        DIAGRAM_LAYOUT_REL,
        "diagramLayout",
        ct::DML_DIAGRAM_LAYOUT,
    )?;
    let quick_style = resolve(
        &ids[2],
        DIAGRAM_QUICK_STYLE_REL,
        "diagramQuickStyle",
        ct::DML_DIAGRAM_STYLE,
    )?;
    let colors = resolve(
        &ids[3],
        DIAGRAM_COLORS_REL,
        "diagramColors",
        ct::DML_DIAGRAM_COLORS,
    )?;
    let outline = diagram_edit::Outline::read(package.get_part(&data)?.blob(), &Limits::default())
        .map_err(|error| drawing_error(error, "SmartArt data model"))?;
    let drawing = match outline.drawing_relationship {
        Some(id) => match owner.rels().get(&id) {
            Some(found) if !found.is_external() && found.reltype() == MS_DIAGRAM_DRAWING_REL => {
                let name = found.target_partname()?;
                package.contains_part(&name).then_some((id, name))
            },
            _ => None,
        },
        None => None,
    };
    Ok(Targets {
        data,
        layout,
        quick_style,
        colors,
        drawing,
    })
}

/// The slide relationship of the cached drawing of the SmartArt whose data
/// model the slide relates through `data_relationship`, when it has one.
pub(crate) fn drawing_relationship(
    package: &OpcPackage,
    slide: &PackURI,
    data_relationship: &str,
) -> Option<String> {
    let owner = package.get_part(slide).ok()?;
    let found = owner.rels().get(data_relationship)?;
    if found.is_external() {
        return None;
    }
    let name = found.target_partname().ok()?;
    let outline =
        diagram_edit::Outline::read(package.get_part(&name).ok()?.blob(), &Limits::default())
            .ok()?;
    let id = outline.drawing_relationship?;
    let drawing = owner.rels().get(&id)?;
    (drawing.reltype() == MS_DIAGRAM_DRAWING_REL).then_some(id)
}
