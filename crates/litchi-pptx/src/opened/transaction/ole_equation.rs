//! OLE equation objects: add, replace and remove.
//!
//! A new object is written in the form ECMA-376 Part 1 defines on its own
//! (§19.3.2, embedded objects: `oleObj`, `embed`, `pic`): a
//! `p:graphicFrame` whose `a:graphicData` holds one `p:oleObj` with
//! `p:embed` and a `p:pic` showing the preview picture. PowerPoint 2010 and
//! later write the object a second time, in an `mc:Choice Requires="v"`
//! branch naming a shape of a legacy VML drawing part, which is how
//! PowerPoint 2007 shows an OLE object; that form needs a VML part for a
//! reader this crate does not target, and its `mc:Fallback` holds the form
//! written here. The preview is a PNG (Office writes EMF or WMF): an
//! ordinary image part the `a:blip` names, which the server redraws when
//! the object is activated.
//!
//! [`Edit::set_ole_equation`] replaces the `Equation Native` stream of the
//! object's compound file, keeping its kind and every other stream, redraws
//! the preview and scales the frame with the equation. An object in
//! PowerPoint's markup-compatibility form becomes the plain form: its
//! Fallback object is kept and its legacy VML shape, whose picture would
//! still show the old equation, is removed; an object without a `p:pic`
//! gets one. [`Edit::remove_ole_equation`]
//! removes the frame, its embedding and preview parts, and its VML shape
//! (with the VML drawing part when no other shape is left in it).
//! Every other byte of the slide stays as it was.

use std::collections::BTreeSet;
use std::ops::Range;

use litchi_equation::{Limits, MtefTarget, PngOptions};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use quick_xml::Reader;
use quick_xml::events::Event;

use litchi_drawingml::math::{
    carry_declarations, max_raw_shape_id, require_declared_prefixes, tag_name,
};

use super::equation::{Splice, resize};
use super::picture::attribute_value;
use super::{
    Edit, conformance_image_relationship, conformance_namespaces, ensure_shape_id, free_part_name,
    invalid, max_shape_reference_id, remove_unreferenced_dependencies, shape_relationship_ids,
};
use crate::shape::math::ole::{self, Branch, FrameSite, OLE_URI, STRICT_OLE_URI};
use crate::shape::math::{EquationWrite, OleEquationSource as EquationSource};
use crate::{Error, Result};

/// The text size a new equation is drawn at when the caller names none, in
/// points: PowerPoint's default for text boxes, as for Office Math
/// equations.
const DEFAULT_FONT_SIZE_PT: f64 = 18.0;
/// The text size the size of an existing equation is measured at.
const MEASURE_POINTS: f64 = 12.0;
/// Pixels per CSS pixel of a preview at the frame's size (384 per inch).
const PREVIEW_SCALE: f64 = 4.0;
/// The smallest scale a preview is drawn at to fit the pixel ceiling.
const MIN_PREVIEW_SCALE: f64 = 1.0 / 64.0;
/// EMU per point (ECMA-376 Part 1, §20.1.10.16).
const EMU_PER_POINT: f64 = 12_700.0;
/// The `EQNOLEFILEHDR` length.
const OLE_HEADER_LEN: usize = 28;
const OLE_OBJECT_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.oleObject";

fn refused(operation: &'static str, reason: &'static str) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        operation,
        reason,
    )
}

fn invalid_value(operation: &'static str, reason: &'static str) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::InvalidValue,
        operation,
        reason,
    )
}

/// The MTEF version an object of `target`'s kind holds: 3 or 5.
fn version_of(target: &MtefTarget) -> u8 {
    if target.prog_id() == "Equation.3" {
        3
    } else {
        5
    }
}

/// Bare MTEF behind an `EQNOLEFILEHDR` ([MathType SDK] "Equation Native"
/// header: `cbHdr` 28, version 0x00020000, the MathType EF clipboard
/// format, the MTEF length and four reserved zero fields).
fn with_ole_header(mtef: &[u8]) -> Result<Vec<u8>> {
    let len = u32::try_from(mtef.len()).map_err(|_error| invalid("the MTEF exceeds 4 GiB"))?;
    let mut out = Vec::with_capacity(OLE_HEADER_LEN + mtef.len());
    out.extend_from_slice(&28u16.to_le_bytes());
    out.extend_from_slice(&0x0002_0000u32.to_le_bytes());
    out.extend_from_slice(&0xC1C6u16.to_le_bytes());
    out.extend_from_slice(&len.to_le_bytes());
    out.extend_from_slice(&[0; 16]);
    out.extend_from_slice(mtef);
    Ok(out)
}

/// The `Equation Native` stream of `source`, the object kind it is written
/// as, and the conversion losses. `kind` is the existing object's when
/// replacing one; a new object is MathType, or for raw MTEF the kind of
/// its version.
fn equation_native(
    source: EquationSource<'_>,
    kind: Option<MtefTarget>,
    max_bytes: usize,
) -> Result<(Vec<u8>, MtefTarget, Vec<String>)> {
    let size = match source {
        EquationSource::Latex(text) | EquationSource::Omml(text) => text.len(),
        EquationSource::Mtef(bytes) => bytes.len(),
    };
    if size > max_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::PatchBytes,
            observed: size,
            maximum: max_bytes,
        });
    }
    let limits = Limits::DEFAULT;
    let expected = kind.as_ref().map(version_of);
    let target = kind.unwrap_or(MtefTarget::MathType);
    match source {
        EquationSource::Latex(latex) => {
            let native = litchi_equation::latex_to_equation_native(latex, &target, &limits)?;
            let losses = native.losses().to_vec();
            Ok((native.into_bytes(), target, losses))
        },
        EquationSource::Omml(omml) => {
            let native = litchi_equation::omml_to_equation_native(omml, &target, &limits)?;
            let losses = native.losses().to_vec();
            Ok((native.into_bytes(), target, losses))
        },
        EquationSource::Mtef(bytes) => {
            let read = litchi_equation::read_mtef(bytes, &limits)?;
            let info = read
                .info()
                .cloned()
                .ok_or_else(|| invalid_value("equation.mtef", "the MTEF has no header"))?;
            let same = MtefTarget::SameAs(info.clone());
            if expected.is_some_and(|expected| expected != version_of(&same)) {
                return Err(invalid_value(
                    "equation.mtef",
                    "the MTEF version differs from the equation's",
                ));
            }
            let native = if info.ole_header.is_some() {
                bytes.to_vec()
            } else {
                with_ole_header(bytes)?
            };
            if native.len() as u64 > litchi_equation::MAX_EQUATION_NATIVE_BYTES {
                return Err(invalid_value("equation.mtef", "the MTEF is too large"));
            }
            Ok((native, same, Vec::new()))
        },
    }
}

/// Render `native` with `options`, halving the scale while the picture
/// exceeds the renderer's pixel ceiling.
fn render_fitting(
    native: &[u8],
    mut options: PngOptions,
) -> std::result::Result<litchi_equation::Png, litchi_equation::Error> {
    loop {
        match litchi_equation::mtef_to_png(native, &options, &Limits::DEFAULT) {
            Err(litchi_equation::Error::Render(_)) if options.scale() > MIN_PREVIEW_SCALE => {
                options = options.with_scale((options.scale() / 2.0).max(MIN_PREVIEW_SCALE))?;
            },
            result => return result,
        }
    }
}

/// The size of `native` drawn at [`MEASURE_POINTS`], in points.
fn measure(native: &[u8]) -> Result<(f64, f64)> {
    let png = render_fitting(native, PngOptions::new(MEASURE_POINTS)?.with_scale(1.0)?)?;
    Ok(png.size_pt())
}

fn emu(points: f64) -> Result<i64> {
    let value = (points * EMU_PER_POINT).round();
    if value.is_finite() && (1.0..=2_147_483_647.0).contains(&value) {
        #[allow(
            clippy::cast_possible_truncation,
            reason = "the value is checked to lie in the 32-bit coordinate range"
        )]
        Ok(value as i64)
    } else {
        Err(invalid("the equation frame is too large"))
    }
}

/// A preview of `native` for a frame of `extent` EMU: drawn at
/// [`PREVIEW_SCALE`] pixels per CSS pixel of the frame.
fn preview(native: &[u8], natural: (f64, f64), extent: (i64, i64)) -> Result<Vec<u8>> {
    #[allow(
        clippy::cast_precision_loss,
        reason = "frame extents are far below 2^52 EMU"
    )]
    let zoom = (extent.0 as f64 / EMU_PER_POINT / natural.0.max(f64::MIN_POSITIVE))
        .min(extent.1 as f64 / EMU_PER_POINT / natural.1.max(f64::MIN_POSITIVE));
    let scale = (PREVIEW_SCALE * zoom).clamp(MIN_PREVIEW_SCALE, PngOptions::MAX_SCALE);
    let png = render_fitting(native, PngOptions::new(MEASURE_POINTS)?.with_scale(scale)?)?;
    Ok(png.into_bytes())
}

/// `extent` scaled by the size change from `old` to `new` (points); when the
/// old equation cannot be measured, the height is kept and the width
/// follows the new aspect ratio.
fn rescaled(extent: (i64, i64), old: Option<(f64, f64)>, new: (f64, f64)) -> Result<(i64, i64)> {
    #![allow(
        clippy::cast_precision_loss,
        reason = "frame extents are far below 2^52 EMU"
    )]
    let (width, height) = (extent.0 as f64, extent.1 as f64);
    let (width, height) = match old {
        Some((old_width, old_height)) if old_width > 0.0 && old_height > 0.0 => {
            (width * new.0 / old_width, height * new.1 / old_height)
        },
        _ => (height * new.0 / new.1.max(f64::MIN_POSITIVE), height),
    };
    Ok((emu(width / EMU_PER_POINT)?, emu(height / EMU_PER_POINT)?))
}

/// Replace the value of attribute `name` in the start tag at `tag`, when
/// the tag has it.
fn set_attribute(
    splice: &mut Splice,
    xml: &[u8],
    tag: &Range<usize>,
    name: &[u8],
    value: &str,
) -> bool {
    match attribute_value(&xml[tag.clone()], name) {
        Some((span, _, _)) => {
            splice.replace(tag.start + span.start..tag.start + span.end, value);
            true
        },
        None => false,
    }
}

/// The object a replacement keeps: the frame's only `p:oleObj`, or in
/// PowerPoint's markup-compatibility form the Fallback's (which holds the
/// preview picture).
fn kept_object(site: &FrameSite) -> Option<&ole::ObjectSite> {
    match &site.block {
        None => site.objects.first(),
        Some(_) => site
            .objects
            .iter()
            .find(|object| object.branch == Branch::Fallback),
    }
}

/// The slide's equation frame at `index` among its OLE equations, refused
/// when the frame sits in a markup-compatibility block of its own.
fn located(sites: &[FrameSite], index: usize, operation: &'static str) -> Result<FrameSite> {
    let site = sites.get(index).ok_or(Error::IndexOutOfBounds {
        index,
        len: sites.len(),
    })?;
    if site.in_block {
        return Err(refused(
            operation,
            "the equation frame sits in a markup-compatibility block of its own",
        ));
    }
    Ok(site.clone())
}

/// The internal target of the slide relationship `id`, when it has one.
fn target(package: &OpcPackage, slide: &PackURI, id: &str) -> Result<Option<PackURI>> {
    let part = package.get_part(slide)?;
    match part.rels().get(id) {
        Some(relationship) if !relationship.is_external() => {
            Ok(Some(relationship.target_partname()?))
        },
        _ => Ok(None),
    }
}

/// How many relationships of the package name `part` as their target.
fn references(package: &OpcPackage, part: &PackURI) -> usize {
    let named = |relationship: &litchi_opc::Relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|target| target == *part)
    };
    package.rels().iter().filter(|r| named(r)).count()
        + package
            .iter_parts()
            .map(|owner| owner.rels().iter().filter(|r| named(r)).count())
            .sum::<usize>()
}

/// The element ranges of the VML shapes in `vml` whose `id` (or `o:spid`)
/// is one of `spids`, and how many shapes the drawing holds.
fn vml_shapes(vml: &[u8], spids: &[&str]) -> Result<(Vec<Range<usize>>, usize)> {
    let mut reader = Reader::from_reader(vml);
    reader.config_mut().check_end_names = true;
    let unreadable = |_error| {
        refused(
            "remove_ole_equation",
            "the legacy VML drawing of the equation cannot be read",
        )
    };
    let mut depth = 0usize;
    // The open matching shape: its start and depth.
    let mut open: Option<(usize, usize)> = None;
    let mut found = Vec::new();
    let mut shapes = 0usize;
    loop {
        let start = usize::try_from(reader.buffer_position()).map_err(|_e| invalid("VML size"))?;
        let event = reader.read_event().map_err(unreadable)?;
        let end = usize::try_from(reader.buffer_position()).map_err(|_e| invalid("VML size"))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = matches!(event, Event::Empty(_));
                if element.local_name().as_ref() == b"shape" {
                    shapes += 1;
                    let matches = element.checked_attributes().flatten().any(|attribute| {
                        let key = attribute.key.as_ref();
                        (key == b"id" || attribute.key.local_name().as_ref() == b"spid")
                            && std::str::from_utf8(&attribute.value)
                                .is_ok_and(|value| spids.contains(&value.trim()))
                    });
                    if matches && open.is_none() {
                        if empty {
                            found.push(start..end);
                        } else {
                            open = Some((start, depth));
                        }
                    }
                }
                if !empty {
                    depth += 1;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced VML"))?;
                if let Some((at, level)) = open
                    && level == depth
                {
                    found.push(at..end);
                    open = None;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok((found, shapes))
}

/// Relationship IDs the attributes of `xml` name, among `ids`.
fn named_ids(xml: &[u8], ids: &BTreeSet<String>) -> Result<BTreeSet<String>> {
    let mut reader = Reader::from_reader(xml);
    let mut out = BTreeSet::new();
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(ref element) | Event::Empty(ref element) => {
                for attribute in element.checked_attributes().flatten() {
                    if let Ok(value) = std::str::from_utf8(&attribute.value)
                        && ids.contains(value.trim())
                    {
                        out.insert(value.trim().to_owned());
                    }
                }
            },
            Event::Eof => return Ok(out),
            _ => {},
        }
    }
}

/// Remove the legacy VML shapes `spids` from the VML drawings `slide`
/// relates in `candidate`. A drawing left without shapes leaves the slide
/// (unless the slide XML `next` still names it); otherwise only the shapes
/// and the relationships they alone used leave. Returns the parts to remove
/// when nothing else uses them.
fn release_vml_shapes(
    candidate: &mut OpcPackage,
    slide: &PackURI,
    next: &[u8],
    spids: &[&str],
) -> Result<Vec<PackURI>> {
    let mut roots = Vec::new();
    if spids.is_empty() {
        return Ok(roots);
    }
    let drawings: Vec<(String, PackURI)> = candidate
        .get_part(slide)?
        .rels()
        .iter()
        .filter(|relationship| {
            !relationship.is_external() && relationship.reltype().ends_with("/vmlDrawing")
        })
        .map(|relationship| {
            Ok((
                relationship.r_id().to_owned(),
                relationship.target_partname()?,
            ))
        })
        .collect::<Result<_>>()?;
    let owner_rels = candidate.get_part(slide)?.rels().clone();
    let named = shape_relationship_ids(next, &owner_rels)?;
    for (id, part) in drawings {
        let vml = match candidate.get_part(&part) {
            Ok(vml) => vml.blob().to_vec(),
            Err(litchi_opc::OpcError::PartNotFound(_)) => continue,
            Err(error) => return Err(error.into()),
        };
        let (found, shapes) = vml_shapes(&vml, spids)?;
        if found.is_empty() {
            continue;
        }
        // A drawing another owner shares keeps its shapes.
        if references(candidate, &part) > 1 {
            continue;
        }
        if found.len() == shapes && !named.contains(&id) {
            if candidate
                .get_part_mut(slide)?
                .rels_mut()
                .remove(&id)
                .is_some()
            {
                roots.push(part);
            }
            continue;
        }
        let mut splice = Splice::default();
        for range in found {
            splice.replace(range, String::new());
        }
        let kept = splice.apply(&vml, usize::MAX)?;
        let drawing = candidate.get_part(&part)?;
        let ids: BTreeSet<String> = drawing
            .rels()
            .iter()
            .map(|relationship| relationship.r_id().to_owned())
            .collect();
        let before = named_ids(&vml, &ids)?;
        let after = named_ids(&kept, &ids)?;
        let drawing = candidate.get_part_mut(&part)?;
        drawing.set_blob(kept);
        for released in before.difference(&after) {
            if let Some(relationship) = drawing.rels_mut().remove(released)
                && !relationship.is_external()
            {
                roots.push(relationship.target_partname()?);
            }
        }
    }
    Ok(roots)
}

/// Drop from the content-types manifest of `candidate` the `Default`
/// entries for the extensions of parts `before` had and `candidate` no
/// longer has, when no remaining part uses the extension: a removed
/// object's `bin`, `wmf` and `vml` declarations leave with it.
fn drop_unused_defaults(before: &OpcPackage, candidate: &mut OpcPackage) -> Result<()> {
    let remaining: BTreeSet<String> = candidate
        .iter_parts()
        .map(|part| part.partname().ext().to_ascii_lowercase())
        .collect();
    let dropped: BTreeSet<String> = before
        .iter_parts()
        .filter(|part| !candidate.contains_part(part.partname()))
        .map(|part| part.partname().ext().to_ascii_lowercase())
        .filter(|extension| !remaining.contains(extension) && extension != "rels")
        .collect();
    if dropped.is_empty() {
        return Ok(());
    }
    let limits = candidate.read_limits();
    let current = candidate.source_content_types(limits)?;
    let xml = current.bytes();
    let mut reader = Reader::from_reader(xml);
    let mut splice = Splice::default();
    let mut changed = false;
    loop {
        let start = usize::try_from(reader.buffer_position()).map_err(|_e| invalid("size"))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = usize::try_from(reader.buffer_position()).map_err(|_e| invalid("size"))?;
        match event {
            Event::Empty(ref element) if element.local_name().as_ref() == b"Default" => {
                let unused = element
                    .checked_attributes()
                    .flatten()
                    .find(|attribute| attribute.key.as_ref() == b"Extension")
                    .and_then(|attribute| {
                        std::str::from_utf8(&attribute.value)
                            .ok()
                            .map(str::to_ascii_lowercase)
                    })
                    .is_some_and(|extension| dropped.contains(&extension));
                if unused {
                    splice.replace(start..end, String::new());
                    changed = true;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if changed {
        let replacement = splice.apply(xml, usize::MAX)?;
        let expected = current.bytes().to_vec();
        candidate.try_replace_content_types_bytes(
            &expected,
            std::sync::Arc::new(replacement),
            limits,
        )?;
    }
    Ok(())
}

impl Edit {
    /// Stage `next` as the XML of `slide` in `candidate`, after checking
    /// that every shape but `shape` keeps its place and ID, that a
    /// `removed` frame is gone, and that the Office Math equations and
    /// `objects` OLE equations remain; the relationships `before` named and
    /// `next` does not are removed, and the legacy VML shapes `spids`, with
    /// the parts nothing else uses.
    #[allow(
        clippy::too_many_arguments,
        reason = "one staging step; the arguments are its inputs and its checks"
    )]
    fn publish_ole_equation_edit(
        &mut self,
        slide: &PackURI,
        mut candidate: super::Working,
        before: &[u8],
        next: Vec<u8>,
        objects: usize,
        shape: u32,
        removed: bool,
        spids: &[&str],
    ) -> Result<()> {
        let shape_limits = self.semantic_limits().shape();
        require_declared_prefixes(&next)?;
        if removed {
            crate::opened::xml::validate_shape_removal_dependencies(
                &next,
                &BTreeSet::from([shape]),
            )?;
        }
        let others = |xml: &[u8]| -> Result<Vec<Option<u32>>> {
            Ok(crate::shape::Scene::read(xml, shape_limits)?
                .iter()
                .map(crate::shape::Shape::id)
                .filter(|id| *id != Some(shape))
                .collect())
        };
        let gone = !crate::shape::Scene::read(&next, shape_limits)?
            .iter()
            .any(|staged| staged.id() == Some(shape));
        if others(before)? != others(&next)?
            || (removed && !gone)
            || ole::scan(&next)?.len() != objects
            || crate::shape::math::equations(&next)?.len()
                != crate::shape::math::equations(before)?.len()
        {
            return Err(invalid("the equation object edit did not round-trip"));
        }
        let owner = candidate.get_part(slide)?;
        let referenced = shape_relationship_ids(before, owner.rels())?;
        let retained = shape_relationship_ids(&next, owner.rels())?;
        let mut roots = release_vml_shapes(&mut candidate, slide, &next, spids)?;
        {
            let owner = candidate.get_part_mut(slide)?;
            owner.set_blob(next);
            for id in referenced.difference(&retained) {
                if let Some(relationship) = owner.rels_mut().remove(id)
                    && !relationship.is_external()
                {
                    roots.push(relationship.target_partname()?);
                }
            }
        }
        remove_unreferenced_dependencies(&mut candidate, roots)?;
        drop_unused_defaults(&self.working, &mut candidate)?;
        self.working = candidate;
        self.note_scene_read(slide);
        Ok(())
    }

    /// Add a MathType equation object (`Equation.DSMT4`, MTEF 5) at
    /// `origin` (`x`, `y` in EMU): an embedding part
    /// (`/ppt/embeddings/oleObjectN.bin`, the compound file Word and
    /// PowerPoint write for MathType), a PNG preview part and the frame, in
    /// the plain form of ECMA-376 (see the module documentation). The frame
    /// is the size of the equation drawn at `font_size_pt` (18 when
    /// `None`). Raw MTEF 3 input makes an Equation Editor 3.0 object
    /// (`Equation.3`), as its version requires.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Equation`] for input that does not convert or
    /// render, a refusal for MTEF without a header, and an error for a
    /// missing slide, malformed slide XML or exceeded bounds. A failure
    /// leaves the edit unchanged.
    pub fn add_ole_equation<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        source: EquationSource<'_>,
        origin: (i64, i64),
        font_size_pt: Option<f64>,
    ) -> Result<EquationWrite> {
        let selected = self.resolve_slide(slide.into())?;
        let max_patch_bytes = self.source_ref()?.limits.max_patch_bytes();
        let shape_limits = self.semantic_limits().shape();
        let font_size = font_size_pt.unwrap_or(DEFAULT_FONT_SIZE_PT);
        if !(font_size.is_finite() && (1.0..=4000.0).contains(&font_size)) {
            return Err(invalid("the equation text size must be 1 to 4000 points"));
        }
        let xml = self.working.get_part(&selected.part_name)?.blob().to_vec();
        let conformance = crate::media_parts::document_conformance_with_limits(
            &xml,
            self.semantic_limits()
                .media_parts()
                .constrained_by_source(self.working.read_limits()),
        )?;
        let (pml, dml, rel) = conformance_namespaces(conformance);
        let (uri, object_relationship) = match conformance {
            crate::media_parts::Conformance::Transitional => (
                OLE_URI,
                litchi_opc::constants::relationship_type::OLE_OBJECT,
            ),
            crate::media_parts::Conformance::Strict => (
                STRICT_OLE_URI,
                litchi_opc::constants::relationship_type::STRICT_OLE_OBJECT,
            ),
        };
        let (native, target, losses) = equation_native(source, None, max_patch_bytes)?;
        let object = litchi_equation::equation_object(&target, &native)?;
        let natural = measure(&native)?;
        let scale = font_size / MEASURE_POINTS;
        let extent = (emu(natural.0 * scale)?, emu(natural.1 * scale)?);
        let png = preview(&native, natural, extent)?;
        let (x, y) = origin;
        if x.checked_add(extent.0).is_none() || y.checked_add(extent.1).is_none() {
            return Err(invalid(
                "the equation frame lies outside the coordinate space",
            ));
        }
        let shape_id = max_raw_shape_id(&xml)?
            .max(max_shape_reference_id(&xml)?.unwrap_or(1))
            .checked_add(1)
            .ok_or_else(|| invalid("opened-presentation equation shape ID overflow"))?;
        let before = ole::scan(&xml)?.len();
        let mut candidate = self.working.clone();
        let object_part = free_part_name(&candidate, "/ppt/embeddings/oleObject", ".bin")?;
        candidate.try_add_part(Box::new(BlobPart::new(
            object_part.clone(),
            OLE_OBJECT_CONTENT_TYPE.to_owned(),
            object,
        )))?;
        let image_part = free_part_name(&candidate, "/ppt/media/image", ".png")?;
        candidate.try_add_part(Box::new(BlobPart::new(
            image_part.clone(),
            "image/png".to_owned(),
            png,
        )))?;
        let base = selected.part_name.base_uri();
        let object_id = candidate
            .get_part_mut(&selected.part_name)?
            .relate_to(&object_part.relative_ref(base), object_relationship);
        let image_id = candidate.get_part_mut(&selected.part_name)?.relate_to(
            &image_part.relative_ref(base),
            conformance_image_relationship(conformance),
        );
        let name = format!("Object {}", shape_id - 1);
        let (cx, cy) = extent;
        let fragment = format!(
            concat!(
                r#"<p:graphicFrame xmlns:p="{pml}" xmlns:a="{dml}" xmlns:r="{rel}">"#,
                r#"<p:nvGraphicFramePr><p:cNvPr id="{id}" name="{name}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>"#,
                r#"<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm>"#,
                r#"<a:graphic><a:graphicData uri="{uri}">"#,
                r#"<p:oleObj name="Equation" r:id="{object_id}" imgW="{cx}" imgH="{cy}" progId="{prog_id}"><p:embed/>"#,
                r#"<p:pic><p:nvPicPr><p:cNvPr id="0" name="{name}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>"#,
                r#"<p:blipFill><a:blip r:embed="{image_id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>"#,
                r#"<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#,
                "</p:oleObj></a:graphicData></a:graphic></p:graphicFrame>",
            ),
            pml = pml,
            dml = dml,
            rel = rel,
            id = shape_id,
            name = name,
            x = x,
            y = y,
            cx = cx,
            cy = cy,
            uri = uri,
            object_id = object_id,
            image_id = image_id,
            prog_id = target.prog_id(),
        );
        let next = crate::opened::xml::append_shape_with_limit(
            &xml,
            fragment.as_bytes(),
            shape_limits.output_bytes(),
        )?;
        ensure_shape_id(&next, shape_id, "equation", shape_limits)?;
        let after = ole::scan(&next)?;
        let index = after
            .iter()
            .position(|site| site.shape_id == Some(shape_id))
            .filter(|_| after.len() == before + 1)
            .ok_or_else(|| invalid("the equation object did not round-trip"))?;
        require_declared_prefixes(&next)?;
        candidate.get_part_mut(&selected.part_name)?.set_blob(next);
        self.working = candidate;
        self.note_scene_read(&selected.part_name);
        Ok(EquationWrite {
            shape_id,
            index,
            losses,
        })
    }

    /// Replace the content of the slide's OLE equation at `equation` (its
    /// position among
    /// [`Slide::ole_equations`](crate::slide::Slide::ole_equations)): the
    /// `Equation Native` stream of its compound file, written in the
    /// object's kind (its MTEF version, or its ProgID's when the old stream
    /// cannot be read), with every other stream kept and cached
    /// presentations dropped. An embedding that is no equation compound
    /// file is replaced by a new one. The preview is redrawn (in place when
    /// it is a PNG nothing else uses) and the frame keeps its position and
    /// scales with the equation. PowerPoint's markup-compatibility form
    /// becomes the plain form, an object without a preview picture gets
    /// one, and no legacy VML shape of the object is left. Returns `None`
    /// when the stream is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing equation, a
    /// refusal for a frame in a markup-compatibility block of its own, one
    /// without an embedded part, MTEF of another version than the object's,
    /// or an unreadable legacy VML drawing, [`Error::Equation`] for input
    /// that does not convert or render, and an error for malformed slide
    /// XML. A failure leaves the edit unchanged.
    pub fn set_ole_equation<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        equation: usize,
        source: EquationSource<'_>,
    ) -> Result<Option<EquationWrite>> {
        const OPERATION: &str = "set_ole_equation";
        let selected = self.resolve_slide(slide.into())?;
        let max_patch_bytes = self.source_ref()?.limits.max_patch_bytes();
        let shape_limits = self.semantic_limits().shape();
        let xml = self.working.get_part(&selected.part_name)?.blob().to_vec();
        let sites = ole::scan(&xml)?;
        let site = located(&sites, equation, OPERATION)?;
        let shape_id = site
            .shape_id
            .ok_or_else(|| invalid("the equation frame has no ID"))?;
        let embedding = match site.relationship() {
            Some(id) => target(&self.working, &selected.part_name, id)?,
            None => None,
        }
        .filter(|part| self.working.contains_part(part))
        .ok_or_else(|| refused(OPERATION, "the equation object has no embedded part"))?;
        let old_object = self.working.get_part(&embedding)?.blob().to_vec();
        let old_native = litchi_equation::equation_object_native(&old_object).ok();
        let limits = Limits::DEFAULT;
        let kind = old_native
            .as_deref()
            .and_then(|native| litchi_equation::read_mtef(native, &limits).ok())
            .and_then(|read| read.info().cloned())
            .map_or_else(
                || site.prog_id().and_then(MtefTarget::from_prog_id),
                |info| Some(MtefTarget::SameAs(info)),
            )
            .unwrap_or(MtefTarget::MathType);
        let (native, target_kind, losses) = equation_native(source, Some(kind), max_patch_bytes)?;
        if old_native.as_deref() == Some(native.as_slice()) {
            return Ok(None);
        }
        let object = if old_native.is_some() {
            litchi_equation::replace_equation_object_native(&old_object, &native)?
        } else {
            litchi_equation::equation_object(&target_kind, &native)?
        };
        let natural = measure(&native)?;
        let old_size = old_native.as_deref().and_then(|old| measure(old).ok());
        let extent = match site.extent {
            Some(extent) if extent.0 > 0 && extent.1 > 0 => rescaled(extent, old_size, natural)?,
            _ => (emu(natural.0)?, emu(natural.1)?),
        };
        let png = preview(&native, natural, extent)?;
        let kept = kept_object(&site)
            .cloned()
            .ok_or_else(|| refused(OPERATION, "the equation object has no picture fallback"))?;

        let conformance = crate::media_parts::document_conformance_with_limits(
            &xml,
            self.semantic_limits()
                .media_parts()
                .constrained_by_source(self.working.read_limits()),
        )?;
        let mut candidate = self.working.clone();
        candidate.get_part_mut(&embedding)?.set_blob(object);
        // The preview: replaced in place when it is a PNG only this frame
        // shows, otherwise a new part the blips name.
        let preview_id = kept.blips.iter().find_map(|(_, embed)| embed.clone());
        let preview_part = match &preview_id {
            Some(id) => target(&candidate, &selected.part_name, id)?,
            None => None,
        };
        let outside = {
            let mut rest = xml[..site.range.start].to_vec();
            rest.extend_from_slice(&xml[site.range.end..]);
            rest
        };
        let owner_rels = candidate.get_part(&selected.part_name)?.rels().clone();
        let in_place = match (&preview_id, &preview_part) {
            (Some(id), Some(part)) => {
                candidate
                    .get_part(part)
                    .is_ok_and(|image| image.content_type() == "image/png")
                    && references(&candidate, part) == 1
                    && !shape_relationship_ids(&outside, &owner_rels)?.contains(id)
            },
            _ => false,
        };
        let new_blip = if in_place {
            if let Some(part) = &preview_part {
                candidate.get_part_mut(part)?.set_blob(png);
            }
            None
        } else {
            let image_part = free_part_name(&candidate, "/ppt/media/image", ".png")?;
            candidate.try_add_part(Box::new(BlobPart::new(
                image_part.clone(),
                "image/png".to_owned(),
                png,
            )))?;
            Some(candidate.get_part_mut(&selected.part_name)?.relate_to(
                &image_part.relative_ref(selected.part_name.base_uri()),
                conformance_image_relationship(conformance),
            ))
        };

        // The kept object's markup: its size attributes and pictures, and
        // without a picture a new `p:pic` (the form PowerPoint 2007 writes
        // names only a VML shape, which would keep showing the old
        // equation).
        let mut inner = Splice::default();
        let (cx, cy) = (extent.0.to_string(), extent.1.to_string());
        let local =
            |range: &Range<usize>| range.start - kept.element.start..range.end - kept.element.start;
        let element = &xml[kept.element.clone()];
        let tag = local(&kept.tag);
        set_attribute(&mut inner, element, &tag, b"imgW", &cx);
        set_attribute(&mut inner, element, &tag, b"imgH", &cy);
        if let Some((_, _, span)) = attribute_value(&element[tag.clone()], b"spid") {
            let start = if span.start > 0 && element[span.start - 1].is_ascii_whitespace() {
                span.start - 1
            } else {
                span.start
            };
            inner.replace(tag.start + start..tag.start + span.end, String::new());
        }
        for (blip, _) in &kept.blips {
            if let Some(id) = &new_blip {
                set_attribute(&mut inner, element, &local(blip), b"embed", id);
            }
        }
        if kept.blips.is_empty()
            && let Some(id) = &new_blip
        {
            let name = tag_name(&element[tag.clone()])?;
            let prefix = name
                .split_once(':')
                .map_or(String::new(), |(prefix, _)| format!("{prefix}:"));
            let (x, y) = site.offset.unwrap_or((0, 0));
            let (_, dml, rel) = conformance_namespaces(conformance);
            let picture = format!(
                concat!(
                    r#"<{p}pic xmlns:a="{dml}" xmlns:r="{rel}"><{p}nvPicPr><{p}cNvPr id="0" name=""/><{p}cNvPicPr/><{p}nvPr/></{p}nvPicPr>"#,
                    r#"<{p}blipFill><a:blip r:embed="{id}"/><a:stretch><a:fillRect/></a:stretch></{p}blipFill>"#,
                    r#"<{p}spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></{p}spPr></{p}pic>"#,
                ),
                p = prefix,
                dml = dml,
                rel = rel,
                id = id,
                x = x,
                y = y,
                cx = cx,
                cy = cy,
            );
            if tag.end == element.len() {
                // An empty element opens to take the picture.
                let open = std::str::from_utf8(&element[tag.clone()])
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .trim_end_matches('>')
                    .trim_end()
                    .trim_end_matches('/')
                    .to_owned();
                inner.replace(tag.clone(), format!("{open}>{picture}</{name}>"));
            } else {
                let close = element.len() - format!("</{name}>").len();
                inner.replace(close..close, picture);
            }
        }
        for ext in &site.extent_tags {
            if kept.element.start <= ext.start && ext.end <= kept.element.end {
                resize(&mut inner, element, &local(ext), extent)?;
            }
        }
        let kept_xml = String::from_utf8(inner.apply(element, usize::MAX)?)
            .map_err(|error| Error::Xml(error.to_string()))?;

        let mut splice = Splice::default();
        match &site.block {
            Some(block) => {
                // The plain form: the Fallback object stands in for the
                // block, with the declarations it relied on.
                let unwrapped =
                    carry_declarations(&xml[block.start..kept.element.start], kept_xml.as_bytes())?;
                splice.replace(block.clone(), unwrapped);
            },
            None => splice.replace(kept.element.clone(), kept_xml),
        }
        // No legacy VML shape of the object survives.
        let spids = site.spids();
        for ext in &site.extent_tags {
            if !(kept.element.start <= ext.start && ext.end <= kept.element.end)
                && site
                    .block
                    .as_ref()
                    .is_none_or(|block| !(block.start <= ext.start && ext.end <= block.end))
            {
                resize(&mut splice, &xml, ext, extent)?;
            }
        }
        let next = splice.apply(&xml, shape_limits.output_bytes())?;
        self.publish_ole_equation_edit(
            &selected.part_name,
            candidate,
            &xml,
            next,
            sites.len(),
            shape_id,
            false,
            &spids,
        )?;
        Ok(Some(EquationWrite {
            shape_id,
            index: equation,
            losses,
        }))
    }

    /// Remove the slide's OLE equation at `equation` (its position among
    /// [`Slide::ole_equations`](crate::slide::Slide::ole_equations)): its
    /// frame, its embedding and preview parts and their relationships when
    /// nothing else uses them, and its legacy VML shape, with the VML
    /// drawing part when no other shape is left in it. Returns the frame's
    /// shape ID.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing equation, a
    /// refusal for a frame in a markup-compatibility block of its own, a
    /// frame other markup refers to (connectors, animations) or an
    /// unreadable legacy VML drawing, and an error for malformed slide XML.
    /// A failure leaves the edit unchanged.
    pub fn remove_ole_equation<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        equation: usize,
    ) -> Result<u32> {
        let selected = self.resolve_slide(slide.into())?;
        let shape_limits = self.semantic_limits().shape();
        let xml = self.working.get_part(&selected.part_name)?.blob().to_vec();
        let sites = ole::scan(&xml)?;
        let site = located(&sites, equation, "remove_ole_equation")?;
        let shape_id = site
            .shape_id
            .ok_or_else(|| invalid("the equation frame has no ID"))?;
        let mut splice = Splice::default();
        splice.replace(site.range.clone(), String::new());
        let next = splice.apply(&xml, shape_limits.output_bytes())?;
        let candidate = self.working.clone();
        self.publish_ole_equation_edit(
            &selected.part_name,
            candidate,
            &xml,
            next,
            sites.len() - 1,
            shape_id,
            true,
            &site.spids(),
        )?;
        Ok(shape_id)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn vml_shapes_are_found_by_id() {
        let vml = br#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="o"><o:shapelayout/><v:shapetype id="t"/><v:shape id="a"><v:imagedata o:relid="rId1"/></v:shape><v:shape o:spid="b"/></xml>"#;
        let (found, shapes) = vml_shapes(vml, &["a", "b"]).unwrap();
        assert_eq!(shapes, 2);
        assert_eq!(found.len(), 2);
        assert!(vml[found[0].clone()].starts_with(br#"<v:shape id="a">"#));
        assert!(vml[found[0].clone()].ends_with(b"</v:shape>"));
        assert_eq!(&vml[found[1].clone()], br#"<v:shape o:spid="b"/>"#);
        let (found, shapes) = vml_shapes(vml, &["c"]).unwrap();
        assert!(found.is_empty());
        assert_eq!(shapes, 2);
        assert!(vml_shapes(b"<xml><v:shape id=\"a\"><br></v:shape></xml>", &["a"]).is_err());
    }

    #[test]
    fn frames_scale_with_the_equation() {
        assert_eq!(
            rescaled((1000, 500), Some((10.0, 5.0)), (20.0, 5.0)).unwrap(),
            (2000, 500)
        );
        // An unmeasured old equation keeps the height.
        assert_eq!(
            rescaled((1000, 500), None, (20.0, 5.0)).unwrap(),
            (2000, 500)
        );
        assert!(emu(f64::NAN).is_err());
        assert!(emu(1e12).is_err());
    }

    #[test]
    fn bare_mtef_gets_the_ole_header() {
        let native =
            litchi_equation::latex_to_equation_native("x", &MtefTarget::MathType, &Limits::DEFAULT)
                .unwrap();
        let bare = &native.bytes()[OLE_HEADER_LEN..];
        let (written, kind, losses) =
            equation_native(EquationSource::Mtef(bare), None, usize::MAX).unwrap();
        assert_eq!(&written[OLE_HEADER_LEN..], bare);
        assert_eq!(kind.prog_id(), "Equation.DSMT4");
        assert!(losses.is_empty());
        // MTEF 3 is refused for a MathType object.
        let v3 = litchi_equation::latex_to_equation_native(
            "x",
            &MtefTarget::EquationEditor3,
            &Limits::DEFAULT,
        )
        .unwrap();
        assert!(
            equation_native(
                EquationSource::Mtef(v3.bytes()),
                Some(MtefTarget::MathType),
                usize::MAX
            )
            .is_err()
        );
        assert!(equation_native(EquationSource::Mtef(b"junk"), None, usize::MAX).is_err());
        assert!(equation_native(EquationSource::Latex("x"), None, 0).is_err());
    }
}
