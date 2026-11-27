//! Staged, failure-atomic edits of worksheet OLE equation objects.

use std::collections::HashSet;
use std::ops::Range;

use litchi_core::edit::RefusalKind;
use litchi_equation::{Limits, MtefTarget, PngOptions};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use litchi_spreadsheet_drawing::{objects, vml};

use super::markup::{self, Collection, EMU_PER_PIXEL, Grid, MAX_OBJECTS, PICTURE_SHAPE_TYPE};
use super::{EquationObject, list};
use crate::Package;
use crate::error::{Result, invalid};
use crate::ole_objects::{OleObject, OleObjectMarker, parse_ole_objects};
use crate::part_edit::{Staged, WorksheetLayout, family};

family!("equation object");

/// The largest shape id Office supports for an OLE object.
const MAX_SHAPE_ID: u32 = 67_098_623;

/// Where an equation's content comes from.
#[derive(Debug, Clone, Copy)]
pub enum Source<'a> {
    /// LaTeX math.
    Latex(&'a str),
    /// An OMML equation (`m:oMath` or `m:oMathPara`).
    Omml(&'a str),
    /// MTEF data: a complete `Equation Native` stream, or bare MTEF that is
    /// given the 28-byte `EQNOLEFILEHDR`.
    Mtef(&'a [u8]),
}

/// What [`Edit::add`] or [`Edit::set`] wrote.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Written {
    shape_id: u32,
    losses: Vec<String>,
}

impl Written {
    /// The object's shape id.
    #[must_use]
    pub const fn shape_id(&self) -> u32 {
        self.shape_id
    }

    /// What the conversion to MTEF could not carry over; empty for MTEF
    /// input and for a lossless conversion.
    #[must_use]
    pub fn losses(&self) -> &[String] {
        &self.losses
    }
}

/// A staged equation-object edit of one [`Package`]. Dropping it discards
/// it.
#[derive(Debug)]
pub struct Edit {
    staged: Staged,
}

/// The part name of the worksheet at catalog position `sheet`.
pub(super) fn worksheet_part(package: &Package, sheet: usize) -> Result<PackURI> {
    let worksheet = package
        .workbook_ref()
        .sheet(sheet)?
        .ok_or_else(|| invalid(format!("worksheet {sheet} does not exist")))?;
    if worksheet.kind() != crate::workbook::WorksheetKind::Worksheet {
        return Err(Error::NotWorksheet {
            sheet: worksheet.name().to_owned(),
        });
    }
    Ok(worksheet.part_uri().clone())
}

impl Package {
    /// Start an equation-object edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused.
    pub fn edit_equation_objects(&self) -> Result<Edit> {
        self.edit_equation_objects_with_options(crate::EditOptions::default())
    }

    /// Start an equation-object edit with an explicit retained-patch budget.
    ///
    /// # Errors
    ///
    /// As [`Self::edit_equation_objects`].
    pub fn edit_equation_objects_with_options(&self, options: crate::EditOptions) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::new_with_options(self, "edit_equation_objects", options)?,
        })
    }

    /// Apply an equation-object patch to a new package; this one is
    /// unchanged. An empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_equation_objects(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_equation_objects")?,
            patch.0.clone(),
        )))
    }
}

fn refused(kind: RefusalKind, operation: &'static str, reason: &'static str) -> Error {
    Error::refused(kind, operation, reason)
}

/// Bare MTEF data behind an `EQNOLEFILEHDR` (MTEF specification, "OLE
/// equation objects"): header size 28, version 0x00020000, clipboard
/// format 0xC295, the data length and 16 reserved bytes.
fn with_ole_header(mtef: &[u8]) -> Result<Vec<u8>> {
    let length = u32::try_from(mtef.len()).map_err(|_| {
        refused(
            RefusalKind::InvalidValue,
            "equation_object_source",
            "the MTEF data is too large",
        )
    })?;
    let mut out = Vec::with_capacity(28 + mtef.len());
    out.extend_from_slice(&28u16.to_le_bytes());
    out.extend_from_slice(&0x0002_0000u32.to_le_bytes());
    out.extend_from_slice(&0xC295u16.to_le_bytes());
    out.extend_from_slice(&length.to_le_bytes());
    out.extend_from_slice(&[0; 16]);
    out.extend_from_slice(mtef);
    Ok(out)
}

/// The `Equation Native` stream of a source written for `target`, the
/// kind of object it belongs in, and what the conversion lost. MTEF data
/// keeps its own kind (MTEF 3 is an Equation Editor 3.0 object).
fn native_of(
    source: Source<'_>,
    target: &MtefTarget,
    limits: &Limits,
) -> Result<(Vec<u8>, MtefTarget, Vec<String>)> {
    let native = match source {
        Source::Latex(latex) => litchi_equation::latex_to_equation_native(latex, target, limits)?,
        Source::Omml(omml) => litchi_equation::omml_to_equation_native(omml, target, limits)?,
        Source::Mtef(bytes) => {
            let read = litchi_equation::read_mtef(bytes, limits)?;
            let info = read.info().cloned().ok_or_else(|| {
                refused(
                    RefusalKind::InvalidValue,
                    "equation_object_source",
                    "the MTEF data is empty",
                )
            })?;
            let native = if info.ole_header.is_some() {
                bytes.to_vec()
            } else {
                with_ole_header(bytes)?
            };
            return Ok((native, MtefTarget::SameAs(info), Vec::new()));
        },
    };
    let losses = native.losses().to_vec();
    let kind = MtefTarget::SameAs(native.info().clone());
    Ok((native.into_bytes(), kind, losses))
}

/// The preview picture of an `Equation Native` stream, drawn at twice the
/// resolution it is shown at, and its size in points.
fn preview_of(native: &[u8], limits: &Limits) -> Result<(Vec<u8>, (f64, f64))> {
    let options = PngOptions::DEFAULT.with_scale(2.0)?;
    let png = litchi_equation::mtef_to_png(native, &options, limits)?;
    let size = png.size_pt();
    Ok((png.into_bytes(), size))
}

impl Edit {
    /// The staged equation objects of the worksheet at `sheet`.
    ///
    /// # Errors
    ///
    /// As [`Package::equation_objects`].
    pub fn equations(&self, sheet: usize) -> Result<Vec<EquationObject>> {
        let worksheet = self.staged.worksheet(sheet)?;
        list(self.staged.opc(), &worksheet)
    }

    fn equation(
        &self,
        sheet: usize,
        shape_id: u32,
        operation: &'static str,
    ) -> Result<EquationObject> {
        self.equations(sheet)?
            .into_iter()
            .find(|object| object.shape_id() == shape_id)
            .ok_or_else(|| {
                refused(
                    RefusalKind::NotFound,
                    operation,
                    "the worksheet has no equation object with this shape id",
                )
            })
    }

    /// Add an equation object whose top-left corner is the top-left corner
    /// of the zero-based cell (`row`, `column`), sized to its preview: a
    /// MathType object (`Equation.DSMT4`, MTEF 5) for LaTeX and OMML, and
    /// an object of the data's own kind for MTEF.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Equation`] when the source does not convert or
    /// render, a refusal for a cell outside the grid, a Strict worksheet or
    /// a signed workbook, and the errors of malformed worksheet, VML or
    /// drawing markup.
    pub fn add(
        &mut self,
        sheet: usize,
        row: u32,
        column: u32,
        source: Source<'_>,
        limits: &Limits,
    ) -> Result<Written> {
        if row > 1_048_575 || column > 16_383 {
            return Err(refused(
                RefusalKind::InvalidValue,
                "add_equation_object",
                "the cell is outside the worksheet grid",
            ));
        }
        let worksheet = self.staged.worksheet(sheet)?;
        let (native, kind, losses) = native_of(source, &MtefTarget::MathType, limits)?;
        let object = litchi_equation::equation_object(&kind, &native)?;
        let (png, size) = preview_of(&native, limits)?;
        crate::signature_policy::ensure_unsigned(self.staged.opc())?;
        let shape_id = self.staged.update(|opc| {
            add_object(
                opc,
                &worksheet,
                sheet,
                (row, column),
                kind.prog_id(),
                object,
                png,
                size,
            )
        })?;
        Ok(Written { shape_id, losses })
    }

    /// Replace the `Equation Native` stream of the object `shape_id`,
    /// keeping its kind (the MTEF version and product of its current
    /// stream, else its ProgID's kind), redraw its preview and resize it
    /// from its top-left corner. Other streams of the object are kept;
    /// cached presentations are dropped. An embedding that is not a
    /// readable compound file is replaced by a new object of the same kind.
    ///
    /// # Errors
    ///
    /// Returns a `NotFound` refusal when the worksheet has no such
    /// equation object, an `InvalidValue` refusal for MTEF data of the other
    /// kind, [`Error::Equation`] when the source does not convert or render,
    /// and the errors of malformed markup.
    pub fn set(
        &mut self,
        sheet: usize,
        shape_id: u32,
        source: Source<'_>,
        limits: &Limits,
    ) -> Result<Written> {
        let worksheet = self.staged.worksheet(sheet)?;
        let current = self.equation(sheet, shape_id, "set_equation_object")?;
        let embedding = current
            .embedding()
            .map(PackURI::new)
            .transpose()
            .map_err(|error| invalid(error.to_string()))?
            .ok_or_else(|| {
                refused(
                    RefusalKind::UnsupportedContent,
                    "set_equation_object",
                    "the equation object has no embedded part to replace",
                )
            })?;
        let target = current
            .native()
            .and_then(|native| litchi_equation::read_mtef(native, limits).ok())
            .and_then(|read| read.info().cloned())
            .map_or_else(|| current.kind().clone(), MtefTarget::SameAs);
        let (native, kind, losses) = native_of(source, &target, limits)?;
        if kind.prog_id() != target.prog_id() {
            return Err(refused(
                RefusalKind::InvalidValue,
                "set_equation_object",
                "the MTEF data is of another equation kind than the object",
            ));
        }
        let old = self.staged.opc().get_part(&embedding)?.blob().to_vec();
        let object = match litchi_equation::replace_equation_object_native(&old, &native) {
            Ok(object) => object,
            Err(litchi_equation::Error::Object(_)) => {
                litchi_equation::equation_object(&target, &native)?
            },
            Err(error) => return Err(error.into()),
        };
        let (png, size) = preview_of(&native, limits)?;
        crate::signature_policy::ensure_unsigned(self.staged.opc())?;
        self.staged
            .update(|opc| set_object(opc, &worksheet, shape_id, &embedding, object, png, size))?;
        Ok(Written { shape_id, losses })
    }

    /// Remove the object `shape_id`: its `<oleObjects>` entry, its VML and
    /// drawing shapes, its embedding and preview parts, and the
    /// relationships, the VML drawing, the drawing and the `<oleObjects>`
    /// element nothing else uses any more.
    ///
    /// # Errors
    ///
    /// Returns a `NotFound` refusal when the worksheet has no such
    /// equation object, and the errors of malformed markup.
    pub fn remove(&mut self, sheet: usize, shape_id: u32) -> Result<()> {
        let worksheet = self.staged.worksheet(sheet)?;
        self.equation(sheet, shape_id, "remove_equation_object")?;
        crate::signature_policy::ensure_unsigned(self.staged.opc())?;
        self.staged
            .update(|opc| remove_object(opc, &worksheet, shape_id))
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged package.
    ///
    /// # Errors
    ///
    /// Returns the errors of the retained-patch budget and of publication.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}

// ---------------------------------------------------------------------------
// Graph helpers

/// The first free part name `{prefix}{N}{suffix}`.
fn free_name(opc: &OpcPackage, prefix: &str, suffix: &str) -> Result<PackURI> {
    for index in 1_u32..=1_000_000 {
        let candidate = PackURI::new(format!("{prefix}{index}{suffix}"))
            .map_err(|error| invalid(error.to_string()))?;
        if opc.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free part name is left"))
}

/// Whether any relationship of the package still targets `part`.
fn referenced(opc: &OpcPackage, part: &PackURI) -> bool {
    let targets = |relationship: &litchi_opc::Relationship| {
        !relationship.is_external() && relationship.target_partname().ok().as_ref() == Some(part)
    };
    opc.rels().iter().any(targets)
        || opc
            .iter_parts()
            .any(|candidate| candidate.rels().iter().any(targets))
}

/// Remove each part of `parts` no relationship targets any more.
fn drop_orphans(opc: &mut OpcPackage, parts: impl IntoIterator<Item = PackURI>) {
    for part in parts {
        if !referenced(opc, &part) {
            opc.remove_part(&part);
        }
    }
}

/// The internal target of relationship `id` of `owner`.
fn target(opc: &OpcPackage, owner: &PackURI, id: &str) -> Result<Option<PackURI>> {
    let Some(relationship) = opc.get_part(owner)?.rels().get(id) else {
        return Ok(None);
    };
    if relationship.is_external() {
        return Ok(None);
    }
    Ok(Some(relationship.target_partname()?))
}

/// The worksheet's VML drawing: its relationship id and part name.
fn legacy_drawing(
    opc: &OpcPackage,
    worksheet: &PackURI,
    layout: &WorksheetLayout,
) -> Result<Option<(String, PackURI)>> {
    let Some(child) = layout.child("legacyDrawing") else {
        return Ok(None);
    };
    let id = child
        .relationship_id
        .clone()
        .ok_or_else(|| invalid("legacyDrawing has no relationship id"))?;
    let part = target(opc, worksheet, &id)?
        .ok_or_else(|| invalid("the legacyDrawing relationship is missing or external"))?;
    Ok(Some((id, part)))
}

fn vml_limits(opc: &OpcPackage) -> vml::Limits {
    vml::Limits::from_source_limits(opc.read_limits())
}

fn drawing_limits(opc: &OpcPackage) -> objects::Limits {
    objects::Limits::from_source_limits(opc.read_limits())
}

/// The span and shape id of each shape of a VML drawing.
fn vml_shapes(opc: &OpcPackage, xml: &[u8]) -> Result<Vec<(Range<usize>, Option<u32>)>> {
    Ok(vml::shapes(xml, &vml_limits(opc))?
        .into_iter()
        .map(|shape| {
            let id = markup::vml_shape_id(&xml[shape.span.clone()]);
            (shape.span, id)
        })
        .collect())
}

/// The worksheet's `<oleObjects>` entries.
fn collection(xml: &[u8], layout: &WorksheetLayout) -> Result<Option<Collection>> {
    layout
        .child("oleObjects")
        .map(|child| Collection::scan(xml, child.span.clone()))
        .transpose()
}

/// The parsed object `shape_id` and its siblings.
fn parsed(xml: &[u8], shape_id: u32) -> Result<(OleObject, Vec<OleObject>)> {
    let objects = parse_ole_objects(xml)?
        .ok_or_else(|| invalid("the worksheet has no oleObjects"))?
        .objects;
    let (mine, others): (Vec<_>, Vec<_>) = objects
        .into_iter()
        .partition(|object| object.shape_id == shape_id);
    let object = mine
        .into_iter()
        .next()
        .ok_or_else(|| invalid("the equation object vanished from oleObjects"))?;
    Ok((object, others))
}

// ---------------------------------------------------------------------------
// Add

#[allow(
    clippy::too_many_arguments,
    reason = "one staged operation's inputs, prepared before the graph is touched"
)]
fn add_object(
    opc: &mut OpcPackage,
    worksheet: &PackURI,
    sheet: usize,
    (row, column): (u32, u32),
    prog_id: &str,
    object: Vec<u8>,
    png: Vec<u8>,
    size: (f64, f64),
) -> Result<u32> {
    let mut xml = opc.get_part(worksheet)?.blob().to_vec();
    let layout = WorksheetLayout::scan(&xml)?;
    if layout.strict {
        return Err(refused(
            RefusalKind::UnsupportedContent,
            "add_equation_object",
            "equation objects are written to Transitional worksheets only",
        ));
    }
    let existing = parse_ole_objects(&xml)?
        .map(|objects| objects.objects)
        .unwrap_or_default();
    if existing.len() >= MAX_OBJECTS {
        return Err(refused(
            RefusalKind::InvalidValue,
            "add_equation_object",
            "the worksheet holds the most OLE objects it may",
        ));
    }
    let vml = legacy_drawing(opc, worksheet, &layout)?;
    let drawing = objects::drawing_of(opc.get_part(worksheet)?.rels())?;

    // A shape id no VML drawing of the workbook, worksheet object or drawing
    // shape uses (`vml_ids`): in the VML drawing's own block when the sheet
    // has one, else in a new block past every block a VML drawing claims.
    let mut taken: HashSet<u32> = existing.iter().map(|object| object.shape_id).collect();
    if let Some(name) = &drawing {
        taken.extend(markup::scan_drawing(opc.get_part(name)?.blob(), None)?.ids);
    }
    let allocation =
        crate::vml_ids::allocate(opc, vml.as_ref().map(|(_, name)| name), sheet, &taken)?;
    let (next, block) = (allocation.id, allocation.block);
    if next > MAX_SHAPE_ID {
        return Err(refused(
            RefusalKind::InvalidValue,
            "add_equation_object",
            "no shape id is left for a new object",
        ));
    }
    let shape_id = next;

    let frame = Grid::read(&xml, row)?.frame(row, column, size);

    // The embedding and the preview.
    let embedding = free_name(opc, "/xl/embeddings/oleObject", ".bin")?;
    opc.add_part(Box::new(BlobPart::new(
        embedding.clone(),
        ct::OFC_OLE_OBJECT.to_owned(),
        object,
    )));
    let image = free_name(opc, "/xl/media/image", ".png")?;
    opc.add_part(Box::new(BlobPart::new(
        image.clone(),
        ct::PNG.to_owned(),
        png,
    )));
    let base = worksheet.base_uri();
    let part = opc.get_part_mut(worksheet)?;
    let object_id = part.relate_to(&embedding.relative_ref(base), rt::OLE_OBJECT);
    let preview_id = part.relate_to(&image.relative_ref(base), rt::IMAGE);

    // The VML picture shape.
    match vml {
        Some((_, name)) => {
            let relationship = opc
                .get_part_mut(&name)?
                .relate_to(&image.relative_ref(name.base_uri()), rt::IMAGE);
            let vml_xml = opc.get_part(&name)?.blob();
            let z = vml::shapes(vml_xml, &vml_limits(opc))?.len() + 1;
            let close = vml::rfind(vml_xml, b"</xml>")
                .ok_or_else(|| invalid("the VML drawing has no closing xml element"))?;
            let mut fragment = String::new();
            if !vml::contains(vml_xml, b"\"_x0000_t75\"")
                && !vml::contains(vml_xml, b"'_x0000_t75'")
            {
                fragment.push_str(PICTURE_SHAPE_TYPE);
            }
            fragment.push_str(&markup::vml_shape(shape_id, &relationship, &frame, z));
            let mut out = vml_xml.to_vec();
            out.splice(close..close, fragment.into_bytes());
            opc.get_part_mut(&name)?.set_blob(out);
        },
        None => {
            let name = free_name(opc, "/xl/drawings/vmlDrawing", ".vml")?;
            opc.add_part(Box::new(BlobPart::new(
                name.clone(),
                ct::OFC_VML_DRAWING.to_owned(),
                Vec::new(),
            )));
            let relationship = opc
                .get_part_mut(&name)?
                .relate_to(&image.relative_ref(name.base_uri()), rt::IMAGE);
            let shape = markup::vml_shape(shape_id, &relationship, &frame, 1);
            opc.get_part_mut(&name)?
                .set_blob(markup::new_vml_drawing(block, &shape).into_bytes());
            let id = opc
                .get_part_mut(worksheet)?
                .relate_to(&name.relative_ref(base), rt::VML_DRAWING);
            let layout = WorksheetLayout::scan(&xml)?;
            xml = layout.insert_relationship_child(&xml, "legacyDrawing", &id)?;
        },
    }

    // The hidden drawing shape.
    let fragment = markup::drawing_anchor(shape_id, &frame.anchor);
    let limits = drawing_limits(opc);
    match drawing {
        Some(name) => {
            let out = objects::append(opc.get_part(&name)?.blob(), &fragment, &limits)?;
            opc.get_part_mut(&name)?.set_blob(out);
        },
        None => {
            let name = free_name(opc, "/xl/drawings/drawing", ".xml")?;
            let out = objects::append(objects::EMPTY_DRAWING.as_bytes(), &fragment, &limits)?;
            opc.add_part(Box::new(BlobPart::new(
                name.clone(),
                ct::OFC_DRAWING.to_owned(),
                out,
            )));
            let id = opc
                .get_part_mut(worksheet)?
                .relate_to(&name.relative_ref(base), rt::DRAWING);
            let layout = WorksheetLayout::scan(&xml)?;
            xml = layout.insert_relationship_child(&xml, "drawing", &id)?;
        },
    }

    // The worksheet entry.
    let layout = WorksheetLayout::scan(&xml)?;
    let entry = markup::worksheet_entry(
        &layout.prefix,
        prog_id,
        shape_id,
        &object_id,
        &preview_id,
        &frame.anchor,
    );
    xml = match collection(&xml, &layout)? {
        Some(Collection {
            close: Some(close), ..
        }) => markup::splice(&xml, vec![(close..close, entry)]),
        Some(Collection { span, .. }) => markup::splice(
            &xml,
            vec![(span, markup::worksheet_collection(&layout.prefix, &entry))],
        ),
        None => {
            let at = markup::collection_position(
                layout
                    .children
                    .iter()
                    .map(|child| (child.name.as_slice(), child.span.start)),
                layout.root_close,
            );
            markup::splice(
                &xml,
                vec![(at..at, markup::worksheet_collection(&layout.prefix, &entry))],
            )
        },
    };
    opc.get_part_mut(worksheet)?.set_blob(xml);
    read_back(opc, worksheet, shape_id)?;
    Ok(shape_id)
}

/// Refuse a staged graph whose worksheet does not list `shape_id` as an
/// equation object with a readable embedding.
fn read_back(opc: &OpcPackage, worksheet: &PackURI, shape_id: u32) -> Result<()> {
    let objects = list(opc, worksheet)?;
    match objects.iter().find(|object| object.shape_id() == shape_id) {
        Some(object) if object.native().is_some() => Ok(()),
        _ => Err(Error::Internal(
            "the written equation object did not read back",
        )),
    }
}

// ---------------------------------------------------------------------------
// Set

/// Where an object's preview is referenced: the worksheet relationship of
/// its `objectPr` and the VML relationship of its shape's `v:imagedata`.
struct PreviewUse {
    worksheet_id: Option<String>,
    vml: Option<(PackURI, Range<usize>, String)>,
}

fn set_object(
    opc: &mut OpcPackage,
    worksheet: &PackURI,
    shape_id: u32,
    embedding: &PackURI,
    object: Vec<u8>,
    png: Vec<u8>,
    size: (f64, f64),
) -> Result<()> {
    opc.get_part_mut(embedding)?.set_blob(object);
    let mut xml = opc.get_part(worksheet)?.blob().to_vec();
    let (current, others) = parsed(&xml, shape_id)?;
    let layout = WorksheetLayout::scan(&xml)?;
    // A VML drawing or drawing whose part is missing is left alone.
    let vml =
        legacy_drawing(opc, worksheet, &layout)?.filter(|(_, name)| opc.get_part(name).is_ok());
    let spid = format!("_x0000_s{shape_id}");

    // The VML shape, its image relationship and the other shapes' ones.
    let mut vml_shape: Option<(PackURI, Range<usize>)> = None;
    let mut other_vml_ids = HashSet::new();
    if let Some((_, name)) = &vml {
        let vml_xml = opc.get_part(name)?.blob();
        for (span, id) in vml_shapes(opc, vml_xml)? {
            if id == Some(shape_id) && vml_shape.is_none() {
                vml_shape = Some((name.clone(), span));
            } else if let Some((_, relid)) = markup::vml_image_relationship(&vml_xml[span]) {
                other_vml_ids.insert(relid);
            }
        }
    }
    let uses = PreviewUse {
        worksheet_id: current
            .properties
            .as_ref()
            .map(|properties| properties.preview_relationship_id.clone()),
        vml: match &vml_shape {
            Some((name, span)) => {
                let shape = &opc.get_part(name)?.blob()[span.clone()];
                markup::vml_image_relationship(shape)
                    .map(|(at, id)| (name.clone(), span.start + at.start..span.start + at.end, id))
            },
            None => None,
        },
    };
    let shared_worksheet_id = uses.worksheet_id.as_ref().is_some_and(|id| {
        others.iter().any(|other| {
            other
                .properties
                .as_ref()
                .is_some_and(|properties| &properties.preview_relationship_id == id)
        })
    });
    let shared_vml_id = uses
        .vml
        .as_ref()
        .is_some_and(|(_, _, id)| other_vml_ids.contains(id));
    let worksheet_target = match &uses.worksheet_id {
        Some(id) => target(opc, worksheet, id)?,
        None => None,
    };
    let vml_target = match &uses.vml {
        Some((name, _, id)) => target(opc, name, id)?,
        None => None,
    };
    let mut targets: Vec<PackURI> = worksheet_target
        .iter()
        .chain(&vml_target)
        .cloned()
        .collect();
    targets.dedup();
    let in_place = match targets.as_slice() {
        [part] => {
            !shared_worksheet_id
                && !shared_vml_id
                && opc.get_part(part)?.content_type() == ct::PNG
                && only_users(
                    opc,
                    part,
                    worksheet,
                    uses.worksheet_id.as_deref(),
                    uses.vml.as_ref(),
                )
        },
        _ => false,
    };
    let mut vml_edits: Vec<(Range<usize>, String)> = Vec::new();
    let mut worksheet_preview: Option<(String, String)> = None;
    if uses.worksheet_id.is_none() && uses.vml.is_none() {
        // Nothing shows a preview of this object.
    } else if in_place {
        opc.get_part_mut(&targets[0])?.set_blob(png);
    } else {
        let image = free_name(opc, "/xl/media/image", ".png")?;
        opc.add_part(Box::new(BlobPart::new(
            image.clone(),
            ct::PNG.to_owned(),
            png,
        )));
        if let Some(old) = &uses.worksheet_id {
            let id = opc
                .get_part_mut(worksheet)?
                .relate_to(&image.relative_ref(worksheet.base_uri()), rt::IMAGE);
            if !shared_worksheet_id {
                opc.get_part_mut(worksheet)?.rels_mut().remove(old);
            }
            worksheet_preview = Some((old.clone(), id));
        }
        if let Some((name, span, old)) = &uses.vml {
            let id = opc
                .get_part_mut(name)?
                .relate_to(&image.relative_ref(name.base_uri()), rt::IMAGE);
            if !shared_vml_id {
                opc.get_part_mut(name)?.rels_mut().remove(old);
            }
            vml_edits.push((span.clone(), id));
        }
        drop_orphans(opc, targets);
    }

    // The new frame, from the object's top-left corner.
    let from = current
        .properties
        .as_ref()
        .map(|properties| properties.anchor.from)
        .or_else(|| {
            let (name, span) = vml_shape.as_ref()?;
            vml_from(&opc.get_part(name).ok()?.blob()[span.clone()])
        });
    let frame = match from {
        Some(from) => Some(Grid::read(&xml, from.row)?.frame_from(from, size)),
        None => None,
    };

    // The worksheet entry: its preview relationship and its `to` marker.
    let entries =
        collection(&xml, &layout)?.ok_or_else(|| invalid("the worksheet has no oleObjects"))?;
    let entry = entries
        .entry(shape_id)
        .ok_or_else(|| invalid("the equation object has no oleObjects entry"))?
        .span
        .clone();
    let fragment = &xml[entry.clone()];
    let mut edits = match &frame {
        Some(frame) => markup::to_marker_edits(fragment, &frame.anchor.to)?,
        None => Vec::new(),
    };
    if let Some((old, new)) = &worksheet_preview {
        edits.extend(object_preview_edits(fragment, old, new)?);
    }
    let fragment = markup::splice(fragment, edits);
    xml = markup::splice(
        &xml,
        vec![(
            entry,
            String::from_utf8(fragment).map_err(|_| invalid("oleObjects entry is not UTF-8"))?,
        )],
    );
    opc.get_part_mut(worksheet)?.set_blob(xml);

    // The VML shape: its image relationship and its size.
    if let Some((name, span)) = &vml_shape {
        let vml_xml = opc.get_part(name)?.blob().to_vec();
        let shape = markup::splice(&vml_xml[span.clone()], {
            vml_edits
                .iter()
                .map(|(at, id)| (at.start - span.start..at.end - span.start, id.clone()))
                .collect()
        });
        let shape = match &frame {
            Some(frame) => markup::resize_vml_shape(&shape, frame),
            None => shape,
        };
        let mut out = vml_xml;
        out.splice(span.clone(), shape);
        opc.get_part_mut(name)?.set_blob(out);
    }

    // The hidden drawing shape's `to` marker.
    if let Some(frame) = &frame
        && let Some(name) = objects::drawing_of(opc.get_part(worksheet)?.rels())?
        && opc.get_part(&name).is_ok()
    {
        let drawing_xml = opc.get_part(&name)?.blob();
        if let Some(span) = markup::scan_drawing(drawing_xml, Some(&spid))?.object {
            let edits = markup::to_marker_edits(&drawing_xml[span.clone()], &frame.anchor.to)?;
            let anchor = markup::splice(&drawing_xml[span.clone()], edits);
            let mut out = drawing_xml.to_vec();
            out.splice(span, anchor);
            opc.get_part_mut(&name)?.set_blob(out);
        }
    }
    read_back(opc, worksheet, shape_id)
}

/// Whether every relationship targeting `part` is the object's own
/// preview relationship of the worksheet or of its VML drawing.
fn only_users(
    opc: &OpcPackage,
    part: &PackURI,
    worksheet: &PackURI,
    worksheet_id: Option<&str>,
    vml: Option<&(PackURI, Range<usize>, String)>,
) -> bool {
    let targets = |relationship: &litchi_opc::Relationship| {
        !relationship.is_external() && relationship.target_partname().ok().as_ref() == Some(part)
    };
    if opc.rels().iter().any(targets) {
        return false;
    }
    opc.iter_parts().all(|owner| {
        owner
            .rels()
            .iter()
            .filter(|r| targets(r))
            .all(|relationship| {
                let id = relationship.r_id();
                (owner.partname() == worksheet && worksheet_id == Some(id))
                    || vml.is_some_and(|(name, _, relid)| owner.partname() == name && relid == id)
            })
    })
}

/// The edits that point the `objectPr` of a worksheet entry at the
/// relationship `new` instead of `old`.
fn object_preview_edits(
    fragment: &[u8],
    old: &str,
    new: &str,
) -> Result<Vec<(Range<usize>, String)>> {
    let mut edits = Vec::new();
    let mut cursor = 0;
    while let Some(at) = markup::find(fragment, b"objectPr", cursor) {
        cursor = at + 8;
        let Some(tag_end) = markup::find(fragment, b">", at) else {
            break;
        };
        let tag = &fragment[at..tag_end];
        for quote in [b'"', b'\''] {
            let mut needle = b":id=".to_vec();
            needle.push(quote);
            needle.extend_from_slice(old.as_bytes());
            needle.push(quote);
            if let Some(found) = markup::find(tag, &needle, 0) {
                let start = at + found + 5;
                edits.push((start..start + old.len(), litchi_core::xml::escape_xml(new)));
            }
        }
    }
    if edits.is_empty() {
        return Err(invalid("the objectPr preview relationship was not found"));
    }
    Ok(edits)
}

/// The top-left marker of a VML object shape's `x:Anchor` (columns, rows
/// and pixel offsets).
fn vml_from(shape: &[u8]) -> Option<OleObjectMarker> {
    let open = markup::find(shape, b"<x:Anchor>", 0)? + b"<x:Anchor>".len();
    let close = markup::find(shape, b"</x:Anchor>", open)?;
    let text = std::str::from_utf8(&shape[open..close]).ok()?;
    let values = text
        .split(',')
        .map(|value| value.trim().parse::<i64>().ok())
        .collect::<Option<Vec<_>>>()?;
    let [column, column_px, row, row_px, ..] = values.as_slice() else {
        return None;
    };
    Some(OleObjectMarker {
        column: u32::try_from(*column).ok()?,
        column_offset: column_px.checked_mul(EMU_PER_PIXEL)?,
        row: u32::try_from(*row).ok()?,
        row_offset: row_px.checked_mul(EMU_PER_PIXEL)?,
    })
}

// ---------------------------------------------------------------------------
// Remove

fn remove_object(opc: &mut OpcPackage, worksheet: &PackURI, shape_id: u32) -> Result<()> {
    let mut xml = opc.get_part(worksheet)?.blob().to_vec();
    let (current, others) = parsed(&xml, shape_id)?;
    let mut orphans = Vec::new();

    // The worksheet entry, or the whole `<oleObjects>` with the last one.
    let layout = WorksheetLayout::scan(&xml)?;
    let entries =
        collection(&xml, &layout)?.ok_or_else(|| invalid("the worksheet has no oleObjects"))?;
    let entry = entries
        .entry(shape_id)
        .ok_or_else(|| invalid("the equation object has no oleObjects entry"))?;
    let span = if entries.entries.len() == 1 {
        entries.span.clone()
    } else {
        entry.span.clone()
    };
    xml.drain(span);

    // Its relationships, kept when another object shares one.
    let mut drop_relationship = |opc: &mut OpcPackage, id: &str| -> Result<()> {
        if let Some(part) = target(opc, worksheet, id)? {
            orphans.push(part);
        }
        opc.get_part_mut(worksheet)?.rels_mut().remove(id);
        Ok(())
    };
    drop_relationship(opc, &current.relationship_id)?;
    if let Some(properties) = &current.properties
        && !others.iter().any(|other| {
            other.properties.as_ref().is_some_and(|theirs| {
                theirs.preview_relationship_id == properties.preview_relationship_id
            })
        })
    {
        drop_relationship(opc, &properties.preview_relationship_id)?;
    }

    // The VML shape, or the VML drawing with its last shape.
    let layout = WorksheetLayout::scan(&xml)?;
    if let Some((id, name)) = legacy_drawing(opc, worksheet, &layout)?
        && opc.get_part(&name).is_ok()
    {
        let vml_xml = opc.get_part(&name)?.blob().to_vec();
        let shapes = vml_shapes(opc, &vml_xml)?;
        if let Some((span, _)) = shapes.iter().find(|(_, id)| *id == Some(shape_id)) {
            let relationship = markup::vml_image_relationship(&vml_xml[span.clone()]);
            if shapes.len() == 1 {
                xml = layout
                    .remove_child(&xml, "legacyDrawing")
                    .ok_or_else(|| invalid("legacyDrawing vanished"))?;
                opc.get_part_mut(worksheet)?.rels_mut().remove(&id);
                orphans.push(name.clone());
            } else {
                if let Some((_, relid)) = relationship {
                    let shared = shapes.iter().any(|(other, id)| {
                        *id != Some(shape_id)
                            && markup::vml_image_relationship(&vml_xml[other.clone()])
                                .is_some_and(|(_, theirs)| theirs == relid)
                    });
                    if !shared {
                        if let Some(part) = target(opc, &name, &relid)? {
                            orphans.push(part);
                        }
                        opc.get_part_mut(&name)?.rels_mut().remove(&relid);
                    }
                }
                let mut out = vml_xml.clone();
                out.drain(span.clone());
                opc.get_part_mut(&name)?.set_blob(out);
            }
        }
    }

    // The hidden drawing shape, or the drawing with its last anchor.
    if let Some(name) = objects::drawing_of(opc.get_part(worksheet)?.rels())?
        && opc.get_part(&name).is_ok()
    {
        let drawing_xml = opc.get_part(&name)?.blob().to_vec();
        let scan = markup::scan_drawing(&drawing_xml, Some(&format!("_x0000_s{shape_id}")))?;
        if let Some(span) = scan.object {
            if scan.children == 1 && opc.get_part(&name)?.rels().iter().next().is_none() {
                let layout = WorksheetLayout::scan(&xml)?;
                if let Some(child) = layout.child("drawing")
                    && let Some(id) = child.relationship_id.clone()
                {
                    xml = layout
                        .remove_child(&xml, "drawing")
                        .ok_or_else(|| invalid("drawing vanished"))?;
                    opc.get_part_mut(worksheet)?.rels_mut().remove(&id);
                    orphans.push(name.clone());
                }
            } else {
                let mut out = drawing_xml;
                out.drain(span);
                opc.get_part_mut(&name)?.set_blob(out);
            }
        }
    }

    opc.get_part_mut(worksheet)?.set_blob(xml);
    // A VML drawing or drawing that is dropped releases its own targets.
    let mut released = Vec::new();
    for part in &orphans {
        if !referenced(opc, part)
            && let Ok(owner) = opc.get_part(part)
        {
            released.extend(
                owner
                    .rels()
                    .iter()
                    .filter(|relationship| !relationship.is_external())
                    .filter_map(|relationship| relationship.target_partname().ok()),
            );
        }
    }
    drop_orphans(opc, orphans);
    drop_orphans(opc, released);
    Ok(())
}
