//! Worksheet pictures and shapes of an XLSB package: read, add, move and
//! remove (unified-ops record 0322).
//!
//! An XLSB worksheet's Drawings part is the same SpreadsheetML drawing XML
//! as XLSX's ([MS-XLSB] section 2.1.7.23); the worksheet names it with a
//! `BrtDrawing` record. [`crate::Workbook::drawing_elements`] and
//! [`crate::SourceBackedWorksheet::drawing_elements`] list the top-level
//! objects of a worksheet's drawing in drawing order, each an [`Element`]
//! with its [`Placement`], names and the properties a picture or a preset
//! shape carries. [`Package::edit_drawings`] starts an [`Edit`]:
//! [`Edit::add_picture`] stages a media part, its relationship and an
//! `xdr:pic`; [`Edit::add_shape`] stages an `xdr:sp` with a preset geometry,
//! fill, outline and text; [`Edit::set`] changes the placement, names,
//! rotation and flips of an element in place; and [`Edit::remove`] deletes
//! an element's anchor with the parts only it refers to. The drawing part
//! (and the worksheet's `BrtDrawing`) is created with the first object.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_drawings`] replays a patch. Every other anchor and
//! worksheet record keeps its bytes; an exact no-op keeps every byte, and a
//! changed edit of a signed workbook is refused with [`Error::Signed`].
//! Colors are `0xRRGGBB`.
//!
//! ```
//! use litchi_xlsb::Package;
//! use litchi_xlsb::drawing::edit::{ElementKind, Marker, Placement, ShapeSpec};
//!
//! # fn main() -> Result<(), litchi_xlsb::package::PackageError> {
//! let package = Package::new()?;
//! let mut edit = package.edit_drawings()?;
//! let placement = Placement::TwoCell {
//!     from: Marker::cell(0, 4),
//!     to: Marker::cell(2, 7),
//! };
//! let mut shape = ShapeSpec::new(placement, "Rectangle 1");
//! shape.text = Some("rect".into());
//! edit.add_shape(0, &shape)?;
//! let after = edit.commit()?.into_package();
//! let elements = after.workbook()?.drawing_elements(0)?;
//! assert_eq!(elements[0].kind, ElementKind::Shape);
//! assert_eq!(elements[0].text.as_deref(), Some("rect"));
//! # Ok(())
//! # }
//! ```

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use litchi_spreadsheet_drawing::objects;

use crate::package::Package;
use crate::package::error::{Error as PackageError, Result};
use crate::part_edit::{Rank, SheetLayout, Staged, family, invalid, snapshot_options, splice};
use crate::raw::{Encoder, kind};

family!("drawing");
snapshot_options!("drawing");

pub use litchi_spreadsheet_drawing::LimitKind;
pub use objects::{
    Align, COLUMN_EMU, Change, DASHES, ElementKind, Image, Limits, MAX_DRAWING_BYTES, MAX_ELEMENTS,
    MAX_IMAGE_BYTES, Marker, PictureSpec, Placement, ROW_EMU, SCHEME_COLORS, VerticalAlign,
    canonical,
};

/// A `0xRRGGBB` color.
pub type Rgb = u32;
/// A shape fill.
pub type Fill = objects::Fill<Rgb>;
/// A shape outline.
pub type Line = objects::Line<Rgb>;
/// The run formatting of a shape's text, read from its first run.
pub type Font = objects::Font<Rgb>;
/// One top-level object of a worksheet drawing.
pub type Element = objects::Element<Rgb>;
/// A preset shape to add.
pub type ShapeSpec = objects::ShapeSpec<Rgb>;

/// The elements of a worksheet part's drawing in an eager package.
pub(crate) fn eager_elements(package: &OpcPackage, worksheet: &PackURI) -> Result<Vec<Element>> {
    let Some(drawing) = objects::drawing_of(package.get_part(worksheet)?.rels())? else {
        return Ok(Vec::new());
    };
    let part = package.get_part(&drawing)?;
    let limits = Limits::from_source_limits(package.read_limits());
    Ok(objects::read(
        part.blob(),
        part.rels(),
        |target| {
            let image = package.get_part(target).ok()?;
            Some((image.content_type().to_owned(), image.blob().len() as u64))
        },
        &limits,
    )?)
}

/// The elements of a worksheet part's drawing in a source-backed package:
/// only the worksheet's relationships, its drawing part and the image parts'
/// entries are read.
pub(crate) fn source_elements(
    package: &litchi_opc::SourceBackedPackage,
    worksheet: &PackURI,
) -> Result<Vec<Element>> {
    let Some(drawing) = objects::drawing_of(package.part(worksheet)?.rels())? else {
        return Ok(Vec::new());
    };
    let part = package.part(&drawing)?;
    let declared = part.declared_uncompressed_size()?;
    let limits = Limits::from_source_limits(package.read_limits());
    let declared = usize::try_from(declared).unwrap_or(usize::MAX);
    if declared > limits.max_drawing_bytes {
        return Err(PackageError::limit(
            LimitKind::DrawingBytes,
            declared,
            limits.max_drawing_bytes,
        ));
    }
    let data = part.data()?;
    Ok(objects::read(
        data.as_bytes(),
        part.rels(),
        |target| {
            let image = package.part(target).ok()?;
            Some((
                image.content_type().to_owned(),
                image.declared_uncompressed_size().ok()?,
            ))
        },
        &limits,
    )?)
}

/// A chart a worksheet drawing hosts: its frame's placement and name, and
/// the chart part's DrawingML (`c:chartSpace`) markup.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct ChartPart {
    /// Where the chart's frame sits.
    pub placement: Option<Placement>,
    /// The frame's `cNvPr/@name`.
    pub name: Option<String>,
    /// The chart part's markup.
    pub xml: Vec<u8>,
}

/// The charts of a worksheet part's drawing in an eager package, in drawing
/// order.
pub(crate) fn eager_charts(package: &OpcPackage, worksheet: &PackURI) -> Result<Vec<ChartPart>> {
    let Some(drawing) = objects::drawing_of(package.get_part(worksheet)?.rels())? else {
        return Ok(Vec::new());
    };
    let part = package.get_part(&drawing)?;
    let mut charts = Vec::new();
    let limits = Limits::from_source_limits(package.read_limits());
    for frame in objects::chart_frames(part.blob(), part.rels(), &limits)? {
        charts.push(ChartPart {
            placement: frame.placement,
            name: frame.name,
            xml: package.get_part(&frame.target)?.blob().to_vec(),
        });
    }
    Ok(charts)
}

/// The charts of a worksheet part's drawing in a source-backed package.
pub(crate) fn source_charts(
    package: &litchi_opc::SourceBackedPackage,
    worksheet: &PackURI,
) -> Result<Vec<ChartPart>> {
    let Some(drawing) = objects::drawing_of(package.part(worksheet)?.rels())? else {
        return Ok(Vec::new());
    };
    let part = package.part(&drawing)?;
    let declared = part.declared_uncompressed_size()?;
    let limits = Limits::from_source_limits(package.read_limits());
    let declared = usize::try_from(declared).unwrap_or(usize::MAX);
    if declared > limits.max_drawing_bytes {
        return Err(PackageError::limit(
            LimitKind::DrawingBytes,
            declared,
            limits.max_drawing_bytes,
        ));
    }
    let data = part.data()?;
    let mut charts = Vec::new();
    for frame in objects::chart_frames(data.as_bytes(), part.rels(), &limits)? {
        let chart = package.part(&frame.target)?;
        let declared = chart.declared_uncompressed_size()?;
        let declared = usize::try_from(declared).unwrap_or(usize::MAX);
        if declared > limits.max_drawing_bytes {
            return Err(PackageError::limit(
                LimitKind::DrawingBytes,
                declared,
                limits.max_drawing_bytes,
            ));
        }
        charts.push(ChartPart {
            placement: frame.placement,
            name: frame.name,
            xml: chart.data()?.as_bytes().to_vec(),
        });
    }
    Ok(charts)
}

impl Package {
    /// Start a worksheet drawing edit.
    ///
    /// # Errors
    ///
    /// Returns the workbook materialization errors of this package.
    pub fn edit_drawings(&self) -> Result<Edit> {
        self.edit_drawings_with_options(EditOptions::default())
    }

    /// Start a worksheet drawing edit with an explicit retained-payload ceiling.
    pub fn edit_drawings_with_options(&self, options: EditOptions) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
            options,
        })
    }

    /// Apply a drawing patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_drawings(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

/// A staged worksheet drawing edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    options: EditOptions,
}

impl Edit {
    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// The staged elements of the worksheet at `sheet`.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet, a chartsheet, or a malformed
    /// drawing graph.
    pub fn elements(&self, sheet: usize) -> Result<Vec<Element>> {
        let worksheet = self.staged.worksheet(sheet)?;
        eager_elements(self.staged.opc(), &worksheet)
    }

    /// The worksheet's drawing part, created (with its relationship and the
    /// worksheet's `BrtDrawing` record) when it has none.
    fn drawing(&mut self, sheet: usize) -> Result<PackURI> {
        let worksheet = self.staged.worksheet(sheet)?;
        if let Some(drawing) = objects::drawing_of(self.staged.opc().get_part(&worksheet)?.rels())?
        {
            return Ok(drawing);
        }
        let drawing = self.staged.free_name("/xl/drawings/drawing", ".xml")?;
        self.staged.add_part(BlobPart::new(
            drawing.clone(),
            ct::OFC_DRAWING.to_owned(),
            objects::EMPTY_DRAWING.as_bytes().to_vec(),
        ))?;
        let id = self
            .staged
            .opc_mut()?
            .get_part_mut(&worksheet)?
            .relate_to(&drawing.relative_ref(worksheet.base_uri()), rt::DRAWING);
        let bytes = self.staged.blob(&worksheet)?;
        let raw_limits = self.staged.raw_limits();
        let layout = SheetLayout::scan_with_limits(bytes, raw_limits)?;
        if layout.units_of(kind::DRAWING).next().is_some() {
            return Err(invalid(
                "the worksheet names a drawing it has no relationship to",
            ));
        }
        let at = layout.insertion(Rank::Drawing);
        let mut payload = Vec::new();
        Encoder::new(&mut payload, raw_limits).write_wide_string(&id)?;
        let updated = splice(
            bytes,
            at..at,
            &crate::part_edit::record_with_limits(kind::DRAWING, &payload, raw_limits)?,
        );
        self.staged.replace_worksheet(&worksheet, updated)?;
        self.staged.touch(&worksheet);
        Ok(drawing)
    }

    /// Append `fragment` to `drawing` and count its elements of `kind`.
    fn append(&mut self, drawing: &PackURI, fragment: &str, kind: ElementKind) -> Result<usize> {
        let limits = Limits::from_source_limits(self.staged.opc().read_limits());
        let xml = objects::append(self.staged.blob(drawing)?, fragment, &limits)?;
        let count = objects::count(&xml, kind, &limits)?;
        self.staged.replace_blob(drawing, xml)?;
        Ok(count)
    }

    /// Add a picture; returns its index among the worksheet's pictures.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, an invalid
    /// placement, name or content type, a full drawing, or a signed workbook.
    pub fn add_picture(&mut self, sheet: usize, spec: &PictureSpec) -> Result<usize> {
        let limits = Limits::from_source_limits(self.staged.opc().read_limits());
        objects::check_picture(spec, &limits)?;
        let drawing = self.drawing(sheet)?;
        let id = objects::next_id(self.staged.blob(&drawing)?, &limits)?;
        let media = self
            .staged
            .free_name("/xl/media/image", &format!(".{}", spec.extension))?;
        self.staged.add_part(BlobPart::new(
            media.clone(),
            spec.content_type.clone(),
            spec.image.clone(),
        ))?;
        let relationship_id = self
            .staged
            .opc_mut()?
            .get_part_mut(&drawing)?
            .relate_to(&media.relative_ref(drawing.base_uri()), rt::IMAGE);
        self.staged.touch(&drawing);
        let fragment = objects::picture_fragment(spec, id, &relationship_id, &limits)?;
        Ok(self.append(&drawing, &fragment, ElementKind::Picture)? - 1)
    }

    /// Add a preset shape; returns its index among the worksheet's shapes.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, an invalid
    /// placement, geometry, text or font, a full drawing, or a signed
    /// workbook.
    pub fn add_shape(&mut self, sheet: usize, spec: &ShapeSpec) -> Result<usize> {
        spec.placement.validate()?;
        let drawing = self.drawing(sheet)?;
        let limits = Limits::from_source_limits(self.staged.opc().read_limits());
        let id = objects::next_id(self.staged.blob(&drawing)?, &limits)?;
        let fragment = objects::shape_fragment(spec, id, &limits)?;
        Ok(self.append(&drawing, &fragment, ElementKind::Shape)? - 1)
    }

    /// The worksheet's drawing part, which an element edit needs.
    fn existing_drawing(&self, sheet: usize) -> Result<PackURI> {
        let worksheet = self.staged.worksheet(sheet)?;
        objects::drawing_of(self.staged.opc().get_part(&worksheet)?.rels())?
            .ok_or_else(|| invalid("the worksheet has no drawing"))
    }

    /// Change the placement, names, rotation or flips of the element at
    /// `index` among the worksheet's elements of `kind`. Returns whether
    /// anything changed; everything else of the element keeps its bytes.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing element, an invalid change, markup this
    /// editor cannot rewrite, or a changed edit of a signed workbook.
    pub fn set(
        &mut self,
        sheet: usize,
        kind: ElementKind,
        index: usize,
        change: &Change,
    ) -> Result<bool> {
        let drawing = self.existing_drawing(sheet)?;
        let limits = Limits::from_source_limits(self.staged.opc().read_limits());
        match objects::change(self.staged.blob(&drawing)?, kind, index, change, &limits)? {
            Some(xml) => self.staged.replace_blob(&drawing, xml),
            None => Ok(false),
        }
    }

    /// Remove the element at `index` among the worksheet's elements of
    /// `kind`, with its relationships and the parts only it refers to.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing element or a signed workbook.
    pub fn remove(&mut self, sheet: usize, kind: ElementKind, index: usize) -> Result<()> {
        let drawing = self.existing_drawing(sheet)?;
        let limits = Limits::from_source_limits(self.staged.opc().read_limits());
        let detached = objects::detach(self.staged.blob(&drawing)?, kind, index, &limits)?;
        let opc = self.staged.opc_mut()?;
        let part = opc.get_part_mut(&drawing)?;
        part.set_blob(detached.xml);
        let mut targets = Vec::new();
        for id in detached.released {
            if let Some(relationship) = part.rels().get(&id)
                && !relationship.is_external()
                && let Ok(target) = relationship.target_partname()
            {
                targets.push(target);
            }
            part.rels_mut().remove(&id);
        }
        self.staged.touch(&drawing);
        for target in targets {
            let opc = self.staged.opc_mut()?;
            let referenced = opc.iter_parts().any(|part| {
                part.rels().iter().any(|relationship| {
                    !relationship.is_external()
                        && relationship.target_partname().ok().as_ref() == Some(&target)
                })
            });
            if !referenced {
                self.staged.remove_part(&target)?;
            }
        }
        Ok(())
    }

    /// Validate the staged package and publish it.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit_with_snapshot_limit(
            self.options.max_snapshot_bytes_value(),
        )?))
    }
}

#[cfg(test)]
mod tests {
    use super::{Change, ElementKind, Fill, Font, Marker, PictureSpec, Placement, ShapeSpec};
    use crate::Package;

    const PNG: &[u8] = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR\0\0\0\x01\0\0\0\x01\x08\x06\0\0\0\x1f\x15\xc4\x89\0\0\0\rIDATx\x9cc\xf8\x0f\0\0\x01\x01\0\x05\x18\xd8\xa8\x00\0\0\0IEND\xaeB`\x82";

    fn placement() -> Placement {
        Placement::TwoCell {
            from: Marker::cell(1, 1),
            to: Marker::cell(4, 6),
        }
    }

    #[test]
    fn shapes_and_pictures_round_trip() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_drawings().unwrap();
        let mut shape = ShapeSpec::new(placement(), "Shape 1");
        shape.geometry = "ellipse".into();
        shape.text = Some("one\ntwo & three".into());
        shape.fill = Some(Fill::Solid(0x4472C4));
        shape.font = Font {
            bold: Some(true),
            color: Some(0xFFFFFF),
            size: Some(14.0),
            ..Font::default()
        };
        shape.rotation = Some(30.0);
        shape.flip_h = true;
        assert_eq!(edit.add_shape(0, &shape).unwrap(), 0);
        let picture = PictureSpec {
            placement: Placement::OneCell {
                from: Marker::cell(6, 2),
                cx: 914_400,
                cy: 914_400,
            },
            name: "Picture 1".into(),
            description: Some("logo".into()),
            title: None,
            image: PNG.to_vec(),
            content_type: "image/png".into(),
            extension: "png".into(),
            rotation: None,
            flip_h: false,
            flip_v: false,
        };
        assert_eq!(edit.add_picture(0, &picture).unwrap(), 0);
        let (after, patch) = edit.commit().unwrap().into_parts();
        let reopened = Package::from_bytes(&after.to_bytes().unwrap()).unwrap();
        let elements = reopened.workbook().unwrap().drawing_elements(0).unwrap();
        assert_eq!(elements.len(), 2);
        let shape = &elements[0];
        assert_eq!(shape.kind, ElementKind::Shape);
        assert_eq!(shape.geometry.as_deref(), Some("ellipse"));
        assert_eq!(shape.text.as_deref(), Some("one\ntwo & three"));
        assert_eq!(shape.rotation, Some(30.0));
        assert!(shape.flip_h);
        assert_eq!(shape.font.bold, Some(true));
        assert_eq!(shape.font.size, Some(14.0));
        assert_eq!(shape.font.color, Some(0xFFFFFF));
        assert_eq!(shape.placement, Some(placement()));
        let picture = &elements[1];
        assert_eq!(picture.kind, ElementKind::Picture);
        assert_eq!(picture.description.as_deref(), Some("logo"));
        assert_eq!(picture.image.as_ref().unwrap().content_type, "image/png");
        assert_eq!(
            package
                .apply_drawings(&patch)
                .map(|commit| commit.into_package())
                .unwrap()
                .to_bytes()
                .unwrap(),
            after.to_bytes().unwrap()
        );

        let mut edit = after.edit_drawings().unwrap();
        let moved = Placement::TwoCell {
            from: Marker::cell(0, 0),
            to: Marker::cell(2, 2),
        };
        assert!(
            edit.set(
                0,
                ElementKind::Picture,
                0,
                &Change {
                    placement: Some(moved),
                    description: Some(None),
                    rotation: Some(90.0),
                    ..Change::default()
                }
            )
            .unwrap()
        );
        let elements = edit.elements(0).unwrap();
        assert_eq!(elements[1].placement, Some(moved));
        assert_eq!(elements[1].description, None);
        assert_eq!(elements[1].rotation, Some(90.0));
        edit.remove(0, ElementKind::Picture, 0).unwrap();
        edit.remove(0, ElementKind::Shape, 0).unwrap();
        let empty = edit.commit().unwrap().into_package();
        assert!(
            empty
                .workbook()
                .unwrap()
                .drawing_elements(0)
                .unwrap()
                .is_empty()
        );
    }
}
