//! Worksheet pictures and shapes: read, add, move and remove (unified-ops
//! record 0320).
//!
//! [`Package::drawing_elements`] and
//! [`crate::workbook::SourceWorksheet::drawing_elements`] list the top-level
//! objects of a worksheet's drawing in drawing order, each an [`Element`]
//! with its [`Placement`], names and the properties a picture or a preset
//! shape carries. [`Package::edit_drawings`] starts an [`Edit`] over an
//! immutable snapshot: [`Edit::add_picture`] stages a media part, its
//! relationship and an `xdr:pic`; [`Edit::add_shape`] stages an `xdr:sp`
//! with a preset geometry, fill, outline and text; [`Edit::set`] changes the
//! placement, names, rotation and flips of an element in place; and
//! [`Edit::remove`] deletes an element's anchor with the parts only it
//! refers to. The worksheet's drawing part is created with the first object.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_drawings`] replays a patch. Every other anchor keeps its
//! bytes; an exact no-op keeps every byte, and a changed edit of a signed
//! workbook is refused with [`Error::Signed`].
//!
//! ```
//! use litchi_xlsx::Package;
//! use litchi_xlsx::drawing::edit::{ElementKind, Marker, Placement, ShapeSpec};
//!
//! # fn main() -> Result<(), litchi_xlsx::Error> {
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
//! let elements = after.drawing_elements(0)?;
//! assert_eq!(elements[0].kind, ElementKind::Shape);
//! assert_eq!(elements[0].text.as_deref(), Some("rect"));
//! # Ok(())
//! # }
//! ```

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use litchi_spreadsheet_drawing::objects;

use crate::Package;
use crate::color::Rgb;
use crate::error::{Result, invalid};
use crate::part_edit::{Staged, WorksheetLayout, family};

family!("drawing");

pub use litchi_spreadsheet_drawing::LimitKind;
pub use objects::{
    Align, COLUMN_EMU, Change, DASHES, ElementKind, Image, Limits, MAX_DRAWING_BYTES, MAX_ELEMENTS,
    MAX_IMAGE_BYTES, Marker, PictureSpec, Placement, ROW_EMU, SCHEME_COLORS, VerticalAlign,
    canonical,
};

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

impl objects::Color for Rgb {
    const BLACK: Self = Self::new(0, 0, 0);

    fn from_hex(value: &str) -> Option<Self> {
        Self::from_hex(value).ok()
    }

    fn hex(self) -> String {
        format!("{:02X}{:02X}{:02X}", self.red(), self.green(), self.blue())
    }
}

/// The elements of a worksheet part's drawing in an eager package.
fn eager_elements(package: &OpcPackage, worksheet: &PackURI) -> Result<Vec<Element>> {
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
    package.check_execution()?;
    let Some(drawing) = objects::drawing_of(package.part(worksheet)?.rels())? else {
        return Ok(Vec::new());
    };
    let part = package.part(&drawing)?;
    let declared = part.declared_uncompressed_size()?;
    let limits = Limits::from_source_limits(package.read_limits());
    let declared = usize::try_from(declared).unwrap_or(usize::MAX);
    if declared > limits.max_drawing_bytes {
        return Err(Error::limit(
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

impl Package {
    /// The top-level objects of a worksheet's drawing, in drawing order.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet, a chartsheet, or a malformed
    /// drawing graph.
    pub fn drawing_elements(&self, sheet: usize) -> Result<Vec<Element>> {
        let worksheet = worksheet_uri(self, sheet)?;
        eager_elements(self.opc(), &worksheet)
    }

    /// Start a worksheet drawing edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused.
    pub fn edit_drawings(&self) -> Result<Edit> {
        self.edit_drawings_with_options(crate::EditOptions::default())
    }

    /// Start a drawing edit with an explicit retained-patch budget.
    pub fn edit_drawings_with_options(&self, options: crate::EditOptions) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::new_with_options(self, "edit_drawings", options)?,
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
        Ok(Commit::new((
            patch.0.apply(self, "apply_drawings")?,
            patch.0.clone(),
        )))
    }
}

/// A replacement image's size, content type and extension.
fn check_image(image: &[u8], content_type: &str, extension: &str, limits: &Limits) -> Result<()> {
    let limits = limits.validate()?;
    if image.len() > limits.max_image_bytes {
        return Err(Error::limit(
            LimitKind::ImageBytes,
            image.len(),
            limits.max_image_bytes,
        ));
    }
    if !content_type.starts_with("image/")
        || extension.is_empty()
        || !extension.bytes().all(|b| b.is_ascii_alphanumeric())
    {
        return Err(invalid(
            "a picture needs an image content type and extension",
        ));
    }
    Ok(())
}

fn worksheet_uri(package: &Package, sheet: usize) -> Result<PackURI> {
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

/// A staged worksheet drawing edit of one [`Package`]. Dropping it discards
/// it.
#[derive(Debug)]
pub struct Edit {
    staged: Staged,
}

impl Edit {
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
    /// worksheet's `drawing` element) when it has none.
    fn drawing(&mut self, sheet: usize) -> Result<PackURI> {
        let worksheet = self.staged.worksheet(sheet)?;
        if let Some(drawing) = objects::drawing_of(self.staged.opc().get_part(&worksheet)?.rels())?
        {
            return Ok(drawing);
        }
        let drawing = self.staged.free_name("/xl/drawings/drawing", ".xml")?;
        let layout = WorksheetLayout::scan(self.staged.blob(&worksheet)?)?;
        let opc = self.staged.opc_mut()?;
        opc.add_part(Box::new(BlobPart::new(
            drawing.clone(),
            ct::OFC_DRAWING.to_owned(),
            objects::EMPTY_DRAWING.as_bytes().to_vec(),
        )));
        let id = opc
            .get_part_mut(&worksheet)?
            .relate_to(&drawing.relative_ref(worksheet.base_uri()), rt::DRAWING);
        let xml =
            layout.insert_relationship_child(opc.get_part(&worksheet)?.blob(), "drawing", &id)?;
        opc.get_part_mut(&worksheet)?.set_blob(xml);
        self.staged.touch(&drawing);
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
        let opc = self.staged.opc_mut()?;
        opc.add_part(Box::new(BlobPart::new(
            media.clone(),
            spec.content_type.clone(),
            spec.image.clone(),
        )));
        let relationship_id = opc
            .get_part_mut(&drawing)?
            .relate_to(&media.relative_ref(drawing.base_uri()), rt::IMAGE);
        self.staged.touch(&media);
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

    /// Replace the image of the picture at `index` among the worksheet's
    /// pictures (record 0338): the bytes become a new media part with
    /// `content_type` and `extension`, the picture's image relationship is
    /// retargeted to it, and the old media part is removed once nothing
    /// refers to it. The picture's anchor, names and geometry keep their
    /// bytes. Returns whether anything changed: the same bytes and content
    /// type are an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing picture, an invalid image, a picture
    /// whose image relationship another object shares (replacing it would
    /// change both), or a signed workbook.
    ///
    /// ```
    /// use litchi_xlsx::Package;
    /// use litchi_xlsx::drawing::edit::{Marker, PictureSpec, Placement};
    ///
    /// # fn main() -> Result<(), litchi_xlsx::Error> {
    /// let placement = Placement::TwoCell { from: Marker::cell(0, 0), to: Marker::cell(2, 2) };
    /// let spec = PictureSpec {
    ///     placement,
    ///     name: "Picture 1".into(),
    ///     description: None,
    ///     title: None,
    ///     image: b"first".to_vec(),
    ///     content_type: "image/png".into(),
    ///     extension: "png".into(),
    ///     rotation: None,
    ///     flip_h: false,
    ///     flip_v: false,
    /// };
    /// let blank = Package::new()?;
    /// let mut edit = blank.edit_drawings()?;
    /// edit.add_picture(0, &spec)?;
    /// let package = edit.commit()?.into_package();
    /// let mut edit = package.edit_drawings()?;
    /// assert!(edit.replace_image(0, 0, b"second".to_vec(), "image/jpeg", "jpeg")?);
    /// let after = edit.commit()?.into_package();
    /// let image = after.drawing_elements(0)?[0].image.clone().expect("image");
    /// assert_eq!(image.content_type, "image/jpeg");
    /// assert_eq!(image.size, 6);
    /// # Ok(())
    /// # }
    /// ```
    pub fn replace_image(
        &mut self,
        sheet: usize,
        index: usize,
        image: Vec<u8>,
        content_type: &str,
        extension: &str,
    ) -> Result<bool> {
        let limits = Limits::from_source_limits(self.staged.opc().read_limits());
        check_image(&image, content_type, extension, &limits)?;
        let elements = self.elements(sheet)?;
        let picture = elements
            .iter()
            .filter(|element| element.kind == ElementKind::Picture)
            .nth(index)
            .ok_or_else(|| invalid("the worksheet has no such picture"))?;
        self.replace_element_image(sheet, index, picture, image, content_type, extension)
    }

    /// [`Self::replace_image`] for a picture its caller has already read:
    /// `picture` is the element at `index` among the worksheet's pictures,
    /// as [`Package::drawing_elements`] or [`Self::elements`] returns it, so
    /// the drawing is not read a second time to find its image (unified-ops
    /// record 0347).
    ///
    /// The picture's image relationship is still resolved in the staged
    /// drawing: a `picture` whose image part the element at `index` does not
    /// name is refused, and nothing changes.
    ///
    /// # Errors
    ///
    /// As [`Self::replace_image`]; also an error when `picture` is not a
    /// picture with an embedded image.
    ///
    /// ```
    /// use litchi_xlsx::Package;
    /// use litchi_xlsx::drawing::edit::{Marker, PictureSpec, Placement};
    ///
    /// # fn main() -> Result<(), litchi_xlsx::Error> {
    /// let picture = |row, image: &[u8]| PictureSpec {
    ///     placement: Placement::TwoCell { from: Marker::cell(row, 0), to: Marker::cell(row + 1, 2) },
    ///     name: format!("Picture {row}"),
    ///     description: None,
    ///     title: None,
    ///     image: image.to_vec(),
    ///     content_type: "image/png".into(),
    ///     extension: "png".into(),
    ///     rotation: None,
    ///     flip_h: false,
    ///     flip_v: false,
    /// };
    /// let blank = Package::new()?;
    /// let mut edit = blank.edit_drawings()?;
    /// edit.add_picture(0, &picture(0, b"first"))?;
    /// edit.add_picture(0, &picture(4, b"other"))?;
    /// let package = edit.commit()?.into_package();
    /// let elements = package.drawing_elements(0)?;
    ///
    /// let mut edit = package.edit_drawings()?;
    /// assert!(edit.replace_element_image(0, 1, &elements[1], b"second".to_vec(), "image/jpeg", "jpeg")?);
    /// let after = edit.commit()?.into_package();
    /// let image = after.drawing_elements(0)?[1].image.clone().expect("image");
    /// assert_eq!((image.content_type.as_str(), image.size), ("image/jpeg", 6));
    ///
    /// // An element that is not the picture at the index is refused.
    /// let mut edit = package.edit_drawings()?;
    /// assert!(edit.replace_element_image(0, 1, &elements[0], b"second".to_vec(), "image/jpeg", "jpeg").is_err());
    /// # Ok(())
    /// # }
    /// ```
    pub fn replace_element_image(
        &mut self,
        sheet: usize,
        index: usize,
        picture: &Element,
        image: Vec<u8>,
        content_type: &str,
        extension: &str,
    ) -> Result<bool> {
        let limits = Limits::from_source_limits(self.staged.opc().read_limits());
        check_image(&image, content_type, extension, &limits)?;
        if picture.kind != ElementKind::Picture {
            return Err(invalid("the element is not a picture"));
        }
        let drawing = self.existing_drawing(sheet)?;
        let old = picture
            .image
            .as_ref()
            .ok_or_else(|| invalid("the picture has no embedded image"))?
            .part
            .clone();
        let old = PackURI::new(&old).map_err(|error| invalid(error.to_string()))?;
        {
            let part = self.staged.opc().get_part(&old)?;
            if part.blob() == image.as_slice() && part.content_type() == content_type {
                return Ok(false);
            }
        }
        // The relationships only this picture names; its image is the one
        // that targets the old media part.
        let released = objects::detach(
            self.staged.blob(&drawing)?,
            ElementKind::Picture,
            index,
            &limits,
        )?
        .released;
        let relationship = {
            let part = self.staged.opc().get_part(&drawing)?;
            released.into_iter().find(|id| {
                part.rels().get(id).is_some_and(|relationship| {
                    !relationship.is_external()
                        && relationship.reltype() == rt::IMAGE
                        && relationship.target_partname().ok().as_ref() == Some(&old)
                })
            })
        }
        .ok_or_else(|| {
            invalid(
                "the picture shares its image with another object; replacing it would change both",
            )
        })?;
        let media = self
            .staged
            .free_name("/xl/media/image", &format!(".{extension}"))?;
        let opc = self.staged.opc_mut()?;
        opc.add_part(Box::new(BlobPart::new(
            media.clone(),
            content_type.to_owned(),
            image,
        )));
        opc.get_part_mut(&drawing)?
            .rels_mut()
            .retarget(&relationship, media.relative_ref(drawing.base_uri()))?;
        self.staged.touch(&media);
        self.staged.touch(&drawing);
        let opc = self.staged.opc_mut()?;
        let referenced = opc.iter_parts().any(|part| {
            part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship.target_partname().ok().as_ref() == Some(&old)
            })
        });
        if !referenced {
            opc.remove_part(&old);
            self.staged.touch(&old);
        }
        Ok(true)
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
                opc.remove_part(&target);
                self.staged.touch(&target);
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
        Ok(Commit::new(self.staged.commit()?))
    }
}

#[cfg(test)]
mod tests {
    use super::{Change, ElementKind, Fill, Font, Marker, PictureSpec, Placement, ShapeSpec};
    use crate::Package;
    use crate::color::Rgb;

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
        shape.fill = Some(Fill::Solid(Rgb::from_hex("4472C4").unwrap()));
        shape.font = Font {
            bold: Some(true),
            color: Some(Rgb::from_hex("FFFFFF").unwrap()),
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
        let elements = reopened.drawing_elements(0).unwrap();
        assert_eq!(elements.len(), 2);
        let shape = &elements[0];
        assert_eq!(shape.kind, ElementKind::Shape);
        assert_eq!(shape.geometry.as_deref(), Some("ellipse"));
        assert_eq!(shape.text.as_deref(), Some("one\ntwo & three"));
        assert_eq!(shape.rotation, Some(30.0));
        assert!(shape.flip_h);
        assert_eq!(shape.font.bold, Some(true));
        assert_eq!(shape.font.size, Some(14.0));
        assert_eq!(shape.font.color, Some(Rgb::from_hex("FFFFFF").unwrap()));
        assert_eq!(shape.placement, Some(placement()));
        let picture = &elements[1];
        assert_eq!(picture.kind, ElementKind::Picture);
        assert_eq!(picture.description.as_deref(), Some("logo"));
        assert_eq!(picture.image.as_ref().unwrap().content_type, "image/png");
        assert!(
            package
                .apply_drawings(&patch)
                .unwrap()
                .package()
                .same_content(&after)
                .unwrap()
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
        assert!(empty.drawing_elements(0).unwrap().is_empty());
    }
}
