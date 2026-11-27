//! Equation objects on slides.
//!
//! A legacy PowerPoint equation is an embedded OLE object ([MS-PPT]
//! 2.10.27 `ExOleEmbedContainer`, subtype Equation, with the ProgID
//! `Equation.3` for Equation Editor 3.0 or `Equation.DSMT4` for MathType)
//! whose storage (2.10.34 `ExOleObjStg`, often zlib-compressed) holds the
//! equation as MTEF in its `Equation Native` stream. A slide shows it
//! through a picture-frame shape (`msosptPictureFrame` with an
//! `ExObjRefAtom`) whose `pib` property names the preview picture in the
//! drawing group's picture store.
//!
//! [`Presentation::equations`](crate::Presentation::equations) lists the
//! equations of every slide. With the `equation` feature,
//! `Package::edit_equations` adds,
//! replaces and removes them; see `equation::Edit`.

mod content;
mod read;
mod tree;

#[cfg(feature = "equation")]
mod edit;
#[cfg(feature = "equation")]
mod family;
#[cfg(feature = "equation")]
mod object;
#[cfg(all(test, feature = "equation"))]
mod tests;

#[cfg(feature = "equation")]
pub use edit::{Edit, Source, Written};

#[cfg(feature = "equation")]
pub use family::{Commit, Diagnostics, Error, Patch, Result};
#[cfg(feature = "equation")]
use family::{apply_patch, published};

pub(crate) use content::Memo;
pub use content::{Native, Unreadable};

use crate::Presentation;

/// A shape's anchor rectangle: master units (576 per inch) on a slide, or
/// the group's coordinates for a shape inside a group.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Bounds {
    /// Left edge.
    pub left: i32,
    /// Top edge.
    pub top: i32,
    /// Right edge.
    pub right: i32,
    /// Bottom edge.
    pub bottom: i32,
}

impl Bounds {
    pub(crate) const fn from_edges(left: i32, top: i32, right: i32, bottom: i32) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    /// The width (`right - left`).
    #[must_use]
    pub const fn width(&self) -> i64 {
        self.right as i64 - self.left as i64
    }

    /// The height (`bottom - top`).
    #[must_use]
    pub const fn height(&self) -> i64 {
        self.bottom as i64 - self.top as i64
    }
}

/// One equation object shown on a slide.
///
/// Listing reads no storage: [`Equation::equation_native`] reads the
/// object's stream on first use, once per object, and an object whose
/// storage cannot be read is still listed with the reason.
#[derive(Clone, Debug)]
pub struct Equation {
    slide: usize,
    shape_id: u32,
    object_id: u32,
    persist_id: u32,
    program_id: Option<String>,
    bounds: Option<Bounds>,
    grouped: bool,
    compressed: bool,
    content: content::Content,
}

impl PartialEq for Equation {
    fn eq(&self, other: &Self) -> bool {
        self.slide == other.slide
            && self.shape_id == other.shape_id
            && self.object_id == other.object_id
            && self.persist_id == other.persist_id
            && self.program_id == other.program_id
            && self.bounds == other.bounds
            && self.grouped == other.grouped
            && self.compressed == other.compressed
            && self.equation_native() == other.equation_native()
    }
}

impl Eq for Equation {}

impl Equation {
    /// The zero-based position of its slide.
    #[must_use]
    pub const fn slide(&self) -> usize {
        self.slide
    }

    /// The shape identifier of its picture frame, unique on its slide.
    #[must_use]
    pub const fn shape_id(&self) -> u32 {
        self.shape_id
    }

    /// Its external-object identifier (`exObjId`).
    #[must_use]
    pub const fn object_id(&self) -> u32 {
        self.object_id
    }

    /// The persist identifier of its `ExOleObjStg` storage.
    #[must_use]
    pub const fn persist_id(&self) -> u32 {
        self.persist_id
    }

    /// The object's ProgID (`Equation.3`, `Equation.DSMT4`, ...).
    #[must_use]
    pub fn program_id(&self) -> Option<&str> {
        self.program_id.as_deref()
    }

    /// The frame's anchor, when it has one.
    #[must_use]
    pub const fn bounds(&self) -> Option<Bounds> {
        self.bounds
    }

    /// Whether the frame is inside a group (its bounds are then in the
    /// group's coordinates).
    #[must_use]
    pub const fn is_grouped(&self) -> bool {
        self.grouped
    }

    /// Whether the object's storage record is zlib-compressed.
    #[must_use]
    pub const fn is_compressed(&self) -> bool {
        self.compressed
    }

    /// The `Equation Native` stream, exactly as stored: the 28-byte
    /// `EQNOLEFILEHDR` followed by MTEF. Shared by every frame of the
    /// object.
    ///
    /// # Errors
    ///
    /// Returns [`Unreadable`] when the storage is missing, corrupt, holds no
    /// such stream, or would exceed the presentation's aggregate payload
    /// limit ([`crate::SourceLimits::max_aggregate_payload_bytes`]) together
    /// with the storages already read.
    pub fn equation_native(&self) -> Native {
        self.content.native(self.persist_id)
    }

    /// The equation as LaTeX.
    ///
    /// # Errors
    ///
    /// Returns the conversion error when the MTEF does not read within
    /// `limits`, and [`litchi_equation::Error::Malformed`] when the stream
    /// is [`Unreadable`].
    #[cfg(feature = "equation")]
    pub fn latex(&self, limits: &litchi_equation::Limits) -> litchi_equation::Result<String> {
        let native = self
            .equation_native()
            .map_err(|error| litchi_equation::Error::Malformed {
                format: litchi_equation::Format::Mtef,
                message: error.reason().to_owned(),
            })?;
        litchi_equation::mtef_to_latex(&native, limits)
    }
}

impl Presentation {
    /// Every equation object on every slide, in slide order and, within a
    /// slide, in drawing order (shapes inside groups included). An object
    /// counts as an equation when its definition is an embedded object of
    /// the Equation subtype or with an `Equation.*` ProgID. No storage is
    /// read; see [`Equation::equation_native`].
    ///
    /// # Errors
    ///
    /// Returns an error when a slide or the object list cannot be read.
    pub fn equations(&self) -> crate::package::Result<Vec<Equation>> {
        let Some(objects) = self.ole_objects()? else {
            return Ok(Vec::new());
        };
        let mut out = Vec::new();
        for slide in 0..self.slide_count() {
            out.extend(self.equations_with(slide, &objects)?);
        }
        Ok(out)
    }

    /// The equations of the slide at zero-based `slide`, as
    /// [`Self::equations`] lists them; empty for a position past the last
    /// slide.
    ///
    /// # Errors
    ///
    /// As [`Self::equations`].
    pub fn slide_equations(&self, slide: usize) -> crate::package::Result<Vec<Equation>> {
        if slide >= self.slide_count() {
            return Ok(Vec::new());
        }
        let Some(objects) = self.ole_objects()? else {
            return Ok(Vec::new());
        };
        self.equations_with(slide, &objects)
    }

    fn equations_with(
        &self,
        slide: usize,
        objects: &crate::embedded::object::Collection,
    ) -> crate::package::Result<Vec<Equation>> {
        let entry = &self.slide_directory().entries()[slide];
        let record = self.persisted_record_bytes(entry.persist_id())?;
        read::slide_equations(slide, record, objects, &mut |persist_id| {
            // The storage header alone says whether it is compressed; a
            // missing or foreign record makes only this equation unreadable.
            match self.persisted_record_bytes(persist_id) {
                Ok(record) if read::is_storage_record(record) => (
                    read::is_compressed_storage(record),
                    content::Content::Deferred(self.clone()),
                ),
                Ok(_) => (
                    false,
                    content::Content::Loaded(Err(Unreadable::new(
                        "the object's persist identifier does not name an ExOleObjStg",
                    ))),
                ),
                Err(error) => (
                    false,
                    content::Content::Loaded(Err(Unreadable::new(error.to_string()))),
                ),
            }
        })
    }
}
