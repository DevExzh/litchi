//! Package and flat-document access to named drawing style resources.

use crate::{
    FlatOpenDocument, OpenDocumentPackage,
    drawing_gradient::{OdfDrawingGradients, parse_drawing_gradients},
    drawing_hatch::{OdfDrawingHatches, parse_drawing_hatches},
    drawing_stroke_dash::{OdfDrawingStrokeDashes, parse_drawing_stroke_dashes},
};
use litchi_core::Result;

impl OpenDocumentPackage {
    /// Return named legacy and SVG drawing gradients from `styles.xml`.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites, load external data, or render gradients.
    pub fn drawing_gradients(&self) -> Result<OdfDrawingGradients> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |styles| parse_drawing_gradients(&styles),
        )
    }

    /// Return named drawing hatch resources from `styles.xml`.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render hatches.
    pub fn drawing_hatches(&self) -> Result<OdfDrawingHatches> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |styles| parse_drawing_hatches(&styles),
        )
    }

    /// Return named drawing stroke-dash resources from `styles.xml`.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render strokes.
    pub fn drawing_stroke_dashes(&self) -> Result<OdfDrawingStrokeDashes> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |styles| parse_drawing_stroke_dashes(&styles),
        )
    }
}

impl FlatOpenDocument {
    /// Return named legacy and SVG drawing gradients from the flat document.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites, load external data, or render gradients.
    pub fn drawing_gradients(&self) -> Result<OdfDrawingGradients> {
        parse_drawing_gradients(self.xml())
    }

    /// Return named drawing hatch resources from the flat document.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render hatches.
    pub fn drawing_hatches(&self) -> Result<OdfDrawingHatches> {
        parse_drawing_hatches(self.xml())
    }

    /// Return named drawing stroke-dash resources from the flat document.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render strokes.
    pub fn drawing_stroke_dashes(&self) -> Result<OdfDrawingStrokeDashes> {
        parse_drawing_stroke_dashes(self.xml())
    }
}
