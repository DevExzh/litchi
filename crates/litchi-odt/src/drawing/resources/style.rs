//! Package and flat-document access to named drawing style resources.

use crate::{
    FlatOpenDocument, OpenDocumentPackage,
    drawing::resources::gradient::{Collection as GradientCollection, parse_drawing_gradients},
    drawing::resources::hatch::{Collection as HatchCollection, parse_drawing_hatches},
    drawing::resources::stroke_dash::{
        Collection as StrokeDashCollection, parse_drawing_stroke_dashes,
    },
};
use litchi_core::Result;

impl OpenDocumentPackage {
    /// Return named legacy and SVG drawing gradients from `styles.xml`.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites, load external data, or render gradients.
    pub fn drawing_gradients(&self) -> Result<GradientCollection> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |styles| parse_drawing_gradients(&styles),
        )
    }

    /// Return named drawing hatch resources from `styles.xml`.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render hatches.
    pub fn drawing_hatches(&self) -> Result<HatchCollection> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |styles| parse_drawing_hatches(&styles),
        )
    }

    /// Return named drawing stroke-dash resources from `styles.xml`.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render strokes.
    pub fn drawing_stroke_dashes(&self) -> Result<StrokeDashCollection> {
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
    pub fn drawing_gradients(&self) -> Result<GradientCollection> {
        parse_drawing_gradients(self.xml())
    }

    /// Return named drawing hatch resources from the flat document.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render hatches.
    pub fn drawing_hatches(&self) -> Result<HatchCollection> {
        parse_drawing_hatches(self.xml())
    }

    /// Return named drawing stroke-dash resources from the flat document.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render strokes.
    pub fn drawing_stroke_dashes(&self) -> Result<StrokeDashCollection> {
        parse_drawing_stroke_dashes(self.xml())
    }
}
