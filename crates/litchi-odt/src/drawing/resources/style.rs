//! Package and flat-document access to named drawing style resources.

use crate::{FlatDocument, Package};
use litchi_core::Result;
use litchi_odf_common::drawing::resources::{
    fill_image::{Collection as FillImageCollection, Image, Source, parse_drawing_fill_images},
    gradient::{Collection as GradientCollection, parse_drawing_gradients},
    hatch::{Collection as HatchCollection, parse_drawing_hatches},
    marker::{Collection as MarkerCollection, parse_drawing_markers},
    opacity::{Collection as OpacityCollection, parse_drawing_opacities},
    stroke_dash::{Collection as StrokeDashCollection, parse_drawing_stroke_dashes},
};
use std::borrow::Cow;

impl Package {
    /// Return named drawing fill-image resources from `styles.xml`.
    pub fn drawing_fill_images(&self) -> Result<FillImageCollection> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |styles| parse_drawing_fill_images(&styles),
        )
    }

    /// Load a safe package fill image or borrow inline bytes without copying.
    pub fn drawing_fill_image_bytes<'a>(&self, image: &'a Image) -> Result<Option<Cow<'a, [u8]>>> {
        match &image.source {
            Source::Inline { bytes, .. } => Ok(Some(Cow::Borrowed(bytes))),
            Source::Linked(link) => {
                let Some(path) = link.package_path() else {
                    return Ok(None);
                };
                if self.has_file(path)? {
                    self.get_file(path).map(Cow::Owned).map(Some)
                } else {
                    Ok(None)
                }
            },
            _ => Ok(None),
        }
    }

    /// Return named legacy and SVG drawing gradients from `styles.xml`.
    pub fn drawing_gradients(&self) -> Result<GradientCollection> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |styles| parse_drawing_gradients(&styles),
        )
    }

    /// Return named drawing hatch resources from `styles.xml`.
    pub fn drawing_hatches(&self) -> Result<HatchCollection> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |styles| parse_drawing_hatches(&styles),
        )
    }

    /// Return named drawing marker resources from `styles.xml`.
    pub fn drawing_markers(&self) -> Result<MarkerCollection> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |styles| parse_drawing_markers(&styles),
        )
    }

    /// Return named drawing opacity resources from `styles.xml`.
    pub fn drawing_opacities(&self) -> Result<OpacityCollection> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |styles| parse_drawing_opacities(&styles),
        )
    }

    /// Return named drawing stroke-dash resources from `styles.xml`.
    pub fn drawing_stroke_dashes(&self) -> Result<StrokeDashCollection> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |styles| parse_drawing_stroke_dashes(&styles),
        )
    }
}

impl FlatDocument {
    /// Return named drawing fill-image resources from the flat document.
    pub fn drawing_fill_images(&self) -> Result<FillImageCollection> {
        parse_drawing_fill_images(self.xml())
    }

    /// Return named legacy and SVG drawing gradients from the flat document.
    pub fn drawing_gradients(&self) -> Result<GradientCollection> {
        parse_drawing_gradients(self.xml())
    }

    /// Return named drawing hatch resources from the flat document.
    pub fn drawing_hatches(&self) -> Result<HatchCollection> {
        parse_drawing_hatches(self.xml())
    }

    /// Return named drawing marker resources from the flat document.
    pub fn drawing_markers(&self) -> Result<MarkerCollection> {
        parse_drawing_markers(self.xml())
    }

    /// Return named drawing opacity resources from the flat document.
    pub fn drawing_opacities(&self) -> Result<OpacityCollection> {
        parse_drawing_opacities(self.xml())
    }

    /// Return named drawing stroke-dash resources from the flat document.
    pub fn drawing_stroke_dashes(&self) -> Result<StrokeDashCollection> {
        parse_drawing_stroke_dashes(self.xml())
    }
}
