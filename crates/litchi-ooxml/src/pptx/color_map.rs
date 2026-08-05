//! Temporary host adapter for PresentationML color maps.
//!
//! The owner stores only the map-to-role relationship. Theme resolution stays
//! here until the host theme graph is migrated, because the old host theme
//! model is not part of this bounded semantic slice.

use crate::error::{OoxmlError, Result};
use crate::pptx::parts::{Theme, ThemeColor};

pub use litchi_pptx::presentation_properties::metadata::color_map::{
    Role as ThemeColorRole, Slot as ColorMapSlot,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ColorMap {
    inner: litchi_pptx::presentation_properties::metadata::color_map::Map,
}

impl ColorMap {
    pub const fn color(&self, slot: ColorMapSlot) -> ThemeColorRole {
        self.inner.color(slot)
    }

    pub fn resolve_theme_color<'a>(
        &self,
        theme: &'a Theme,
        slot: ColorMapSlot,
    ) -> Option<&'a ThemeColor> {
        let name = self.color(slot).as_str();
        theme.colors.iter().find(|color| color.name == name)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColorMapOverride {
    Master,
    Override(ColorMap),
}

pub(crate) fn parse_master_color_map(xml: &[u8]) -> Result<ColorMap> {
    litchi_pptx::presentation_properties::metadata::color_map::parse_master(xml)
        .map(|inner| ColorMap { inner })
        .map_err(OoxmlError::from)
}

pub(crate) fn parse_color_map_override(
    xml: &[u8],
    root_name: &[u8],
    root_label: &str,
) -> Result<Option<ColorMapOverride>> {
    litchi_pptx::presentation_properties::metadata::color_map::parse_override(
        xml, root_name, root_label,
    )
    .map(|value| {
        value.map(|value| match value {
            litchi_pptx::presentation_properties::metadata::color_map::Override::Master => {
                ColorMapOverride::Master
            },
            litchi_pptx::presentation_properties::metadata::color_map::Override::Override(map) => {
                ColorMapOverride::Override(ColorMap { inner: map })
            },
        })
    })
    .map_err(OoxmlError::from)
}
