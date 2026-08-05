//! Migration adapter for the standalone PresentationML theme owner.
//!
//! The owner API is contextual (`shape::theme::Palette`, `Slot`, `Face`, and
//! `Override`). These host spellings are retained only so the package facade
//! can be migrated independently; no semantic implementation remains here.

use crate::error::{OoxmlError, Result};
use litchi_opc::{OpcPackage, PackURI};

pub use litchi_pptx::shape::theme::{
    Authored as AuthoredTheme, Color as ThemeColorValue, Face as ThemeFontFace,
    FontSet as ThemeFontScheme, Override as ThemeOverride, Palette as ThemeColorScheme,
    Script as ThemeScriptFont, Slot as ThemeColorSlot, System as SystemColorKind,
};

pub fn add_theme(
    package: &mut OpcPackage,
    name: &str,
    colors: &ThemeColorScheme,
    fonts: &ThemeFontScheme,
) -> Result<AuthoredTheme> {
    litchi_pptx::shape::theme::add(package, name, colors, fonts).map_err(OoxmlError::from)
}

pub fn attach_theme_to_master(
    package: &mut OpcPackage,
    master_name: &str,
    theme_name: &str,
) -> Result<String> {
    litchi_pptx::shape::theme::attach(package, master_name, theme_name).map_err(OoxmlError::from)
}

pub fn store_theme_color_scheme(
    package: &mut OpcPackage,
    theme_name: &str,
    colors: &ThemeColorScheme,
) -> Result<()> {
    litchi_pptx::shape::theme::put_colors(package, theme_name, colors).map_err(OoxmlError::from)
}

pub fn store_theme_font_scheme(
    package: &mut OpcPackage,
    theme_name: &str,
    fonts: &ThemeFontScheme,
) -> Result<()> {
    litchi_pptx::shape::theme::put_fonts(package, theme_name, fonts).map_err(OoxmlError::from)
}

pub fn store_theme_override(
    package: &mut OpcPackage,
    parent_name: &str,
    value: &ThemeOverride,
) -> Result<String> {
    litchi_pptx::shape::theme::put_override(package, parent_name, value).map_err(OoxmlError::from)
}

pub fn theme_override(package: &OpcPackage, parent_name: &str) -> Result<Option<ThemeOverride>> {
    litchi_pptx::shape::theme::load_override(package, parent_name).map_err(OoxmlError::from)
}

pub fn remove_theme_override(package: &mut OpcPackage, parent_name: &str) -> Result<bool> {
    litchi_pptx::shape::theme::remove_override(package, parent_name).map_err(OoxmlError::from)
}

pub fn validate_theme_graph(package: &OpcPackage) -> Result<()> {
    litchi_pptx::shape::theme::validate(package).map_err(OoxmlError::from)
}

pub(crate) fn next_theme_part_uri(package: &OpcPackage) -> Result<PackURI> {
    litchi_pptx::shape::theme::next_part_uri(package, "/ppt/theme/theme", ".xml")
        .map_err(OoxmlError::from)
}
