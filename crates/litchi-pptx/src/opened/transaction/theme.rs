//! The presentation theme's color scheme and font scheme (record 0321).
//!
//! [`Edit::theme`] reads the theme the presentation's colors and fonts come
//! from (the presentation part's own theme, else the first slide master's);
//! [`Edit::set_theme_colors`] and [`Edit::set_theme_fonts`] replace its
//! `a:clrScheme` or `a:fontScheme` in place, keeping every other byte of the
//! theme part.

use litchi_drawingml::theme::{FontSet, Palette, Theme};

use super::{Edit, invalid};
use crate::Result;

impl Edit {
    /// The staged presentation theme, or `None` when the presentation has
    /// none.
    ///
    /// ```
    /// let source = litchi_pptx::Package::new()?;
    /// let mut edit = source.edit()?;
    /// assert!(edit.theme()?.is_some());
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed presentation graph or theme part.
    pub fn theme(&mut self) -> Result<Option<Theme>> {
        self.ensure_ready()?;
        let working = self.working_ref()?;
        let Some(name) = crate::shape::theme::presentation_theme(working)? else {
            return Ok(None);
        };
        crate::shape::theme::load(working, name.as_str()).map(Some)
    }

    /// Replace the presentation theme's color scheme (`a:clrScheme`),
    /// keeping the rest of the theme part byte for byte. Returns whether
    /// anything changed; a palette naming the same colors (in any slot
    /// order) is an exact no-op.
    ///
    /// ```
    /// use litchi_pptx::shape::theme::{Color, Slot};
    /// let source = litchi_pptx::Package::new()?;
    /// let mut edit = source.edit()?;
    /// let colors = edit.theme()?.unwrap().colors.with(Slot::Accent1, Color::rgb("FF0000")?);
    /// assert!(edit.set_theme_colors(&colors)?);
    /// let package = edit.commit()?.into_package();
    /// assert_eq!(package.theme()?.unwrap().colors, colors);
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006).
    ///
    /// # Errors
    ///
    /// Returns an error when the presentation has no theme, the palette is
    /// invalid, or the source scheme holds content the typed model cannot
    /// carry (it is then left untouched).
    pub fn set_theme_colors(&mut self, colors: &Palette) -> Result<bool> {
        self.ensure_ready()?;
        let name = self.theme_name()?;
        let current = crate::shape::theme::load(&self.working, name.as_str())?;
        if same_palette(&current.colors, colors) {
            return Ok(false);
        }
        let before = self.working.get_part(&name)?.blob_arc();
        let mut candidate = self.working.clone();
        crate::shape::theme::package::put_colors(&mut candidate, name.as_str(), colors)?;
        let changed = candidate.get_part(&name)?.blob() != before.as_slice();
        if changed {
            self.working = candidate;
        }
        Ok(changed)
    }

    /// Replace the presentation theme's font scheme (`a:fontScheme`),
    /// keeping the rest of the theme part byte for byte. Returns whether
    /// anything changed; an equal font set is an exact no-op.
    ///
    /// Staging a change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006).
    ///
    /// # Errors
    ///
    /// Returns an error when the presentation has no theme, the font set is
    /// invalid, or the source scheme holds content the typed model cannot
    /// carry (it is then left untouched).
    pub fn set_theme_fonts(&mut self, fonts: &FontSet) -> Result<bool> {
        self.ensure_ready()?;
        let name = self.theme_name()?;
        if crate::shape::theme::load(&self.working, name.as_str())?.fonts == *fonts {
            return Ok(false);
        }
        let before = self.working.get_part(&name)?.blob_arc();
        let mut candidate = self.working.clone();
        crate::shape::theme::package::put_fonts(&mut candidate, name.as_str(), fonts)?;
        let changed = candidate.get_part(&name)?.blob() != before.as_slice();
        if changed {
            self.working = candidate;
        }
        Ok(changed)
    }

    fn theme_name(&self) -> Result<litchi_opc::PackURI> {
        crate::shape::theme::presentation_theme(&self.working)?
            .ok_or_else(|| invalid("the presentation has no theme"))
    }
}

/// Whether two palettes name the same colors, whatever their slot order.
fn same_palette(left: &Palette, right: &Palette) -> bool {
    left.name() == right.name()
        && litchi_drawingml::theme::Slot::ALL
            .into_iter()
            .all(|slot| left.color(slot) == right.color(slot))
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use litchi_drawingml::theme::{Color, Face, FontSet, Slot};

    #[test]
    fn theme_colors_and_fonts_round_trip_and_no_op_is_exact() {
        let source = crate::Package::new().unwrap();
        let theme = source.theme().unwrap().unwrap();
        let mut edit = source.edit().unwrap();
        assert!(!edit.set_theme_colors(&theme.colors).unwrap());
        assert!(!edit.set_theme_fonts(&theme.fonts).unwrap());
        let colors = theme
            .colors
            .clone()
            .with(Slot::Accent2, Color::rgb("00FF00").unwrap());
        let fonts = FontSet::new(
            theme.fonts.name(),
            Face::new("Georgia"),
            theme.fonts.minor().clone(),
        );
        assert!(edit.set_theme_colors(&colors).unwrap());
        assert!(edit.set_theme_fonts(&fonts).unwrap());
        let package = edit.commit().unwrap().into_package();
        let read = package.theme().unwrap().unwrap();
        assert_eq!(read.colors, colors);
        assert_eq!(read.fonts.major().latin, "Georgia");
        assert_eq!(read.fonts.minor(), theme.fonts.minor());
    }
}
