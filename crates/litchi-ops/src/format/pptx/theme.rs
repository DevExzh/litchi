//! The presentation theme (record 0321, L-PPTX-R4).
//!
//! `/theme` reads the theme the presentation's colors and fonts come from
//! (the presentation part's own theme, else the first slide master's, as
//! OfficeCLI's `GetThemePart`): the color scheme's twelve slots, the major
//! (`headingFont`) and minor (`bodyFont`) Latin, East Asian and complex
//! script typefaces, and the color scheme's `name`. `set` replaces scheme
//! colors (opaque RGB), typefaces and the scheme name; the theme part keeps
//! every other byte. `add` and `remove` are not offered (OfficeCLI neither).
//!
//! Source owners are promoted to an eager package for a theme write, as for
//! charts and SmartArt: the source-backed edit stages slide parts only.

use std::borrow::Cow;
use std::sync::Arc;

use litchi_drawingml::theme::{Color, Face, FontSet, Palette, Slot, Theme};

use super::spec::{THEME_WRITE_TABLE, ThemeKey};
use super::structure::{bound, error};
use super::write::{Commit, Proof, native_snapshot_budget, opened_edit_options};
use super::{Editable, Pptx, PptxOwner};
use crate::node::Value;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, Props, Result};

/// The theme a node shows.
#[derive(Clone)]
pub(super) struct ThemeItem {
    pub(super) theme: Arc<Theme>,
}

impl Pptx {
    /// The presentation theme, or `None` when the deck has none.
    pub(super) fn presentation_theme(&self) -> Result<Option<Theme>> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                Ok(presentation.theme()?)
            },
            PptxOwner::Eager(package) => Ok(package.theme()?),
        })
    }
}

const fn slot(key: ThemeKey) -> Option<Slot> {
    Some(match key {
        ThemeKey::Dark1 => Slot::Dark1,
        ThemeKey::Light1 => Slot::Light1,
        ThemeKey::Dark2 => Slot::Dark2,
        ThemeKey::Light2 => Slot::Light2,
        ThemeKey::Accent1 => Slot::Accent1,
        ThemeKey::Accent2 => Slot::Accent2,
        ThemeKey::Accent3 => Slot::Accent3,
        ThemeKey::Accent4 => Slot::Accent4,
        ThemeKey::Accent5 => Slot::Accent5,
        ThemeKey::Accent6 => Slot::Accent6,
        ThemeKey::Hyperlink => Slot::Hyperlink,
        ThemeKey::FollowedHyperlink => Slot::FollowedHyperlink,
        _ => return None,
    })
}

/// A typeface key's face (`true`: major) and script.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Writing {
    Latin,
    EastAsian,
    ComplexScript,
}

const fn typeface(key: ThemeKey) -> Option<(bool, Writing)> {
    Some(match key {
        ThemeKey::HeadingFont => (true, Writing::Latin),
        ThemeKey::BodyFont => (false, Writing::Latin),
        ThemeKey::HeadingFontEastAsian => (true, Writing::EastAsian),
        ThemeKey::BodyFontEastAsian => (false, Writing::EastAsian),
        ThemeKey::HeadingFontComplexScript => (true, Writing::ComplexScript),
        ThemeKey::BodyFontComplexScript => (false, Writing::ComplexScript),
        _ => return None,
    })
}

fn face_text(face: &Face, script: Writing) -> &str {
    match script {
        Writing::Latin => &face.latin,
        Writing::EastAsian => &face.east_asian,
        Writing::ComplexScript => &face.complex_script,
    }
}

/// The RGB of a hex color, `RRGGBB`.
fn hex_rgb(value: &str) -> Option<crate::value::Color> {
    if value.len() != 6 || !value.is_ascii() {
        return None;
    }
    let byte = |at: usize| u8::from_str_radix(&value[at..at + 2], 16).ok();
    Some(crate::value::Color::rgb(byte(0)?, byte(2)?, byte(4)?))
}

/// The value of the theme property at `index` (OfficeCLI's readback: a
/// system color reads as its last computed RGB; empty typefaces are not
/// reported).
pub(super) fn value(theme: &Theme, index: usize) -> Option<Value<'_>> {
    let key = ThemeKey::at(index)?;
    if let Some(slot) = slot(key) {
        return match theme.colors.color(slot)? {
            Color::Rgb(hex) => hex_rgb(hex).map(Value::Color),
            Color::System { last, .. } => last.as_deref().and_then(hex_rgb).map(Value::Color),
        };
    }
    if let Some((major, script)) = typeface(key) {
        let face = if major {
            theme.fonts.major()
        } else {
            theme.fonts.minor()
        };
        let text = face_text(face, script);
        return (!text.is_empty()).then_some(Value::Text(Cow::Borrowed(text)));
    }
    let name = theme.colors.name();
    (!name.is_empty()).then_some(Value::Text(Cow::Borrowed(name)))
}

/// The palette with `name` and the slots of `palette`, changed by `colors`.
fn palette(source: &Palette, name: Option<&str>, colors: &[(Slot, Color)]) -> Palette {
    let mut out = Palette::new(name.unwrap_or(source.name()));
    for slot in Slot::ALL {
        let color = colors
            .iter()
            .find(|(at, _)| *at == slot)
            .map(|(_, color)| color)
            .or_else(|| source.color(slot));
        if let Some(color) = color {
            out = out.with(slot, color.clone());
        }
    }
    out
}

fn rgb_hex(color: &crate::value::Color) -> Result<String> {
    match color.base() {
        crate::value::Base::Rgb(value) if value.alpha() == 255 && color.transforms().is_empty() => {
            Ok(format!(
                "{:02X}{:02X}{:02X}",
                value.red(),
                value.green(),
                value.blue()
            ))
        },
        _ => Err(error(
            Code::InvalidValue,
            "a theme color must be an opaque RGB color",
        )
        .with_suggestion("use a hex color such as 4472C4")),
    }
}

fn leaf_error(error: impl std::fmt::Display) -> crate::Error {
    crate::Error::new(Code::InvalidValue, error.to_string())
        .with_format(litchi_core::FileFormat::Pptx)
}

impl Pptx {
    /// A `set` on `/theme`; `None` for other targets.
    pub(super) fn try_set_theme(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        let steps = path.path().steps();
        if steps.len() != 1 || !steps[0].name().eq_ignore_ascii_case("theme") {
            return Ok(None);
        }
        // Resolve the path for its diagnostics (a predicate, a missing theme).
        let node = crate::path::resolve::get(&super::Nav::new(self), path)?;
        let super::node::Item::Theme(item) = node.item() else {
            return Ok(None);
        };
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let theme = Arc::clone(&item.theme);
        drop(node);
        let mut colors: Vec<(Slot, Color)> = Vec::new();
        let mut faces: Vec<(bool, Writing, String)> = Vec::new();
        let mut name = None;
        let mut staged = 0_usize;
        let mut seen = 0_u32;
        for (key, scalar) in props.iter() {
            let (index, entry) = THEME_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            let bit = 1_u32 << index;
            if seen & bit != 0 {
                return Err(error(
                    Code::InvalidArgument,
                    "a theme property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            seen |= bit;
            entry.prepare(Operation::Set, scalar, |key, value| {
                if let Some(Scalar::Text(text)) = value {
                    staged = staged.saturating_add(text.len());
                    bound(LimitKind::StagedBytes, staged, limits.max_staged_bytes())?;
                    super::write::validate_text(text)?;
                }
                if let Some(slot) = slot(key) {
                    let Some(Scalar::Color(color)) = value else {
                        unreachable!("theme color coercion")
                    };
                    let hex = rgb_hex(&color)?;
                    colors.push((slot, Color::rgb(&hex).map_err(leaf_error)?));
                } else if let Some((major, script)) = typeface(key) {
                    let text = match value {
                        Some(Scalar::Text(text)) => text,
                        _ => "",
                    };
                    // OfficeCLI: `none` and `default` clear a typeface.
                    let text = if text.eq_ignore_ascii_case("none")
                        || text.eq_ignore_ascii_case("default")
                    {
                        ""
                    } else {
                        text
                    };
                    faces.push((major, script, text.to_owned()));
                } else if let Some(Scalar::Text(text)) = value {
                    name = Some(text.to_owned());
                }
                Ok(())
            })?;
        }
        let source = self.editable_owner_for_write()?;
        let base = super::smartart::eager(&source, self.0.ooxml_limits)?;
        let mut native = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        if !colors.is_empty() || name.is_some() {
            let next = palette(&theme.colors, name.as_deref(), &colors);
            native
                .set_theme_colors(&next)
                .map_err(|error| match error {
                    litchi_pptx::Error::Drawing(error) => leaf_error(error),
                    error => crate::Error::from(error),
                })?;
        }
        if !faces.is_empty() {
            let mut major = theme.fonts.major().clone();
            let mut minor = theme.fonts.minor().clone();
            for (is_major, script, text) in faces {
                let face = if is_major { &mut major } else { &mut minor };
                match script {
                    Writing::Latin => face.latin = text,
                    Writing::EastAsian => face.east_asian = text,
                    Writing::ComplexScript => face.complex_script = text,
                }
            }
            let fonts = FontSet::new(theme.fonts.name(), major, minor);
            native
                .set_theme_fonts(&fonts)
                .map_err(|error| match error {
                    litchi_pptx::Error::Drawing(error) => leaf_error(error),
                    error => crate::Error::from(error),
                })?;
        }
        let (target, patch) = native.commit()?.into_parts();
        let target = Arc::new(target);
        let after = if patch.is_noop() {
            self.clone()
        } else {
            self.adopt_eager(Arc::clone(&target))
        };
        Ok(Some(Commit::new(
            after,
            source,
            Editable::Eager(target),
            Proof::Eager(patch),
            1,
        )))
    }
}
