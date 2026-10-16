//! Automatic text and paragraph styles of `content.xml`, and the character
//! formatting edits that derive new ones (record 0308, G-ODP-1).
//!
//! Character formatting in a presentation lives in automatic styles
//! (`office:automatic-styles`): a `text:span` names a `text`-family style,
//! a `text:p` a `paragraph`-family style, and either style's
//! `style:text-properties` holds the bold, italic, underline, size, colour
//! and font. The scene indexes those styles once while it scans; a
//! [`TextPropertyEdit`] never changes a style other markup may share, it
//! derives a new automatic style from the node's own and points the node at
//! it.

use std::ops::Range;

use litchi_core::{Error, Result};

/// Character properties an automatic style sets locally
/// (`style:text-properties`), as written.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TextProperties {
    pub(crate) font_weight: Option<String>,
    pub(crate) font_style: Option<String>,
    pub(crate) underline: Option<String>,
    pub(crate) font_size: Option<String>,
    pub(crate) color: Option<String>,
    pub(crate) font_name: Option<String>,
    pub(crate) font_family: Option<String>,
}

impl TextProperties {
    /// Whether the weight is bold (`fo:font-weight` `bold` or 600 and up).
    #[must_use]
    pub fn bold(&self) -> Option<bool> {
        self.font_weight.as_deref().map(|weight| match weight {
            "bold" => true,
            "normal" => false,
            numeric => numeric.parse::<u16>().is_ok_and(|weight| weight >= 600),
        })
    }

    /// Whether the posture is italic or oblique (`fo:font-style`).
    #[must_use]
    pub fn italic(&self) -> Option<bool> {
        self.font_style
            .as_deref()
            .map(|style| matches!(style, "italic" | "oblique"))
    }

    /// Whether text is underlined (`style:text-underline-style` other than
    /// `none`).
    #[must_use]
    pub fn underline(&self) -> Option<bool> {
        self.underline.as_deref().map(|style| style != "none")
    }

    /// `fo:font-size` as written: a length (`18pt`) or a percentage.
    #[must_use]
    pub fn size(&self) -> Option<&str> {
        self.font_size.as_deref()
    }

    /// `fo:color` as written (`#rrggbb`).
    #[must_use]
    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    /// The font: `fo:font-family` without its quotes, else the
    /// `style:font-name` font-face reference.
    #[must_use]
    pub fn font(&self) -> Option<&str> {
        self.font_family
            .as_deref()
            .map(|family| family.trim_matches(|character| character == '\'' || character == '"'))
            .or(self.font_name.as_deref())
    }

    /// Whether every property `edit` sets already has that value here.
    pub(crate) fn satisfies(&self, edit: &TextPropertyEdit) -> bool {
        edit.bold.is_none_or(|bold| self.bold() == Some(bold))
            && edit
                .italic
                .is_none_or(|italic| self.italic() == Some(italic))
            && edit
                .underline
                .is_none_or(|underline| self.underline() == Some(underline))
            && edit
                .size
                .as_deref()
                .is_none_or(|size| self.size() == Some(size))
            && edit.color.as_deref().is_none_or(|color| {
                self.color()
                    .is_some_and(|held| held.eq_ignore_ascii_case(color))
            })
            && edit
                .font
                .as_deref()
                .is_none_or(|font| self.font_name.is_none() && self.font() == Some(font))
    }
}

/// One `style:style` of `office:automatic-styles` in `content.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutomaticStyle {
    pub(crate) range: Range<usize>,
    pub(crate) start_tag: Range<usize>,
    pub(crate) empty: bool,
    pub(crate) name: String,
    pub(crate) family: String,
    pub(crate) parent: Option<String>,
    /// The start tag of the style's `style:text-properties`.
    pub(crate) text_tag: Option<Range<usize>>,
    /// Where a new `style:text-properties` child goes: after the last
    /// other properties child, or after the start tag.
    pub(crate) text_insertion: usize,
    pub(crate) text: TextProperties,
}

impl AutomaticStyle {
    /// The bytes of the whole element.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// `style:name`.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// `style:family` (`text`, `paragraph`, `graphic`, ...).
    #[must_use]
    pub fn family(&self) -> &str {
        &self.family
    }

    /// `style:parent-style-name`.
    #[must_use]
    pub fn parent(&self) -> Option<&str> {
        self.parent.as_deref()
    }

    /// The character properties the style sets locally.
    #[must_use]
    pub fn text_properties(&self) -> &TextProperties {
        &self.text
    }
}

/// Where new automatic styles go in `content.xml`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) enum StylesInsertion {
    /// Before this offset, the `</office:automatic-styles>` end tag.
    Before(usize),
    /// Replace this empty `<office:automatic-styles/>`.
    Expand(Range<usize>),
    /// Create the element at this offset (the `office:body` start tag).
    Create(usize),
    /// The content has no place for automatic styles.
    #[default]
    None,
}

/// Character properties to set on a run or paragraph. Each field left
/// `None` keeps what the node's style had.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct TextPropertyEdit {
    /// Bold (`fo:font-weight`, with its Asian and complex twins).
    pub bold: Option<bool>,
    /// Italic (`fo:font-style`, with its twins).
    pub italic: Option<bool>,
    /// Single underline, or none.
    pub underline: Option<bool>,
    /// The font size, an ODF length such as `18pt`.
    pub size: Option<String>,
    /// The colour, `#rrggbb`.
    pub color: Option<String>,
    /// The font family name.
    pub font: Option<String>,
}

impl TextPropertyEdit {
    /// An edit that changes nothing.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether the edit sets nothing.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }

    /// Merge `later` over this edit: its set fields win.
    pub fn merge(&mut self, later: &Self) {
        if later.bold.is_some() {
            self.bold = later.bold;
        }
        if later.italic.is_some() {
            self.italic = later.italic;
        }
        if later.underline.is_some() {
            self.underline = later.underline;
        }
        if later.size.is_some() {
            self.size.clone_from(&later.size);
        }
        if later.color.is_some() {
            self.color.clone_from(&later.color);
        }
        if later.font.is_some() {
            self.font.clone_from(&later.font);
        }
    }

    pub(crate) fn validate(&self) -> Result<()> {
        if let Some(size) = &self.size {
            let number = size.strip_suffix("pt").unwrap_or("");
            let valid = number
                .parse::<f64>()
                .is_ok_and(|value| value.is_finite() && value > 0.0 && value <= 4000.0);
            if !valid {
                return Err(Error::InvalidFormat(format!(
                    "ODP font sizes are positive point lengths such as `18pt`, not `{size}`"
                )));
            }
        }
        if let Some(color) = &self.color {
            let valid = color.len() == 7
                && color.starts_with('#')
                && color[1..].bytes().all(|byte| byte.is_ascii_hexdigit());
            if !valid {
                return Err(Error::InvalidFormat(format!(
                    "ODP colours are `#rrggbb`, not `{color}`"
                )));
            }
        }
        if let Some(font) = &self.font {
            let valid = !font.is_empty()
                && font.len() <= 256
                && !font
                    .chars()
                    .any(|character| character.is_control() || matches!(character, '\'' | '"'));
            if !valid {
                return Err(Error::InvalidFormat(
                    "ODP font names hold 1 to 256 bytes without quotes or control characters"
                        .to_string(),
                ));
            }
        }
        Ok(())
    }

    /// The `(local name with its prefix, value)` attribute changes of a
    /// `style:text-properties` start tag; `None` removes the attribute.
    pub(crate) fn attributes(&self, style: &str, fo: &str) -> Vec<(String, Option<String>)> {
        let mut out = Vec::new();
        let mut set = |prefix: &str, local: &str, value: Option<String>| {
            out.push((format!("{prefix}:{local}"), value));
        };
        if let Some(bold) = self.bold {
            let value = if bold { "bold" } else { "normal" };
            set(fo, "font-weight", Some(value.to_string()));
            set(style, "font-weight-asian", Some(value.to_string()));
            set(style, "font-weight-complex", Some(value.to_string()));
        }
        if let Some(italic) = self.italic {
            let value = if italic { "italic" } else { "normal" };
            set(fo, "font-style", Some(value.to_string()));
            set(style, "font-style-asian", Some(value.to_string()));
            set(style, "font-style-complex", Some(value.to_string()));
        }
        if let Some(underline) = self.underline {
            if underline {
                set(style, "text-underline-style", Some("solid".to_string()));
                set(style, "text-underline-width", Some("auto".to_string()));
                set(
                    style,
                    "text-underline-color",
                    Some("font-color".to_string()),
                );
            } else {
                set(style, "text-underline-style", Some("none".to_string()));
            }
        }
        if let Some(size) = &self.size {
            set(fo, "font-size", Some(size.clone()));
            set(style, "font-size-asian", Some(size.clone()));
            set(style, "font-size-complex", Some(size.clone()));
        }
        if let Some(color) = &self.color {
            set(fo, "color", Some(color.to_ascii_lowercase()));
        }
        if let Some(font) = &self.font {
            let quoted = if font.contains(' ') {
                format!("'{font}'")
            } else {
                font.clone()
            };
            set(fo, "font-family", Some(quoted));
            // A font-face reference would win over the family.
            set(style, "font-name", None);
        }
        out
    }
}
