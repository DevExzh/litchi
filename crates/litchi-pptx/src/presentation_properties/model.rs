//! Typed PresentationML property values.

use crate::{Error, Result};
use litchi_ooxml_common::xml::is_ncname;

use super::{MAX_EXTENSIONS, MAX_STRING};

/// Relationship projection used by the HTML publishing property.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HtmlTarget {
    pub relationship_id: String,
    pub target: Option<String>,
    pub relationship_type: Option<String>,
    pub external: Option<bool>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BrowserSupport {
    V3,
    V4,
    V3V4,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum WebScreenSize {
    S544x376,
    S640x480,
    S720x512,
    S800x600,
    S1024x768,
    S1152x882,
    S1152x900,
    S1280x1024,
    S1600x1200,
    S1800x1400,
    S1920x1200,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum WebColor {
    None,
    Browser,
    PresentationText,
    PresentationAccent,
    WhiteTextOnBlack,
    BlackTextOnWhite,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PrintOutput {
    Slides,
    Handouts1,
    Handouts2,
    Handouts3,
    Handouts4,
    Handouts6,
    Handouts9,
    Notes,
    Outline,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PrintColorMode {
    BlackWhite,
    Gray,
    Color,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SlideSelection {
    All,
    Range { start: u32, end: u32 },
    CustomShow(u32),
}

impl SlideSelection {
    /// Validate the typed selection constraints from `CT_SlideRange` and
    /// `CT_CustomShowId` before a caller snapshots it into XML.
    pub fn validate(&self) -> Result<()> {
        if let Self::Range { start, end } = self
            && start > end
        {
            return Err(invalid("slide range start exceeds end"));
        }
        Ok(())
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ShowMode {
    Present,
    Browse { show_scrollbar: Option<bool> },
    Kiosk { restart: Option<u32> },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ColorKind {
    ScRgb,
    Srgb,
    Hsl,
    System,
    Scheme,
    Preset,
}

/// DrawingML color plus its bounded source fragment.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Color {
    pub kind: ColorKind,
    pub attributes: Vec<(String, String)>,
    pub xml: Vec<u8>,
}

/// Extension payload preserved without interpreting unknown content.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OpaqueExtension {
    pub uri: String,
    pub xml: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Extension {
    DiscardImageEditData(bool),
    DefaultImageDpi(u32),
    ChartTrackingReferenceBased(bool),
    Math(crate::presentation_properties::math::Properties),
    Unknown(OpaqueExtension),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ShowExtension {
    BrowseMode { show_status: Option<bool> },
    LaserColor(Color),
    ShowMediaControls(bool),
    Unknown(OpaqueExtension),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HtmlPublish {
    pub show_speaker_notes: Option<bool>,
    pub browser: Option<BrowserSupport>,
    pub target: HtmlTarget,
    pub slides: SlideSelection,
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct Web {
    pub show_animation: Option<bool>,
    pub resize_graphics: Option<bool>,
    pub allow_png: Option<bool>,
    pub rely_on_vml: Option<bool>,
    pub organize_in_folders: Option<bool>,
    pub use_long_filenames: Option<bool>,
    pub image_size: Option<WebScreenSize>,
    pub encoding: Option<String>,
    pub color: Option<WebColor>,
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct Print {
    pub output: Option<PrintOutput>,
    pub color_mode: Option<PrintColorMode>,
    pub hidden_slides: Option<bool>,
    pub scale_to_fit_paper: Option<bool>,
    pub frame_slides: Option<bool>,
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct Show {
    pub loop_show: Option<bool>,
    pub show_narration: Option<bool>,
    pub show_animation: Option<bool>,
    pub use_timings: Option<bool>,
    pub mode: Option<ShowMode>,
    pub slides: Option<SlideSelection>,
    pub pen_color: Option<Color>,
    pub extensions: Vec<ShowExtension>,
}

#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct Properties {
    pub html_publish: Option<HtmlPublish>,
    pub web: Option<Web>,
    pub print: Option<Print>,
    pub show: Option<Show>,
    pub recent_colors: Vec<Color>,
    pub extensions: Vec<Extension>,
}

impl Properties {
    /// Validate the package-independent PresentationML property snapshot.
    ///
    /// XML fragments are checked by the codec, while this method covers the
    /// constraints that can be established from typed values alone. It is
    /// useful for callers that edit a loaded snapshot before serializing it.
    pub fn validate(&self) -> Result<()> {
        if self.recent_colors.len() > 10 {
            return Err(invalid("clrMru permits at most ten colors"));
        }
        if self.extensions.len() > MAX_EXTENSIONS {
            return Err(invalid("presentation extension count exceeds limit"));
        }
        validate_extension_uris(&self.extensions)?;

        if let Some(html) = &self.html_publish {
            if !is_ncname(&html.target.relationship_id) {
                return Err(invalid("HTML publish relationship ID is not an XML NCName"));
            }
            html.slides.validate()?;
        }
        if let Some(show) = &self.show {
            if show.extensions.len() > MAX_EXTENSIONS {
                return Err(invalid("slide-show extension count exceeds limit"));
            }
            if let Some(selection) = &show.slides {
                selection.validate()?;
            }
            validate_show_extension_uris(&show.extensions)?;
        }
        if let Some(web) = &self.web
            && web
                .encoding
                .as_ref()
                .is_some_and(|value| value.len() > MAX_STRING)
        {
            return Err(invalid("presentation property string exceeds 1 MiB"));
        }
        Ok(())
    }

    /// Borrow the typed document-level math defaults, if present.
    pub fn math(&self) -> Option<&crate::presentation_properties::math::Properties> {
        self.extensions
            .iter()
            .find_map(|extension| match extension {
                Extension::Math(value) => Some(value),
                _ => None,
            })
    }

    /// Replace the typed document-level math defaults and return the prior snapshot.
    pub fn replace_math(
        &mut self,
        value: crate::presentation_properties::math::Properties,
    ) -> Option<crate::presentation_properties::math::Properties> {
        if let Some(extension) = self
            .extensions
            .iter_mut()
            .find(|extension| matches!(extension, Extension::Math(_)))
        {
            let Extension::Math(previous) = std::mem::replace(extension, Extension::Math(value))
            else {
                unreachable!("math extension selector changed during replacement")
            };
            Some(previous)
        } else {
            self.extensions.push(Extension::Math(value));
            None
        }
    }

    /// Remove the typed document-level math defaults, if present.
    pub fn remove_math(&mut self) -> Option<crate::presentation_properties::math::Properties> {
        let index = self
            .extensions
            .iter()
            .position(|extension| matches!(extension, Extension::Math(_)))?;
        match self.extensions.remove(index) {
            Extension::Math(value) => Some(value),
            _ => unreachable!("math extension selector changed during removal"),
        }
    }
}

fn validate_extension_uris(values: &[Extension]) -> Result<()> {
    let mut seen = [false; 4];
    for value in values {
        let Some(slot) = (match value {
            Extension::DiscardImageEditData(_) => Some(0),
            Extension::DefaultImageDpi(_) => Some(1),
            Extension::ChartTrackingReferenceBased(_) => Some(2),
            Extension::Math(value) => {
                value.validate()?;
                Some(3)
            },
            Extension::Unknown(value) => {
                if value.uri.is_empty() || value.uri.len() > MAX_STRING {
                    return Err(invalid("opaque presentation extension URI is invalid"));
                }
                None
            },
        }) else {
            continue;
        };
        if std::mem::replace(&mut seen[slot], true) {
            return Err(invalid("duplicate typed presentation extension"));
        }
    }
    Ok(())
}

fn validate_show_extension_uris(values: &[ShowExtension]) -> Result<()> {
    let mut seen = [false; 3];
    for value in values {
        let Some(slot) = (match value {
            ShowExtension::BrowseMode { .. } => Some(0),
            ShowExtension::LaserColor(_) => Some(1),
            ShowExtension::ShowMediaControls(_) => Some(2),
            ShowExtension::Unknown(value) => {
                if value.uri.is_empty() || value.uri.len() > MAX_STRING {
                    return Err(invalid("opaque show extension URI is invalid"));
                }
                None
            },
        }) else {
            continue;
        };
        if std::mem::replace(&mut seen[slot], true) {
            return Err(invalid("duplicate typed slide-show extension"));
        }
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
