//! Semantic Keynote slide-background access.

use super::slide_background_color::{color_from_native, validate_color};
use super::slide_background_gradient::{KeynoteGradient, validate_gradient};
use super::slide_background_gradient_wire::gradient_from_fill;
use super::*;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const SLIDE_STYLE_MESSAGE_TYPE: u32 = 9;

/// Effective background of one Keynote slide.
#[derive(Debug, Clone, PartialEq)]
pub enum KeynoteSlideBackground {
    /// Native explicit “No Fill”.
    None,
    /// Native RGB color fill.
    Solid(super::slide_background_color::KeynoteRgbaColor),
    /// Native simple or advanced linear/radial gradient fill.
    Gradient(KeynoteGradient),
    /// Byte-exact native fill for images and future or extended fill kinds.
    Opaque(Vec<u8>),
}

pub(super) struct ResolvedSlideBackground {
    pub(super) background: KeynoteSlideBackground,
    pub(super) fill_payload: Vec<u8>,
}

impl KeynoteEditor {
    /// Read the effective background, following the native slide-style parent chain.
    pub fn slide_background(&self, slide_index: usize) -> Result<KeynoteSlideBackground> {
        Ok(resolve_slide_background(self, slide_index)?.background)
    }

    /// Read the background stored directly on the slide's variation style.
    ///
    /// `None` means the slide inherits its effective background from its
    /// layout style. This is distinct from [`KeynoteSlideBackground::None`],
    /// which is an explicit native “No Fill” override.
    pub fn slide_background_override(
        &self,
        slide_index: usize,
    ) -> Result<Option<KeynoteSlideBackground>> {
        direct_slide_background_override(self, slide_index)
    }

    /// Set a slide background through a native, cullable slide-style variation.
    pub fn set_slide_background(
        &mut self,
        slide_index: usize,
        background: KeynoteSlideBackground,
    ) -> Result<()> {
        validate_background(&background)?;
        let resolved = resolve_slide_background(self, slide_index)?;
        if resolved.background == background {
            return Ok(());
        }
        super::slide_background_wire::set_slide_background(
            self,
            slide_index,
            background,
            &resolved.fill_payload,
        )
    }

    /// Delete a direct slide-background override and restore layout inheritance.
    ///
    /// Returns `true` when an override was removed and `false` when the slide
    /// already inherited its background.
    pub fn reset_slide_background(&mut self, slide_index: usize) -> Result<bool> {
        if direct_slide_background_override(self, slide_index)?.is_none() {
            return Ok(false);
        }
        super::slide_background_reset::reset_slide_background(self, slide_index)?;
        Ok(true)
    }
}

fn direct_slide_background_override(
    editor: &KeynoteEditor,
    slide_index: usize,
) -> Result<Option<KeynoteSlideBackground>> {
    let slides = editor.slides()?;
    let slide = slides.get(slide_index).ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote slide index {slide_index} is out of range for {} slides",
            slides.len()
        ))
    })?;
    let graph = ObjectGraph::read(editor.package())?;
    let native: kn::SlideArchive =
        graph.decode_type(slide.slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
    let style_id = native.style.identifier;
    let style: kn::SlideStyleArchive =
        graph.decode_type(style_id, SLIDE_STYLE_MESSAGE_TYPE, "KN.SlideStyleArchive")?;
    if style.super_.is_variation != Some(true) {
        return Ok(None);
    }
    let raw =
        graph.message_data_type(style_id, SLIDE_STYLE_MESSAGE_TYPE, "KN.SlideStyleArchive")?;
    let properties_payload = optional_length_delimited_payload(raw, 11)?;
    if properties_payload.is_some() != style.slide_properties.is_some() {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide style {style_id} has inconsistent slide-properties wire data"
        )));
    }
    let Some(properties_payload) = properties_payload else {
        return Ok(None);
    };
    let fill_payload = optional_length_delimited_payload(properties_payload, 1)?;
    if fill_payload.is_some()
        != style
            .slide_properties
            .as_ref()
            .is_some_and(|properties| properties.fill.is_some())
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide style {style_id} has inconsistent fill wire data"
        )));
    }
    fill_payload.map(background_from_fill).transpose()
}

pub(super) fn resolve_slide_background(
    editor: &KeynoteEditor,
    slide_index: usize,
) -> Result<ResolvedSlideBackground> {
    let slides = editor.slides()?;
    let slide = slides.get(slide_index).ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote slide index {slide_index} is out of range for {} slides",
            slides.len()
        ))
    })?;
    let graph = ObjectGraph::read(editor.package())?;
    let native: kn::SlideArchive =
        graph.decode_type(slide.slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
    let mut style_id = native.style.identifier;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(style_id) {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {slide_index} has a cyclic slide-style parent chain at {style_id}"
            )));
        }
        let style: kn::SlideStyleArchive =
            graph.decode_type(style_id, SLIDE_STYLE_MESSAGE_TYPE, "KN.SlideStyleArchive")?;
        let raw =
            graph.message_data_type(style_id, SLIDE_STYLE_MESSAGE_TYPE, "KN.SlideStyleArchive")?;
        let properties_payload = optional_length_delimited_payload(raw, 11)?;
        if properties_payload.is_some() != style.slide_properties.is_some() {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide style {style_id} has inconsistent slide-properties wire data"
            )));
        }
        if let (Some(properties), Some(properties_payload)) =
            (style.slide_properties.as_ref(), properties_payload)
        {
            let fill_payload = optional_length_delimited_payload(properties_payload, 1)?;
            if fill_payload.is_some() != properties.fill.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "Keynote slide style {style_id} has inconsistent fill wire data"
                )));
            }
            if let Some(fill_payload) = fill_payload {
                return Ok(ResolvedSlideBackground {
                    background: background_from_fill(fill_payload)?,
                    fill_payload: fill_payload.to_vec(),
                });
            }
        }
        let Some(parent) = style.super_.parent else {
            return Ok(ResolvedSlideBackground {
                background: KeynoteSlideBackground::None,
                fill_payload: Vec::new(),
            });
        };
        style_id = parent.identifier;
    }
}

pub(super) fn background_from_fill(fill_payload: &[u8]) -> Result<KeynoteSlideBackground> {
    let fill = tsd::FillArchive::decode(fill_payload)?;
    if fill_payload.is_empty()
        && fill.color.is_none()
        && fill.gradient.is_none()
        && fill.image.is_none()
    {
        return Ok(KeynoteSlideBackground::None);
    }
    if fill.color.is_none() && fill.gradient.is_some() && fill.image.is_none() {
        return Ok(match gradient_from_fill(fill_payload)? {
            Some(gradient) => KeynoteSlideBackground::Gradient(gradient),
            None => KeynoteSlideBackground::Opaque(fill_payload.to_vec()),
        });
    }
    let Some(color) = fill.color.as_ref() else {
        return Ok(KeynoteSlideBackground::Opaque(fill_payload.to_vec()));
    };
    if fill.gradient.is_some() || fill.image.is_some() {
        return Ok(KeynoteSlideBackground::Opaque(fill_payload.to_vec()));
    }
    Ok(match color_from_native(color) {
        Some(color) => KeynoteSlideBackground::Solid(color),
        None => KeynoteSlideBackground::Opaque(fill_payload.to_vec()),
    })
}

pub(super) fn validate_background(background: &KeynoteSlideBackground) -> Result<()> {
    match background {
        KeynoteSlideBackground::None => Ok(()),
        KeynoteSlideBackground::Solid(color) => validate_color(*color, "Keynote slide-background"),
        KeynoteSlideBackground::Gradient(gradient) => validate_gradient(gradient),
        KeynoteSlideBackground::Opaque(payload) => {
            if payload.is_empty() {
                return Err(Error::ParseError(
                    "an opaque Keynote slide background cannot encode native No Fill".to_owned(),
                ));
            }
            tsd::FillArchive::decode(payload.as_slice())?;
            Ok(())
        },
    }
}
