//! Per-slide visibility for layout-provided title and body placeholders.

use super::*;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const PLACEHOLDER_MESSAGE_TYPE: u32 = 7;
const TITLE_PLACEHOLDER_FIELD: u32 = 5;
const BODY_PLACEHOLDER_FIELD: u32 = 6;

impl KeynoteSlideTextPlaceholder {
    const fn label(self) -> &'static str {
        match self {
            Self::Title => "title",
            Self::Body => "body",
        }
    }

    const fn reference_field(self) -> u32 {
        match self {
            Self::Title => TITLE_PLACEHOLDER_FIELD,
            Self::Body => BODY_PLACEHOLDER_FIELD,
        }
    }

    fn reference(self, slide: &kn::SlideArchive) -> Option<&tsp::Reference> {
        match self {
            Self::Title => slide.title_placeholder.as_ref(),
            Self::Body => slide.body_placeholder.as_ref(),
        }
    }

    fn visibility(self, info: &KeynoteSlideInfo) -> Option<bool> {
        match self {
            Self::Title => info.is_title_visible,
            Self::Body => info.is_body_visible,
        }
    }
}

impl KeynoteEditor {
    /// Show or hide a layout-provided title or body placeholder on one slide.
    ///
    /// Keynote retains the placeholder object and its text while hidden. Only
    /// the slide's drawable ownership and z-order lists are changed. Slides
    /// whose selected layout lacks the requested placeholder are rejected.
    pub fn set_slide_text_placeholder_visible(
        &mut self,
        slide_index: usize,
        placeholder: KeynoteSlideTextPlaceholder,
        visible: bool,
    ) -> Result<()> {
        let slides = self.slides()?;
        let info = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let graph = ObjectGraph::read(self.package())?;
        let slide: kn::SlideArchive =
            graph.decode_type(info.slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
        let placeholder_id = placeholder
            .reference(&slide)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote slide {slide_index} has no layout-provided {} placeholder",
                    placeholder.label()
                ))
            })?
            .identifier;
        graph.decode_type::<kn::PlaceholderArchive>(
            placeholder_id,
            PLACEHOLDER_MESSAGE_TYPE,
            "KN.PlaceholderArchive",
        )?;
        let archive_name = graph.archive_name(info.slide_id)?.to_owned();
        if graph.archive_name(placeholder_id)? != archive_name {
            return Err(Error::InvalidFormat(format!(
                "Keynote {} placeholder {placeholder_id} is outside slide component {}",
                placeholder.label(),
                info.slide_id
            )));
        }
        let current = placeholder_ownership::validate(
            slide_index,
            &slide,
            placeholder_id,
            placeholder.label(),
        )?;
        if current == visible {
            return Ok(());
        }

        let mut staged = self.package().clone();
        placeholder_ownership::patch(
            &mut staged,
            &archive_name,
            info.slide_id,
            placeholder.reference_field(),
            placeholder_id,
            visible,
            placeholder.label(),
        )?;
        let verified = Self::from_package(staged)?;
        let verified_info = verified
            .slides()?
            .into_iter()
            .nth(slide_index)
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Keynote slide disappeared during placeholder update".to_owned(),
                )
            })?;
        if placeholder.visibility(&verified_info) != Some(visible) {
            return Err(Error::InvalidFormat(format!(
                "Keynote {} placeholder visibility failed round-trip validation",
                placeholder.label()
            )));
        }
        *self = verified;
        Ok(())
    }

    /// Show or hide the layout-provided title placeholder on one slide.
    pub fn set_slide_title_visible(&mut self, slide_index: usize, visible: bool) -> Result<()> {
        self.set_slide_text_placeholder_visible(
            slide_index,
            KeynoteSlideTextPlaceholder::Title,
            visible,
        )
    }

    /// Show or hide the layout-provided body placeholder on one slide.
    pub fn set_slide_body_visible(&mut self, slide_index: usize, visible: bool) -> Result<()> {
        self.set_slide_text_placeholder_visible(
            slide_index,
            KeynoteSlideTextPlaceholder::Body,
            visible,
        )
    }
}

pub(super) fn validate_placeholder_ownership(
    slide_index: usize,
    slide: &kn::SlideArchive,
    placeholder_id: u64,
    placeholder: KeynoteSlideTextPlaceholder,
) -> Result<bool> {
    placeholder_ownership::validate(slide_index, slide, placeholder_id, placeholder.label())
}
