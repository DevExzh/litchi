//! Semantic edits for the in-memory presentation snapshot.

use super::Builder;
use crate::{Slide, Transition};
use litchi_core::Result;

pub(super) fn titled_slide(index: usize, title: &str, text: &str) -> Slide {
    Slide {
        title: Some(title.to_string()),
        text: text.to_string(),
        index,
        notes: None,
        transition: None,
        animations: Vec::new(),
        legacy_animation: None,
        shapes: Vec::new(),
    }
}

pub(super) fn text_slide(index: usize, text: &str) -> Slide {
    Slide {
        title: None,
        text: text.to_string(),
        index,
        notes: None,
        transition: None,
        animations: Vec::new(),
        legacy_animation: None,
        shapes: Vec::new(),
    }
}

pub(super) fn set_slide_transition(
    slides: &mut [Slide],
    slide_index: usize,
    transition: Option<Transition>,
) -> Result<()> {
    let slide = slides.get_mut(slide_index).ok_or_else(|| {
        litchi_core::Error::InvalidFormat(format!("Slide index {slide_index} out of bounds"))
    })?;
    slide.transition = transition;
    Ok(())
}

impl Builder {
    /// Set or clear the typed transition attached to a slide.
    ///
    /// The transition is emitted as the slide's ODF drawing-page style. Pass
    /// `None` to remove the transition and restore the default drawing-page
    /// style.
    pub fn set_slide_transition(
        &mut self,
        slide_index: usize,
        transition: Option<Transition>,
    ) -> Result<&mut Self> {
        set_slide_transition(&mut self.slides, slide_index, transition)?;
        Ok(self)
    }

    pub(super) fn append_slide(&mut self, mut slide: Slide) -> Result<&mut Self> {
        slide.index = self.slides.len();
        self.slides.push(slide);
        Ok(self)
    }
}

#[cfg(test)]
mod tests {
    use super::Builder;
    use crate::{Presentation, Speed, Style, Transition, Type};

    #[test]
    fn sets_and_clears_a_typed_slide_transition() {
        let mut builder = Builder::new();
        builder.add_slide("Transition slide").unwrap();

        let mut transition = Transition::new();
        transition
            .set_transition_type(Some(Type::Automatic))
            .set_style(Some(Style::new("fade-from-left").unwrap()))
            .set_speed(Some(Speed::Fast));
        transition.set_duration(Some("PT2S")).unwrap();

        builder
            .set_slide_transition(0, Some(transition.clone()))
            .unwrap();
        assert_eq!(builder.slides[0].transition.as_ref(), Some(&transition));

        let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        let slide = presentation.slides().unwrap().remove(0);
        assert_eq!(slide.transition().unwrap(), &transition);

        let mut builder = Builder::new();
        builder.add_slide("Transition slide").unwrap();
        builder
            .set_slide_transition(0, Some(transition.clone()))
            .unwrap();
        assert!(
            builder
                .set_slide_transition(1, Some(transition.clone()))
                .is_err()
        );
        assert_eq!(builder.slides[0].transition.as_ref(), Some(&transition));
        builder.set_slide_transition(0, None).unwrap();
        assert!(builder.slides[0].transition.is_none());

        let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        let slide = presentation.slides().unwrap().remove(0);
        assert!(slide.transition().is_none());
    }
}
