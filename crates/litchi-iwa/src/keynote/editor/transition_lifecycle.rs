//! Semantic read and delete operations for Keynote slide transitions.

use super::*;
use litchi_keynote::transition::Effect;

const TRANSITION_ANIMATION_TYPE: &str = "Transition";
const NO_EFFECT_DURATION_SECONDS: f64 = 1.0;

impl KeynoteTransitionSettings {
    /// Whether the slide has a visible transition effect.
    pub fn has_effect(&self) -> bool {
        !matches!(self.effect, None | Some(Effect::None))
    }

    pub(super) fn without_effect(&self) -> Self {
        Self {
            animation_type: Some(TRANSITION_ANIMATION_TYPE.to_owned()),
            effect: Some(Effect::None),
            duration: Some(NO_EFFECT_DURATION_SECONDS),
            direction: None,
            delay: self.delay,
            is_automatic: self.is_automatic,
            animation_parameters: KeynoteTransitionAnimationParameters {
                random_number_seed: self.animation_parameters.random_number_seed,
                writing_direction_is_rtl: self.animation_parameters.writing_direction_is_rtl,
                ..KeynoteTransitionAnimationParameters::default()
            },
            custom_parameters: KeynoteTransitionCustomParameters::default(),
        }
    }
}

impl KeynoteEditor {
    /// Read one slide's modern transition settings.
    pub fn slide_transition(
        &self,
        slide_index: usize,
    ) -> Result<Option<KeynoteTransitionSettings>> {
        let slides = self.slides()?;
        slides
            .get(slide_index)
            .map(|slide| slide.transition.clone())
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Keynote slide index {slide_index} is out of range for {} slides",
                    slides.len()
                ))
            })
    }

    /// Remove a slide's transition effect using Keynote's native `none` state.
    ///
    /// Start timing and the native random seed are retained. Effect-specific
    /// direction, animation payloads, and custom parameters are cleared. The
    /// return value is `true` when the package changed.
    pub fn clear_slide_transition(&mut self, slide_index: usize) -> Result<bool> {
        let Some(settings) = self.slide_transition(slide_index)? else {
            return Ok(false);
        };
        let cleared = settings.without_effect();
        if cleared == settings {
            return Ok(false);
        }
        self.set_slide_transition(slide_index, cleared)?;
        Ok(true)
    }
}
