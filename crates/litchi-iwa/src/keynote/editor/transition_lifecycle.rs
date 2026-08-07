//! Semantic read and delete operations for Keynote slide transitions.

use super::*;
use litchi_keynote::transition::{
    AnimationParameters, CustomParameters, Effect, Settings as TransitionSettings,
};

const TRANSITION_ANIMATION_TYPE: &str = "Transition";
const NO_EFFECT_DURATION_SECONDS: f64 = 1.0;

fn without_effect(settings: &TransitionSettings) -> Result<TransitionSettings> {
    let mut animation_parameters = AnimationParameters::new();
    animation_parameters
        .set_random_number_seed(settings.animation_parameters().random_number_seed());
    animation_parameters
        .set_writing_direction_is_rtl(settings.animation_parameters().writing_direction_is_rtl());

    let mut cleared = TransitionSettings::new();
    cleared
        .set_animation_type(Some(TRANSITION_ANIMATION_TYPE))
        .map_err(transition_error)?;
    cleared
        .set_effect(Some(Effect::None))
        .map_err(transition_error)?;
    cleared
        .set_duration(Some(NO_EFFECT_DURATION_SECONDS))
        .map_err(transition_error)?;
    cleared.set_direction(None);
    cleared
        .set_delay(settings.delay())
        .map_err(transition_error)?;
    cleared.set_is_automatic(settings.is_automatic());
    cleared
        .set_animation_parameters(animation_parameters)
        .map_err(transition_error)?;
    cleared
        .set_custom_parameters(CustomParameters::new())
        .map_err(transition_error)?;
    Ok(cleared)
}

fn transition_error(error: litchi_keynote::Error) -> Error {
    Error::ParseError(format!(
        "invalid Keynote transition clear settings: {error}"
    ))
}

impl KeynoteEditor {
    /// Read one slide's modern transition settings.
    pub fn slide_transition(&self, slide_index: usize) -> Result<Option<TransitionSettings>> {
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
        let cleared = without_effect(&settings)?;
        if cleared == settings {
            return Ok(false);
        }
        self.set_slide_transition(slide_index, cleared)?;
        Ok(true)
    }
}
