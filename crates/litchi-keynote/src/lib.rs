//! Keynote semantic value models.
//!
//! Archive parsing, object topology, and mutation remain owned by the Keynote
//! implementation. This crate owns the compact slide and show values used by
//! readers and higher-level presentation consumers.

#![forbid(unsafe_code)]

pub mod transition;

use litchi_iwa_text::TextStorage;

/// A Keynote slide and its extracted semantic content.
#[derive(Debug, Clone)]
pub struct Slide {
    /// Zero-based slide index.
    pub index: usize,
    /// Optional slide title.
    pub title: Option<String>,
    /// Text blocks on the slide.
    pub text_content: Vec<String>,
    /// Optional speaker notes.
    pub notes: Option<String>,
    /// Rich-text storages belonging to the slide.
    pub text_storages: Vec<TextStorage>,
    /// Build animations on the slide.
    pub builds: Vec<BuildAnimation>,
    /// Optional slide transition.
    pub transition: Option<SlideTransition>,
    /// Optional master-slide identifier.
    pub master_slide_id: Option<u64>,
}

impl Slide {
    /// Creates an empty slide at `index`.
    #[must_use]
    pub fn new(index: usize) -> Self {
        Self {
            index,
            title: None,
            text_content: Vec::new(),
            notes: None,
            text_storages: Vec::new(),
            builds: Vec::new(),
            transition: None,
            master_slide_id: None,
        }
    }

    /// Returns all non-empty text values in slide order.
    #[must_use]
    pub fn all_text(&self) -> Vec<String> {
        let mut all = Vec::with_capacity(
            usize::from(self.title.is_some())
                .saturating_add(self.text_content.len())
                .saturating_add(usize::from(self.notes.is_some()))
                .saturating_add(self.text_storages.len()),
        );
        if let Some(title) = &self.title {
            all.push(title.clone());
        }
        all.extend(self.text_content.iter().cloned());
        if let Some(notes) = &self.notes {
            all.push(notes.clone());
        }
        all.extend(
            self.text_storages
                .iter()
                .filter(|storage| !storage.is_empty())
                .map(|storage| storage.plain_text().to_owned()),
        );

        all
    }

    /// Returns all slide text joined with newlines.
    #[must_use]
    pub fn plain_text(&self) -> String {
        self.all_text().join("\n")
    }

    /// Returns whether the slide has no modeled content.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.title.is_none()
            && self.text_content.is_empty()
            && self.notes.is_none()
            && self.text_storages.is_empty()
    }

    /// Returns the number of build animations.
    #[must_use]
    pub const fn build_count(&self) -> usize {
        self.builds.len()
    }
}

/// A build animation attached to a slide.
#[derive(Debug, Clone)]
pub struct BuildAnimation {
    /// Animation kind.
    pub animation_type: BuildAnimationType,
    /// Target object identifier, when known.
    pub target_id: Option<u64>,
    /// Animation duration in seconds.
    pub duration: f32,
}

/// Supported semantic build animation kinds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BuildAnimationType {
    /// Appear.
    Appear,
    /// Dissolve.
    Dissolve,
    /// Move in.
    MoveIn,
    /// Scale.
    Scale,
    /// Fade and scale.
    FadeAndScale,
    /// An unrecognized producer-specific animation.
    Other,
}

impl BuildAnimationType {
    /// Returns a stable human-readable name.
    #[must_use]
    pub const fn name(self) -> &'static str {
        match self {
            Self::Appear => "Appear",
            Self::Dissolve => "Dissolve",
            Self::MoveIn => "Move In",
            Self::Scale => "Scale",
            Self::FadeAndScale => "Fade and Scale",
            Self::Other => "Other",
        }
    }
}

/// A slide transition effect.
#[derive(Debug, Clone)]
pub struct SlideTransition {
    /// Transition kind.
    pub transition_type: TransitionType,
    /// Transition duration in seconds.
    pub duration: f32,
}

/// Supported semantic slide transition kinds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TransitionType {
    /// No transition.
    None,
    /// Dissolve.
    Dissolve,
    /// Push.
    Push,
    /// Wipe.
    Wipe,
    /// Flip.
    Flip,
    /// Cube.
    Cube,
    /// An unrecognized producer-specific transition.
    Other,
}

impl TransitionType {
    /// Returns a stable human-readable name.
    #[must_use]
    pub const fn name(self) -> &'static str {
        match self {
            Self::None => "None",
            Self::Dissolve => "Dissolve",
            Self::Push => "Push",
            Self::Wipe => "Wipe",
            Self::Flip => "Flip",
            Self::Cube => "Cube",
            Self::Other => "Other",
        }
    }
}

/// A Keynote presentation container.
#[derive(Debug, Clone)]
pub struct Show {
    /// Optional presentation title.
    pub title: Option<String>,
    /// Slides in presentation order.
    pub slides: Vec<Slide>,
    /// Slide size in points.
    pub slide_size: Option<(f32, f32)>,
    /// Whether the presentation auto-plays.
    pub auto_play: bool,
    /// Whether the presentation loops.
    pub loop_presentation: bool,
}

impl Show {
    /// Creates an empty show with the standard 4:3 slide size.
    #[must_use]
    pub fn new() -> Self {
        Self {
            title: None,
            slides: Vec::new(),
            slide_size: Some((1024.0, 768.0)),
            auto_play: false,
            loop_presentation: false,
        }
    }

    /// Appends a slide to the presentation.
    pub fn add_slide(&mut self, slide: Slide) {
        self.slides.push(slide);
    }

    /// Returns the total number of slides.
    #[must_use]
    pub const fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Returns a slide by zero-based index.
    #[must_use]
    pub fn get_slide(&self, index: usize) -> Option<&Slide> {
        self.slides.get(index)
    }

    /// Returns all show and slide text in presentation order.
    #[must_use]
    pub fn all_text(&self) -> Vec<String> {
        let mut all_text = Vec::with_capacity(
            usize::from(self.title.is_some())
                .saturating_add(self.slides.iter().map(|slide| slide.all_text().len()).sum()),
        );

        if let Some(title) = &self.title {
            all_text.push(title.clone());
        }
        for slide in &self.slides {
            all_text.extend(slide.all_text());
        }

        all_text
    }

    /// Returns whether the show contains no slides.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.slides.is_empty()
    }
}

impl Default for Show {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn slide_models_preserve_text_order() {
        let mut slide = Slide::new(0);
        slide.title = Some("Introduction".to_owned());
        slide.text_content.push("Point 1".to_owned());
        slide.notes = Some("Speaker notes".to_owned());
        slide
            .text_storages
            .push(TextStorage::from_text("Storage".to_owned()));

        assert_eq!(
            slide.all_text(),
            ["Introduction", "Point 1", "Speaker notes", "Storage"]
        );
        assert_eq!(slide.build_count(), 0);
    }

    #[test]
    fn animation_and_transition_names_are_stable() {
        assert_eq!(BuildAnimationType::MoveIn.name(), "Move In");
        assert_eq!(TransitionType::Dissolve.name(), "Dissolve");
    }

    #[test]
    fn show_owns_slides_and_text() {
        let mut show = Show::new();
        show.title = Some("My Presentation".to_owned());
        show.add_slide(Slide {
            title: Some("Slide 1".to_owned()),
            ..Slide::new(0)
        });

        assert_eq!(show.slide_count(), 1);
        assert_eq!(
            show.get_slide(0).and_then(|slide| slide.title.as_deref()),
            Some("Slide 1")
        );
        assert_eq!(show.all_text(), ["My Presentation", "Slide 1"]);
    }
}
