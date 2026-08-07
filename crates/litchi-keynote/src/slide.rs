//! Immutable Keynote slide values and detached builders.

pub mod audio;
pub mod image;
pub mod media;
pub mod movie;
pub mod table;

use std::collections::TryReserveError;

use litchi_iwa_text::storage::Storage;

use crate::{Build, Effect, Seconds, SlideSelector};

/// A semantic transition attached to one slide.
#[derive(Debug, Clone, PartialEq)]
pub struct Transition {
    effect: Effect,
    duration: Seconds,
}

impl Transition {
    /// Construct a transition from validated semantic values.
    #[must_use]
    pub const fn new(effect: Effect, duration: Seconds) -> Self {
        Self { effect, duration }
    }

    /// Return the transition effect.
    #[must_use]
    pub const fn effect(&self) -> &Effect {
        &self.effect
    }

    /// Return the transition duration.
    #[must_use]
    pub const fn duration(&self) -> Seconds {
        self.duration
    }
}

/// An immutable semantic slide snapshot.
#[derive(Debug, Clone, PartialEq)]
pub struct Slide {
    index: usize,
    is_skipped: bool,
    name: Option<Box<str>>,
    title: Option<Box<str>>,
    text_content: Box<[String]>,
    notes: Option<Box<str>>,
    text_storages: Box<[Storage]>,
    builds: Box<[Build]>,
    transition: Option<Transition>,
}

impl Slide {
    /// Start a detached builder for a zero-based slide position.
    #[must_use]
    pub fn builder(index: usize) -> Builder {
        Builder::new(index)
    }

    /// Return the zero-based position in the semantic show snapshot.
    #[must_use]
    pub const fn index(&self) -> usize {
        self.index
    }

    /// Return whether Keynote omits this slide during presentation playback.
    #[must_use]
    pub const fn is_skipped(&self) -> bool {
        self.is_skipped
    }

    /// Return the optional developer-facing navigator name.
    ///
    /// This is distinct from [`Self::title`], which is visible slide content.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Return a semantic selector, preferring the navigator name when present.
    ///
    /// Name resolution can report ambiguity when malformed or producer-authored
    /// input repeats a name. Use [`Self::position_selector`] when an
    /// unambiguous snapshot-local selector is required.
    #[must_use]
    pub fn selector(&self) -> SlideSelector<'_> {
        self.name
            .as_deref()
            .map_or_else(|| self.position_selector(), SlideSelector::name)
    }

    /// Return a typed selector for this slide's zero-based source position.
    #[must_use]
    pub const fn position_selector(&self) -> SlideSelector<'static> {
        SlideSelector::index(self.index)
    }

    /// Return the optional slide title.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Borrow text blocks in source order.
    #[must_use]
    pub fn text_content(&self) -> &[String] {
        &self.text_content
    }

    /// Return optional speaker notes.
    #[must_use]
    pub fn notes(&self) -> Option<&str> {
        self.notes.as_deref()
    }

    /// Borrow rich-text storages without copying them.
    #[must_use]
    pub fn text_storages(&self) -> &[Storage] {
        &self.text_storages
    }

    /// Borrow builds in presentation order.
    #[must_use]
    pub fn builds(&self) -> &[Build] {
        &self.builds
    }

    /// Return the optional slide transition.
    #[must_use]
    pub const fn transition(&self) -> Option<&Transition> {
        self.transition.as_ref()
    }

    /// Return all modeled non-empty text values in semantic order.
    #[must_use]
    pub fn all_text(&self) -> Vec<String> {
        let capacity = usize::from(self.title.is_some())
            .saturating_add(self.text_content.len())
            .saturating_add(usize::from(self.notes.is_some()))
            .saturating_add(self.text_storages.len());
        let mut all = Vec::with_capacity(capacity);
        if let Some(title) = &self.title {
            all.push(title.to_string());
        }
        all.extend(self.text_content.iter().cloned());
        all.extend(
            self.text_storages
                .iter()
                .filter(|storage| !storage.is_empty())
                .map(|storage| storage.text().to_owned()),
        );
        if let Some(notes) = &self.notes {
            all.push(notes.to_string());
        }
        all
    }

    /// Return all modeled text joined with newlines.
    #[must_use]
    pub fn plain_text(&self) -> String {
        self.all_text().join("\n")
    }

    /// Return whether the snapshot contains no modeled content.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.title.is_none()
            && self.text_content.is_empty()
            && self.notes.is_none()
            && self.text_storages.is_empty()
            && self.builds.is_empty()
            && self.transition.is_none()
    }
}

/// A detached, mutable slide builder.
#[derive(Debug, Default)]
pub struct Builder {
    index: usize,
    is_skipped: bool,
    name: Option<Box<str>>,
    title: Option<Box<str>>,
    text_content: Vec<String>,
    notes: Option<Box<str>>,
    text_storages: Vec<Storage>,
    builds: Vec<Build>,
    transition: Option<Transition>,
}

impl Builder {
    /// Create an empty builder at `index`.
    #[must_use]
    pub fn new(index: usize) -> Self {
        Self {
            index,
            ..Self::default()
        }
    }

    /// Set whether the detached slide is skipped during presentation playback.
    pub fn set_skipped(&mut self, is_skipped: bool) {
        self.is_skipped = is_skipped;
    }

    /// Set or clear the developer-facing navigator name.
    ///
    /// Empty producer names are normalized to absence. The name remains
    /// separate from visible title content.
    pub fn set_name(&mut self, name: Option<String>) {
        self.name = name
            .filter(|candidate| !candidate.is_empty())
            .map(String::into_boxed_str);
    }

    /// Set or clear the title without exposing mutable attached state.
    pub fn set_title(&mut self, title: Option<String>) {
        self.title = title.map(String::into_boxed_str);
    }

    /// Set or clear speaker notes.
    pub fn set_notes(&mut self, notes: Option<String>) {
        self.notes = notes.map(String::into_boxed_str);
    }

    pub(crate) fn try_reserve_text_storages(
        &mut self,
        additional: usize,
    ) -> Result<(), TryReserveError> {
        self.text_storages.try_reserve_exact(additional)
    }

    pub(crate) fn try_reserve_builds(&mut self, additional: usize) -> Result<(), TryReserveError> {
        self.builds.try_reserve_exact(additional)
    }

    /// Append one text block in source order.
    pub fn push_text(&mut self, text: String) {
        self.text_content.push(text);
    }

    /// Append one rich-text storage.
    pub fn push_text_storage(&mut self, storage: Storage) {
        self.text_storages.push(storage);
    }

    /// Append one build animation.
    pub fn push_build(&mut self, build: Build) {
        self.builds.push(build);
    }

    /// Set or clear the slide transition.
    pub fn set_transition(&mut self, transition: Option<Transition>) {
        self.transition = transition;
    }

    /// Finish the detached builder as an immutable snapshot.
    #[must_use]
    pub fn build(self) -> Slide {
        Slide {
            index: self.index,
            is_skipped: self.is_skipped,
            name: self.name,
            title: self.title,
            text_content: self.text_content.into_boxed_slice(),
            notes: self.notes,
            text_storages: self.text_storages.into_boxed_slice(),
            builds: self.builds.into_boxed_slice(),
            transition: self.transition,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::AnimationType;

    #[test]
    fn builder_reservations_preserve_slide_content() -> Result<(), TryReserveError> {
        let mut builder = Slide::builder(4);
        builder.try_reserve_text_storages(0)?;
        builder.try_reserve_builds(0)?;
        builder.try_reserve_text_storages(1)?;
        builder.try_reserve_builds(1)?;

        builder.set_title(Some("title".to_owned()));
        builder.set_notes(Some("notes".to_owned()));
        builder.push_text("first".to_owned());
        builder.push_text("second".to_owned());
        builder.push_text_storage(Storage::from_text("rich".to_owned()));
        builder.push_build(Build::new(AnimationType::Appear, Seconds::ZERO));

        let slide = builder.build();
        assert_eq!(slide.text_content(), ["first", "second"]);
        assert_eq!(slide.text_storages()[0].text(), "rich");
        assert_eq!(slide.builds()[0].animation_type(), &AnimationType::Appear);
        assert_eq!(
            slide.all_text(),
            ["title", "first", "second", "rich", "notes"]
        );
        Ok(())
    }
}
