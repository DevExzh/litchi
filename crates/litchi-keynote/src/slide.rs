//! Immutable Keynote slide values and detached builders.

pub mod media;

use litchi_iwa_text::TextStorage;

use crate::{Build, Effect, Seconds};

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
#[derive(Debug, Clone)]
pub struct Slide {
    index: usize,
    title: Option<Box<str>>,
    text_content: Box<[String]>,
    notes: Option<Box<str>>,
    text_storages: Box<[TextStorage]>,
    builds: Box<[Build]>,
    transition: Option<Transition>,
}

impl PartialEq for Slide {
    fn eq(&self, other: &Self) -> bool {
        self.index == other.index
            && self.title == other.title
            && self.text_content == other.text_content
            && self.notes == other.notes
            && self.text_storages.iter().zip(&other.text_storages).all(
                |(left_storage, right_storage)| {
                    left_storage.text == right_storage.text
                        && left_storage.identifier == right_storage.identifier
                        && left_storage.runs.len() == right_storage.runs.len()
                        && left_storage.runs.iter().zip(&right_storage.runs).all(
                            |(left_run, right_run)| {
                                left_run.offset == right_run.offset
                                    && left_run.length == right_run.length
                                    && left_run.style == right_run.style
                            },
                        )
                },
            )
            && self.text_storages.len() == other.text_storages.len()
            && self.builds == other.builds
            && self.transition == other.transition
    }
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
    pub fn text_storages(&self) -> &[TextStorage] {
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
        if let Some(notes) = &self.notes {
            all.push(notes.to_string());
        }
        all.extend(
            self.text_storages
                .iter()
                .filter(|storage| !storage.is_empty())
                .map(|storage| storage.plain_text().to_owned()),
        );
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
    title: Option<Box<str>>,
    text_content: Vec<String>,
    notes: Option<Box<str>>,
    text_storages: Vec<TextStorage>,
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

    /// Set or clear the title without exposing mutable attached state.
    pub fn set_title(&mut self, title: Option<String>) {
        self.title = title.map(String::into_boxed_str);
    }

    /// Set or clear speaker notes.
    pub fn set_notes(&mut self, notes: Option<String>) {
        self.notes = notes.map(String::into_boxed_str);
    }

    /// Append one text block in source order.
    pub fn push_text(&mut self, text: String) {
        self.text_content.push(text);
    }

    /// Append one rich-text storage.
    pub fn push_text_storage(&mut self, storage: TextStorage) {
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
            title: self.title,
            text_content: self.text_content.into_boxed_slice(),
            notes: self.notes,
            text_storages: self.text_storages.into_boxed_slice(),
            builds: self.builds.into_boxed_slice(),
            transition: self.transition,
        }
    }
}
