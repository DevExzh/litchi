//! Immutable Keynote slide values and detached builders.

pub mod audio;
pub mod comment;
pub mod delete;
pub mod image;
pub mod media;
pub mod movie;
pub mod placeholder;
pub mod table;

use std::collections::TryReserveError;

use litchi_core::Position;
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
    movies: Box<[media::MovieInfo]>,
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

    /// Return the checked zero-based position in the semantic show snapshot.
    ///
    /// [`Self::index`] remains available as a compatibility accessor, while
    /// new selector-first code can keep collection positions typed all the
    /// way to the public boundary.
    #[must_use]
    pub const fn position(&self) -> Position {
        Position::new(self.index)
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
        SlideSelector::position(self.position())
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

    /// Borrow movie and audio drawables in their original source order.
    ///
    /// Audio-only controls are retained in this collection and can be
    /// distinguished with [`media::MovieInfo::is_audio`]. The values contain
    /// no native object or media-data identifiers.
    #[must_use]
    pub fn movies(&self) -> &[media::MovieInfo] {
        &self.movies
    }

    /// Borrow the same source-ordered collection through its media-oriented
    /// name.
    #[must_use]
    pub fn media(&self) -> &[media::MovieInfo] {
        self.movies()
    }

    /// Iterate over independently positioned audio controls in source order.
    pub fn audio(&self) -> impl Iterator<Item = &media::MovieInfo> {
        self.movies.iter().filter(|movie| movie.is_audio())
    }

    /// Iterate over non-audio movie drawables in source order.
    pub fn video_movies(&self) -> impl Iterator<Item = &media::MovieInfo> {
        self.movies.iter().filter(|movie| !movie.is_audio())
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
        let mut text = String::new();
        if let Some(length) = self.checked_text_len() {
            text.reserve(length);
        }
        self.append_plain_text(&mut text);
        text
    }

    fn checked_text_len(&self) -> Option<usize> {
        let mut length = 0usize;
        let mut values = 0usize;

        if let Some(title) = &self.title {
            length = length.checked_add(title.len())?;
            values = values.checked_add(1)?;
        }
        for content in &self.text_content {
            length = length.checked_add(content.len())?;
            values = values.checked_add(1)?;
        }
        for storage in &self.text_storages {
            if !storage.is_empty() {
                length = length.checked_add(storage.len())?;
                values = values.checked_add(1)?;
            }
        }
        if let Some(notes) = &self.notes {
            length = length.checked_add(notes.len())?;
            values = values.checked_add(1)?;
        }

        length.checked_add(values.saturating_sub(1))
    }

    fn append_plain_text(&self, output: &mut String) {
        let mut first = true;
        if let Some(title) = self.title.as_deref() {
            append_value(output, &mut first, title);
        }
        for content in &self.text_content {
            append_value(output, &mut first, content);
        }
        for storage in &self.text_storages {
            if !storage.is_empty() {
                append_value(output, &mut first, storage.text());
            }
        }
        if let Some(notes) = self.notes.as_deref() {
            append_value(output, &mut first, notes);
        }
    }

    /// Return whether the snapshot contains no modeled content.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.title.is_none()
            && self.text_content.is_empty()
            && self.notes.is_none()
            && self.text_storages.is_empty()
            && self.movies.is_empty()
            && self.builds.is_empty()
            && self.transition.is_none()
    }
}

fn append_value(output: &mut String, first: &mut bool, value: &str) {
    if !*first {
        output.push('\n');
    }
    output.push_str(value);
    *first = false;
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
    movies: Vec<media::MovieInfo>,
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

    pub(crate) fn try_reserve_movies(&mut self, additional: usize) -> Result<(), TryReserveError> {
        self.movies.try_reserve_exact(additional)
    }

    /// Append one text block in source order.
    pub fn push_text(&mut self, text: String) {
        self.text_content.push(text);
    }

    /// Append one rich-text storage.
    pub fn push_text_storage(&mut self, storage: Storage) {
        self.text_storages.push(storage);
    }

    /// Append one movie or audio drawable in source order.
    pub fn push_movie(&mut self, movie: media::MovieInfo) {
        self.movies.push(movie);
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
            movies: self.movies.into_boxed_slice(),
            builds: self.builds.into_boxed_slice(),
            transition: self.transition,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::AnimationType;
    use litchi_core::Position;

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
        builder.push_movie(media::MovieInfo::from_parts(
            media::MovieKind::File,
            Some(media::Point { x: 1.0, y: 2.0 }),
            None,
            None,
            None,
        ));
        builder.push_build(Build::new(AnimationType::Appear, Seconds::ZERO));

        let slide = builder.build();
        assert_eq!(slide.text_content(), ["first", "second"]);
        assert_eq!(slide.text_storages()[0].text(), "rich");
        assert_eq!(slide.movies().len(), 1);
        assert_eq!(slide.media(), slide.movies());
        assert_eq!(slide.video_movies().count(), 1);
        assert_eq!(slide.audio().count(), 0);
        assert_eq!(slide.builds()[0].animation_type(), &AnimationType::Appear);
        assert_eq!(
            slide.all_text(),
            ["title", "first", "second", "rich", "notes"]
        );
        Ok(())
    }

    #[test]
    fn plain_text_matches_all_text_join_for_empty_slide() {
        let slide = Slide::builder(0).build();

        assert_eq!(slide.plain_text(), slide.all_text().join("\n"));
        assert_eq!(slide.plain_text(), "");
        assert_eq!(slide.checked_text_len(), Some(0));
    }

    #[test]
    fn plain_text_matches_all_text_join_for_title_content_storages_and_notes() {
        let mut builder = Slide::builder(0);
        builder.set_title(Some("Title".to_owned()));
        builder.push_text("Content".to_owned());
        builder.push_text_storage(Storage::new());
        builder.push_text_storage(Storage::from_text("Storage".to_owned()));
        builder.set_notes(Some("Notes".to_owned()));
        let slide = builder.build();

        assert_eq!(slide.plain_text(), slide.all_text().join("\n"));
        assert_eq!(slide.plain_text(), "Title\nContent\nStorage\nNotes");
        assert_eq!(slide.checked_text_len(), Some(slide.plain_text().len()));
    }

    #[test]
    fn plain_text_preserves_separators_for_empty_modeled_values() {
        let mut builder = Slide::builder(0);
        builder.set_title(Some(String::new()));
        builder.push_text(String::new());
        builder.push_text_storage(Storage::new());
        builder.push_text_storage(Storage::from_text("Storage".to_owned()));
        builder.set_notes(Some(String::new()));
        let slide = builder.build();

        assert_eq!(slide.all_text(), ["", "", "Storage", ""]);
        assert_eq!(slide.plain_text(), slide.all_text().join("\n"));
        assert_eq!(slide.plain_text(), "\n\nStorage\n");
    }

    #[test]
    fn position_accessor_preserves_the_checked_semantic_position() {
        let slide = Slide::builder(4).build();

        assert_eq!(slide.position(), Position::new(4));
        assert_eq!(
            slide.position_selector(),
            SlideSelector::position(Position::new(4))
        );
    }
}
