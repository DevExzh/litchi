//! Package-independent custom-show values and typed snapshot edits.

use crate::{Error, Result};

/// A custom slide show definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Show {
    /// Unique ID for the custom show.
    pub id: u32,
    /// Display name of the custom show.
    pub name: String,
    /// List of presentation slide IDs included in source order.
    pub slide_ids: Vec<u32>,
}

impl Show {
    /// Create a new custom show.
    pub fn new(id: u32, name: impl Into<String>) -> Self {
        Self {
            id,
            name: name.into(),
            slide_ids: Vec::new(),
        }
    }

    /// Add a slide to the custom show.
    pub fn add_slide(&mut self, slide_id: u32) {
        self.slide_ids.push(slide_id);
    }

    /// Add multiple slides to the custom show.
    pub fn add_slides(&mut self, slide_ids: impl IntoIterator<Item = u32>) {
        self.slide_ids.extend(slide_ids);
    }

    /// Set slides with builder pattern.
    #[must_use]
    pub fn with_slides(mut self, slide_ids: Vec<u32>) -> Self {
        self.slide_ids = slide_ids;
        self
    }

    /// Get the number of slides in the custom show.
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.slide_ids.len()
    }
}

/// Collection of custom slide shows for a presentation.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct List {
    /// List of custom shows.
    pub shows: Vec<Show>,
    /// Next available ID for new shows.
    next_id: u32,
}

impl List {
    /// Create a new empty custom-show list.
    #[must_use]
    pub fn new() -> Self {
        Self {
            shows: Vec::new(),
            next_id: 0,
        }
    }

    /// Add a custom show to the list.
    pub fn add(&mut self, show: Show) {
        if show.id >= self.next_id {
            self.next_id = show.id + 1;
        }
        self.shows.push(show);
    }

    /// Create and add a new custom show.
    pub fn create(&mut self, name: impl Into<String>, slide_ids: Vec<u32>) -> &Show {
        let show = Show::new(self.next_id, name).with_slides(slide_ids);
        self.next_id += 1;
        let index = self.shows.len();
        self.shows.push(show);
        &self.shows[index]
    }

    /// Get a custom show by name.
    #[must_use]
    pub fn get_by_name(&self, name: &str) -> Option<&Show> {
        self.shows.iter().find(|show| show.name == name)
    }

    /// Get a custom show by ID.
    #[must_use]
    pub fn get_by_id(&self, id: u32) -> Option<&Show> {
        self.shows.iter().find(|show| show.id == id)
    }

    /// Get mutable access to a custom show by ID.
    pub fn get_by_id_mut(&mut self, id: u32) -> Option<&mut Show> {
        self.shows.iter_mut().find(|show| show.id == id)
    }

    /// Replace a custom show while retaining its stable ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_by_id(&mut self, id: u32, mut replacement: Show) -> Result<()> {
        let target = self
            .shows
            .iter_mut()
            .find(|show| show.id == id)
            .ok_or_else(|| Error::Invalid(format!("custom show {id} was not found")))?;
        replacement.id = id;
        *target = replacement;
        Ok(())
    }

    /// Remove a custom show by ID.
    pub fn remove_by_id(&mut self, id: u32) -> Option<Show> {
        self.shows
            .iter()
            .position(|show| show.id == id)
            .map(|offset| self.shows.remove(offset))
    }

    /// Reorder custom shows by a complete ID permutation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reorder(&mut self, ordered_ids: &[u32]) -> Result<()> {
        let expected = self
            .shows
            .iter()
            .map(|show| show.id)
            .collect::<std::collections::HashSet<_>>();
        let actual = ordered_ids
            .iter()
            .copied()
            .collect::<std::collections::HashSet<_>>();
        if expected != actual || ordered_ids.len() != self.shows.len() {
            return Err(Error::Invalid(
                "custom-show reorder is not a permutation".into(),
            ));
        }
        self.shows = ordered_ids
            .iter()
            .map(|id| self.get_by_id(*id).cloned())
            .collect::<Option<Vec<_>>>()
            .ok_or_else(|| Error::Invalid("custom-show reorder lost a validated ID".into()))?;
        Ok(())
    }

    /// Remove a custom show by name.
    pub fn remove_by_name(&mut self, name: &str) -> Option<Show> {
        self.shows
            .iter()
            .position(|show| show.name == name)
            .map(|offset| self.shows.remove(offset))
    }

    /// Get the number of custom shows.
    #[must_use]
    pub fn len(&self) -> usize {
        self.shows.len()
    }

    /// Return whether this list has no custom shows.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.shows.is_empty()
    }
}
