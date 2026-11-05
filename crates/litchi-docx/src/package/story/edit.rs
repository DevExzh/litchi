//! The story family (record 0355): hyperlink text across the main, header
//! and footer stories of an owning package.
//!
//! [`Package::story_inventory`](crate::Package::story_inventory) reads the
//! stories; [`Package::edit_stories`](crate::Package::edit_stories) stages an
//! [`Edit`] whose commit publishes a new package, and
//! [`Package::apply_stories`](crate::Package::apply_stories) replays its
//! [`Patch`].

use super::{StoryHyperlinkTextReplacement, StoryInventory};
use crate::package::family::{graph_edit, graph_family};

graph_family!("story", own_limits);

pub use super::limit::LimitKind;
graph_edit!("story", edit_stories, apply_stories);

impl Edit {
    /// Replace direct hyperlink text across the reachable stories; returns
    /// the resulting story inventory. Replacements must be non-empty,
    /// unique, and strictly increasing by part name and then by
    /// paragraph/hyperlink address. Every unselected byte is kept.
    ///
    /// # Errors
    ///
    /// Returns a typed document error when a selector is ambiguous, a story
    /// is not an editable main, header or footer story, or a leaf cannot be
    /// represented safely; the staged package is then unchanged.
    pub fn replace_hyperlink_texts(
        &mut self,
        replacements: &[StoryHyperlinkTextReplacement],
    ) -> crate::document::Result<StoryInventory> {
        let mut next = self.staged.working().clone();
        let inventory = next.replace_story_hyperlink_texts(replacements)?;
        *self.staged.working_mut() = next;
        Ok(inventory)
    }
}
