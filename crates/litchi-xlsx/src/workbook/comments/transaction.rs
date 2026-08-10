//! Clone-staged semantic transactions for one worksheet's classic comments.

use litchi_opc::{OpcPackage, PackURI};
use litchi_sheet::At;

use super::model::{Comment, Comments};
use super::package::{remove_from_worksheet, replace_on_worksheet, validate_graph};
use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;
use super::validation;
use crate::error::{Result, invalid};

/// A failure-atomic edit over one worksheet's legacy-note graph.
///
/// Semantic edits are held separately from the source package. Commit clones
/// the package, reuses the existing comments relationship planner and codec
/// validation, and publishes only after the complete graph passes validation.
/// Dropping the transaction rolls back every pending operation.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    worksheet: PackURI,
    before: Snapshot,
    draft: Option<Comments>,
}

impl<'a> Transaction<'a> {
    /// Start a transaction after validating the existing worksheet comments
    /// graph and capturing its exact source bytes.
    pub fn new(target: &'a mut OpcPackage, worksheet: &PackURI) -> Result<Self> {
        let before = Snapshot::load(target, worksheet)?;
        let draft = before.comments().cloned();
        Ok(Self {
            target,
            worksheet: worksheet.clone(),
            before,
            draft,
        })
    }

    /// Immutable source snapshot used for conflict checks and inverse patches.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the current staged typed graph.
    #[must_use]
    pub fn comments(&self) -> Option<&Comments> {
        self.draft.as_ref()
    }

    /// Replace the complete graph. `None` removes the comments part and its
    /// worksheet relationship; VML drawing parts remain untouched.
    pub fn replace(&mut self, value: Option<Comments>) -> Result<bool> {
        if let Some(value) = &value {
            validation::comments(value)?;
        }
        if self.draft == value {
            return Ok(false);
        }
        self.draft = value;
        Ok(true)
    }

    /// Add or replace a note at a checked cell address.
    ///
    /// Existing GUID and inert VML shape ID metadata are retained when the
    /// note is replaced. New notes do not invent a VML shape ID.
    pub fn set<'cell>(
        &mut self,
        cell: impl Into<At<'cell>>,
        author: impl Into<String>,
        text: impl Into<String>,
    ) -> Result<bool> {
        let cell_ref = validation::cell(cell)?;
        let author = author.into();
        let text = text.into();
        validation::text(&author, "author")?;
        validation::text(&text, "text")?;

        let mut comments = self.draft.clone().unwrap_or_default();
        let previous = comments.comments.get(&cell_ref).cloned();
        let previous_author_id = previous
            .as_ref()
            .filter(|comment| comment.author == author)
            .map(|comment| comment.author_id);
        let author_id = if let Some(author_id) = previous_author_id {
            author_id
        } else if let Some(index) = comments
            .authors
            .iter()
            .position(|candidate| candidate == &author)
        {
            u32::try_from(index)
                .map_err(|_source| invalid("classic-comments author index exceeds u32"))?
        } else {
            comments.authors.push(author.clone());
            u32::try_from(comments.authors.len() - 1)
                .map_err(|_source| invalid("classic-comments author index exceeds u32"))?
        };
        let replacement = Comment {
            cell_ref: cell_ref.clone(),
            author,
            author_id,
            text,
            guid: previous.as_ref().and_then(|comment| comment.guid.clone()),
            shape_id: previous.as_ref().and_then(|comment| comment.shape_id),
        };
        if previous.as_ref() == Some(&replacement) {
            return Ok(false);
        }
        comments.comments.insert(cell_ref, replacement);
        validation::comments(&comments)?;
        self.draft = Some(comments);
        Ok(true)
    }

    /// Add or replace a fully typed note graph entry.
    ///
    /// The cell reference and author relationship are checked by the existing
    /// comments validator. Shape IDs remain opaque and inert.
    pub fn set_comment(&mut self, value: Comment) -> Result<bool> {
        litchi_sheet::Cell::from_a1(&value.cell_ref)?;
        validation::text(&value.author, "author")?;
        validation::text(&value.text, "text")?;
        let mut comments = self.draft.clone().unwrap_or_default();
        let changed = comments.comments.get(&value.cell_ref) != Some(&value);
        if !changed {
            return Ok(false);
        }
        comments.comments.insert(value.cell_ref.clone(), value);
        validation::comments(&comments)?;
        self.draft = Some(comments);
        Ok(true)
    }

    /// Rename one checked author-table entry and update every referencing note.
    pub fn set_author(&mut self, author_id: u32, author: impl Into<String>) -> Result<bool> {
        let author = author.into();
        validation::text(&author, "author")?;
        let mut comments = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit an author on an absent comments part"))?;
        let index = usize::try_from(author_id)
            .ok()
            .filter(|index| *index < comments.authors.len())
            .ok_or_else(|| invalid(format!("classic comments author ID {author_id} is absent")))?;
        if comments.authors[index] == author {
            return Ok(false);
        }
        comments.authors[index].clone_from(&author);
        for comment in comments.comments.values_mut() {
            if comment.author_id == author_id {
                comment.author.clone_from(&author);
            }
        }
        validation::comments(&comments)?;
        self.draft = Some(comments);
        Ok(true)
    }

    /// Rename a unique semantic author entry without exposing its native
    /// author-table index. Ambiguous names are rejected rather than guessed.
    pub fn rename_author(&mut self, current: &str, replacement: impl Into<String>) -> Result<bool> {
        validation::text(current, "author")?;
        let comments = self
            .draft
            .as_ref()
            .ok_or_else(|| invalid("cannot edit an author on an absent comments part"))?;
        let mut found = None;
        for (index, author) in comments.authors.iter().enumerate() {
            if author == current {
                if found.is_some() {
                    return Err(invalid(format!(
                        "classic comments author '{current}' is ambiguous"
                    )));
                }
                found = Some(index);
            }
        }
        let index = found
            .ok_or_else(|| invalid(format!("classic comments author '{current}' is absent")))?;
        let index = u32::try_from(index)
            .map_err(|_source| invalid("classic-comments author index exceeds u32"))?;
        self.set_author(index, replacement)
    }

    /// Remove the note at a checked cell address.
    pub fn remove<'cell>(&mut self, cell: impl Into<At<'cell>>) -> Result<Option<Comment>> {
        let cell_ref = validation::cell(cell)?;
        let Some(mut comments) = self.draft.clone() else {
            return Ok(None);
        };
        let removed = comments.comments.remove(&cell_ref);
        let Some(removed) = removed else {
            return Ok(None);
        };
        self.draft = if comments.comments.is_empty() {
            None
        } else {
            validation::comments(&comments)?;
            Some(comments)
        };
        Ok(Some(removed))
    }

    /// Remove all classic notes from the worksheet.
    pub fn clear(&mut self) -> bool {
        self.draft.take().is_some()
    }

    /// Whether the staged semantic graph differs from the source graph.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.comments() != self.draft.as_ref()
    }

    /// Validate and atomically publish the transaction.
    pub fn commit(self) -> Result<Commit> {
        let changed = self.is_changed();
        if !changed {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit {
                snapshot: self.before,
                patch,
                changed: false,
            });
        }

        let mut candidate = self.target.clone();
        match &self.draft {
            Some(value) => {
                replace_on_worksheet(&mut candidate, &self.worksheet, value)?;
            },
            None => {
                remove_from_worksheet(&mut candidate, &self.worksheet)?;
            },
        }
        validate_graph(&candidate)?;
        let snapshot = Snapshot::load(&candidate, &self.worksheet)?;
        if snapshot.comments() != self.draft.as_ref() {
            return Err(invalid(
                "classic comments package publication changed the staged graph",
            ));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }
}
