//! Immutable edits of XLSB threaded comments and workbook persons (record
//! 0360b).
//!
//! [`crate::Package::threaded_comments`] reads the typed graph;
//! [`crate::Package::edit_threaded_comments`] starts an [`Edit`] that stages
//! people, worksheet collections, comments and threads against it, each
//! staging validated as a complete graph. [`Edit::commit`] stores the graph,
//! reads it back and returns a [`Commit`] with the new package, the exact,
//! reversible part-level [`Patch`] and the [`Diagnostics`];
//! [`crate::Package::apply_threaded_comments`] replays a patch onto an equal
//! package. An unchanged graph publishes nothing, and a changed edit of a
//! signed workbook is refused with [`Error::Signed`].

use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, TargetMode};

use super::package;
use super::semantic::{Comment, Comments, CommentsPart, Graph, People, PeoplePart, Person, Thread};
use super::validation;
use crate::package::error::Result;
use crate::part_edit::{Staged, family};

family!("threaded-comment");

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Serialized bytes of one threaded-comments or persons part
    /// ([`super::MAX_PART_BYTES`]).
    PartBytes,
    /// Persons of one workbook ([`super::MAX_PERSONS`]).
    Persons,
}

impl crate::Package {
    /// The typed threaded-comments/persons graph of this package.
    ///
    /// # Errors
    ///
    /// Returns the graph, decoding and validation errors.
    pub fn threaded_comments(&self) -> Result<Snapshot> {
        Snapshot::read(self.opc_package())
    }

    /// Start a threaded-comments edit. Starting one copies no part payload.
    ///
    /// The graph is read on the first operation or commit, so dropping an
    /// untouched edit does not decode threaded-comments parts.
    pub fn edit_threaded_comments(&self) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
            base: None,
            graph: Graph::default(),
        })
    }

    /// Apply a threaded-comments patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_threaded_comments(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

/// An immutable package-backed threaded-comments/persons snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    graph: Graph,
    source: Option<Arc<SourceState>>,
}

impl Snapshot {
    /// Read and validate the complete threaded-comments/persons owner.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        let graph = package::load_graph(package)?;
        Ok(Self {
            graph,
            source: Some(Arc::new(SourceState::capture(package)?)),
        })
    }

    /// Borrow the package-neutral typed graph.
    #[must_use]
    pub const fn graph(&self) -> &Graph {
        &self.graph
    }

    /// Borrow workbook people, when the optional persons part exists.
    #[must_use]
    pub fn people(&self) -> Option<&People> {
        self.graph.persons.as_ref().map(|part| &part.persons)
    }

    /// Borrow the exact XML source parts and owner relationship context.
    ///
    /// Workbook and worksheet source parts are included when they own a
    /// threaded-comments/person relationship, but their potentially large
    /// BIFF12 blobs are intentionally not copied into this XML-owner view.
    #[must_use]
    pub fn source_parts(&self) -> &[SourcePart] {
        self.source
            .as_deref()
            .map_or(&[], |source| source.parts.as_slice())
    }

    /// Whether this snapshot is bound to exact package source context.
    #[must_use]
    pub fn is_source_bound(&self) -> bool {
        self.source.is_some()
    }

    /// Borrow one worksheet's threaded comments by absolute worksheet part.
    #[must_use]
    pub fn worksheet(&self, worksheet_part_name: &str) -> Option<&Comments> {
        self.graph
            .worksheets
            .iter()
            .find(|part| part.worksheet_part_name == worksheet_part_name)
            .map(|part| &part.comments)
    }

    /// Whether the package owner has no threaded-comments/persons parts.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.graph.persons.is_none() && self.graph.worksheets.is_empty()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.graph == other.graph && self.source == other.source
    }
}

impl Eq for Snapshot {}

/// A staged threaded-comments/persons edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    base: Option<Graph>,
    graph: Graph,
}

impl Edit {
    fn ensure_prepared(&mut self) -> Result<()> {
        if self.base.is_some() {
            return Ok(());
        }
        let base = package::load_graph(self.staged.opc())?;
        self.graph = base.clone();
        self.base = Some(base);
        Ok(())
    }

    fn base(&self) -> Result<&Graph> {
        self.base
            .as_ref()
            .ok_or_else(|| invalid_state("threaded-comments edit is not prepared"))
    }

    /// Discard every staged change.
    pub fn reset(&mut self) -> Result<()> {
        self.ensure_prepared()?;
        self.graph = self.base()?.clone();
        Ok(())
    }

    /// Replace the complete staged graph.
    ///
    /// # Errors
    ///
    /// Returns the graph validation errors.
    pub fn replace(&mut self, graph: Graph) -> Result<()> {
        self.ensure_prepared()?;
        validate(&graph)?;
        self.graph = graph;
        Ok(())
    }

    /// Remove every threaded-comments and persons part. Returns whether the
    /// staged graph had any.
    pub fn clear(&mut self) -> Result<bool> {
        self.ensure_prepared()?;
        let present = self.graph.persons.is_some() || !self.graph.worksheets.is_empty();
        self.graph = Graph::default();
        Ok(present)
    }

    /// Remove the workbook persons part when no comment or mention still
    /// references a person. Returns whether one was staged.
    ///
    /// # Errors
    ///
    /// Returns the graph validation errors.
    pub fn remove_people(&mut self) -> Result<bool> {
        self.ensure_prepared()?;
        if self.graph.persons.is_none() {
            return Ok(false);
        }
        let mut candidate = self.graph.clone();
        candidate.persons = None;
        validate(&candidate)?;
        self.graph = candidate;
        Ok(true)
    }

    /// Borrow the currently staged graph.
    #[must_use]
    pub fn graph(&mut self) -> Result<&Graph> {
        self.ensure_prepared()?;
        Ok(&self.graph)
    }

    /// Borrow the currently staged workbook people.
    #[must_use]
    pub fn people(&mut self) -> Result<Option<&People>> {
        self.ensure_prepared()?;
        Ok(self.graph.persons.as_ref().map(|part| &part.persons))
    }

    /// Replace workbook people, retaining the existing package identity.
    pub fn set_people(&mut self, people: People) -> Result<()> {
        self.ensure_prepared()?;
        let mut candidate = self.graph.clone();
        if let Some(part) = candidate.persons.as_mut() {
            part.persons = people;
        } else {
            candidate.persons = Some(PeoplePart {
                relationship_id: String::new(),
                part_name: String::new(),
                persons: people,
            });
        }
        validate(&candidate)?;
        self.graph = candidate;
        Ok(())
    }

    /// Add or replace one workbook person by stable GUID.
    pub fn upsert_person(&mut self, person: Person) -> Result<()> {
        self.ensure_prepared()?;
        let mut people = self.people()?.cloned().unwrap_or_default();
        if let Some(existing) = people.persons.iter_mut().find(|item| item.id == person.id) {
            *existing = person;
        } else {
            people.persons.push(person);
        }
        self.set_people(people)
    }

    /// Remove a workbook person when no comment or mention still references it.
    pub fn remove_person(&mut self, person_id: &str) -> Result<Option<Person>> {
        self.ensure_prepared()?;
        let Some(current) = self.graph.persons.as_ref() else {
            return Ok(None);
        };
        let Some(index) = current
            .persons
            .persons
            .iter()
            .position(|person| person.id == person_id)
        else {
            return Ok(None);
        };
        let mut candidate = self.graph.clone();
        let part = candidate
            .persons
            .as_mut()
            .ok_or_else(|| invalid_state("people disappeared while removing a person"))?;
        if index >= part.persons.persons.len() {
            return Err(invalid_state(
                "person index disappeared while removing a person",
            ));
        }
        let removed = part.persons.persons.remove(index);
        if candidate
            .persons
            .as_ref()
            .is_some_and(|part| part.persons.persons.is_empty())
        {
            candidate.persons = None;
        }
        validate(&candidate)?;
        self.graph = candidate;
        Ok(Some(removed))
    }

    /// Borrow one staged worksheet threaded-comments collection.
    #[must_use]
    pub fn worksheet(&mut self, worksheet_part_name: &str) -> Result<Option<&Comments>> {
        self.ensure_prepared()?;
        Ok(self
            .graph
            .worksheets
            .iter()
            .find(|part| part.worksheet_part_name == worksheet_part_name)
            .map(|part| &part.comments))
    }

    /// Replace one worksheet's threaded comments, retaining package identity.
    pub fn set_worksheet(
        &mut self,
        worksheet_part_name: impl Into<String>,
        comments: Comments,
    ) -> Result<()> {
        self.ensure_prepared()?;
        let worksheet_part_name = worksheet_part_name.into();
        let mut candidate = self.graph.clone();
        if let Some(part) = candidate
            .worksheets
            .iter_mut()
            .find(|part| part.worksheet_part_name == worksheet_part_name)
        {
            part.comments = comments;
        } else {
            candidate.worksheets.push(CommentsPart {
                worksheet_part_name,
                relationship_id: String::new(),
                part_name: String::new(),
                comments,
            });
            candidate
                .worksheets
                .sort_by(|left, right| left.worksheet_part_name.cmp(&right.worksheet_part_name));
        }
        validate(&candidate)?;
        self.graph = candidate;
        Ok(())
    }

    /// Append one root or reply comment to an existing worksheet collection.
    pub fn add_comment(&mut self, worksheet_part_name: &str, comment: Comment) -> Result<()> {
        self.ensure_prepared()?;
        let mut candidate = self.graph.clone();
        let part = candidate
            .worksheets
            .iter_mut()
            .find(|part| part.worksheet_part_name == worksheet_part_name)
            .ok_or_else(|| missing_worksheet(worksheet_part_name))?;
        part.comments.comments.push(comment);
        validate(&candidate)?;
        self.graph = candidate;
        Ok(())
    }

    /// Remove a comment and all replies rooted at it.
    pub fn remove_thread(
        &mut self,
        worksheet_part_name: &str,
        root_id: &str,
    ) -> Result<Option<Thread>> {
        self.ensure_prepared()?;
        let Some(current) = self.worksheet(worksheet_part_name)? else {
            return Ok(None);
        };
        let threads = validation::group_threads(current).map_err(map_validation_error)?;
        let Some(thread) = threads.into_iter().find(|thread| thread.root.id == root_id) else {
            return Ok(None);
        };
        let ids = std::iter::once(thread.root.id.as_str())
            .chain(thread.replies.iter().map(|reply| reply.id.as_str()))
            .collect::<Vec<_>>();
        let mut candidate = self.graph.clone();
        let part = candidate
            .worksheets
            .iter_mut()
            .find(|part| part.worksheet_part_name == worksheet_part_name)
            .ok_or_else(|| invalid_state("worksheet disappeared while removing a thread"))?;
        part.comments
            .comments
            .retain(|comment| !ids.contains(&comment.id.as_str()));
        validate(&candidate)?;
        self.graph = candidate;
        Ok(Some(thread))
    }

    /// Remove one worksheet's complete threaded-comments part from the graph.
    pub fn remove_worksheet(&mut self, worksheet_part_name: &str) -> Result<bool> {
        self.ensure_prepared()?;
        let mut candidate = self.graph.clone();
        let old_len = candidate.worksheets.len();
        candidate
            .worksheets
            .retain(|part| part.worksheet_part_name != worksheet_part_name);
        if candidate.worksheets.len() == old_len {
            return Ok(false);
        }
        validate(&candidate)?;
        self.graph = candidate;
        Ok(true)
    }

    /// Validate the staged graph, store it, read it back and publish a new
    /// package. An unchanged graph returns the source package.
    ///
    /// # Errors
    ///
    /// Returns the graph validation errors, [`Error::Signed`] for a changed
    /// edit of a signed workbook and the workbook validation errors of the
    /// result.
    pub fn commit(mut self) -> Result<Commit> {
        self.ensure_prepared()?;
        validate(&self.graph)?;
        let changed = {
            let base = self.base()?;
            self.graph != *base
        };
        if changed {
            let graph = self.graph.clone();
            self.staged.update(|candidate| {
                package::store_graph(candidate, &graph)?;
                let stored = package::load_graph(candidate)?;
                validate(&stored)
            })?;
        }
        Ok(Commit::new(self.staged.commit()?))
    }
}

fn validate(graph: &Graph) -> Result<()> {
    validation::validate_graph(graph).map_err(map_validation_error)
}

fn map_validation_error(_error: validation::Error) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::InvalidValue,
        "edit_threaded_comments",
        "the threaded-comments graph is invalid",
    )
}

fn missing_worksheet(_worksheet_part_name: &str) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::NotFound,
        "edit_threaded_comments",
        "the threaded-comments target was not found",
    )
}

fn invalid_state(_message: impl Into<String>) -> Error {
    Error::InvalidFormat("XLSB threaded-comments edit")
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct SourceState {
    parts: Vec<SourcePart>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourcePart {
    part_name: String,
    content_type: String,
    blob: Option<Arc<Vec<u8>>>,
    relationships: Vec<SourceRelationship>,
}

impl SourcePart {
    /// Absolute OPC part identity.
    #[must_use]
    pub fn part_name(&self) -> &str {
        &self.part_name
    }

    /// Part content type.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Exact source bytes for an XML owner part, or empty for a source
    /// workbook/worksheet context entry.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.blob.as_deref().map_or(&[], Vec::as_slice)
    }

    /// Whether this entry carries exact source bytes.
    #[must_use]
    pub fn has_bytes(&self) -> bool {
        self.blob.is_some()
    }

    /// Relationships in the source owner context.
    #[must_use]
    pub fn relationships(&self) -> &[SourceRelationship] {
        &self.relationships
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceRelationship {
    id: String,
    relationship_type: String,
    target: String,
    mode: TargetMode,
}

impl SourceRelationship {
    /// Relationship identity.
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Relationship type URI.
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    /// Original OPC target reference.
    #[must_use]
    pub fn target(&self) -> &str {
        &self.target
    }

    /// Target mode.
    #[must_use]
    pub const fn mode(&self) -> TargetMode {
        self.mode
    }
}

impl SourceState {
    fn capture(package: &OpcPackage) -> Result<Self> {
        package::validate_graph(package)?;
        let mut parts = Vec::new();
        for part in package.iter_parts() {
            let owner = matches!(
                part.content_type(),
                package::PERSONS_CONTENT_TYPE | package::COMMENTS_CONTENT_TYPE
            );
            let source = matches!(
                part.content_type(),
                ct::XLSB_BIN | package::WORKSHEET_CONTENT_TYPE
            );
            let mut relationships = part
                .rels()
                .iter()
                .filter(|relationship| {
                    matches!(relationship.reltype(), rt::PERSONS | rt::THREADED_COMMENTS)
                })
                .map(|relationship| {
                    Ok(SourceRelationship {
                        id: relationship.r_id().to_string(),
                        relationship_type: relationship.reltype().to_string(),
                        target: relationship.target_ref().to_string(),
                        mode: relationship.target_mode(),
                    })
                })
                .collect::<Result<Vec<_>>>()?;
            relationships.sort_by(|left, right| left.id.cmp(&right.id));
            let relevant_source = source && !relationships.is_empty();
            if !owner && !relevant_source {
                continue;
            }
            parts.push(SourcePart {
                part_name: part.partname().to_string(),
                content_type: part.content_type().to_string(),
                // Only an owner part's payload is captured, so only an
                // owner part's payload is decoded (ADR 0030).
                blob: if owner {
                    Some(package.get_part(part.partname())?.blob_arc())
                } else {
                    None
                },
                relationships,
            });
        }
        parts.sort_by(|left, right| left.part_name.cmp(&right.part_name));
        Ok(Self { parts })
    }
}
