//! Snapshot-based edits for XLSB threaded comments and workbook persons.
//!
//! The editor owns only the typed threaded-comments/persons graph. Package
//! application is source checked and clone staged so semantic or relationship
//! failures cannot partially publish an OPC mutation.

use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, TargetMode};

use super::package;
use super::semantic::{Comment, Comments, CommentsPart, Graph, People, PeoplePart, Person, Thread};
use super::validation;
use crate::package::error::{Error, Result};

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
    ///
    /// A changed [`Commit`] is intentionally detached because package part
    /// names and relationship IDs may be allocated by `store_graph`. The
    /// successful package [`Patch::apply`] result is source bound again.
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

    /// Start a detached transaction against this snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            graph: self.graph.clone(),
        }
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

/// A detached, source-checked threaded-comments/persons transaction.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    graph: Graph,
}

impl Transaction {
    /// Borrow the currently staged graph.
    #[must_use]
    pub const fn graph(&self) -> &Graph {
        &self.graph
    }

    /// Borrow the currently staged workbook people.
    #[must_use]
    pub fn people(&self) -> Option<&People> {
        self.graph.persons.as_ref().map(|part| &part.persons)
    }

    /// Replace workbook people, retaining the existing package identity.
    pub fn set_people(&mut self, people: People) -> Result<()> {
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
        let mut people = self.people().cloned().unwrap_or_default();
        if let Some(existing) = people.persons.iter_mut().find(|item| item.id == person.id) {
            *existing = person;
        } else {
            people.persons.push(person);
        }
        self.set_people(people)
    }

    /// Remove a workbook person when no comment or mention still references it.
    pub fn remove_person(&mut self, person_id: &str) -> Result<Option<Person>> {
        let Some(current) = self.people() else {
            return Ok(None);
        };
        let Some(index) = current
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
    pub fn worksheet(&self, worksheet_part_name: &str) -> Option<&Comments> {
        self.graph
            .worksheets
            .iter()
            .find(|part| part.worksheet_part_name == worksheet_part_name)
            .map(|part| &part.comments)
    }

    /// Replace one worksheet's threaded comments, retaining package identity.
    pub fn set_worksheet(
        &mut self,
        worksheet_part_name: impl Into<String>,
        comments: Comments,
    ) -> Result<()> {
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
        let Some(current) = self.worksheet(worksheet_part_name) else {
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

    /// Validate and produce an immutable commit and reversible patch.
    pub fn commit(self) -> Result<Commit> {
        validate(&self.graph)?;
        let changed = self.graph != self.base.graph;
        let source = Arc::clone(self.base.source.as_ref().ok_or_else(|| {
            Error::UnsupportedFeature(
                "threaded-comments edit snapshot is not source bound".to_string(),
            )
        })?);
        Ok(Commit {
            snapshot: Snapshot {
                graph: self.graph.clone(),
                source: (!changed).then_some(Arc::clone(&source)),
            },
            patch: Patch {
                before: source,
                after: self.graph,
                changed,
            },
        })
    }
}

/// A successful immutable transaction result.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the committed semantic snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the source-checked package patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Split the commit into snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A source-checked, clone-staged package patch.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Arc<SourceState>,
    after: Graph,
    changed: bool,
}

impl Patch {
    /// Whether applying this patch leaves the owner byte streams untouched.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        !self.changed
    }

    /// Apply this patch only to the exact source owner graph.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        let current = SourceState::capture(package)?;
        if current != *self.before {
            return Err(Error::UnsupportedFeature(
                "threaded-comments patch source snapshot does not match".to_string(),
            ));
        }
        if !self.changed {
            return Snapshot::read(package);
        }

        let mut candidate = package.clone();
        package::store_graph(&mut candidate, &self.after)?;
        let snapshot = Snapshot::read(&candidate)?;
        *package = candidate;
        Ok(snapshot)
    }

    /// Apply the patch's exact before-image guard without publishing a partial
    /// package mutation. This is an alias suited to transaction pipelines.
    pub fn commit(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        self.apply(package)
    }
}

/// Read the threaded-comments/persons snapshot for an OPC package.
pub fn read(package: &OpcPackage) -> Result<Snapshot> {
    Snapshot::read(package)
}

/// Apply a previously committed patch atomically to an OPC package.
pub fn apply(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    patch.apply(package)
}

fn validate(graph: &Graph) -> Result<()> {
    validation::validate_graph(graph).map_err(map_validation_error)
}

fn map_validation_error(error: validation::Error) -> Error {
    Error::Unrecognized {
        typ: "XLSB threaded-comments edit".to_string(),
        val: error.to_string(),
    }
}

fn missing_worksheet(worksheet_part_name: &str) -> Error {
    Error::Unrecognized {
        typ: "XLSB threaded-comments edit".to_string(),
        val: format!("worksheet '{worksheet_part_name}' has no threaded-comments part"),
    }
}

fn invalid_state(message: impl Into<String>) -> Error {
    Error::InvalidFormat(format!("XLSB threaded-comments edit: {}", message.into()))
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
                blob: owner.then(|| part.blob_arc()),
                relationships,
            });
        }
        parts.sort_by(|left, right| left.part_name.cmp(&right.part_name));
        Ok(Self { parts })
    }
}
