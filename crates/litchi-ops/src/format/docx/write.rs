//! DOCX `set` (record 0015 W5a).
//!
//! The handler resolves the target once, proves every node against the main
//! story's direct transaction positions, admits every property before it
//! touches an editor, and then applies one isolated leaf edit.  A failed
//! admission or edit drops that transaction, so the input handle remains an
//! immutable snapshot in every case.

pub(crate) mod coalesce;

use litchi_core::{FileFormat, Position};
use litchi_docx::source_backed;

use super::Docx;
use super::DocxEditor;
use super::Nav;
use super::element;
use super::few::Few;
use super::node::{Node, WriteSelection};
use super::setters::Setter;
use super::spec::{
    CELL_REGISTRY, HYPERLINK_REGISTRY, Kind, PARAGRAPH_REGISTRY, ROW_REGISTRY, RUN_REGISTRY,
    TABLE_REGISTRY,
};
use crate::error::{Code, Error, Result};
use crate::props::{IntoProps, Key, PropValue, Props};
use crate::registry::{Entry, Operation, Table};
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// Evaluate `$body` against whichever owner's leaf edit `$work` holds.
macro_rules! both {
    ($work:expr, |$edit:ident| $body:expr) => {
        match $work {
            WorkEdit::Eager($edit) => $body,
            WorkEdit::Source($edit) => $body,
        }
    };
}

mod find;

/// Complete result of one DOCX leaf transaction.
///
/// Single-operation sugar takes only the file. The transaction executor can
/// retain the exact patch and diagnostics without reconstructing a change from
/// the read projection or serializing the package.
pub(crate) struct Commit {
    file: Docx,
    patch: litchi_docx::document::Patch,
    diagnostics: litchi_docx::document::Diagnostics,
    operations: usize,
}

impl Commit {
    pub(crate) fn into_parts(
        self,
    ) -> (
        Docx,
        litchi_docx::document::Patch,
        litchi_docx::document::Diagnostics,
    ) {
        (self.file, self.patch, self.diagnostics)
    }

    pub(crate) fn into_file(self) -> Docx {
        self.into_parts().0
    }

    /// The committed handle.
    pub(super) const fn file(&self) -> &Docx {
        &self.file
    }

    pub(crate) const fn operations(&self) -> usize {
        self.operations
    }

    /// A commit from its parts, for a route that publishes outside the
    /// main-document edit (record 0338: metadata).
    pub(super) const fn from_parts(
        file: Docx,
        patch: litchi_docx::document::Patch,
        diagnostics: litchi_docx::document::Diagnostics,
    ) -> Self {
        let operations = diagnostics.operations();
        Self {
            file,
            patch,
            diagnostics,
            operations,
        }
    }
}

impl std::fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}

/// One admitted property. Immediate edits borrow the original input; deferred
/// formatting plans explicitly take ownership before retaining assignments.
#[derive(Clone)]
struct Assignment<'a> {
    key: Key<'a>,
    value: PropValue<'a>,
    entry: Entry<Setter>,
}

impl Assignment<'_> {
    fn into_owned(self) -> Assignment<'static> {
        Assignment {
            key: self.key.into_owned(),
            value: self.value.into_owned(),
            entry: self.entry,
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum TargetKind {
    Paragraph,
    Run,
    Hyperlink,
    Cell,
    /// A table's borders (record 0342).
    Table,
    /// A row's height and flags (record 0342).
    Row,
}

/// Where a set lands: a direct-body paragraph or run on the body-position
/// route, or any other element through its leaf element path (record 0251a).
#[derive(Clone)]
enum Selection {
    Body(WriteSelection),
    Element(litchi_docx::document::ElementPath),
}

#[derive(Clone)]
struct Target {
    selection: Selection,
    kind: TargetKind,
}

/// The two leaf edit types have the same operation surface but different
/// package owners.  Keeping this small enum local avoids serializing and
/// reopening a committed package merely to regain a `Docx` handle.
pub(super) enum WorkEdit {
    Eager(Box<litchi_docx::document::Edit>),
    Source(Box<litchi_docx::document::Edit<source_backed::Package>>),
}

impl Docx {
    #[cfg(test)]
    pub(crate) fn set<'p, 'v>(
        &self,
        target: &TargetArg<'p>,
        props: impl IntoProps<'v>,
    ) -> Result<Commit> {
        self.set_with_limits(target, props, &SemanticLimits::default())
    }

    /// Apply one setter under a caller-selected remaining native snapshot
    /// budget. The budget is attached to the leaf edit before any projected
    /// main-document XML can be retained.
    pub(crate) fn set_with_limits<'p, 'v>(
        &self,
        target: &TargetArg<'p>,
        props: impl IntoProps<'v>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        // The final source fence belongs outside the whole operation.  This
        // makes a source revision win over every failure path too: target
        // resolution, property admission, leaf staging, and commit errors
        // cannot leak a stale leaf error after the input changed.
        let result = self.set_inner(target, props, false, limits.max_snapshot_bytes());
        self.owner().finish(result)
    }

    /// [`Self::set`] for single-operation `File::set`, which keeps only the
    /// committed handle: the leaf's patch is dropped where the commit
    /// returns instead of travelling up the call chain, and a direct body
    /// paragraph's text skips the generic resolver (record 0305).
    pub(crate) fn set_file(&self, target: &TargetArg<'_>, props: Props<'_>) -> Result<Self> {
        self.set_file_with_limits(target, props, &SemanticLimits::default())
    }

    pub(crate) fn set_file_with_limits(
        &self,
        target: &TargetArg<'_>,
        props: Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Self> {
        let result = match self.direct_paragraph(target, &props) {
            Some(position) => {
                self.set_paragraph_text(position, &props, limits.max_snapshot_bytes())
            },
            None => self.set_inner(target, props, false, limits.max_snapshot_bytes()),
        };
        self.owner().finish(result)
    }

    /// The transaction position of the paragraph a `text`-only set on an
    /// eager owner names by the direct path `/body/p[N]` (exactly the step
    /// names `body` and `p`, no predicate on `body`, one positional index),
    /// read from the body index the resolver reads: the `N`-th paragraph
    /// block of the body. `None` sends anything else, a missing paragraph
    /// and any read failure to the general route, which reports it.
    fn direct_paragraph(&self, target: &TargetArg<'_>, props: &Props<'_>) -> Option<Position> {
        if !matches!(self.owner(), super::DocxOwner::Eager(_)) {
            return None;
        }
        self.direct_body_paragraph(target, props)
    }

    /// [`Self::direct_paragraph`] on either owner (the batch text runs of
    /// record 0355a read the body index both owners cache).
    fn direct_body_paragraph(&self, target: &TargetArg<'_>, props: &Props<'_>) -> Option<Position> {
        let mut keys = props.iter();
        let (Some((key, _)), None) = (keys.next(), keys.next()) else {
            return None;
        };
        if key.as_str() != "text" {
            return None;
        }
        let path = target.path()?.path();
        let [body, paragraph] = path.steps() else {
            return None;
        };
        if !path.is_absolute()
            || body.name() != "body"
            || !body.predicates().is_empty()
            || paragraph.name() != "p"
        {
            return None;
        }
        let index = paragraph.index()?.get();
        let blocks = super::nav::load_body(self).ok()?;
        let block = blocks.nth(litchi_docx::BlockKind::Paragraph, Position::new(index))?;
        block.transaction_position().ok()
    }

    /// The general route's work for one direct body paragraph and one
    /// `text` value: the editor first (including the signed exact-no-op
    /// exception), the same property admission, then the paragraph text
    /// replacement.
    fn set_paragraph_text(
        &self,
        position: Position,
        props: &Props<'_>,
        snapshot_budget: usize,
    ) -> Result<Self> {
        let editor = self.edit_for_write_with_snapshot_budget(snapshot_budget)?;
        let assignments = admit_assignments(TargetKind::Paragraph, props)?;
        let target = Target {
            selection: Selection::Body(WriteSelection::Paragraph(position)),
            kind: TargetKind::Paragraph,
        };
        let mut editor = WorkEdit::from(editor)?;
        for assignment in &assignments {
            editor.apply_with_policy(&target, assignment, false)?;
        }
        editor.commit_file(self)
    }

    pub(crate) fn set_command_with_limits<'p, 'v>(
        &self,
        target: &TargetArg<'p>,
        props: impl IntoProps<'v>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.owner()
            .finish(self.set_inner(target, props, true, limits.max_snapshot_bytes()))
    }

    fn set_inner<'p, 'v, O: Publish>(
        &self,
        target: &TargetArg<'p>,
        props: impl IntoProps<'v>,
        command_text: bool,
        snapshot_budget: usize,
    ) -> Result<O> {
        // Build the editor first. Besides preserving the signed exact-no-op
        // exception and refusing read-only handles, this makes an empty
        // property list honor the same security policy as a non-empty write.
        // Core document properties on `/` or `/metadata` (record 0338); the
        // metadata route applies the same edit policy without building a
        // main-document editor.
        let props = if crate::format::metadata::is_target(target)
            || crate::format::metadata::is_root(target)
        {
            let props = Props::from_props(props)?;
            if Self::is_metadata_set(target, &props) {
                return O::metadata(self, &props, snapshot_budget);
            }
            Ok(props)
        } else {
            Err(props)
        };
        let editor = self.edit_for_write_with_snapshot_budget(snapshot_budget)?;
        let nodes = self.resolve_targets(target)?;
        let props = match props {
            Ok(props) => props,
            Err(props) => Props::from_props(props)?,
        };
        let editor = match self.try_set_chart(editor, &nodes, &props)? {
            Ok(commit) => return Ok(O::from_commit(commit)),
            Err(editor) => editor,
        };
        // Pictures, comments, SmartArt graphics and their nodes (record
        // 0335).
        let editor = match self.try_set_picture(editor, &nodes, &props)? {
            Ok(commit) => return Ok(O::from_commit(commit)),
            Err(editor) => editor,
        };
        let editor = match self.try_set_comment(editor, &nodes, &props)? {
            Ok(commit) => return Ok(O::from_commit(commit)),
            Err(editor) => editor,
        };
        let editor = match self.try_set_smartart(editor, &nodes, &props)? {
            Ok(commit) => return Ok(O::from_commit(commit)),
            Err(editor) => editor,
        };
        // Table grid columns (record 0345).
        let editor = match self.try_set_column(editor, &nodes, &props)? {
            Ok(commit) => return Ok(O::from_commit(commit)),
            Err(editor) => editor,
        };
        // Styles, numbering, headers, footers and notes (record 0319).
        let editor = match self.try_set_parts(editor, &nodes, &props)? {
            Ok(commit) => return Ok(O::from_commit(commit)),
            Err(editor) => editor,
        };
        // Tab stops (record 0319).
        let editor = match self.try_set_marks(editor, &nodes, &props)? {
            Ok(commit) => return Ok(O::from_commit(commit)),
            Err(editor) => editor,
        };
        // Header, footer and note paragraphs and runs (record 0325).
        let editor = match self.try_set_story(editor, &nodes, &props)? {
            Ok(commit) => return Ok(O::from_commit(commit)),
            Err(editor) => editor,
        };
        let controls = command_text && crate::batch::has_text_controls(&props);
        if controls
            && nodes
                .iter()
                .any(|node| !matches!(node.element_kind(), Kind::Paragraph | Kind::Run))
        {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "OfficeCLI text controls require paragraph or run targets",
            )
            .with_format(FileFormat::Docx));
        }
        let targets = admit_targets(&nodes)?;
        let props = if controls && matches!(targets[0].kind, TargetKind::Paragraph) {
            crate::batch::inline_text_props(&props)
        } else {
            std::borrow::Cow::Borrowed(&props)
        };

        if props.is_empty() {
            return O::publish(WorkEdit::from(editor)?, self);
        }
        let mut assignments = admit_assignments(targets[0].kind, &props)?;
        self.resolve_style_names(&mut assignments)?;
        let request = list_request(&mut assignments)?;
        self.owner().check_source()?;
        // A hyperlink target change stages a relationship, which only the
        // owning package publishes.
        // So does a new list definition (record 0251e).
        let new_lists = request
            .is_some_and(|request| !matches!(request.style, Some(super::lists::Style::None)));
        let editor = if new_lists
            || assignments.iter().any(|assignment| {
                matches!(
                    assignment.entry.access,
                    crate::registry::Access::Set(Setter::HyperlinkUrl)
                        | crate::registry::Access::SetAndAdd(Setter::HyperlinkUrl)
                )
            }) {
            self.eager_editor(editor)?
        } else {
            editor
        };

        let mut editor = WorkEdit::from(editor)?;
        let Some(request) = request else {
            if editor.bulk_paragraph_text(&targets, &assignments)? {
                return O::publish(editor, self);
            }
            // A table's or row's assignments are spliced in one property
            // edit per target (record 0342).
            if matches!(targets[0].kind, TargetKind::Table | TargetKind::Row) {
                for target in &targets {
                    editor.format_structure(target, &assignments)?;
                }
                return O::publish(editor, self);
            }
            for target in &targets {
                for assignment in &assignments {
                    editor.apply_with_policy(target, assignment, command_text)?;
                }
            }
            return O::publish(editor, self);
        };
        // `listStyle`: plan each paragraph's list (staging new definitions),
        // then write its numbering after the other properties.
        let planned = if new_lists {
            let WorkEdit::Eager(eager) = &mut editor else {
                return Err(Error::new(
                    Code::InternalError,
                    "a DOCX list edit needs the owning package editor",
                ));
            };
            super::lists::plan(&nodes, request, eager)?
        } else {
            vec![None; targets.len()]
        };
        let explicit_level = assignments.iter().any(|assignment| {
            matches!(
                assignment.entry.access,
                crate::registry::Access::SetAndAdd(Setter::ParagraphNumLevel)
                    | crate::registry::Access::Set(Setter::ParagraphNumLevel)
            )
        });
        for (target, planned) in targets.iter().zip(planned) {
            for assignment in &assignments {
                editor.apply_with_policy(target, assignment, command_text)?;
            }
            editor.set_numbering(target, planned, explicit_level)?;
        }
        O::publish(editor, self)
    }

    /// `set` on paragraphs and runs of headers, footers and notes (record
    /// 0325): the main document's element-path setters, run on each story
    /// through the leaf's story edit. `Err(editor)` hands the editor back
    /// when no target is in a story.
    fn try_set_story(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
        props: &Props<'_>,
    ) -> Result<std::result::Result<Commit, DocxEditor>> {
        let mut stories = Vec::with_capacity(nodes.len());
        for node in nodes {
            if let Some(found) = element::story_path(node)? {
                stories.push((found, node.element_kind()));
            }
        }
        if stories.is_empty() {
            return Ok(Err(editor));
        }
        let unsupported = |message: &'static str| {
            Error::new(Code::UnsupportedCapability, message).with_format(FileFormat::Docx)
        };
        if stories.len() != nodes.len() {
            return Err(unsupported(
                "one set writes story elements or main-story elements, not both",
            ));
        }
        let kind = match stories[0].1 {
            Kind::Paragraph => TargetKind::Paragraph,
            Kind::Run => TargetKind::Run,
            _ => return Err(unsupported("story paragraphs and runs are set")),
        };
        if stories.iter().any(|(_, other)| *other != stories[0].1) {
            return Err(unsupported("one DOCX mutation cannot mix target kinds"));
        }
        let mut assignments = admit_assignments(kind, props)?;
        self.resolve_style_names(&mut assignments)?;
        if list_request(&mut assignments)?.is_some()
            || assignments.iter().any(|assignment| {
                matches!(
                    assignment.entry.access,
                    crate::registry::Access::Set(Setter::HyperlinkUrl)
                        | crate::registry::Access::SetAndAdd(Setter::HyperlinkUrl)
                )
            })
        {
            return Err(unsupported(
                "list styles and hyperlink targets are not written in headers, footers and notes",
            ));
        }
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(Error::new(
                Code::InternalError,
                "a story edit needs the owning package editor",
            ));
        };
        let mut index = 0;
        while index < stories.len() {
            let key = stories[index].0.0.clone();
            let end = index
                + stories[index..]
                    .iter()
                    .take_while(|((other, _), _)| *other == key)
                    .count();
            let group = &stories[index..end];
            eager.edit_story(super::story::target(&key), |story| -> Result<()> {
                for ((_, path), _) in group {
                    for assignment in &assignments {
                        apply_story(story, path, kind, assignment)
                            .map_err(|error| property_error(error, &assignment.key))?;
                    }
                }
                Ok(())
            })?;
            index = end;
        }
        Ok(Ok(edit.commit(self)?))
    }

    /// An editor over the owning package: package-graph changes
    /// (relationships, new parts) publish only through it. A source-backed
    /// owner is materialized once; the committed handle is eager.
    pub(super) fn eager_editor(&self, editor: DocxEditor) -> Result<DocxEditor> {
        match editor {
            DocxEditor::Source(source_edit) => {
                let snapshot_budget = source_edit.max_snapshot_bytes_value();
                let super::DocxOwner::Source(package) = self.owner() else {
                    return Err(Error::new(
                        Code::InternalError,
                        "a source editor came from an eager DOCX owner",
                    ));
                };
                let owned = self.promoted(package)?;
                let edit = match snapshot_budget {
                    Some(maximum) => owned.edit_with_options(
                        litchi_docx::document::EditOptions::default().max_snapshot_bytes(maximum),
                    )?,
                    None => owned.edit()?,
                };
                Ok(DocxEditor::Edit(Box::new(edit)))
            },
            editor => Ok(editor),
        }
    }

    /// A paragraph or run style named by display name is written as the
    /// style's identifier; an unknown name is written as given (record 0319,
    /// M1-Δ4).
    fn resolve_style_names(&self, assignments: &mut [Assignment<'_>]) -> Result<()> {
        for assignment in assignments {
            if !matches!(
                assignment.entry.access,
                crate::registry::Access::SetAndAdd(Setter::ParagraphStyle | Setter::RunStyle)
            ) {
                continue;
            }
            if let PropValue::Text(text) = &assignment.value {
                let id = super::parts::write::resolve_style_id(self, text)?;
                if id != text.as_ref() {
                    assignment.value = PropValue::from(id);
                }
            }
        }
        Ok(())
    }

    pub(super) fn resolve_targets<'f>(&'f self, target: &TargetArg<'_>) -> Result<Few<Node<'f>>> {
        let nav = Nav::new(self);
        let nodes = if let Some(path) = target.path() {
            Few::One([crate::path::resolve::get(&nav, path)?])
        } else if let Some(query) = target.query() {
            let walk = crate::query::run(nav, query)?;
            // Query diagnostics are harmless for reads because an unknown
            // predicate simply evaluates as absent.  A mutation must refuse
            // them: otherwise a typo could silently widen or empty a write.
            if !walk.plan().diagnostics().is_empty() {
                return Err(Error::new(
                    Code::InvalidSelector,
                    "the mutation selector contains an unsupported predicate",
                )
                .with_format(FileFormat::Docx));
            }
            walk.collect::<Result<Few<_>>>()?
        } else {
            return Err(
                Error::new(Code::InvalidArgument, "the mutation target is empty")
                    .with_format(FileFormat::Docx),
            );
        };
        self.owner().check_source()?;
        if nodes.is_empty() {
            return Err(Error::new(
                Code::ZeroMatches,
                "the mutation selector matched no element",
            )
            .with_format(FileFormat::Docx));
        }
        Ok(nodes)
    }
}

/// What a `set` publishes: the complete [`Commit`] (transactions keep its
/// patch and diagnostics), or only the committed handle (record 0305).
pub(super) trait Publish: Sized {
    /// Commit `edit` over `source`.
    fn publish(edit: WorkEdit, source: &Docx) -> Result<Self>;

    /// A commit another route already made.
    fn from_commit(commit: Commit) -> Self;

    /// A core-properties set on `source` (record 0338).
    fn metadata(source: &Docx, props: &Props<'_>, snapshot_budget: usize) -> Result<Self>;
}

impl Publish for Commit {
    fn publish(edit: WorkEdit, source: &Docx) -> Result<Self> {
        edit.commit(source)
    }

    fn from_commit(commit: Commit) -> Self {
        commit
    }

    fn metadata(source: &Docx, props: &Props<'_>, snapshot_budget: usize) -> Result<Self> {
        source.set_metadata(props, snapshot_budget)
    }
}

impl Publish for Docx {
    fn publish(edit: WorkEdit, source: &Docx) -> Result<Self> {
        edit.commit_file(source)
    }

    fn from_commit(commit: Commit) -> Self {
        commit.into_file()
    }

    fn metadata(source: &Docx, props: &Props<'_>, _snapshot_budget: usize) -> Result<Self> {
        source.set_metadata_file(props)
    }
}

impl WorkEdit {
    pub(super) fn from(editor: DocxEditor) -> Result<Self> {
        Ok(match editor {
            DocxEditor::Edit(edit) => Self::Eager(edit),
            DocxEditor::Source(edit) => Self::Source(edit),
            DocxEditor::Unsigned(package) => {
                Self::Eager(Box::new(package.edit().map_err(Error::from)?))
            },
        })
    }

    /// Canonical homogeneous fanout can use one native reconstruction. Preserve
    /// ordered fallback for mixed assignments and noncanonical target order.
    fn bulk_paragraph_text(
        &mut self,
        targets: &[Target],
        assignments: &[Assignment<'_>],
    ) -> Result<bool> {
        let [assignment] = assignments else {
            return Ok(false);
        };
        if targets.len() < 2
            || !matches!(assignment.entry.access,
                crate::registry::Access::Set(setter) | crate::registry::Access::SetAndAdd(setter)
                if setter.is_text())
        {
            return Ok(false);
        }
        let mut previous = None;
        for target in targets {
            let Selection::Body(WriteSelection::Paragraph(position)) = target.selection else {
                return Ok(false);
            };
            if previous.is_some_and(|prior| position.get() <= prior) {
                return Ok(false);
            }
            previous = Some(position.get());
        }
        let text = assignment
            .entry
            .prepare(Operation::Set, &assignment.value, |setter, value| {
                if setter.is_text() {
                    text_value(value).map(Some)
                } else {
                    Ok(None)
                }
            })
            .map_err(|error| property_error(error, &assignment.key))?;
        let Some(text) = text else {
            return Ok(false);
        };
        let mut replacements = Vec::new();
        replacements
            .try_reserve_exact(targets.len())
            .map_err(|source| Error::from(crate::transaction::Error::Allocation { source }))?;
        for target in targets {
            if let Selection::Body(WriteSelection::Paragraph(position)) = target.selection {
                replacements.push(litchi_docx::document::ParagraphTextReplacement::new(
                    position,
                    text.clone(),
                ));
            }
        }
        let result = match self {
            Self::Eager(edit) => edit
                .replace_body_paragraph_texts(&replacements)
                .map(|_| ())
                .map_err(Error::from),
            Self::Source(edit) => edit
                .replace_body_paragraph_texts(&replacements)
                .map(|_| ())
                .map_err(Error::from),
        };
        result.map_err(|error| property_error(error, &assignment.key))?;
        Ok(true)
    }

    fn apply(&mut self, target: &Target, assignment: &Assignment<'_>) -> Result<()> {
        self.apply_with_policy(target, assignment, false)
    }

    fn apply_with_policy(
        &mut self,
        target: &Target,
        assignment: &Assignment<'_>,
        command_text: bool,
    ) -> Result<()> {
        let selection = match &target.selection {
            Selection::Body(selection) => *selection,
            Selection::Element(path) => {
                return self
                    .apply_element(path, target.kind, assignment)
                    .map_err(|error| property_error(error, &assignment.key));
            },
        };
        let result = match selection {
            WriteSelection::Paragraph(position) => {
                assignment
                    .entry
                    .prepare(Operation::Set, &assignment.value, |setter, value| {
                        if setter.is_text() {
                            let text = text_value(value)?;
                            self.replace_paragraph_text(position, text)
                        } else {
                            self.format_paragraph(position, setter, value)
                        }
                    })
            },
            WriteSelection::Run { paragraph, run } => {
                assignment
                    .entry
                    .prepare(Operation::Set, &assignment.value, |setter, value| {
                        if setter.is_text() {
                            let text = text_value(value)?;
                            if command_text {
                                self.replace_run_text_node(paragraph, run, text)
                            } else {
                                self.replace_run_text(paragraph, run, text)
                            }
                        } else {
                            self.format_run(paragraph, run, setter, value)
                        }
                    })
            },
        };
        result.map_err(|error| property_error(error, &assignment.key))
    }

    /// Apply every admitted assignment of a table or row target in one
    /// property edit (record 0342).
    fn format_structure(&mut self, target: &Target, assignments: &[Assignment<'_>]) -> Result<()> {
        let Selection::Element(path) = &target.selection else {
            return Err(Error::new(
                Code::InternalError,
                "a DOCX table or row target has an element path",
            ));
        };
        let mut failure = None;
        let mut stage =
            |apply: &mut dyn FnMut(Setter, Option<crate::value::Scalar<'_>>) -> Result<()>| {
                for assignment in assignments {
                    if failure.is_some() {
                        break;
                    }
                    if let Err(error) = assignment.entry.prepare(
                        Operation::Set,
                        &assignment.value,
                        |setter, value| apply(setter, value),
                    ) {
                        failure = Some(property_error(error, &assignment.key));
                    }
                }
            };
        let result = match target.kind {
            TargetKind::Table => both!(self, |edit| edit
                .format_table_at(path, |table| {
                    stage(&mut |setter, value| setter.apply_table(table, value));
                })
                .map(|_| ())
                .map_err(Error::from)),
            _ => both!(self, |edit| edit
                .format_row_at(path, |row| {
                    stage(&mut |setter, value| setter.apply_row(row, value));
                })
                .map(|_| ())
                .map_err(Error::from)),
        };
        // A refused value wins over what the staged change then did.
        failure.map_or(result, Err)
    }

    /// Apply one admitted assignment to an element reached by its leaf path.
    /// Join a paragraph to a list, or remove its numbering, in one
    /// formatting operation (record 0251e). An explicit `level` wins over
    /// the planned one.
    fn set_numbering(
        &mut self,
        target: &Target,
        numbering: Option<(u32, u8)>,
        keep_level: bool,
    ) -> Result<()> {
        let apply = |properties: &mut litchi_docx::properties::Paragraph| match numbering {
            Some((id, level)) => {
                let numbering = properties.numbering.get_or_insert_default();
                numbering.id = Some(id);
                if !keep_level || numbering.level.is_none() {
                    numbering.level = Some(level);
                }
            },
            None => properties.numbering = None,
        };
        match &target.selection {
            Selection::Body(WriteSelection::Paragraph(position)) => {
                both!(self, |edit| edit
                    .format_paragraph(*position, apply)
                    .map(|_| ())
                    .map_err(Error::from))
            },
            Selection::Element(path) => both!(self, |edit| edit
                .format_paragraph_at(path, apply)
                .map(|_| ())
                .map_err(Error::from)),
            Selection::Body(WriteSelection::Run { .. }) => Err(Error::new(
                Code::UnsupportedCapability,
                "listStyle applies to paragraphs",
            )
            .with_property("listStyle")),
        }
    }

    fn apply_element(
        &mut self,
        path: &litchi_docx::document::ElementPath,
        kind: TargetKind,
        assignment: &Assignment<'_>,
    ) -> Result<()> {
        assignment
            .entry
            .prepare(Operation::Set, &assignment.value, |setter, value| {
                if setter.is_text() {
                    let text = text_value(value)?;
                    return if kind == TargetKind::Cell {
                        both!(self, |edit| edit
                            .replace_cell_text_at(path, text)
                            .map(|_| ())
                            .map_err(Error::from))
                    } else {
                        both!(self, |edit| edit
                            .replace_text_at(path, text)
                            .map(|_| ())
                            .map_err(Error::from))
                    };
                }
                if setter == Setter::HyperlinkTooltip {
                    let tooltip = match value {
                        None => None,
                        value => Some(text_value(value)?),
                    };
                    return both!(self, |edit| edit
                        .set_hyperlink_tooltip_at(path, tooltip.as_deref())
                        .map(|_| ())
                        .map_err(Error::from));
                }
                if setter == Setter::HyperlinkUrl {
                    let url = text_value(value)?;
                    let Self::Eager(edit) = self else {
                        return Err(Error::new(
                            Code::InternalError,
                            "a hyperlink target change needs the owning package editor",
                        ));
                    };
                    edit.set_hyperlink_target_at(
                        path,
                        &litchi_docx::document::HyperlinkTarget::Url(url),
                    )
                    .map_err(Error::from)?;
                    return Ok(());
                }
                let mut failure = None;
                match kind {
                    TargetKind::Paragraph => both!(self, |edit| edit
                        .format_paragraph_at(path, |paragraph| {
                            if let Err(error) = setter.apply_paragraph(paragraph, value) {
                                failure = Some(error);
                            }
                        })
                        .map(|_| ())
                        .map_err(Error::from))?,
                    TargetKind::Run => both!(self, |edit| edit
                        .format_run_at(path, |run| {
                            if let Err(error) = setter.apply_run(run, value) {
                                failure = Some(error);
                            }
                        })
                        .map(|_| ())
                        .map_err(Error::from))?,
                    TargetKind::Table => both!(self, |edit| edit
                        .format_table_at(path, |table| {
                            if let Err(error) = setter.apply_table(table, value) {
                                failure = Some(error);
                            }
                        })
                        .map(|_| ())
                        .map_err(Error::from))?,
                    TargetKind::Row => both!(self, |edit| edit
                        .format_row_at(path, |row| {
                            if let Err(error) = setter.apply_row(row, value) {
                                failure = Some(error);
                            }
                        })
                        .map(|_| ())
                        .map_err(Error::from))?,
                    TargetKind::Hyperlink | TargetKind::Cell => {
                        return Err(Error::new(
                            Code::UnsupportedProperty,
                            "the DOCX element accepts only its text property",
                        ));
                    },
                }
                failure.map_or(Ok(()), Err)
            })
    }

    fn replace_paragraph_text(&mut self, position: Position, text: String) -> Result<()> {
        match self {
            Self::Eager(edit) => edit
                .replace_paragraph_text(position, text)
                .map(|_| ())
                .map_err(Error::from),
            Self::Source(edit) => edit
                .replace_paragraph_text(position, text)
                .map(|_| ())
                .map_err(Error::from),
        }
    }

    fn replace_run_text_node(
        &mut self,
        paragraph: Position,
        run: Position,
        text: String,
    ) -> Result<()> {
        match self {
            Self::Eager(edit) => {
                edit.replace_run_text_node(paragraph, run, text)?;
                edit.mark_cached_run_field_dirty(paragraph, run)?;
            },
            Self::Source(edit) => {
                edit.replace_run_text_node(paragraph, run, text)?;
                edit.mark_cached_run_field_dirty(paragraph, run)?;
            },
        }
        Ok(())
    }

    fn replace_run_text(&mut self, paragraph: Position, run: Position, text: String) -> Result<()> {
        match self {
            Self::Eager(edit) => edit
                .replace_run_text(paragraph, run, text)
                .map(|_| ())
                .map_err(Error::from),
            Self::Source(edit) => edit
                .replace_run_text(paragraph, run, text)
                .map(|_| ())
                .map_err(Error::from),
        }
    }

    fn format_paragraph(
        &mut self,
        position: Position,
        setter: Setter,
        value: Option<crate::value::Scalar<'_>>,
    ) -> Result<()> {
        let mut failure = None;
        match self {
            Self::Eager(edit) => {
                edit.format_paragraph(position, |paragraph| {
                    if let Err(error) = setter.apply_paragraph(paragraph, value) {
                        failure = Some(error);
                    }
                })
                .map_err(Error::from)?;
            },
            Self::Source(edit) => {
                edit.format_paragraph(position, |paragraph| {
                    if let Err(error) = setter.apply_paragraph(paragraph, value) {
                        failure = Some(error);
                    }
                })
                .map_err(Error::from)?;
            },
        }
        failure.map_or(Ok(()), Err)
    }

    fn format_run(
        &mut self,
        paragraph: Position,
        run: Position,
        setter: Setter,
        value: Option<crate::value::Scalar<'_>>,
    ) -> Result<()> {
        let mut failure = None;
        match self {
            Self::Eager(edit) => {
                edit.format_run(paragraph, run, |run| {
                    if let Err(error) = setter.apply_run(run, value) {
                        failure = Some(error);
                    }
                })
                .map_err(Error::from)?;
            },
            Self::Source(edit) => {
                edit.format_run(paragraph, run, |run| {
                    if let Err(error) = setter.apply_run(run, value) {
                        failure = Some(error);
                    }
                })
                .map_err(Error::from)?;
            },
        }
        failure.map_or(Ok(()), Err)
    }

    pub(super) fn insert_paragraph(
        &mut self,
        anchor: litchi_docx::document::BodyAnchor,
        paragraph: &litchi_docx::document::NewParagraph,
    ) -> Result<Position> {
        match self {
            Self::Eager(edit) => edit
                .insert_formatted_paragraph(anchor, paragraph)
                .map_err(Error::from),
            Self::Source(edit) => edit
                .insert_formatted_paragraph(anchor, paragraph)
                .map_err(Error::from),
        }
    }

    pub(super) fn remove_paragraph(&mut self, position: Position) -> Result<()> {
        match self {
            Self::Eager(edit) => edit
                .remove_paragraph(position)
                .map(|_| ())
                .map_err(Error::from),
            Self::Source(edit) => edit
                .remove_paragraph(position)
                .map(|_| ())
                .map_err(Error::from),
        }
    }

    pub(super) fn move_paragraph(
        &mut self,
        position: Position,
        anchor: litchi_docx::document::BodyAnchor,
    ) -> Result<Position> {
        match self {
            Self::Eager(edit) => edit.move_paragraph(position, anchor).map_err(Error::from),
            Self::Source(edit) => edit.move_paragraph(position, anchor).map_err(Error::from),
        }
    }

    pub(super) fn swap_paragraphs(&mut self, first: Position, second: Position) -> Result<()> {
        match self {
            Self::Eager(edit) => edit
                .swap_paragraphs(first, second)
                .map(|_| ())
                .map_err(Error::from),
            Self::Source(edit) => edit
                .swap_paragraphs(first, second)
                .map(|_| ())
                .map_err(Error::from),
        }
    }

    pub(super) fn insert_run(
        &mut self,
        paragraph: Position,
        run: Option<Position>,
        new_run: &litchi_docx::document::NewRun,
    ) -> Result<Position> {
        match self {
            Self::Eager(edit) => match run {
                Some(run) => edit
                    .insert_run(paragraph, run, new_run)
                    .map_err(Error::from),
                None => edit.append_run(paragraph, new_run).map_err(Error::from),
            },
            Self::Source(edit) => match run {
                Some(run) => edit
                    .insert_run(paragraph, run, new_run)
                    .map_err(Error::from),
                None => edit.append_run(paragraph, new_run).map_err(Error::from),
            },
        }
    }

    pub(super) fn remove_run(&mut self, paragraph: Position, run: Position) -> Result<()> {
        match self {
            Self::Eager(edit) => edit
                .remove_run(paragraph, run)
                .map(|_| ())
                .map_err(Error::from),
            Self::Source(edit) => edit
                .remove_run(paragraph, run)
                .map(|_| ())
                .map_err(Error::from),
        }
    }

    pub(super) fn move_run(
        &mut self,
        paragraph: Position,
        run: Position,
        anchor: litchi_docx::document::RunAnchor,
    ) -> Result<Position> {
        match self {
            Self::Eager(edit) => edit.move_run(paragraph, run, anchor).map_err(Error::from),
            Self::Source(edit) => edit.move_run(paragraph, run, anchor).map_err(Error::from),
        }
    }

    pub(super) fn move_run_to(
        &mut self,
        from_paragraph: Position,
        run: Position,
        to_paragraph: Position,
        anchor: litchi_docx::document::RunAnchor,
    ) -> Result<Position> {
        match self {
            Self::Eager(edit) => edit
                .move_run_to(from_paragraph, run, to_paragraph, anchor)
                .map_err(Error::from),
            Self::Source(edit) => edit
                .move_run_to(from_paragraph, run, to_paragraph, anchor)
                .map_err(Error::from),
        }
    }

    pub(super) fn swap_runs(
        &mut self,
        paragraph: Position,
        first: Position,
        second: Position,
    ) -> Result<()> {
        match self {
            Self::Eager(edit) => edit
                .swap_runs(paragraph, first, second)
                .map(|_| ())
                .map_err(Error::from),
            Self::Source(edit) => edit
                .swap_runs(paragraph, first, second)
                .map(|_| ())
                .map_err(Error::from),
        }
    }

    pub(super) fn swap_runs_to(
        &mut self,
        first_paragraph: Position,
        first: Position,
        second_paragraph: Position,
        second: Position,
    ) -> Result<()> {
        match self {
            Self::Eager(edit) => edit
                .swap_runs_to(first_paragraph, first, second_paragraph, second)
                .map(|_| ())
                .map_err(Error::from),
            Self::Source(edit) => edit
                .swap_runs_to(first_paragraph, first, second_paragraph, second)
                .map(|_| ())
                .map_err(Error::from),
        }
    }

    pub(super) fn insert_paragraph_at(
        &mut self,
        container: &litchi_docx::document::ElementPath,
        anchor: litchi_docx::document::ElementAnchor,
        paragraph: &litchi_docx::document::NewParagraph,
    ) -> Result<Position> {
        both!(self, |edit| edit
            .insert_paragraph_at(container, anchor, paragraph)
            .map_err(Error::from))
    }

    pub(super) fn insert_table_at(
        &mut self,
        container: &litchi_docx::document::ElementPath,
        anchor: litchi_docx::document::ElementAnchor,
        table: &litchi_docx::document::NewTable,
    ) -> Result<Position> {
        both!(self, |edit| edit
            .insert_table_at(container, anchor, table)
            .map_err(Error::from))
    }

    /// Change the direct properties of the table at `path` (record 0342).
    pub(super) fn format_table_at(
        &mut self,
        path: &litchi_docx::document::ElementPath,
        change: impl FnOnce(&mut litchi_docx::properties::Table),
    ) -> Result<()> {
        both!(self, |edit| edit
            .format_table_at(path, change)
            .map(|_| ())
            .map_err(Error::from))
    }

    pub(super) fn insert_row_at(
        &mut self,
        table: &litchi_docx::document::ElementPath,
        at: Option<Position>,
        cells: Option<usize>,
    ) -> Result<Position> {
        both!(self, |edit| edit
            .insert_row_at(table, at, cells, &[])
            .map_err(Error::from))
    }

    /// Insert a grid column into the table at `table` (record 0345).
    pub(super) fn insert_column_at(
        &mut self,
        table: &litchi_docx::document::ElementPath,
        at: Option<Position>,
        width: Option<u32>,
        text: &str,
    ) -> Result<Position> {
        both!(self, |edit| edit
            .insert_column_at(table, at, width, text)
            .map_err(Error::from))
    }

    /// Set a grid column's width (record 0345).
    pub(super) fn set_column_width_at(
        &mut self,
        table: &litchi_docx::document::ElementPath,
        column: Position,
        width: u32,
    ) -> Result<()> {
        both!(self, |edit| edit
            .set_column_width_at(table, column, width)
            .map(|_| ())
            .map_err(Error::from))
    }

    /// Remove a grid column and its cells (record 0345).
    pub(super) fn remove_column_at(
        &mut self,
        table: &litchi_docx::document::ElementPath,
        column: Position,
    ) -> Result<()> {
        both!(self, |edit| edit
            .remove_column_at(table, column)
            .map(|_| ())
            .map_err(Error::from))
    }

    /// Insert a cell into the row at `row` (record 0345).
    pub(super) fn insert_cell_at(
        &mut self,
        row: &litchi_docx::document::ElementPath,
        at: Option<Position>,
        cell: &litchi_docx::document::NewCell,
    ) -> Result<Position> {
        both!(self, |edit| edit
            .insert_cell_at(row, at, cell)
            .map_err(Error::from))
    }

    pub(super) fn insert_run_at(
        &mut self,
        owner: &litchi_docx::document::ElementPath,
        at: Option<Position>,
        run: &litchi_docx::document::NewRun,
    ) -> Result<Position> {
        both!(self, |edit| edit
            .insert_run_at(owner, at, run)
            .map_err(Error::from))
    }

    pub(super) fn remove_element_at(
        &mut self,
        path: &litchi_docx::document::ElementPath,
    ) -> Result<()> {
        both!(self, |edit| edit
            .remove_element_at(path)
            .map(|_| ())
            .map_err(Error::from))
    }

    pub(super) fn move_element_at(
        &mut self,
        path: &litchi_docx::document::ElementPath,
        container: &litchi_docx::document::ElementPath,
        anchor: litchi_docx::document::ElementAnchor,
    ) -> Result<litchi_docx::document::ElementPath> {
        both!(self, |edit| edit
            .move_element_at(path, container, anchor)
            .map_err(Error::from))
    }

    /// Commit and adopt the published package, dropping the patch.
    pub(super) fn commit_file(self, source: &Docx) -> Result<Docx> {
        match self {
            Self::Eager(edit) => {
                let package = edit.commit().map_err(Error::from)?.into_package();
                Ok(source.adopt_eager_package(package))
            },
            Self::Source(edit) => {
                let package = edit.commit().map_err(Error::from)?.into_package();
                Ok(source.adopt_source_package(package))
            },
        }
    }

    /// Commit and keep only the main-document patch and diagnostics.
    pub(super) fn commit_patch(
        self,
    ) -> Result<(
        litchi_docx::document::Patch,
        litchi_docx::document::Diagnostics,
    )> {
        Ok(match self {
            Self::Eager(edit) => {
                let commit = edit.commit().map_err(Error::from)?;
                let diagnostics = commit.diagnostics();
                (commit.into_parts().1, diagnostics)
            },
            Self::Source(edit) => {
                let commit = edit.commit().map_err(Error::from)?;
                let diagnostics = commit.diagnostics();
                (commit.into_parts().1, diagnostics)
            },
        })
    }

    pub(super) fn commit(self, source: &Docx) -> Result<Commit> {
        match self {
            Self::Eager(edit) => {
                let commit = edit.commit().map_err(Error::from)?;
                let diagnostics = commit.diagnostics();
                let operations = diagnostics.operations();
                let (package, patch) = commit.into_parts();
                Ok(Commit {
                    file: source.adopt_eager_package(package),
                    patch,
                    diagnostics,
                    operations,
                })
            },
            Self::Source(edit) => {
                let commit = edit.commit().map_err(Error::from)?;
                let diagnostics = commit.diagnostics();
                let operations = diagnostics.operations();
                let (package, patch) = commit.into_parts();
                Ok(Commit {
                    file: source.adopt_source_package(package),
                    patch,
                    diagnostics,
                    operations,
                })
            },
        }
    }
}

fn admit_targets(nodes: &[Node<'_>]) -> Result<Few<Target>> {
    let mut targets = Few::with_capacity(nodes.len());
    // Every node comes from one resolve_targets call on one immutable owner.
    // Cache only within this admission; later operations may have a new body.
    let mut body = None;
    for node in nodes {
        let kind = match node.element_kind() {
            Kind::Paragraph => TargetKind::Paragraph,
            Kind::Run => TargetKind::Run,
            Kind::Hyperlink => TargetKind::Hyperlink,
            Kind::Cell => TargetKind::Cell,
            Kind::Table => TargetKind::Table,
            Kind::Row => TargetKind::Row,
            _ => {
                return Err(Error::new(
                    Code::UnsupportedElement,
                    "the DOCX write target must be a paragraph, run, hyperlink, cell, table or row",
                )
                .with_format(FileFormat::Docx));
            },
        };
        let selection = match kind {
            TargetKind::Paragraph | TargetKind::Run => {
                match node.write_selection_cached(&mut body) {
                    Ok(selection) => Selection::Body(selection),
                    // Not a direct-body paragraph or direct run: try the
                    // element path, which proves nested positions itself.
                    Err(error) if error.code() == Code::UnsupportedCapability => {
                        Selection::Element(element::element_path(node)?)
                    },
                    Err(error) => return Err(error),
                }
            },
            TargetKind::Hyperlink | TargetKind::Cell | TargetKind::Table | TargetKind::Row => {
                Selection::Element(element::element_path(node)?)
            },
        };
        targets.push(Target { selection, kind });
    }
    let first = targets[0].kind;
    if targets.iter().any(|target| target.kind != first) {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "one DOCX mutation cannot mix target kinds",
        )
        .with_format(FileFormat::Docx));
    }
    Ok(targets)
}

/// Apply one admitted assignment to a story element on the story's edit,
/// as [`WorkEdit::apply_element`] does on the main document's.
fn apply_story(
    story: &mut litchi_docx::document::raw::Tree,
    path: &litchi_docx::document::ElementPath,
    kind: TargetKind,
    assignment: &Assignment<'_>,
) -> Result<()> {
    assignment
        .entry
        .prepare(Operation::Set, &assignment.value, |setter, value| {
            if setter.is_text() {
                let text = text_value(value)?;
                story.replace_text_at(path, text).map_err(Error::from)?;
                return Ok(());
            }
            let mut failure = None;
            match kind {
                TargetKind::Paragraph => {
                    story
                        .format_paragraph_at(path, |paragraph| {
                            if let Err(error) = setter.apply_paragraph(paragraph, value) {
                                failure = Some(error);
                            }
                        })
                        .map_err(Error::from)?;
                },
                TargetKind::Run => {
                    story
                        .format_run_at(path, |run| {
                            if let Err(error) = setter.apply_run(run, value) {
                                failure = Some(error);
                            }
                        })
                        .map_err(Error::from)?;
                },
                TargetKind::Hyperlink | TargetKind::Cell | TargetKind::Table | TargetKind::Row => {
                    return Err(Error::new(
                        Code::UnsupportedProperty,
                        "the DOCX element accepts only its text property",
                    ));
                },
            }
            failure.map_or(Ok(()), Err)
        })
}

fn admit_assignments<'a>(kind: TargetKind, props: &'a Props<'_>) -> Result<Few<Assignment<'a>>> {
    match kind {
        TargetKind::Paragraph => admit_table(&PARAGRAPH_REGISTRY, kind, props),
        TargetKind::Run => admit_table(&RUN_REGISTRY, kind, props),
        TargetKind::Hyperlink => admit_table(&HYPERLINK_REGISTRY, kind, props),
        TargetKind::Cell => admit_table(&CELL_REGISTRY, kind, props),
        TargetKind::Table => admit_table(&TABLE_REGISTRY, kind, props),
        TargetKind::Row => admit_table(&ROW_REGISTRY, kind, props),
    }
}

fn admit_table<'a, const N: usize, const A: usize>(
    table: &Table<Setter, N, A>,
    kind: TargetKind,
    props: &'a Props<'_>,
) -> Result<Few<Assignment<'a>>> {
    let mut assignments = Few::with_capacity(props.len());
    for (key, value) in props.iter() {
        let (_, entry) = table
            .resolve(key.as_str())
            .map_err(|error| property_error(error, key))?;
        let entry = *entry;
        entry
            .prepare(Operation::Set, value, |setter, value| {
                // Run the format conversion on a disposable typed property
                // now.  This keeps leaf-domain failures before any editor
                // operation, preserving all-or-nothing assignment admission.
                if setter.is_text() {
                    Ok(setter)
                } else {
                    match kind {
                        TargetKind::Paragraph => {
                            let mut probe = litchi_docx::properties::Paragraph::default();
                            setter.apply_paragraph(&mut probe, value)?;
                        },
                        TargetKind::Run => {
                            let mut probe = litchi_docx::properties::Run::default();
                            setter.apply_run(&mut probe, value)?;
                        },
                        TargetKind::Table => {
                            let mut probe = litchi_docx::properties::Table::default();
                            setter.apply_table(&mut probe, value)?;
                        },
                        TargetKind::Row => {
                            let mut probe = litchi_docx::properties::Row::default();
                            setter.apply_row(&mut probe, value)?;
                        },
                        // Their registries hold only `text` as writable.
                        TargetKind::Hyperlink | TargetKind::Cell => {},
                    }
                    Ok(setter)
                }
            })
            .map_err(|error| property_error(error, key))?;
        assignments.push(Assignment {
            key: key.reborrow(),
            value: value.reborrow(),
            entry,
        });
    }
    Ok(assignments)
}

/// Split the `listStyle` and `start` assignments (record 0251e) from the
/// others, which stay in `assignments`. The list is changed only when it
/// holds one of them (a moved list is a copy of its inline element).
fn list_request(assignments: &mut Few<Assignment<'_>>) -> Result<Option<super::lists::Request>> {
    let is_list =
        |assignment: &Assignment<'_>| super::lists::list_setter(assignment.entry.access).is_some();
    if !assignments.iter().any(is_list) {
        return Ok(None);
    }
    let mut request = None::<super::lists::Request>;
    let mut rest = Few::with_capacity(assignments.len());
    for assignment in std::mem::take(assignments) {
        match super::lists::list_setter(assignment.entry.access) {
            Some(Setter::ParagraphListStyle) => {
                let style = assignment
                    .entry
                    .prepare(Operation::Set, &assignment.value, |_, value| {
                        Ok(super::lists::style(&text_value(value)?))
                    })
                    .map_err(|error| property_error(error, &assignment.key))?;
                request
                    .get_or_insert(super::lists::Request {
                        style: None,
                        start: None,
                    })
                    .style = Some(style);
            },
            Some(Setter::ParagraphListStart) => {
                let start = assignment
                    .entry
                    .prepare(Operation::Set, &assignment.value, |_, value| {
                        super::lists::start(value)
                    })
                    .map_err(|error| property_error(error, &assignment.key))?;
                request
                    .get_or_insert(super::lists::Request {
                        style: None,
                        start: None,
                    })
                    .start = Some(start);
            },
            _ => rest.push(assignment),
        }
    }
    *assignments = rest;
    Ok(request)
}

fn text_value(value: Option<crate::value::Scalar<'_>>) -> Result<String> {
    match value {
        None => Ok(String::new()),
        Some(crate::value::Scalar::Text(value) | crate::value::Scalar::Token(value)) => {
            Ok(value.to_owned())
        },
        Some(_) => Err(Error::new(Code::InvalidValue, "the text value is invalid")),
    }
}

fn property_error(error: Error, key: &Key<'_>) -> Error {
    error
        .with_property(key.as_str().to_owned())
        .with_format(FileFormat::Docx)
}

#[cfg(test)]
#[path = "commit_tests.rs"]
mod commit_tests;
