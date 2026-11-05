//! Exact-markup formatting runs, and direct body-paragraph text sets
//! (record 0355a). Structural operations are barriers.
//!
//! A body paragraph's ordinal is stable across text replacement, so `text`
//! on a direct `/body/p[N]` path shares the run's native edit, resolved
//! against the run's base. Replacing a paragraph's text rewrites its runs,
//! so a later target inside such a paragraph, or any element-path target,
//! closes the run and resolves against the published candidate.

use std::collections::BTreeSet;

use super::{
    Assignment, Commit, Docx, Selection, Target, TargetKind, WorkEdit, WriteSelection,
    admit_assignments, admit_targets,
};
use crate::registry::Operation;
#[cfg(test)]
use crate::transaction::SemanticLimits;
use crate::{Code, Error, Props, Result, TargetArg};

pub(crate) struct Plan {
    target: Target,
    assignments: Vec<Assignment<'static>>,
}
impl Plan {
    pub(crate) fn operation_bound(&self) -> usize {
        self.assignments.len()
    }
    fn stage(&self, edit: &mut WorkEdit) -> Result<()> {
        for assignment in &self.assignments {
            edit.apply(&self.target, assignment)?;
        }
        Ok(())
    }
}
struct Item {
    plan: Plan,
    index: usize,
}

pub(crate) struct Run {
    base: Docx,
    edit: WorkEdit,
    snapshot_budget: usize,
    items: Vec<Item>,
    last_target: Option<(Box<str>, Target)>,
    /// Transaction positions of the body paragraphs whose text this run
    /// replaced: their run ordinals no longer match the base.
    texted: BTreeSet<usize>,
}
impl Docx {
    /// Formatting leaves paragraph/run ordinals unchanged. Only structural
    /// direct paths can resolve against the run's base; content selectors need
    /// a committed staged view. Text also closes the run because its commit
    /// compaction policy differs from exact-markup formatting publication.
    pub(crate) fn formatting_plan(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<Plan>> {
        // A metadata set publishes through its own writer (record 0338).
        if target.path().is_none() || props.is_empty() || Self::is_metadata_set(target, props) {
            return Ok(None);
        }
        let result = (|| {
            let nodes = self.resolve_targets(target)?;
            // Charts have their own setter outside the formatting surface
            // (record 0283): the ordinary set path writes them (0289).
            // So do the part elements (record 0319), and pictures,
            // comments, SmartArt and chart series and axes (record 0335).
            if nodes.iter().any(|node| {
                use crate::format::docx::Kind;
                matches!(
                    node.element_kind(),
                    Kind::Chart
                        | Kind::Series
                        | Kind::Axis
                        | Kind::Picture
                        | Kind::Comment
                        | Kind::SmartArt
                        | Kind::Node
                        | Kind::Column
                ) || super::super::parts::is_part_kind(node.element_kind())
                    || super::super::marks::is_mark_kind(node.element_kind())
            }) {
                return Ok(None);
            }
            let mut targets = admit_targets(&nodes)?;
            let target = targets
                .pop()
                .ok_or_else(|| Error::new(Code::InternalError, "direct target is missing"))?;
            let mut planned = plan(target, props)?;
            if let Some(planned) = planned.as_mut() {
                self.resolve_style_names(&mut planned.assignments)?;
            }
            Ok(planned)
        })();
        self.owner().finish(result)
    }

    /// A `text`-only set on a direct body paragraph `/body/p[N]`, the route [`Docx::set_file`] takes for one operation;
    /// `None` sends anything else to the ordinary setter.
    pub(crate) fn text_plan(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<Plan>> {
        let Some(position) = self.direct_body_paragraph(target, props) else {
            return Ok(None);
        };
        let result = admit_assignments(TargetKind::Paragraph, props).map(|assignments| {
            Some(Plan {
                target: Target {
                    selection: Selection::Body(WriteSelection::Paragraph(position)),
                    kind: TargetKind::Paragraph,
                },
                assignments: assignments
                    .into_iter()
                    .map(Assignment::into_owned)
                    .collect(),
            })
        });
        self.owner().finish(result)
    }
}
fn plan(target: Target, props: &Props<'_>) -> Result<Option<Plan>> {
    let assignments = admit_assignments(target.kind, props)?;
    for assignment in &assignments {
        if assignment
            .entry
            .prepare(Operation::Set, &assignment.value, |setter, _| {
                // Lists stage numbering definitions: a barrier (0251e).
                // A hyperlink target changes a relationship, which only the
                // owning package editor can do: a barrier (0289).
                Ok(setter.is_text()
                    || matches!(
                        setter,
                        super::Setter::ParagraphListStyle
                            | super::Setter::ParagraphListStart
                            | super::Setter::HyperlinkUrl
                    ))
            })?
        {
            return Ok(None);
        }
    }
    Ok(Some(Plan {
        target,
        assignments: assignments
            .into_iter()
            .map(Assignment::into_owned)
            .collect(),
    }))
}

impl Run {
    #[cfg(test)]
    pub(crate) fn new(base: &Docx) -> Result<Self> {
        Self::new_with_snapshot_budget(base, SemanticLimits::default().max_snapshot_bytes())
    }

    /// Start a formatting run with the caller's remaining semantic snapshot
    /// budget already attached to the native edit.  Formatting is coalesced
    /// before commit, so admitting only the final pair would let an
    /// intermediate `Snapshot` escape the facade's retention ceiling.
    pub(crate) fn new_with_snapshot_budget(base: &Docx, snapshot_budget: usize) -> Result<Self> {
        let editor = if base.signature() == crate::format::signature::Signature::Signed {
            base.edit_for_write_with_snapshot_budget(snapshot_budget)?
        } else {
            base.edit_with_snapshot_budget(snapshot_budget)?
        };
        Ok(Self {
            base: base.clone(),
            edit: WorkEdit::from(editor)?,
            snapshot_budget,
            items: Vec::new(),
            last_target: None,
            texted: BTreeSet::new(),
        })
    }
    /// [`Docx::text_plan`] against the run's base.
    pub(crate) fn text_plan(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<Plan>> {
        self.base.text_plan(target, props)
    }
    /// Whether the run has replaced a paragraph's text, after which
    /// targets planned against its base may be stale.
    pub(crate) fn has_text(&self) -> bool {
        !self.texted.is_empty()
    }
    /// Whether `plan`, resolved against the run's base, still names the
    /// same content in the run's staged state.
    pub(crate) fn admits(&self, plan: &Plan) -> bool {
        match &plan.target.selection {
            Selection::Body(WriteSelection::Paragraph(_)) => true,
            Selection::Body(WriteSelection::Run { paragraph, .. }) => {
                !self.texted.contains(&paragraph.get())
            },
            Selection::Element(_) => self.texted.is_empty(),
        }
    }
    /// Formatting cannot change paragraph/run ordinals. Reuse only the last
    /// exact direct-path proof, scoped to this run, and keep the source fence.
    pub(crate) fn formatting_plan(
        &self,
        spelling: &str,
        target: &TargetArg<'_>,
        props: &Props<'_>,
    ) -> Result<Option<Plan>> {
        if let Some((key, proven)) = &self.last_target {
            if key.as_ref() == spelling {
                return self.base.owner().finish(plan(proven.clone(), props));
            }
        }
        self.base.formatting_plan(target, props)
    }
    pub(crate) fn first_index(&self) -> Option<usize> {
        self.items.first().map(|item| item.index)
    }
    pub(crate) fn remaining_operations(&self) -> usize {
        match &self.edit {
            WorkEdit::Eager(edit) => edit.remaining_operations(),
            WorkEdit::Source(edit) => edit.remaining_operations(),
        }
    }
    fn projected(&self) -> Result<litchi_docx::document::Snapshot> {
        match &self.edit {
            WorkEdit::Eager(edit) => Ok(edit.projected()?),
            WorkEdit::Source(edit) => Ok(edit.projected()?),
        }
    }
    pub(crate) fn snapshot_bytes(&self) -> Result<usize> {
        let source = match &self.edit {
            WorkEdit::Eager(edit) => edit.source(),
            WorkEdit::Source(edit) => edit.source(),
        }?;
        Ok(source
            .xml_bytes()
            .len()
            .saturating_add(self.projected()?.xml_bytes().len()))
    }
    /// A staging failure poisons the caller transaction or discards its whole
    /// isolated item; no partially staged command can be published.
    pub(crate) fn stage(
        &mut self,
        plan: Plan,
        index: usize,
        spelling: Option<&str>,
    ) -> Result<bool> {
        let result = (|| {
            self.items
                .try_reserve(1)
                .map_err(|source| crate::transaction::Error::Allocation { source })?;
            let before = self.projected()?;
            plan.stage(&mut self.edit)?;
            let changed = before.xml_bytes() != self.projected()?.xml_bytes();
            if let Selection::Body(WriteSelection::Paragraph(position)) = plan.target.selection
                && plan
                    .assignments
                    .iter()
                    .any(|assignment| assignment.key.as_str().eq_ignore_ascii_case("text"))
            {
                self.texted.insert(position.get());
            }
            if let Some(spelling) = spelling {
                if self
                    .last_target
                    .as_ref()
                    .is_none_or(|(key, _)| key.as_ref() != spelling)
                {
                    self.last_target = Some((spelling.into(), plan.target.clone()));
                }
            } else {
                self.last_target = None;
            }
            self.items.push(Item { plan, index });
            Ok(changed)
        })();
        self.base.owner().finish(result)
    }
    pub(crate) fn commit(self) -> Result<(Docx, Commit)> {
        let Self {
            base,
            edit,
            snapshot_budget,
            items,
            ..
        } = self;
        let authored_operations = items.len();
        let result = base.owner().finish(edit.commit(&base));
        match result {
            Ok(commit)
                if base.signature() == crate::format::signature::Signature::Signed
                    && commit.patch.is_noop() =>
            {
                // The staged patch may contain cancelled native operations.
                // Replace that operation-bearing patch with a fresh empty
                // edit over the original signed owner, so the leaf history,
                // inverse and unified facade all carry the same canonical
                // no-op endpoint.  The staged commit has validated the operations; release its
                // proof before constructing the canonical replacement.
                drop(commit);
                let editor = base.edit_for_write_with_snapshot_budget(snapshot_budget)?;
                let mut empty = WorkEdit::from(editor)?.commit(&base)?;
                empty.operations = authored_operations;
                Ok((base, empty))
            },
            Ok(commit) => Ok((base, commit)),
            Err(error) => Err(locate_failure(&base, &items, error, snapshot_budget)),
        }
    }
}

#[cfg(test)]
fn replay(base: &Docx, items: &[Item]) -> Result<Commit> {
    replay_with_snapshot_budget(base, items, SemanticLimits::default().max_snapshot_bytes())
}

fn replay_with_snapshot_budget(
    base: &Docx,
    items: &[Item],
    snapshot_budget: usize,
) -> Result<Commit> {
    let result = (|| {
        let editor = if base.signature() == crate::format::signature::Signature::Signed {
            base.edit_for_write_with_snapshot_budget(snapshot_budget)?
        } else {
            base.edit_with_snapshot_budget(snapshot_budget)?
        };
        let mut edit = WorkEdit::from(editor)?;
        for item in items {
            item.plan
                .stage(&mut edit)
                .map_err(|error| error.with_item(item.index))?;
        }
        edit.commit(base)
    })();
    base.owner().finish(result)
}

/// Error-path-only bisection. Each probe owns a new native editor; successful
/// prefixes become the exact base of later probes. Resource/source failures
/// describe the run and are never retried as smaller operations.
fn locate_failure(base: &Docx, items: &[Item], error: Error, snapshot_budget: usize) -> Error {
    let mut replay =
        |base: &Docx, items: &[Item]| replay_with_snapshot_budget(base, items, snapshot_budget);
    locate_with(base, items, error, &mut replay)
}

fn locate_with(
    base: &Docx,
    items: &[Item],
    error: Error,
    probe: &mut impl FnMut(&Docx, &[Item]) -> Result<Commit>,
) -> Error {
    let Some(first) = items.first() else {
        return error;
    };
    if items.len() == 1
        || matches!(
            error.code(),
            Code::IoError
                | Code::Cancelled
                | Code::LimitExceeded
                | Code::InternalError
                | Code::DocumentProtected
        )
    {
        return error.with_item(first.index);
    }
    let middle = items.len() / 2;
    let prefix = match probe(base, &items[..middle]) {
        Ok(commit) => commit.into_file(),
        Err(error) => return locate_with(base, &items[..middle], error, probe),
    };
    match probe(&prefix, &items[middle..]) {
        Err(error) => locate_with(&prefix, &items[middle..], error, probe),
        // A group-only failure has no uniquely failing command. Attribute it
        // to the first item of the run, as the settled batch error contract.
        Ok(_) => error.with_item(first.index),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{File, IntoTarget, NewNode};

    fn fixture() -> (Docx, Vec<Item>) {
        let file = crate::create(litchi_core::FileFormat::Docx)
            .unwrap()
            .add("/body", NewNode::new("paragraph").prop("text", "one"))
            .unwrap()
            .file;
        let File::Docx(base) = file else {
            panic!("DOCX expected")
        };
        let items = (10..18)
            .map(|index| Item {
                index,
                plan: base
                    .formatting_plan(
                        &"/body/p[1]/r[1]".into_target().unwrap(),
                        &Props::new().with("bold", index % 2 == 0),
                    )
                    .unwrap()
                    .unwrap(),
            })
            .collect();
        (base, items)
    }

    #[test]
    fn commit_failure_bisection_keeps_original_wire_index() {
        let (base, items) = fixture();
        let error = locate_with(
            &base,
            &items,
            Error::new(Code::InvalidInput, "injected commit failure"),
            &mut |base, items| {
                if items.iter().any(|item| item.index == 16) {
                    Err(Error::new(Code::InvalidInput, "injected commit failure"))
                } else {
                    replay(base, items)
                }
            },
        );
        assert_eq!(error.item(), Some(16));
    }

    #[test]
    fn global_commit_failures_never_probe_and_group_only_failures_keep_run_start() {
        let (base, items) = fixture();
        for code in [
            Code::IoError,
            Code::Cancelled,
            Code::LimitExceeded,
            Code::InternalError,
            Code::DocumentProtected,
        ] {
            let error = locate_with(
                &base,
                &items,
                Error::new(code, "global failure"),
                &mut |_, _| panic!("global failure was retried"),
            );
            assert_eq!(error.item(), Some(10));
            assert_eq!(error.code(), code);
        }
        let error = locate_with(
            &base,
            &items,
            Error::new(Code::InvalidInput, "group failure"),
            &mut replay,
        );
        assert_eq!(error.item(), Some(10));
    }

    #[test]
    fn failure_probes_keep_the_run_snapshot_budget() {
        let (base, items) = fixture();
        let run = Run::new(&base).unwrap();
        let budget = run.snapshot_bytes().unwrap().saturating_sub(1);
        let error = match replay_with_snapshot_budget(&base, &items, budget) {
            Ok(_) => panic!("an over-budget replay probe was admitted"),
            Err(error) => error,
        };
        assert_eq!(error.code(), Code::LimitExceeded);
    }
}
