//! Exact-markup formatting runs. Text and structural operations are barriers.

use super::{Assignment, Commit, Docx, Target, WorkEdit, admit_assignments, admit_targets};
use crate::registry::Operation;
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
            edit.apply(self.target.selection, assignment)?;
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
    items: Vec<Item>,
    last_target: Option<(Box<str>, Target)>,
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
        if target.path().is_none() || props.is_empty() {
            return Ok(None);
        }
        let result = (|| {
            let mut targets = admit_targets(self.resolve_targets(target)?)?;
            let target = targets
                .pop()
                .ok_or_else(|| Error::new(Code::InternalError, "direct target is missing"))?;
            plan(target, props)
        })();
        self.owner().finish(result)
    }
}
fn plan(target: Target, props: &Props<'_>) -> Result<Option<Plan>> {
    let assignments = admit_assignments(target.kind, props)?;
    for assignment in &assignments {
        if assignment
            .entry
            .prepare(Operation::Set, &assignment.value, |setter, _| {
                Ok(setter.is_text())
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
    pub(crate) fn new(base: &Docx) -> Result<Self> {
        Ok(Self {
            base: base.clone(),
            edit: WorkEdit::from(base.edit()?)?,
            items: Vec::new(),
            last_target: None,
        })
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
                return self.base.owner().finish(plan(*proven, props));
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
    fn projected(&self) -> &litchi_docx::document::Snapshot {
        match &self.edit {
            WorkEdit::Eager(edit) => edit.projected(),
            WorkEdit::Source(edit) => edit.projected(),
        }
    }
    pub(crate) fn snapshot_bytes(&self) -> usize {
        let source = match &self.edit {
            WorkEdit::Eager(edit) => edit.source(),
            WorkEdit::Source(edit) => edit.source(),
        };
        source
            .xml_bytes()
            .len()
            .saturating_add(self.projected().xml_bytes().len())
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
            let before = self.projected().clone();
            plan.stage(&mut self.edit)?;
            let changed = before.xml_bytes() != self.projected().xml_bytes();
            if let Some(spelling) = spelling {
                if self
                    .last_target
                    .as_ref()
                    .is_none_or(|(key, _)| key.as_ref() != spelling)
                {
                    self.last_target = Some((spelling.into(), plan.target));
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
            base, edit, items, ..
        } = self;
        let result = base.owner().finish(edit.commit(&base));
        match result {
            Ok(commit) => Ok((base, commit)),
            Err(error) => Err(locate_failure(&base, &items, error)),
        }
    }
}

fn replay(base: &Docx, items: &[Item]) -> Result<Commit> {
    let result = (|| {
        let mut edit = WorkEdit::from(base.edit()?)?;
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
fn locate_failure(base: &Docx, items: &[Item], error: Error) -> Error {
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
}
