//! Immutable Office Add-in task-pane edits of an XLSB [`Package`] (record
//! 0360b).
//!
//! [`Package::edit_task_panes`] starts an [`Edit`] of the package-level
//! task-pane graph; [`Edit::replace`] stores a complete pane collection
//! after validating every binary worksheet `appRef` against it, and
//! [`Edit::clear`] removes the graph when no worksheet binding would
//! dangle. Every operation works on a private candidate graph, so an error
//! stages nothing. [`Edit::commit`] returns a [`Commit`] with the new
//! package, the exact part-level [`Patch`] (the task-pane parts and the
//! package-level relationships) and the [`Diagnostics`];
//! [`Package::apply_task_panes`] replays a patch onto an equal package.
//! Panes equal to the stored graph publish nothing, so a signed workbook
//! keeps its signature byte for byte; any other change of a signed workbook
//! is refused with [`Error::Signed`].

use litchi_ooxml_common::web::{self as common_web, Conformance, Panes};
use litchi_opc::OpcPackage;

use crate::package::Package;
use crate::package::error::Result;
use crate::part_edit::{Staged, family};

family!("task-pane");

pub use crate::opc::LimitKind;

/// Configuration of a task-pane edit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    conformance: Option<Conformance>,
}

impl EditOptions {
    /// The OOXML conformance class new task-pane parts are written in. The
    /// default is the class of the stored graph, or transitional when there
    /// is none.
    #[must_use]
    pub const fn conformance(mut self, conformance: Conformance) -> Self {
        self.conformance = Some(conformance);
        self
    }

    /// The configured conformance class, if any.
    #[must_use]
    pub const fn conformance_value(&self) -> Option<Conformance> {
        self.conformance
    }
}

/// A staged task-pane edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    conformance: Option<Conformance>,
}

impl Package {
    /// The inert persisted Office Add-in task panes of this package.
    ///
    /// # Errors
    ///
    /// Returns the task-pane graph and decoding errors.
    pub fn task_panes(&self) -> Result<Option<Panes>> {
        Ok(common_web::load(
            self.opc_package(),
            &common_web::Limits::standard(),
        )?)
    }

    /// Start a task-pane edit with the default options.
    ///
    /// # Errors
    ///
    /// Returns the task-pane graph errors of this package.
    pub fn edit_task_panes(&self) -> Result<Edit> {
        self.edit_task_panes_with_options(EditOptions::default())
    }

    /// Start a task-pane edit with explicit options.
    ///
    /// # Errors
    ///
    /// Returns the task-pane graph errors of this package.
    pub fn edit_task_panes_with_options(&self, options: EditOptions) -> Result<Edit> {
        let staged = Staged::unparsed(self)?;
        Ok(Edit {
            staged,
            conformance: options.conformance,
        })
    }

    /// Apply a task-pane patch to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_task_panes(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply_unparsed(self)?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// The staged task panes; empty when there are none.
    ///
    /// # Errors
    ///
    /// Returns the task-pane graph and decoding errors.
    pub fn panes(&self) -> Result<Panes> {
        Ok(
            common_web::load(self.staged.opc(), &common_web::Limits::standard())?
                .unwrap_or_default(),
        )
    }

    /// Store a complete task-pane collection after validating every binary
    /// worksheet `appRef` against it; an empty collection removes the graph.
    /// Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a worksheet binding the panes do not declare,
    /// invalid panes, and a changed edit of a signed workbook.
    pub fn replace(&mut self, panes: Panes) -> Result<bool> {
        if panes.is_empty() {
            return self.clear();
        }
        let conformance = self.conformance()?;
        self.staged
            .workbook()?
            .validate_task_pane_bindings(&panes)?;
        let patch = common_web::plan_put(
            self.staged.opc(),
            panes,
            conformance,
            &common_web::Limits::standard(),
        )?;
        self.stage(&patch)
    }

    /// Remove the task-pane graph when no binary worksheet binding would
    /// dangle. Returns whether a graph was removed.
    ///
    /// # Errors
    ///
    /// Returns an error for a worksheet binding that would dangle and a
    /// changed edit of a signed workbook.
    pub fn clear(&mut self) -> Result<bool> {
        if !self
            .staged
            .opc()
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == common_web::raw::TASK_PANES_RELATIONSHIP)
        {
            return Ok(false);
        }
        self.staged
            .workbook()?
            .validate_task_pane_bindings(&Panes::new())?;
        let patch = common_web::plan_remove(self.staged.opc(), &common_web::Limits::standard())?;
        self.stage(&patch)
    }

    fn stage(&mut self, patch: &common_web::Patch) -> Result<bool> {
        if patch.is_noop() {
            return Ok(false);
        }
        crate::signature_policy::ensure_unsigned(self.staged.opc())?;
        self.staged.update(|candidate| {
            let commit = patch.apply(candidate)?;
            let changed = commit.diagnostics().changed();
            *candidate = commit.into_package();
            Ok(changed)
        })
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged graph to a new package. An unchanged edit returns
    /// the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed edit of a signed workbook.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }

    fn conformance(&mut self) -> Result<Conformance> {
        if let Some(conformance) = self.conformance {
            return Ok(conformance);
        }
        let conformance = existing_conformance(self.staged.opc())?;
        self.conformance = Some(conformance);
        Ok(conformance)
    }
}

/// The conformance class of the stored task-pane graph.
fn existing_conformance(package: &OpcPackage) -> Result<Conformance> {
    let mut relationships = package
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == common_web::raw::TASK_PANES_RELATIONSHIP);
    let Some(relationship) = relationships.next() else {
        return Ok(Conformance::Transitional);
    };
    if relationships.next().is_some() {
        return Err(crate::part_edit::invalid(
            "package has multiple task-pane relationships",
        ));
    }
    if relationship.is_external() {
        return Err(crate::part_edit::invalid(
            "task-pane relationship cannot be external",
        ));
    }
    let part = package.get_part(&relationship.target_partname()?)?;
    let strict_namespace = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
    Ok(
        if part
            .blob()
            .windows(strict_namespace.len())
            .any(|window| window == strict_namespace)
        {
            Conformance::Strict
        } else {
            Conformance::Transitional
        },
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn edit_constructor_defers_default_conformance_detection() {
        let cold = Package::from(Package::new().expect("package").into_opc());
        let edit = cold.edit_task_panes().expect("task-pane edit entry");
        assert!(edit.conformance.is_none());
        assert!(!cold.projection_cached());
    }
}
