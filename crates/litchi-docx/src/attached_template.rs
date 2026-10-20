//! The attached-template family (record 0355): the settings part's
//! `w:attachedTemplate` and its external relationship. The template is never
//! dereferenced.
//!
//! [`Package::attached_template`](crate::Package::attached_template) reads it;
//! [`Package::edit_attached_template`](crate::Package::edit_attached_template)
//! stages an [`Edit`] whose commit publishes a new package, and
//! [`Package::apply_attached_template`](crate::Package::apply_attached_template)
//! replays its [`Patch`].
//!
//! ```
//! use litchi_docx::Package;
//!
//! let package = Package::new()?;
//! let mut edit = package.edit_attached_template();
//! edit.set_uri("https://example.invalid/report.dotx")?;
//! let commit = edit.commit()?;
//! assert!(commit.package().attached_template()?.is_some());
//! assert!(package.attached_template()?.is_none());
//! # Ok::<(), litchi_docx::Error>(())
//! ```

pub use crate::settings::AttachedTemplate;

use crate::package::family::{graph_edit, graph_family};

graph_family!("attached-template", own_limits);

/// The budgets of this family: the settings part it rewrites.
pub use crate::settings::LimitKind;
graph_edit!(
    "attached-template",
    edit_attached_template,
    apply_attached_template
);

impl Edit {
    /// Attach the external template at `target_uri`, replacing any attached
    /// template while keeping its relationship identifier.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid target URI or settings part.
    pub fn set_uri(&mut self, target_uri: impl Into<String>) -> Result<&mut Self> {
        let target_uri = target_uri.into();
        self.stage(|package| package.set_attached_template_uri(target_uri).map(|_| ()))?;
        Ok(self)
    }

    /// Detach the template; returns the detached template, if any.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid settings part.
    pub fn remove(&mut self) -> Result<Option<AttachedTemplate>> {
        self.stage(crate::Package::remove_attached_template)
    }
}
