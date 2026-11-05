//! The mail-merge family (record 0355): the settings part's `w:mailMerge`,
//! its data-source relationships and the recipient-data part.
//!
//! [`Package::mail_merge_settings`](crate::Package::mail_merge_settings)
//! reads the settings; [`Package::edit_mail_merge`](crate::Package::edit_mail_merge)
//! stages an [`Edit`] whose commit publishes a new package, and
//! [`Package::apply_mail_merge`](crate::Package::apply_mail_merge) replays its
//! [`Patch`]. Data sources are never connected to or queried.

use super::{Conformance, Recipients, Settings, Source};
use crate::package::family::{graph_edit, graph_family};

graph_family!("mail-merge", own_limits);

pub use super::limit::LimitKind;
graph_edit!("mail-merge", edit_mail_merge, apply_mail_merge);

impl Edit {
    /// Replace the mail-merge settings, data sources and recipients.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid settings or sources and the
    /// signed-package refusal for a change to a signed package.
    pub fn set(
        &mut self,
        settings: Settings,
        data_source: Option<Source>,
        header_source: Option<Source>,
        recipients: Option<Recipients>,
        conformance: Conformance,
    ) -> Result<&mut Self> {
        self.stage(|package| {
            package.set_mail_merge(
                settings,
                data_source,
                header_source,
                recipients,
                conformance,
            )
        })?;
        Ok(self)
    }

    /// Update the existing mail-merge settings, data sources and recipients.
    ///
    /// # Errors
    ///
    /// As [`Self::set`], and an error when the document has no mail merge.
    pub fn update(
        &mut self,
        settings: Settings,
        data_source: Option<Source>,
        header_source: Option<Source>,
        recipients: Option<Recipients>,
        conformance: Conformance,
    ) -> Result<&mut Self> {
        self.stage(|package| {
            package.update_mail_merge(
                settings,
                data_source,
                header_source,
                recipients,
                conformance,
            )
        })?;
        Ok(self)
    }

    /// Replace only the recipient data of the existing mail merge.
    ///
    /// # Errors
    ///
    /// Returns an error when the document has no mail merge.
    pub fn update_recipients(
        &mut self,
        recipients: Recipients,
        conformance: Conformance,
    ) -> Result<&mut Self> {
        self.stage(|package| package.update_mail_merge_recipients(recipients, conformance))?;
        Ok(self)
    }

    /// Remove the mail merge and its graph; returns whether it was present.
    ///
    /// # Errors
    ///
    /// Returns the signed-package refusal for a change to a signed package.
    pub fn clear(&mut self) -> Result<bool> {
        self.stage(crate::Package::clear_mail_merge)
    }
}
