//! [`SummaryProperties`] over the property-set streams of an existing
//! compound file.
//!
//! A metadata transaction on an opened document reads the three sections
//! through the shared [`Edit`], edits them as [`SummaryProperties`], and
//! stages back only the sections that changed. An unchanged section is never
//! re-serialized, so an exact no-op leaves the editor without a staged stream
//! and [`Edit::finish`] returns the source bytes.

use super::super::binding::Binding;
use super::super::codec::Edit;
use super::model::SummaryProperties;
use crate::Error;

impl SummaryProperties {
    /// Reads the current properties of the compound file an editor holds.
    ///
    /// See [`Self::from_sections`] for how missing and blank sections read.
    ///
    /// # Errors
    ///
    /// Returns an error if a stream cannot be parsed or a present standard
    /// section breaks its typed owner's invariants.
    pub fn from_editor(editor: &Edit) -> Result<Self, Error> {
        let summary = editor.property_set(Binding::SummaryInformation)?;
        let document_summary = editor.property_set(Binding::DocumentSummaryInformation)?;
        let user_defined = editor.property_set(Binding::UserDefinedProperties)?;
        Ok(Self::from_sections(
            summary.as_ref(),
            document_summary.as_ref(),
            user_defined.as_ref(),
        )?)
    }

    /// Stages into `editor` every section of `self` that differs from `base`,
    /// the value [`Self::from_editor`] read from the same editor.
    ///
    /// A user-defined section that `self` dropped is removed from the
    /// `DocumentSummaryInformation` stream; a stream the source lacked is
    /// created with its code page. Returns whether anything was staged.
    ///
    /// # Errors
    ///
    /// Returns an error if the editor refuses a replacement (for example a
    /// source stream whose read repair would not preserve its content).
    pub fn stage_changes(&self, base: &Self, editor: &mut Edit) -> Result<bool, Error> {
        let (summary, document_summary, user_defined) = self.sections();
        let (base_summary, base_document_summary, base_user_defined) = base.sections();
        let mut changed = false;
        if summary != base_summary {
            editor.replace(
                Binding::SummaryInformation,
                super::super::model::try_clone_property_set(summary)?,
            )?;
            changed = true;
        }
        if document_summary != base_document_summary {
            editor.replace(
                Binding::DocumentSummaryInformation,
                super::super::model::try_clone_property_set(document_summary)?,
            )?;
            changed = true;
        }
        if user_defined != base_user_defined {
            match user_defined {
                Some(section) => {
                    editor.replace(
                        Binding::UserDefinedProperties,
                        super::super::model::try_clone_property_set(section)?,
                    )?;
                },
                None => {
                    editor.remove(Binding::UserDefinedProperties)?;
                },
            }
            changed = true;
        }
        Ok(changed)
    }
}
