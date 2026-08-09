//! Failure-atomic typed toolbar-control transactions.

use std::borrow::Cow;

use super::control::{Body, Control};
use super::patch::Patch;
use super::snapshot::Snapshot;
use super::validation;
use super::{
    ControlFlags, ControlHeader, Data, Dimensions, Error, GeneralInfo, SpecificFlags, TextIcon,
    WString,
};

/// An isolated typed edit over one toolbar-control snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    candidate: Control<'static>,
}

impl Transaction {
    pub(crate) fn new(source: Snapshot) -> Self {
        Self {
            candidate: source.control().clone(),
            source,
        }
    }

    /// Borrow the immutable source snapshot used by this transaction.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Borrow the current typed candidate.
    #[must_use]
    pub const fn control(&self) -> &Control<'static> {
        &self.candidate
    }

    /// Whether the candidate currently serializes differently from its source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.candidate.to_bytes().as_slice() != self.source.bytes()
    }

    /// Replace all general `TBCFlags`, retaining no invalid reserved bits.
    ///
    /// # Errors
    ///
    /// Returns an error if `flags` or the resulting header is invalid.
    pub fn set_control_flags(&mut self, flags: ControlFlags) -> Result<&mut Self, Error> {
        flags.validate()?;
        let header = self.header_with(flags, self.candidate.header().specifics())?;
        self.replace_header(header)?;
        Ok(self)
    }

    /// Replace all `TBCSFlags`, retaining the body and opaque prefix.
    ///
    /// # Errors
    ///
    /// Returns an error if `specifics` or the resulting header is invalid.
    pub fn set_specific_flags(&mut self, specifics: SpecificFlags) -> Result<&mut Self, Error> {
        specifics.validate()?;
        let header = self.header_with(self.candidate.header().flags(), specifics)?;
        self.replace_header(header)?;
        Ok(self)
    }

    /// Change the shared `textIcon` visibility mode without touching unknown bits.
    ///
    /// # Errors
    ///
    /// Returns an error if the resulting header is invalid.
    pub fn set_text_icon(&mut self, value: TextIcon) -> Result<&mut Self, Error> {
        let specifics = self.candidate.header().specifics().with_text_icon(value);
        let header = self.header_with(self.candidate.header().flags(), specifics)?;
        self.replace_header(header)?;
        Ok(self)
    }

    /// Set or clear the custom control text.
    ///
    /// Adding text also enables `fSaveUIStrings`, as required by
    /// `[MS-OSHARED]`; clearing it leaves unrelated UI-string state intact.
    ///
    /// # Errors
    ///
    /// Returns an error if the text is invalid or the resulting control is
    /// invalid.
    pub fn set_custom_text(&mut self, value: Option<&str>) -> Result<&mut Self, Error> {
        let custom_text = value.map(WString::new).transpose()?;
        let enable_ui_strings = custom_text.is_some();
        self.update_general(
            |general| {
                let flags = general.flags().with_save_text(custom_text.is_some());
                GeneralInfo::new(
                    flags,
                    custom_text,
                    general.description().cloned(),
                    general.tooltip().cloned(),
                    general.extra().cloned(),
                )
            },
            enable_ui_strings,
        )?;
        Ok(self)
    }

    /// Set or clear the paired description and tooltip UI strings.
    ///
    /// # Errors
    ///
    /// Returns an error if only one string is supplied, a string is invalid,
    /// or the resulting control is invalid.
    pub fn set_ui_strings(
        &mut self,
        description: Option<&str>,
        tooltip: Option<&str>,
    ) -> Result<&mut Self, Error> {
        if description.is_some() != tooltip.is_some() {
            return Err(Error::invalid(
                "toolbar description and tooltip must be set or cleared together",
            ));
        }
        let decoded_description = description.map(WString::new).transpose()?;
        let decoded_tooltip = tooltip.map(WString::new).transpose()?;
        let enable_ui_strings = decoded_description.is_some();
        self.update_general(
            |general| {
                let flags = general
                    .flags()
                    .with_save_misc_ui_strings(description.is_some());
                GeneralInfo::new(
                    flags,
                    general.custom_text().cloned(),
                    decoded_description,
                    decoded_tooltip,
                    general.extra().cloned(),
                )
            },
            enable_ui_strings,
        )?;
        Ok(self)
    }

    /// Set the disabled bit in the common general metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the control body is opaque or the resulting control
    /// is invalid.
    pub fn set_disabled(&mut self, value: bool) -> Result<&mut Self, Error> {
        self.update_general(
            |general| {
                GeneralInfo::new(
                    general.flags().with_disabled(value),
                    general.custom_text().cloned(),
                    general.description().cloned(),
                    general.tooltip().cloned(),
                    general.extra().cloned(),
                )
            },
            false,
        )
    }

    /// Replace the priority while preserving all other header fields.
    ///
    /// # Errors
    ///
    /// Returns an error if `priority` or the resulting header is invalid.
    pub fn set_priority(&mut self, priority: u8) -> Result<&mut Self, Error> {
        let header = self.header_with_priority(priority)?;
        self.replace_header(header)?;
        Ok(self)
    }

    /// Set or clear dimensions and keep `fSaveDxy` consistent.
    ///
    /// # Errors
    ///
    /// Returns an error if the resulting header is invalid.
    pub fn set_dimensions(&mut self, dimensions: Option<Dimensions>) -> Result<&mut Self, Error> {
        let flags = self
            .candidate
            .header()
            .flags()
            .with_save_dimensions(dimensions.is_some());
        let header = ControlHeader::from_decoded(
            self.candidate.header().control_type(),
            self.candidate.header().control_id(),
            flags,
            self.candidate.header().specifics(),
            self.candidate.header().priority(),
            dimensions,
        );
        self.replace_header(header)?;
        Ok(self)
    }

    /// Capture the candidate as a snapshot without publishing it.
    ///
    /// # Errors
    ///
    /// Returns an error if the candidate cannot be materialized as a valid
    /// toolbar control.
    pub fn snapshot(&self) -> Result<Snapshot, Error> {
        self.materialize()
    }

    /// Publish the candidate and its reversible source-checked patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the candidate cannot be materialized as a valid
    /// toolbar control.
    pub fn commit(self) -> Result<Commit, Error> {
        let snapshot = self.materialize()?;
        let patch = Patch::new(&self.source, &snapshot);
        Ok(Commit { snapshot, patch })
    }

    /// Alias for move-oriented writer terminology.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::commit`].
    pub fn finish(self) -> Result<Commit, Error> {
        self.commit()
    }

    /// Discard staged changes and recover the source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    fn update_general<F>(&mut self, edit: F, enable_ui_strings: bool) -> Result<&mut Self, Error>
    where
        F: FnOnce(&GeneralInfo<'static>) -> Result<GeneralInfo<'static>, Error>,
    {
        let data = self
            .candidate
            .data()
            .ok_or_else(|| Error::invalid("toolbar control body is opaque or empty"))?;
        let general = edit(&data.general().clone().into_owned())?;
        let replacement = Data::new(general, data.specific().to_vec())?;
        let specifics = if enable_ui_strings {
            self.candidate
                .header()
                .specifics()
                .with_save_ui_strings(true)
        } else {
            self.candidate.header().specifics()
        };
        let header = self.header_with(self.candidate.header().flags(), specifics)?;
        let candidate = Control::from_edited(
            header,
            Cow::Owned(self.candidate.prefix().to_vec()),
            Body::Data(replacement),
        )?;
        self.candidate = candidate;
        Ok(self)
    }

    fn replace_header(&mut self, header: ControlHeader) -> Result<(), Error> {
        let candidate = Control::from_edited(
            header,
            Cow::Owned(self.candidate.prefix().to_vec()),
            self.candidate.body().clone().into_owned(),
        )?;
        self.candidate = candidate;
        Ok(())
    }

    fn header_with(
        &self,
        flags: ControlFlags,
        specifics: SpecificFlags,
    ) -> Result<ControlHeader, Error> {
        let header = ControlHeader::from_decoded(
            self.candidate.header().control_type(),
            self.candidate.header().control_id(),
            flags,
            specifics,
            self.candidate.header().priority(),
            self.candidate.header().dimensions(),
        );
        validation::validate_header(&header)?;
        Ok(header)
    }

    fn header_with_priority(&self, priority: u8) -> Result<ControlHeader, Error> {
        let header = ControlHeader::from_decoded(
            self.candidate.header().control_type(),
            self.candidate.header().control_id(),
            self.candidate.header().flags(),
            self.candidate.header().specifics(),
            priority,
            self.candidate.header().dimensions(),
        );
        validation::validate_header(&header)?;
        Ok(header)
    }

    fn materialize(&self) -> Result<Snapshot, Error> {
        let bytes = self.candidate.to_bytes();
        if bytes.as_slice() == self.source.bytes() {
            return Ok(self.source.clone());
        }
        Snapshot::from_control(self.candidate.clone())
    }
}

/// A successful toolbar-control publication.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Whether the publication changed source bytes.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Borrow the immutable target snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its target snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}
