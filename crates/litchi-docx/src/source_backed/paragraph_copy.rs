//! Raw, bounded XML copy and removal for a plain Word main story.
//!
//! Package editing lives in [`super::plain_paragraph_copy`] and
//! [`super::plain_paragraph_removal`].  This module keeps the detached XML
//! codec useful for callers that already own a document part: [`raw::Tree`]
//! admits the bytes under an explicit [`raw::Limits`] policy, applies one
//! exact copy or removal, and writes the resulting bytes through `&self`.
//! Its staging state is private to the tree, so the package families remain
//! the only public edit/commit API.

pub mod raw {
    use std::io::Write;

    use litchi_core::{Position, SourceVersion};

    use super::super::paragraph_copy_record as record;

    /// Result returned by raw plain-paragraph XML operations.
    pub type Result<T> = std::result::Result<T, Error>;

    /// The bounded raw XML codec error.
    pub type Error = record::Error;

    /// A refusal from the exact plain-paragraph closure.
    pub type Refusal = record::Refusal;

    /// The finite policy used by raw XML admission, editing and output.
    pub type Limits = record::Limits;

    /// A resource dimension carried by the raw codec's limit error.
    pub type LimitKind = record::LimitKind;

    /// The effect of a copy operation.
    pub type EffectReport = record::EffectReport;

    /// The effect of a removal operation.
    pub type RemovalEffectReport = record::RemovalEffectReport;

    /// A durable exact copy patch over a plain main-story XML tree.
    #[derive(Debug, Clone)]
    pub struct Patch(record::Patch);

    impl Patch {
        /// Whether this patch leaves the XML unchanged.
        #[must_use]
        pub fn is_noop(&self) -> bool {
            self.0.is_noop()
        }

        /// Return the exact patch that restores the source XML.
        #[must_use]
        pub fn inverse(&self) -> Self {
            Self(self.0.inverse())
        }

        /// Return the deterministic copy effect.
        #[must_use]
        pub fn effect_report(&self) -> EffectReport {
            self.0.effect_report()
        }

        /// Encode the patch under the required caller policy.
        pub fn to_bytes(&self, limits: Limits) -> Result<Vec<u8>> {
            self.0.to_bytes(limits)
        }

        /// Decode and validate a patch under the required caller policy.
        pub fn from_bytes(bytes: &[u8], limits: Limits) -> Result<Self> {
            Ok(Self(record::Patch::from_bytes(bytes, limits)?))
        }

        pub(crate) fn into_record(self) -> record::Patch {
            self.0
        }

        pub(crate) fn from_record(patch: record::Patch) -> Self {
            Self(patch)
        }
    }

    /// A durable exact removal patch over a plain main-story XML tree.
    #[derive(Debug, Clone)]
    pub struct RemovalPatch(record::RemovalPatch);

    impl RemovalPatch {
        /// Whether this patch leaves the XML unchanged.
        #[must_use]
        pub fn is_noop(&self) -> bool {
            self.0.is_noop()
        }

        /// Return the exact patch that restores the source XML.
        #[must_use]
        pub fn inverse(&self) -> Self {
            Self(self.0.inverse())
        }

        /// Return the deterministic removal effect.
        #[must_use]
        pub fn effect_report(&self) -> RemovalEffectReport {
            self.0.effect_report()
        }

        /// Encode the patch under the required caller policy.
        pub fn to_bytes(&self, limits: Limits) -> Result<Vec<u8>> {
            self.0.to_bytes(limits)
        }

        /// Decode and validate a patch under the required caller policy.
        pub fn from_bytes(bytes: &[u8], limits: Limits) -> Result<Self> {
            Ok(Self(record::RemovalPatch::from_bytes(bytes, limits)?))
        }

        pub(crate) fn into_record(self) -> record::RemovalPatch {
            self.0
        }

        pub(crate) fn from_record(patch: record::RemovalPatch) -> Self {
            Self(patch)
        }
    }

    /// An owned, bounded plain main-story XML tree.
    ///
    /// The tree is a raw codec value rather than a package owner.  Its
    /// mutation methods use a private one-operation staging value and replace
    /// the tree only after the exact readback proof succeeds.
    #[derive(Debug, Clone)]
    pub struct Tree {
        snapshot: record::Snapshot,
    }

    impl Tree {
        pub(crate) fn from_record(snapshot: record::Snapshot) -> Self {
            Self { snapshot }
        }

        /// Admit owned XML under the required finite policy before scanning it.
        ///
        /// Taking ownership keeps admission bounded without an implicit copy;
        /// callers that start with borrowed bytes choose whether and when to
        /// allocate before calling this constructor.
        pub fn from_xml(xml: Vec<u8>, limits: Limits) -> Result<Self> {
            Ok(Self {
                snapshot: record::Snapshot::from_raw_xml(xml, limits)?,
            })
        }

        /// Borrow the exact XML bytes currently held by the tree.
        #[must_use]
        pub fn xml_bytes(&self) -> &[u8] {
            self.snapshot.xml_bytes()
        }

        /// Return the number of direct plain paragraphs.
        #[must_use]
        pub fn paragraph_count(&self) -> usize {
            self.snapshot.paragraph_count()
        }

        /// Return the raw source identity used by exact patch freshness.
        #[must_use]
        pub const fn source_version(&self) -> SourceVersion {
            self.snapshot.source_version()
        }

        /// Copy one source-order paragraph before one insertion slot.
        ///
        /// The tree changes only after the copy and its exact source-fragment
        /// readback proof succeed.  Dropping the returned patch does not undo
        /// the tree; use [`Patch::inverse`] and [`Self::apply_copy`] when an
        /// explicit reversible operation is needed.
        pub fn copy_plain_paragraph(
            &mut self,
            source: Position,
            before: Position,
        ) -> Result<Patch> {
            let mut edit = self.snapshot.edit();
            edit.copy_plain_paragraph(source, before)?;
            let commit = edit.commit();
            self.snapshot = commit.projected().clone();
            Ok(Patch::from_record(commit.patch().clone()))
        }

        /// Remove one source-order paragraph.
        pub fn remove_plain_paragraph(&mut self, position: Position) -> Result<RemovalPatch> {
            let mut edit = self.snapshot.removal_edit();
            edit.remove_plain_paragraph(position)?;
            let commit = edit.commit();
            self.snapshot = commit.projected().clone();
            Ok(RemovalPatch::from_record(commit.patch().clone()))
        }

        /// Apply a copy patch to this tree after exact content and freshness
        /// validation.
        pub fn apply_copy(&mut self, patch: &Patch) -> Result<()> {
            self.snapshot = patch.0.apply_raw(&self.snapshot)?;
            Ok(())
        }

        /// Apply a removal patch to this tree after exact content and
        /// freshness validation.
        pub fn apply_removal(&mut self, patch: &RemovalPatch) -> Result<()> {
            self.snapshot = patch.0.apply_raw(&self.snapshot)?;
            Ok(())
        }

        /// Copy the current XML under the required output policy.
        pub fn to_bytes(&self, limits: Limits) -> Result<Vec<u8>> {
            self.snapshot.to_bytes(limits)
        }

        /// Stream the current XML under the required output policy.
        pub fn write_to(&self, mut writer: impl Write, limits: Limits) -> Result<()> {
            self.snapshot.write_to(&mut writer, limits)
        }
    }
}
