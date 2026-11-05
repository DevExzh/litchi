//! Handle for Rich Text Format (`.rtf`) inputs.
//!
//! The handle retains the native immutable `Package`. Existing family views
//! borrow its advanced read-only model; unified editing can use the same owner.

pub(crate) mod children;
mod nav;
mod node;
mod parts;
mod spec;
mod stories;
pub(crate) mod views;
pub(crate) mod write;

pub(crate) use nav::Nav;
pub use node::Node;
pub(crate) use node::Props;
pub use spec::Kind;
pub(crate) use spec::{ADDABLE, REMOVABLE, SETTABLE, SPEC, WRITE_GAPS};
pub(crate) use spec::{alignment_name, underline_name};

use std::fmt;
use std::path::Path;
use std::sync::Arc;

use litchi_core::Metadata;
use litchi_core::metadata::Timestamp as CoreTimestamp;
use litchi_rtf::Package;
use litchi_rtf::leniency::ToleranceReport;
use litchi_rtf::raw::Tree;

use crate::transaction::SemanticLimits;
use crate::{IntoProps, OpenOptions, Result, TargetArg};

/// An opened RTF document. Cheap to clone (one `Arc` bump); `Send + Sync`.
#[derive(Clone)]
pub struct Rtf(Arc<RtfState>);

#[derive(Clone)]
struct RtfState {
    origin: Option<Arc<Path>>,
    /// The native immutable owner, including retained open policies.
    package: Package,
}

const fn assert_handle<T: Send + Sync + Clone>() {}
const _: () = {
    assert_handle::<Rtf>();
    assert!(size_of::<Rtf>() == size_of::<usize>());
};

impl Rtf {
    pub(crate) fn with_origin(mut self, origin: Option<Arc<Path>>) -> Self {
        Arc::make_mut(&mut self.0).origin = origin;
        self
    }

    pub(crate) fn origin(&self) -> Option<&Arc<Path>> {
        self.0.origin.as_ref()
    }

    /// Adopt the immutable native endpoint produced by one checked edit.
    ///
    /// The package carries the native open options and source retention policy
    /// through the edit.  The facade sidecar is copied separately so an edit
    /// keeps saving to the same origin as its input.
    pub(crate) fn with_package(&self, package: Package) -> Self {
        Self(Arc::new(RtfState {
            origin: self.0.origin.clone(),
            package,
        }))
    }

    /// Build the bounded validation inventory from the opened package.
    ///
    /// The package retains the caller's native [`litchi_rtf::OpenOptions`],
    /// so this owner route never replaces its limits with a fresh default.
    pub(crate) fn validation_report(&self) -> litchi_rtf::ValidationReport {
        self.0.package.validation_report()
    }

    /// Whether two handles retain the same immutable native snapshot.
    pub(crate) fn same_snapshot(&self, other: &Self) -> bool {
        self.0.package.same_snapshot(&other.0.package)
    }

    /// Whether two handles serialize to the same exact RTF artifact.
    pub(crate) fn same_content(&self, other: &Self) -> Result<bool> {
        Ok(
            self.same_snapshot(other)
                || self.0.package.to_bytes()? == other.0.package.to_bytes()?,
        )
    }

    /// RTF packages retain their immutable source and native edit options.
    /// There is no external source revision to fence at this layer.
    pub(crate) fn check_source(&self) -> Result<()> {
        Ok(())
    }

    /// Admit a native RTF edit before a unified transaction starts.
    pub(crate) fn admit_edit(&self) -> Result<()> {
        self.check_source()
    }

    /// Set paragraph properties using the default unified limits.
    pub(crate) fn set<'p, 'v>(
        &self,
        target: TargetArg<'p>,
        props: impl IntoProps<'v>,
    ) -> Result<write::Commit> {
        let limits = SemanticLimits::default();
        self.set_with_limits(target, props, &limits)
    }

    /// Set paragraph properties through one bounded native RTF transaction.
    pub(crate) fn set_with_limits<'p, 'v>(
        &self,
        target: TargetArg<'p>,
        props: impl IntoProps<'v>,
        limits: &SemanticLimits,
    ) -> Result<write::Commit> {
        write::set_with_limits(self, target, props, limits)
    }

    /// Add one paragraph directly under the main body through a bounded native
    /// lifecycle edit.
    pub(crate) fn add<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::NewNode<'n>,
    ) -> Result<(write::Commit, crate::path::Path<'static>)> {
        self.add_with_limits(parent, node, &SemanticLimits::default())
    }

    /// Add one paragraph with the caller's transaction limits.
    pub(crate) fn add_with_limits<'p, 'n>(
        &self,
        parent: crate::path::PathArg<'p>,
        node: crate::NewNode<'n>,
        limits: &SemanticLimits,
    ) -> Result<(write::Commit, crate::path::Path<'static>)> {
        write::add_with_limits(self, parent, node, limits)
    }

    /// Remove one paragraph or every paragraph selected by a scoped query.
    pub(crate) fn remove<'p>(&self, target: TargetArg<'p>) -> Result<write::Commit> {
        let limits = SemanticLimits::default();
        self.remove_with_limits(target, &limits)
    }

    /// Remove selected body paragraphs through one bounded native transaction.
    pub(crate) fn remove_with_limits<'p>(
        &self,
        target: TargetArg<'p>,
        limits: &SemanticLimits,
    ) -> Result<write::Commit> {
        write::remove_with_limits(self, target, limits)
    }

    /// Move one ordinary main-story paragraph to a final body position.
    pub(crate) fn move_to<'p, 'q>(
        &self,
        path: crate::path::PathArg<'p>,
        placement: crate::Placement<'q>,
    ) -> Result<write::Commit> {
        let limits = SemanticLimits::default();
        self.move_to_with_limits(path, placement, &limits)
    }

    /// Move one paragraph with the caller's target, staging and snapshot limits.
    pub(crate) fn move_to_with_limits<'p, 'q>(
        &self,
        path: crate::path::PathArg<'p>,
        placement: crate::Placement<'q>,
        limits: &SemanticLimits,
    ) -> Result<write::Commit> {
        write::move_to_with_limits(self, path, placement, limits)
    }

    /// Exchange two ordinary main-story paragraphs through one bounded native
    /// lifecycle edit.
    pub(crate) fn swap<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
    ) -> Result<write::Commit> {
        self.swap_with_limits(left, right, &SemanticLimits::default())
    }

    /// Exchange two paragraphs with the caller's target, staging and snapshot
    /// limits.
    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        left: crate::path::PathArg<'p>,
        right: crate::path::PathArg<'q>,
        limits: &SemanticLimits,
    ) -> Result<write::Commit> {
        write::swap_with_limits(self, left, right, limits)
    }

    pub(crate) fn write_to(&self, sink: impl std::io::Write) -> Result<()> {
        self.0
            .package
            .write_to(sink)
            .map_err(|error| crate::Error::from(error).with_format(litchi_core::FileFormat::Rtf))
    }

    pub(crate) fn save_to(&self, path: &Path, options: &crate::SaveOptions) -> Result<u64> {
        let overwrite = match options.overwrite {
            crate::Overwrite::Replace => litchi_rtf::save::Overwrite::Replace,
            crate::Overwrite::Refuse => litchi_rtf::save::Overwrite::Refuse,
        };
        let policy = litchi_rtf::save::Options::new()
            .durability(options.durability)
            .overwrite(overwrite);
        self.0
            .package
            .save_with_options(path, policy)
            .map_err(|error| crate::Error::from(error).with_format(litchi_core::FileFormat::Rtf))
    }

    /// Adopt a fresh native package without a detector or serialization round trip.
    pub(crate) fn create(options: &crate::CreateOptions) -> Result<Self> {
        use crate::{Code, Conformance, Error, Flavor};
        let reason = if options.flavor_value() != Flavor::Document {
            Some("RTF template creation is not supported")
        } else if options.conformance_value() == Conformance::Strict {
            Some("RTF does not support OOXML Strict conformance")
        } else if options.metadata_value().is_some() {
            Some("initial RTF metadata authoring is not integrated")
        } else {
            None
        };
        if let Some(reason) = reason {
            return Err(Error::new(Code::UnsupportedCapability, reason)
                .with_format(litchi_core::FileFormat::Rtf)
                .with_suggestion("use default CreateOptions for an ordinary empty RTF document"));
        }
        let package = Package::new()?;
        Ok(Self(Arc::new(RtfState {
            package,
            origin: None,
        })))
    }

    /// Parses the RTF bytes detection retained (`DetectedFormat::Rtf`): plain,
    /// literal code-page, and compressed (LZFu/MELA) transports alike.
    ///
    /// Every facade route (path, bytes, reader) reaches RTF through these
    /// detected bytes, so this is the only constructor. Native source limits
    /// and repair policy apply in addition to the detection input ceiling.
    pub(crate) fn from_bytes(bytes: Vec<u8>, options: &OpenOptions) -> Result<Self> {
        let package =
            Package::from_owned_bytes_with_options(bytes, options.rtf_options.unwrap_or_default())?;
        Ok(Self(Arc::new(RtfState {
            package,
            origin: None,
        })))
    }

    /// The document's plain text.
    ///
    /// # Errors
    ///
    /// None today; the text was decoded at open. The `Result` keeps the
    /// read surface uniform across formats.
    pub fn text(&self) -> Result<String> {
        Ok(self.0.package.text().to_owned())
    }

    /// The number of body paragraphs.
    ///
    /// # Errors
    ///
    /// None today; see [`Self::text`].
    pub fn paragraph_count(&self) -> Result<usize> {
        Ok(self.0.package.paragraph_count())
    }

    /// Document properties projected from the `\info` group and the
    /// `\generator` destination. Always `Some` for RTF.
    ///
    /// # Errors
    ///
    /// None today; see [`Self::text`].
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        Ok(Some(metadata(self.0.package.raw())))
    }

    /// Everything the reader repaired while parsing, which happens entirely
    /// at open.
    ///
    /// # Errors
    ///
    /// None today; see [`Self::text`].
    pub fn tolerance_report(&self) -> Result<ToleranceReport> {
        Ok(self.0.package.tolerance_report().clone())
    }

    /// The retained leaf owner, for the crate's family views.
    pub(crate) fn document(&self) -> &Tree<'static> {
        self.0.package.raw()
    }
}

impl fmt::Debug for Rtf {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Rtf").finish_non_exhaustive()
    }
}

fn timestamp_to_naive(
    timestamp: Option<litchi_rtf::metadata::Timestamp>,
) -> Option<chrono::NaiveDateTime> {
    let timestamp = timestamp?;
    let year = timestamp.year?;
    let month = u32::try_from(timestamp.month?).ok()?;
    let day = u32::try_from(timestamp.day?).ok()?;
    let hour = u32::try_from(timestamp.hour.unwrap_or(0)).ok()?;
    let minute = u32::try_from(timestamp.minute.unwrap_or(0)).ok()?;
    let second = u32::try_from(timestamp.second.unwrap_or(0)).ok()?;
    chrono::NaiveDate::from_ymd_opt(year, month, day)?.and_hms_opt(hour, minute, second)
}

/// The facade's RTF metadata projection (`litchi::document::doc`), moved
/// verbatim.
fn metadata(document: &Tree<'_>) -> Metadata {
    let info = document.info();
    let text = |value: Option<&str>| value.map(str::to_owned);
    Metadata {
        title: text(info.title.as_deref()),
        subject: text(info.subject.as_deref()),
        author: text(info.author.as_deref()),
        keywords: info
            .keywords
            .as_deref()
            .map(litchi_core::metadata::keywords::split)
            .unwrap_or_default(),
        description: text(info.document_comment.as_deref().or(info.comment.as_deref())),
        identifier: info.id.map(|value| value.to_string()),
        language: None,
        template: None,
        last_modified_by: text(info.operator.as_deref()),
        revision: info.version.map(|value| value.to_string()),
        created: timestamp_to_naive(info.creation_timestamp).map(CoreTimestamp::Unzoned),
        modified: timestamp_to_naive(info.revision_timestamp).map(CoreTimestamp::Unzoned),
        page_count: info.pages,
        word_count: info.words,
        character_count: info.characters,
        character_count_with_spaces: info.characters_with_spaces,
        editing_time_minutes: info.editing_time,
        application: document
            .generator()
            .map(|generator| generator.value.to_string()),
        category: text(info.category.as_deref()),
        company: text(info.company.as_deref()),
        manager: text(info.manager.as_deref()),
        content_status: None,
        content_type: None,
        version: info.revision.map(|value| value.to_string()),
        last_printed: timestamp_to_naive(info.print_timestamp).map(CoreTimestamp::Unzoned),
        last_backup: timestamp_to_naive(info.backup_timestamp).map(CoreTimestamp::Unzoned),
        hyperlink_base: text(info.hyperlink_base.as_deref()),
        security: None,
        codepage: None,
    }
}

#[cfg(test)]
mod tests;
