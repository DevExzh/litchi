//! Handle for Rich Text Format (`.rtf`) inputs.
//!
//! The handle retains the native immutable `Package`. Existing family views
//! borrow its advanced read-only model; unified editing can use the same owner.

use std::fmt;
use std::sync::Arc;

use litchi_core::Metadata;
use litchi_rtf::leniency::ToleranceReport;
use litchi_rtf::{Package, RtfDocument};

use crate::{OpenOptions, Result};

/// An opened RTF document. Cheap to clone (one `Arc` bump); `Send + Sync`.
#[derive(Clone)]
pub struct Rtf(Arc<RtfState>);

struct RtfState {
    /// The native immutable owner, including retained open policies.
    package: Package,
}

const fn assert_handle<T: Send + Sync + Clone>() {}
const _: () = {
    assert_handle::<Rtf>();
    assert!(size_of::<Rtf>() == size_of::<usize>());
};

impl Rtf {
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
        Ok(Self(Arc::new(RtfState { package })))
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
        Ok(Self(Arc::new(RtfState { package })))
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
        Ok(self.0.package.raw().paragraph_count())
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
    pub(crate) fn document(&self) -> &RtfDocument<'static> {
        self.0.package.raw()
    }
}

impl fmt::Debug for Rtf {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Rtf").finish_non_exhaustive()
    }
}

fn timestamp_to_naive(
    timestamp: Option<litchi_rtf::RtfTimestamp>,
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
fn metadata(document: &RtfDocument<'_>) -> Metadata {
    let info = document.info();
    let text = |value: Option<&str>| value.map(str::to_owned);
    Metadata {
        title: text(info.title.as_deref()),
        subject: text(info.subject.as_deref()),
        author: text(info.author.as_deref()),
        keywords: text(info.keywords.as_deref()),
        description: text(info.document_comment.as_deref().or(info.comment.as_deref())),
        identifier: info.id.map(|value| value.to_string()),
        language: None,
        template: None,
        last_modified_by: text(info.operator.as_deref()),
        revision: info.version.map(|value| value.to_string()),
        created: None,
        created_local: timestamp_to_naive(info.creation_timestamp),
        modified: None,
        modified_local: timestamp_to_naive(info.revision_timestamp),
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
        last_printed_time: None,
        last_printed_local: timestamp_to_naive(info.print_timestamp),
        last_backup_local: timestamp_to_naive(info.backup_timestamp),
        hyperlink_base: text(info.hyperlink_base.as_deref()),
        security: None,
        codepage: None,
    }
}

#[cfg(test)]
mod tests;
