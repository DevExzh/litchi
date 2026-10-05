//! Handle for Rich Text Format (`.rtf`) inputs.
//!
//! RTF is tier 3 and read-only in M0a (record 0003 §2.3). The handle carries
//! the facade's eager `RtfDocument` over unchanged. That owner already keeps
//! the input it was parsed from and exposes `Document::edit()`, so M6 can add
//! the write handler without reshaping this state.

use std::fmt;
use std::sync::Arc;

use litchi_core::Metadata;
use litchi_rtf::RtfDocument;
use litchi_rtf::leniency::ToleranceReport;

use crate::{OpenOptions, Result};

/// An opened RTF document. Cheap to clone (one `Arc` bump); `Send + Sync`.
#[derive(Clone)]
pub struct Rtf(Arc<RtfState>);

struct RtfState {
    /// The eager reader the facade's document pipeline builds today.
    document: RtfDocument<'static>,
}

const fn assert_handle<T: Send + Sync + Clone>() {}
const _: () = {
    assert_handle::<Rtf>();
    assert!(size_of::<Rtf>() == size_of::<usize>());
};

impl Rtf {
    /// Parses the RTF bytes detection retained (`DetectedFormat::Rtf`): plain,
    /// literal code-page, and compressed (LZFu/MELA) transports alike.
    ///
    /// Every facade route (path, bytes, reader) reaches RTF through these
    /// detected bytes, so this is the only constructor. RTF has no open
    /// sub-options; the input ceiling was applied by detection.
    pub(crate) fn from_bytes(bytes: Vec<u8>, _options: &OpenOptions) -> Result<Self> {
        let document = RtfDocument::from_owned_bytes(bytes)?;
        Ok(Self(Arc::new(RtfState { document })))
    }

    /// The document's plain text.
    ///
    /// # Errors
    ///
    /// None today; the text was decoded at open. The `Result` keeps the
    /// read surface uniform across formats.
    pub fn text(&self) -> Result<String> {
        Ok(self.0.document.text())
    }

    /// The number of body paragraphs.
    ///
    /// # Errors
    ///
    /// None today; see [`Self::text`].
    pub fn paragraph_count(&self) -> Result<usize> {
        Ok(self.0.document.paragraph_count())
    }

    /// Document properties projected from the `\info` group and the
    /// `\generator` destination. Always `Some` for RTF.
    ///
    /// # Errors
    ///
    /// None today; see [`Self::text`].
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        Ok(Some(metadata(&self.0.document)))
    }

    /// Everything the reader repaired while parsing, which happens entirely
    /// at open.
    ///
    /// # Errors
    ///
    /// None today; see [`Self::text`].
    pub fn tolerance_report(&self) -> Result<ToleranceReport> {
        Ok(self.0.document.tolerance_report().clone())
    }

    /// The retained leaf owner, for the crate's family views.
    pub(crate) fn document(&self) -> &RtfDocument<'static> {
        &self.0.document
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
