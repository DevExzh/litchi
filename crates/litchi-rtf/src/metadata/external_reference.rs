//! Inert external document names from RTF document-format destinations.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;
use std::ops::Range;

/// Maximum decoded UTF-8 byte length of one external document name.
pub const MAX_DOCUMENT_EXTERNAL_REFERENCE_BYTES: usize = 65_536;
/// Maximum aggregate decoded UTF-8 bytes retained for document references.
pub const MAX_DOCUMENT_EXTERNAL_REFERENCE_TOTAL_BYTES: usize =
    2 * MAX_DOCUMENT_EXTERNAL_REFERENCE_BYTES;

/// Passive names from the RTF `nextfile` and `template` destinations.
///
/// The names are metadata only. Parsing never opens, resolves, or invokes them.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentExternalReferences<'a> {
    /// Name of a file to print or index next (`nextfile`).
    pub next_file: Option<Cow<'a, str>>,
    /// Name of the document's related template (`template`).
    pub template: Option<Cow<'a, str>>,
}

/// Exact source ownership for the two passive document-reference groups.
///
/// The parser fills this only while token spans still address the original
/// uncompressed ASCII transport.  It is deliberately crate-private: callers
/// receive the checked ranges through the redaction inventory rather than
/// being able to use offsets without the source snapshot that owns them.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct DocumentExternalReferenceSpans {
    pub(crate) next_file: Option<Range<usize>>,
    pub(crate) template: Option<Range<usize>>,
}

impl DocumentExternalReferences<'_> {
    /// Validate name and aggregate resource bounds.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        let mut total = 0usize;
        for (kind, raw_value) in [
            ("next-file", self.next_file.as_deref()),
            ("template", self.template.as_deref()),
        ] {
            let Some(value) = raw_value else { continue };
            if value.trim().is_empty() {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {kind} name cannot be empty"
                )));
            }
            if value.len() > MAX_DOCUMENT_EXTERNAL_REFERENCE_BYTES {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {kind} name exceeds the safety limit"
                )));
            }
            if value.contains(['\0', '\r', '\n']) {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {kind} name contains a forbidden control character"
                )));
            }
            total = total.checked_add(value.len()).ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF external-reference aggregate size overflow".to_string(),
                )
            })?;
        }
        if total > MAX_DOCUMENT_EXTERNAL_REFERENCE_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF external-reference aggregate text exceeds the safety limit".to_string(),
            ));
        }
        Ok(())
    }

    /// Return whether neither external-reference destination is present.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.next_file.is_none() && self.template.is_none()
    }

    pub(crate) fn into_owned(self) -> DocumentExternalReferences<'static> {
        DocumentExternalReferences {
            next_file: self.next_file.map(|value| Cow::Owned(value.into_owned())),
            template: self.template.map(|value| Cow::Owned(value.into_owned())),
        }
    }
}
