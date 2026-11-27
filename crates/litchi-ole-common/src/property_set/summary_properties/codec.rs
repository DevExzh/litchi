//! Serialization of [`SummaryProperties`] into property-set streams.

use super::super::binding::Binding;
use super::super::model::{Section, Stream, try_clone_property_set};
use super::model::SummaryProperties;
use litchi_cfb::{OleError, OleWriter};

/// The `PropertySetStream.SystemIdentifier` the writers record: the
/// [MS-OSHARED] 2.3.3.1.1 `PropertySetSystemIdentifier` with `OSType` 2 (the
/// only value the structure allows) and operating-system version 10.0. The
/// value is fixed, so the bytes never depend on the machine that writes them.
pub const SYSTEM_IDENTIFIER: u32 = 0x0002_000A;

/// The serialized `\005SummaryInformation` and `\005DocumentSummaryInformation`
/// streams of a [`SummaryProperties`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Streams {
    summary_information: Vec<u8>,
    document_summary_information: Vec<u8>,
}

impl Streams {
    /// The `\005SummaryInformation` stream.
    #[must_use]
    pub fn summary_information(&self) -> &[u8] {
        &self.summary_information
    }

    /// The `\005DocumentSummaryInformation` stream, with the user-defined
    /// section second when there is one.
    #[must_use]
    pub fn document_summary_information(&self) -> &[u8] {
        &self.document_summary_information
    }

    /// Consumes the streams into the `\005SummaryInformation` and
    /// `\005DocumentSummaryInformation` bytes, in that order.
    #[must_use]
    pub fn into_parts(self) -> (Vec<u8>, Vec<u8>) {
        (self.summary_information, self.document_summary_information)
    }

    /// Creates both root streams in a compound file being written, moving the
    /// bytes into the writer.
    ///
    /// # Errors
    ///
    /// Returns an error if the writer refuses a stream, for example because
    /// one of that name already exists.
    pub fn write_into(self, writer: &mut OleWriter) -> Result<(), OleError> {
        let summary = Binding::SummaryInformation.name();
        writer.create_stream_owned(&[summary.as_str()], self.summary_information)?;
        let document_summary = Binding::DocumentSummaryInformation.name();
        writer.create_stream_owned(
            &[document_summary.as_str()],
            self.document_summary_information,
        )
    }
}

impl SummaryProperties {
    /// Serializes the two property-set streams, or returns `None` when
    /// [`Self::is_empty`]: a writer then writes no property stream.
    ///
    /// Both streams are version 0 with a null class identifier and the fixed
    /// [`SYSTEM_IDENTIFIER`], as [MS-OSHARED] 2.3.3.2 requires;
    /// `DocumentSummaryInformation` is always written with its code page, and
    /// carries the user-defined section second when one exists.
    ///
    /// # Errors
    ///
    /// Returns an error if a section cannot be serialized (a value too large
    /// for the wire format, or an allocation failure).
    pub fn to_streams(&self) -> Result<Option<Streams>, OleError> {
        if self.is_empty() {
            return Ok(None);
        }
        let (summary, document_summary, user_defined) = self.sections();
        Ok(Some(Streams {
            summary_information: serialize(summary, None)?,
            document_summary_information: serialize(document_summary, user_defined)?,
        }))
    }
}

fn serialize(first: &Section, second: Option<&Section>) -> Result<Vec<u8>, OleError> {
    let mut stream = Stream::new(try_clone_property_set(first)?);
    stream.system_identifier = SYSTEM_IDENTIFIER;
    if let Some(section) = second {
        stream.add_section(try_clone_property_set(section)?)?;
    }
    stream.to_bytes()
}
