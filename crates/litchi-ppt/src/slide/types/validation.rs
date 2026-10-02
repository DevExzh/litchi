//! Narrow binary invariants shared by slide codecs.

use crate::records::Record;

pub(super) const MAX_SHAPE_DEPTH: usize = 256;

pub(super) fn is_ppt10_tag_name(record: &Record) -> bool {
    const PPT10: [u16; 8] = [0x5F, 0x5F, 0x5F, 0x50, 0x50, 0x54, 0x31, 0x30];
    record.version == 0
        && record.instance == 0
        && record.data.len() == 16
        && record
            .data
            .as_chunks::<2>()
            .0
            .iter()
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
            .eq(PPT10)
}

/// Whether a failure converting one `OfficeArt` shape is confined to that
/// shape's content, so the shape can be dropped and its siblings kept
/// (finding F5.11).
///
/// Content checks surface as `Corrupted`/`InvalidFormat` or as an
/// `OfficeArt` property/geometry/shape/image validation failure. I/O,
/// container, encryption, allocation, and resource-limit failures are never
/// confined to one shape. Several legacy parsers report an exceeded resource
/// bound as `Corrupted("... limit")`; those stay fatal too.
pub(crate) fn is_isolatable_shape_error(error: &crate::package::Error) -> bool {
    use crate::package::Error;
    match error {
        Error::Corrupted(message) | Error::InvalidFormat(message) => !message.contains(" limit"),
        Error::OfficeArt(error) => matches!(
            error,
            litchi_odraw::Error::MalformedProperties { .. }
                | litchi_odraw::Error::MalformedGeometry { .. }
                | litchi_odraw::Error::MalformedShape { .. }
                | litchi_odraw::Error::MalformedImage { .. }
                | litchi_odraw::Error::InvalidBlipId { .. }
        ),
        Error::Io(_)
        | Error::Ole(_)
        | Error::Graph(_)
        | Error::StreamNotFound(_)
        | Error::ResourceLimit(_)
        | Error::AllocationFailed(_)
        | Error::PasswordRequired
        | Error::InvalidPassword
        | Error::UnsupportedEncryption(_)
        | Error::MalformedEncryptionHeader(_) => false,
    }
}
