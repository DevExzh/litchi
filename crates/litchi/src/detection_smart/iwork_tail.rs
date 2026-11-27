//! The iWork arms of smart detection.
//!
//! These ran last among the ZIP candidates when detection lived in the
//! facade, and they still do: `litchi_ops::detect` invokes this tail only
//! after OOXML and ODF have declined a ZIP input, and never after a hard OPC
//! policy error or a failed cursor restore.

use std::io::{Read, Seek};

use litchi_core::detection::FileFormat;
use litchi_ops::detect::DetectedFormat;
use litchi_ops::detect::tail::ZipTail;

/// Classifies the ZIP candidates no other detector claimed as Pages, Keynote
/// or Numbers.
#[derive(Debug, Clone, Copy)]
pub(crate) struct IworkTail;

impl ZipTail for IworkTail {
    fn classify_bytes(&self, bytes: &[u8]) -> Option<FileFormat> {
        if let Ok(Some(format)) =
            litchi_iwa_detect::bytes(bytes, litchi_iwa_detect::Limits::default())
            && let Some(result) = iwork_format(format)
        {
            return Some(result);
        }
        None
    }

    fn classify_reader<R: Read + Seek>(&self, reader: &mut R) -> Option<FileFormat> {
        if let Ok(Some(format)) =
            litchi_iwa_detect::reader(reader, litchi_iwa_detect::Limits::default())
            && let Some(format) = iwork_format(format)
        {
            return Some(format);
        }
        None
    }

    fn detect_owned(&self, bytes: Vec<u8>) -> Option<DetectedFormat> {
        if let Ok(Some(format)) =
            litchi_iwa_detect::bytes(&bytes, litchi_iwa_detect::Limits::default())
        {
            #[allow(
                unreachable_patterns,
                reason = "match arms are feature-gated; some are unreachable depending on the enabled features"
            )]
            let detected = match format {
                #[cfg(feature = "pages")]
                litchi_iwa_detect::Format::Pages => DetectedFormat::Pages(bytes),
                #[cfg(feature = "keynote")]
                litchi_iwa_detect::Format::Keynote => DetectedFormat::Keynote(bytes),
                #[cfg(feature = "numbers")]
                litchi_iwa_detect::Format::Numbers => DetectedFormat::Numbers(bytes),
                _ => return None,
            };
            return Some(detected);
        }
        None
    }
}

#[allow(
    unreachable_patterns,
    reason = "match arms are feature-gated; some are unreachable depending on the enabled features"
)]
fn iwork_format(format: litchi_iwa_detect::Format) -> Option<FileFormat> {
    match format {
        #[cfg(feature = "pages")]
        litchi_iwa_detect::Format::Pages => Some(FileFormat::Pages),
        #[cfg(feature = "keynote")]
        litchi_iwa_detect::Format::Keynote => Some(FileFormat::Keynote),
        #[cfg(feature = "numbers")]
        litchi_iwa_detect::Format::Numbers => Some(FileFormat::Numbers),
        _ => None,
    }
}
