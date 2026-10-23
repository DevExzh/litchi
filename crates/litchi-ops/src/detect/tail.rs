//! The lowest-precedence ZIP candidate slot.
//!
//! iWork is the last ZIP candidate every detector tries: OOXML first (unless
//! the input is a validated ordinary ODF package), then ODF, then iWork.
//! The openers classify iWork themselves since record 0296; the facade's
//! detection functions still reach their iWork arm through [`ZipTail`]. A
//! detector invokes the tail at exactly the point where the iWork arm
//! belongs: after every other ZIP candidate has declined the input, and
//! never after a hard OPC policy error or a failed cursor restore.
//!
//! The `_with_tail` detectors of [`crate::detect`] and
//! [`open_with_tail`](crate::open_with_tail) /
//! [`open_bytes_with_tail`](crate::open_bytes_with_tail) take a tail. The
//! detectors and openers without one pass [`NoZipTail`], so they classify
//! every format except iWork, with unchanged precedence.
//!
//! The detectors consult a tail only in a build with at least one of the
//! `pages`, `keynote` and `numbers` features, which shape [`DetectedFormat`]
//! and carry no dependency. The openers consult it in every build, and
//! refuse a ZIP it names with
//! [`Error::iwork_unsupported`](crate::Error::iwork_unsupported).

use std::io::{Read, Seek};

use litchi_core::detection::FileFormat;

use super::DetectedFormat;

/// A classifier for a ZIP candidate that every other detector declined.
pub trait ZipTail {
    /// Classify borrowed bytes.
    fn classify_bytes(&self, bytes: &[u8]) -> Option<FileFormat>;

    /// Classify a reader. The detector has already rewound it to the start
    /// and restores the caller's position afterwards.
    fn classify_reader<R: Read + Seek>(&self, reader: &mut R) -> Option<FileFormat>;

    /// Classify moved bytes, retaining them in the result.
    fn detect_owned(&self, bytes: Vec<u8>) -> Option<DetectedFormat>;
}

/// The empty tail: nothing after ODF is tried.
#[derive(Debug, Clone, Copy, Default)]
pub struct NoZipTail;

impl ZipTail for NoZipTail {
    #[inline]
    fn classify_bytes(&self, _bytes: &[u8]) -> Option<FileFormat> {
        None
    }

    #[inline]
    fn classify_reader<R: Read + Seek>(&self, _reader: &mut R) -> Option<FileFormat> {
        None
    }

    #[inline]
    fn detect_owned(&self, bytes: Vec<u8>) -> Option<DetectedFormat> {
        drop(bytes);
        None
    }
}

/// The tail runs only after every other candidate declined the input. The
/// detectors call it only in builds with an iWork feature.
#[cfg(all(test, any(feature = "pages", feature = "keynote", feature = "numbers")))]
mod tests {
    use std::cell::Cell;
    use std::io::{Cursor, Read, Seek, Write};

    use litchi_core::detection::FileFormat;

    use super::super::SourceLimits;
    use super::super::functions::{
        detect_file_format_from_bytes_with_tail, detect_format_from_reader_with_tail,
    };
    use super::{DetectedFormat, ZipTail};

    /// Claims every input it sees as Pages and counts the calls.
    #[derive(Default)]
    struct RecordingTail {
        calls: Cell<usize>,
    }

    impl ZipTail for RecordingTail {
        fn classify_bytes(&self, _bytes: &[u8]) -> Option<FileFormat> {
            self.calls.set(self.calls.get() + 1);
            Some(FileFormat::Pages)
        }

        fn classify_reader<R: Read + Seek>(&self, reader: &mut R) -> Option<FileFormat> {
            assert_eq!(
                reader.stream_position().unwrap(),
                0,
                "tail must see a rewound reader"
            );
            self.calls.set(self.calls.get() + 1);
            Some(FileFormat::Pages)
        }

        fn detect_owned(&self, bytes: Vec<u8>) -> Option<DetectedFormat> {
            self.calls.set(self.calls.get() + 1);
            drop(bytes);
            None
        }
    }

    fn zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut output = Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for (name, data) in entries {
            writer.start_file(*name, options).unwrap();
            writer.write_all(data).unwrap();
        }
        writer.finish().unwrap();
        output.into_inner()
    }

    /// The classification and tail-call count of the byte and reader paths.
    fn classify(bytes: &[u8]) -> [(Option<FileFormat>, usize); 2] {
        let tail = RecordingTail::default();
        let limits = SourceLimits::default();
        let from_bytes = detect_file_format_from_bytes_with_tail(bytes, &limits, &tail);
        let byte_calls = tail.calls.replace(0);
        let from_reader =
            detect_format_from_reader_with_tail(&mut Cursor::new(bytes), &limits, &tail);
        [(from_bytes, byte_calls), (from_reader, tail.calls.get())]
    }

    #[test]
    fn an_unclaimed_zip_reaches_the_tail_once() {
        let bytes = zip(&[("Index/Document.iwa", b"not an office part")]);
        for result in classify(&bytes) {
            assert_eq!(result, (Some(FileFormat::Pages), 1));
        }
    }

    #[test]
    fn non_zip_inputs_never_reach_the_tail() {
        let mut ole2 = litchi_core::detection::utils::OLE2_SIGNATURE.to_vec();
        ole2.resize(512, 0);
        for bytes in [ole2.as_slice(), br"{\rtf1 plain}".as_slice(), b"plain text"] {
            for (_, calls) in classify(bytes) {
                assert_eq!(calls, 0);
            }
        }
    }

    #[cfg(feature = "docx")]
    fn docx() -> Vec<u8> {
        zip(&[
            (
                "[Content_Types].xml",
                br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
            ),
            (
                "word/document.xml",
                br#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#,
            ),
        ])
    }

    #[cfg(feature = "docx")]
    #[test]
    fn an_ooxml_package_never_reaches_the_tail() {
        for result in classify(&docx()) {
            assert_eq!(result, (Some(FileFormat::Docx), 0));
        }

        let tail = RecordingTail::default();
        let detected = super::super::detected::detect_format_smart_with_tail(
            docx(),
            &SourceLimits::default(),
            &tail,
        );
        assert!(matches!(detected, Some(DetectedFormat::Docx(_))));
        assert_eq!(tail.calls.get(), 0);
    }

    #[cfg(feature = "docx")]
    #[test]
    fn a_hard_opc_policy_error_does_not_fall_through_to_the_tail() {
        let bytes = zip(&[("Index/Document.iwa", b"not an office part")]);
        let limits = crate::opc::SourceLimits::builder()
            .max_input_bytes(8)
            .unwrap()
            .build()
            .unwrap();
        let tail = RecordingTail::default();
        let limits = SourceLimits::default().ooxml_source_limits(limits);
        let detected = super::super::detected::detect_ooxml_smart(bytes, &limits, &tail);
        assert!(detected.is_none());
        assert_eq!(tail.calls.get(), 0);
    }

    #[cfg(feature = "odt")]
    #[test]
    fn an_odf_package_never_reaches_the_tail() {
        let mut writer = litchi_odf_common::core::PackageWriter::new(
            litchi_odf_common::core::StreamingArchiveLimits::default(),
        );
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_TEXT)
            .unwrap();
        writer
            .add_file(
                litchi_odf_common::constants::ODF_CONTENT,
                br#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:text><t:p>odf</t:p></o:text></o:body></o:document-content>"#,
            )
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        for result in classify(&bytes) {
            assert_eq!(result, (Some(FileFormat::Odt), 0));
        }
    }
}
