use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, ReadAt,
    SourceVersion, TextOutputError, TextOutputOptions,
};
use litchi_opc::PackURI;
use litchi_pptx::{Error, Package, ReadLimits, SourceBackedPresentation};
use soapberry_zip::office::StreamingArchiveWriter;

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWINGML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn slide_xml(body: &str) -> Vec<u8> {
    format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:grpSp><p:grpSpPr/><p:sp>{body}</p:sp></p:grpSp></p:spTree></p:cSld></p:sld>"#
    )
    .into_bytes()
}

fn package_bytes(slides: &[Vec<u8>]) -> Vec<u8> {
    let mut authored = Package::new().expect("new PPTX package");
    {
        let presentation = authored.presentation_mut().expect("mutable presentation");
        for _ in slides {
            presentation.add_slide().expect("add slide");
        }
    }
    let authored_bytes = authored.to_bytes().expect("serialize authored package");
    let mut package = Package::from_bytes(&authored_bytes).expect("open authored package");
    package
        .edit_opc(|opc| {
            for (index, slide) in slides.iter().enumerate() {
                let uri = PackURI::new(format!("/ppt/slides/slide{}.xml", index + 1))
                    .expect("generated PPTX slide URI is valid");
                opc.get_part_mut(&uri)?.set_blob(slide.clone());
            }
            Ok(())
        })
        .expect("replace slide payloads");
    package.to_bytes().expect("serialize slide fixture")
}

fn zip_u16(bytes: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([bytes[offset], bytes[offset + 1]])
}

fn zip_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        bytes[offset],
        bytes[offset + 1],
        bytes[offset + 2],
        bytes[offset + 3],
    ])
}

fn put_zip_u32(bytes: &mut [u8], offset: usize, value: u32) {
    bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

fn replace_stored_zip_member(
    mut archive: Vec<u8>,
    member_name: &[u8],
    replacement: &[u8],
) -> Vec<u8> {
    const LOCAL_FILE_HEADER: u32 = 0x0403_4b50;
    const CENTRAL_FILE_HEADER: u32 = 0x0201_4b50;
    const DATA_DESCRIPTOR: u32 = 0x0807_4b50;

    let replacement_crc = soapberry_zip::crc32(replacement);
    let mut local_found = false;
    let mut offset = 0usize;
    while offset.saturating_add(30) <= archive.len() {
        if zip_u32(&archive, offset) != LOCAL_FILE_HEADER {
            offset += 1;
            continue;
        }
        let flags = zip_u16(&archive, offset + 6);
        let compression = zip_u16(&archive, offset + 8);
        let compressed_size = zip_u32(&archive, offset + 18) as usize;
        let uncompressed_size = zip_u32(&archive, offset + 22) as usize;
        let name_len = zip_u16(&archive, offset + 26) as usize;
        let extra_len = zip_u16(&archive, offset + 28) as usize;
        let name_start = offset + 30;
        let data_start = name_start
            .saturating_add(name_len)
            .saturating_add(extra_len);
        let data_end = data_start.saturating_add(compressed_size);
        if data_end > archive.len() || name_start.saturating_add(name_len) > archive.len() {
            break;
        }
        if &archive[name_start..name_start + name_len] == member_name {
            assert_eq!(compression, 0, "fixture member must use stored compression");
            assert_eq!(compressed_size, replacement.len());
            assert_eq!(uncompressed_size, replacement.len());
            archive[data_start..data_end].copy_from_slice(replacement);
            put_zip_u32(&mut archive, offset + 14, replacement_crc);
            if flags & 0x0008 != 0 {
                let descriptor_crc = if zip_u32(&archive, data_end) == DATA_DESCRIPTOR {
                    data_end + 4
                } else {
                    data_end
                };
                put_zip_u32(&mut archive, descriptor_crc, replacement_crc);
            }
            local_found = true;
            break;
        }
        offset = data_end;
    }
    assert!(local_found, "stored ZIP member not found in local records");

    let mut central_found = false;
    for offset in 0..archive.len().saturating_sub(46) {
        if zip_u32(&archive, offset) != CENTRAL_FILE_HEADER {
            continue;
        }
        let compressed_size = zip_u32(&archive, offset + 20) as usize;
        let uncompressed_size = zip_u32(&archive, offset + 24) as usize;
        let name_len = zip_u16(&archive, offset + 28) as usize;
        let extra_len = zip_u16(&archive, offset + 30) as usize;
        let comment_len = zip_u16(&archive, offset + 32) as usize;
        let name_start = offset + 46;
        let record_end = name_start
            .saturating_add(name_len)
            .saturating_add(extra_len)
            .saturating_add(comment_len);
        if record_end > archive.len() {
            continue;
        }
        if &archive[name_start..name_start + name_len] == member_name {
            assert_eq!(compressed_size, replacement.len());
            assert_eq!(uncompressed_size, replacement.len());
            put_zip_u32(&mut archive, offset + 16, replacement_crc);
            central_found = true;
            break;
        }
    }
    assert!(
        central_found,
        "stored ZIP member not found in central directory"
    );
    archive
}

fn malformed_later_slide_package_bytes() -> Vec<u8> {
    let first = slide_xml(r#"<a:t>first</a:t>"#);
    let valid_second = slide_xml(r#"<a:t>second</a:t>"#);
    let mut malformed_second = valid_second.clone();
    let closing_text = b"</a:t>";
    let closing_start = malformed_second
        .windows(closing_text.len())
        .position(|window| window == closing_text)
        .expect("valid second slide text closing tag");
    malformed_second[closing_start + closing_text.len() - 1] = b'x';

    let mut writer = StreamingArchiveWriter::new();
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#.to_string();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("write content types");
    writer
        .write_stored(
            "_rels/.rels",
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        )
        .expect("write package relationships");
    writer
        .write_stored(
            "ppt/presentation.xml",
            br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#,
        )
        .expect("write presentation");
    writer
        .write_stored(
            "ppt/_rels/presentation.xml.rels",
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>"#,
        )
        .expect("write presentation relationships");
    writer
        .write_stored("ppt/slides/slide1.xml", &first)
        .expect("write first slide");
    writer
        .write_stored("ppt/slides/slide2.xml", &valid_second)
        .expect("write valid second slide");
    let archive = writer
        .finish_to_bytes()
        .expect("finish stored PPTX archive");
    replace_stored_zip_member(archive, b"ppt/slides/slide2.xml", &malformed_second)
}

fn oversized_slide_declaration(mut bytes: Vec<u8>) -> Vec<u8> {
    let target = b"ppt/slides/slide1.xml";
    let mut offset = 0usize;
    while offset.saturating_add(46) <= bytes.len() {
        if &bytes[offset..offset + 4] == b"PK\x01\x02" {
            let name_len = u16::from_le_bytes([bytes[offset + 28], bytes[offset + 29]]) as usize;
            let name_start = offset + 46;
            let name_end = name_start.saturating_add(name_len);
            if name_end <= bytes.len() && &bytes[name_start..name_end] == target {
                let declared = (64_u64 * 1024 * 1024 + 1) as u32;
                bytes[offset + 24..offset + 28].copy_from_slice(&declared.to_le_bytes());
                return bytes;
            }
        }
        offset += 1;
    }
    panic!("slide central-directory entry not found");
}

fn parity_options(include_empty: bool) -> TextOutputOptions<'static> {
    TextOutputOptions::new("\n", "\n", u64::MAX, u64::MAX).with_empty_objects(include_empty)
}

fn text_fixture() -> Vec<u8> {
    package_bytes(&[
        slide_xml(
            r#"<a:p><a:r><a:t>One &amp; 世界</a:t></a:r><a:r><a:t><![CDATA[Two]]></a:t></a:r></a:p><a:grpSp><a:t>Three</a:t></a:grpSp>"#,
        ),
        slide_xml(r#"<a:p><a:r><a:t>Four</a:t></a:r></a:p>"#),
    ])
}

fn eager_text(
    bytes: &[u8],
    options: TextOutputOptions<'_>,
) -> (Vec<u8>, litchi_core::TextOutputReport) {
    let package = Package::from_bytes(bytes).expect("eager package");
    let presentation = package.presentation().expect("eager presentation");
    let mut output = Vec::new();
    let report = presentation
        .write_text_to(&mut output, options)
        .expect("eager semantic text");
    (output, report)
}

fn source_text(
    bytes: Vec<u8>,
    options: TextOutputOptions<'_>,
) -> (Vec<u8>, litchi_core::TextOutputReport) {
    let presentation = SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("source-backed presentation");
    let mut output = Vec::new();
    let report = presentation
        .write_text_to(&mut output, options)
        .expect("source semantic text");
    (output, report)
}

#[test]
fn eager_and_source_sinks_have_independent_legacy_text_parity() {
    let bytes = text_fixture();
    let expected = "One & 世界\nTwo\nThree\nFour".as_bytes().to_vec();

    let (eager, eager_report) = eager_text(&bytes, parity_options(false));
    let (source, source_report) = source_text(bytes, parity_options(false));

    assert_eq!(eager, expected);
    assert_eq!(source, expected);
    assert_eq!(eager_report.bytes_written(), expected.len() as u64);
    assert_eq!(eager_report.objects_written(), 2);
    assert_eq!(source_report, eager_report);
}

#[test]
fn empty_objects_and_custom_slide_separator_follow_options() {
    let bytes = package_bytes(&[slide_xml(r#"<a:t>A</a:t>"#), slide_xml("")]);

    let (excluded, excluded_report) = eager_text(
        &bytes,
        TextOutputOptions::new("\n", "|", 32, 4).with_empty_objects(false),
    );
    let (source_excluded, source_excluded_report) = source_text(
        bytes.clone(),
        TextOutputOptions::new("\n", "|", 32, 4).with_empty_objects(false),
    );
    assert_eq!(excluded, b"A");
    assert_eq!(source_excluded, excluded);
    assert_eq!(excluded_report.objects_written(), 1);
    assert_eq!(source_excluded_report, excluded_report);

    let (included, included_report) = eager_text(&bytes, TextOutputOptions::new("\n", "|", 32, 4));
    let (source_included, source_included_report) =
        source_text(bytes, TextOutputOptions::new("\n", "|", 32, 4));
    assert_eq!(included, b"A|");
    assert_eq!(source_included, included);
    assert_eq!(included_report.bytes_written(), 2);
    assert_eq!(included_report.objects_written(), 2);
    assert_eq!(source_included_report, included_report);
}

#[test]
fn custom_paragraph_separator_has_eager_and_source_parity() {
    let bytes = package_bytes(&[slide_xml(r#"<a:t>left</a:t><a:t>right</a:t>"#)]);
    let options = TextOutputOptions::new("<p>", "|", u64::MAX, u64::MAX).with_empty_objects(false);

    let (eager, eager_report) = eager_text(&bytes, options);
    let (source, source_report) = source_text(bytes, options);

    assert_eq!(eager, b"left<p>right");
    assert_eq!(source, eager);
    assert_eq!(eager_report, source_report);
}

#[test]
fn mce_fallback_is_processed_and_foreign_text_is_not_accepted() {
    let mce_slide = format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DRAWINGML}" xmlns:mc="{MCE}" xmlns:x="urn:future" mc:Ignorable="x"><p:cSld><p:spTree><mc:AlternateContent><mc:Choice Requires="x"><a:t>ignored</a:t></mc:Choice><mc:Fallback><a:t>fallback</a:t></mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>"#
    )
    .into_bytes();
    let (output, report) = eager_text(&package_bytes(&[mce_slide]), parity_options(true));
    assert_eq!(output, b"fallback");
    assert_eq!(report.objects_written(), 1);

    let foreign = package_bytes(&[slide_xml(
        r#"<x:t xmlns:x="urn:foreign">not DrawingML</x:t><a:t>valid</a:t>"#,
    )]);
    let package = Package::from_bytes(&foreign).expect("foreign-text package");
    let presentation = package.presentation().expect("foreign-text presentation");
    let mut output = Vec::new();
    assert!(matches!(
        presentation.write_text_to(&mut output, parity_options(true)),
        Err(TextOutputError::Document { .. })
    ));
    assert!(output.is_empty());
}

#[test]
fn malformed_later_slide_preserves_only_prior_slide_progress() {
    let bytes = malformed_later_slide_package_bytes();
    let package = Package::from_bytes(&bytes).expect("malformed slide package");
    let presentation = package.presentation().expect("malformed presentation");
    let mut output = Vec::new();
    let error = presentation
        .write_text_to(&mut output, parity_options(false))
        .expect_err("second slide must fail");

    assert_eq!(output, b"first");
    assert!(matches!(
        error,
        TextOutputError::Document {
            progress,
            ..
        } if progress.bytes_written() == 5 && progress.objects_written() == 1
    ));
}

#[test]
fn source_malformed_later_slide_preserves_only_prior_slide_progress() {
    let bytes = malformed_later_slide_package_bytes();
    let presentation = SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("malformed source-backed presentation");
    let mut output = Vec::new();
    let error = presentation
        .write_text_to(&mut output, parity_options(false))
        .expect_err("second source slide must fail");

    assert_eq!(output, b"first");
    assert!(matches!(
        error,
        TextOutputError::Document {
            progress,
            ..
        } if progress.bytes_written() == 5 && progress.objects_written() == 1
    ));
}

#[test]
fn output_and_object_limits_report_exact_progress() {
    let bytes = package_bytes(&[slide_xml("<a:t>one</a:t>"), slide_xml("<a:t>two</a:t>")]);

    let mut output = Vec::new();
    let package = Package::from_bytes(&bytes).expect("test package");
    let presentation = package.presentation().expect("test presentation");
    let error = presentation
        .write_text_to(
            &mut output,
            TextOutputOptions::new("\n", "\n", 3, 2).with_empty_objects(false),
        )
        .unwrap_err();
    match error {
        TextOutputError::Limit { progress, .. } => {
            assert_eq!(progress.bytes_written(), 3);
            assert_eq!(progress.objects_written(), 1);
        },
        other => panic!("unexpected error: {other:?}"),
    }
    assert_eq!(output, b"one");

    let mut output = Vec::new();
    let package = Package::from_bytes(&bytes).expect("test package");
    let presentation = package.presentation().expect("test presentation");
    let error = presentation
        .write_text_to(
            &mut output,
            TextOutputOptions::new("\n", "\n", u64::MAX, 1).with_empty_objects(false),
        )
        .unwrap_err();
    match error {
        TextOutputError::Limit { progress, .. } => {
            assert_eq!(progress.bytes_written(), 3);
            assert_eq!(progress.objects_written(), 1);
        },
        other => panic!("unexpected error: {other:?}"),
    }
    assert_eq!(output, b"one");
}

#[test]
fn source_output_and_object_limits_report_exact_progress() {
    let bytes = package_bytes(&[slide_xml("<a:t>one</a:t>"), slide_xml("<a:t>two</a:t>")]);
    let presentation = SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("source limit presentation");

    let mut output = Vec::new();
    let error = presentation
        .write_text_to(
            &mut output,
            TextOutputOptions::new("\n", "\n", 3, 2).with_empty_objects(false),
        )
        .unwrap_err();
    assert!(matches!(
        error,
        TextOutputError::Limit { progress, .. }
            if progress.bytes_written() == 3 && progress.objects_written() == 1
    ));
    assert_eq!(output, b"one");

    let mut output = Vec::new();
    let error = presentation
        .write_text_to(
            &mut output,
            TextOutputOptions::new("\n", "\n", u64::MAX, 1).with_empty_objects(false),
        )
        .unwrap_err();
    assert!(matches!(
        error,
        TextOutputError::Limit { progress, .. }
            if progress.bytes_written() == 3 && progress.objects_written() == 1
    ));
    assert_eq!(output, b"one");
}

#[test]
fn source_declared_slide_size_limit_precedes_payload_load() {
    let bytes = oversized_slide_declaration(package_bytes(&[slide_xml(r#"<a:t>tiny</a:t>"#)]));
    let source = Arc::new(VersionedSource::new(bytes));
    let presentation = SourceBackedPresentation::from_read_at(source.clone())
        .expect("oversized declaration presentation");
    let reads_before_sink = source.reads.load(Ordering::SeqCst);
    let mut output = Vec::new();
    let error = presentation
        .write_text_to(&mut output, parity_options(true))
        .expect_err("declared semantic raw limit must fail");
    assert!(matches!(
        error,
        TextOutputError::Document {
            source: Error::Limit { .. },
            progress,
        } if progress.bytes_written() == 0 && progress.objects_written() == 0
    ));
    assert!(output.is_empty());
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before_sink);
}

#[derive(Default)]
struct ShortWriter {
    bytes: Vec<u8>,
    chunk: usize,
}

impl Write for ShortWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let written = bytes.len().min(self.chunk.max(1));
        self.bytes.extend_from_slice(&bytes[..written]);
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct ZeroWriter;

impl Write for ZeroWriter {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        Ok(0)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct OverreportingWriter;

impl Write for OverreportingWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        Ok(bytes.len().saturating_add(1))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct InterruptedOnce {
    bytes: Vec<u8>,
    interrupted: bool,
}

impl Write for InterruptedOnce {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.interrupted {
            self.interrupted = true;
            return Err(io::Error::from(io::ErrorKind::Interrupted));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct FailAfter {
    bytes: Vec<u8>,
    remaining: usize,
}

impl Write for FailAfter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::other("injected sink failure"));
        }
        let written = bytes.len().min(self.remaining);
        self.bytes.extend_from_slice(&bytes[..written]);
        self.remaining -= written;
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn sink_progress_handles_short_zero_overreporting_interrupted_and_failure_writers() {
    let bytes = package_bytes(&[slide_xml(r#"<a:t>sink</a:t>"#)]);
    let package = Package::from_bytes(&bytes).expect("sink package");
    let presentation = package.presentation().expect("sink presentation");

    let mut short = ShortWriter {
        chunk: 2,
        ..ShortWriter::default()
    };
    let report = presentation
        .write_text_to(&mut short, parity_options(true))
        .expect("short writes are retried");
    assert_eq!(short.bytes, b"sink");
    assert_eq!(report.bytes_written(), 4);

    let mut interrupted = InterruptedOnce {
        bytes: Vec::new(),
        interrupted: false,
    };
    presentation
        .write_text_to(&mut interrupted, parity_options(true))
        .expect("interrupted writes are retried");
    assert_eq!(interrupted.bytes, b"sink");

    let mut zero = ZeroWriter;
    assert!(matches!(
        presentation.write_text_to(&mut zero, parity_options(true)),
        Err(TextOutputError::Sink { progress, .. }) if progress.bytes_written() == 0
    ));

    let mut overreporting = OverreportingWriter;
    assert!(matches!(
        presentation.write_text_to(&mut overreporting, parity_options(true)),
        Err(TextOutputError::Sink { progress, .. }) if progress.bytes_written() == 0
    ));

    let mut failing = FailAfter {
        bytes: Vec::new(),
        remaining: 2,
    };
    assert!(matches!(
        presentation.write_text_to(&mut failing, parity_options(true)),
        Err(TextOutputError::Sink { progress, .. }) if progress.bytes_written() == 2
    ));
}

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
    reads: AtomicUsize,
    armed: AtomicBool,
    change_on_read: AtomicBool,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
            reads: AtomicUsize::new(0),
            armed: AtomicBool::new(false),
            change_on_read: AtomicBool::new(false),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }

    fn arm_change_on_read(&self) {
        self.change_on_read.store(true, Ordering::SeqCst);
        self.armed.store(true, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        if self.armed.load(Ordering::SeqCst) && self.change_on_read.swap(false, Ordering::SeqCst) {
            self.changed();
        }
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            801,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn execution_context() -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "pptx-sequential-text-test",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("worker limit"),
        NonZeroUsize::new(1).expect("task limit"),
        NonZeroU64::new(u64::MAX).expect("memory limit"),
        0,
    )
    .expect("execution limits");
    (source, ExecutionContext::new(budget, cancellation, limits))
}

struct CancelOnWrite {
    source: CancellationSource,
    bytes: Vec<u8>,
    cancelled: bool,
}

struct ChangeOnWrite {
    source: Arc<VersionedSource>,
    bytes: Vec<u8>,
    changed: bool,
}

impl Write for ChangeOnWrite {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.changed {
            self.changed = true;
            self.source.changed();
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct ChangeAndFail {
    source: Arc<VersionedSource>,
    changed: bool,
}

impl Write for ChangeAndFail {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        if !self.changed {
            self.changed = true;
            self.source.changed();
        }
        Err(io::Error::other("injected sink failure"))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Write for CancelOnWrite {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.cancelled {
            self.cancelled = true;
            self.source.cancel();
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn source_stale_before_and_during_parse_is_document_progress() {
    let source = Arc::new(VersionedSource::new(package_bytes(&[slide_xml(
        r#"<a:t>stale</a:t>"#,
    )])));
    let presentation =
        SourceBackedPresentation::from_read_at(source.clone()).expect("source-backed package");
    source.changed();
    let mut output = Vec::new();
    assert!(matches!(
        presentation.write_text_to(&mut output, parity_options(true)),
        Err(TextOutputError::Document { progress, .. }) if progress.bytes_written() == 0
    ));

    let source = Arc::new(VersionedSource::new(package_bytes(&[slide_xml(
        r#"<a:t>during</a:t>"#,
    )])));
    let presentation =
        SourceBackedPresentation::from_read_at(source.clone()).expect("source-backed package");
    source.arm_change_on_read();
    let mut output = Vec::new();
    assert!(matches!(
        presentation.write_text_to(&mut output, parity_options(true)),
        Err(TextOutputError::Document { progress, .. }) if progress.bytes_written() == 0
    ));
    assert!(output.is_empty());
}

#[test]
fn source_stale_after_accepted_output_overrides_sink_result() {
    let source = Arc::new(VersionedSource::new(package_bytes(&[slide_xml(
        r#"<a:t>stale</a:t>"#,
    )])));
    let presentation =
        SourceBackedPresentation::from_read_at(source.clone()).expect("source-backed package");
    let mut sink = ChangeOnWrite {
        source,
        bytes: Vec::new(),
        changed: false,
    };
    let error = presentation
        .write_text_to(&mut sink, parity_options(true))
        .expect_err("source change must be observed after output");
    assert!(matches!(
        error,
        TextOutputError::Document { progress, .. }
            if progress.bytes_written() == 5 && progress.objects_written() == 1
    ));
    assert_eq!(sink.bytes, b"stale");
}

#[test]
fn source_cancellation_before_output_is_document_progress() {
    let (cancellation_source, context) = execution_context();
    let presentation = SourceBackedPresentation::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(package_bytes(&[slide_xml(
            r#"<a:t>before</a:t>"#,
        )]))),
        ReadLimits::default(),
        context,
    )
    .expect("managed source-backed package");
    cancellation_source.cancel();
    let mut output = Vec::new();
    let error = presentation
        .write_text_to(&mut output, parity_options(true))
        .expect_err("cancellation must be observed before output");
    assert!(matches!(
        error,
        TextOutputError::Document { progress, .. } if progress.bytes_written() == 0
    ));
    assert!(output.is_empty());
}

#[test]
fn source_cancellation_after_sink_acceptance_overrides_sink_result() {
    let (cancellation_source, context) = execution_context();
    let presentation = SourceBackedPresentation::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(package_bytes(&[slide_xml(
            r#"<a:t>cancel</a:t>"#,
        )]))),
        ReadLimits::default(),
        context,
    )
    .expect("managed source-backed package");
    let mut sink = CancelOnWrite {
        source: cancellation_source,
        bytes: Vec::new(),
        cancelled: false,
    };
    let error = presentation
        .write_text_to(&mut sink, parity_options(true))
        .expect_err("cancellation must be observed after sink output");
    assert!(
        matches!(error, TextOutputError::Document { progress, .. } if progress.bytes_written() == 6)
    );
    assert_eq!(sink.bytes, b"cancel");
}

#[test]
fn source_change_overrides_simultaneous_sink_failure() {
    let source = Arc::new(VersionedSource::new(package_bytes(&[slide_xml(
        r#"<a:t>sink</a:t>"#,
    )])));
    let presentation =
        SourceBackedPresentation::from_read_at(source.clone()).expect("source-backed package");
    let mut sink = ChangeAndFail {
        source,
        changed: false,
    };
    let error = presentation
        .write_text_to(&mut sink, parity_options(true))
        .expect_err("source change must take precedence");
    assert!(matches!(
        error,
        TextOutputError::Document { progress, .. } if progress.bytes_written() == 0
    ));
}

#[test]
fn xml_depth_limit_is_document_error() {
    let mut body = String::new();
    for _ in 0..129 {
        body.push_str("<a:grpSp>");
    }
    body.push_str("<a:t>deep</a:t>");
    for _ in 0..129 {
        body.push_str("</a:grpSp>");
    }

    let mut output = Vec::new();
    let package = Package::from_bytes(&package_bytes(&[slide_xml(&body)])).expect("test package");
    let presentation = package.presentation().expect("test presentation");
    let error = presentation
        .write_text_to(
            &mut output,
            TextOutputOptions::new("\n", "\n", u64::MAX, u64::MAX),
        )
        .unwrap_err();
    assert!(matches!(error, TextOutputError::Document { .. }));
    assert!(output.is_empty());
}
