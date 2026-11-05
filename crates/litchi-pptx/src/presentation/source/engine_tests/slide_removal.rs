#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "the in-memory fixtures and assertions intentionally panic"
)]

//! Source-backed, dependency-free whole-slide removal.
//!
//! The fixture is authored by the native builder so each slide has the normal
//! single-layout closure.  The tests then add one targeted topology defect at
//! a time and verify that planning refuses before the output stream receives
//! a byte.

use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::ops::Range;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

use crate::builder::Builder;
use crate::{Error, SlideRemovalRefusal, SourceBackedPresentationEditor, SourceLimits};
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, Position,
    ReadAt, SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, TargetMode};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const MAIN: &str = "/ppt/presentation.xml";
const REMOVED_SLIDE: &str = "/ppt/slides/slide2.xml";
const REMOVED_SLIDE_RELS: &str = "ppt/slides/_rels/slide2.xml.rels";
const RETAINED_SLIDE_RELS: &str = "ppt/slides/_rels/slide1.xml.rels";
const LARGE_MEMBER: &str = "ppt/media/untouched.bin";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn fixture() -> Vec<u8> {
    let mut builder = Builder::new().unwrap();
    {
        let presentation = builder.presentation().unwrap();
        for text in ["A", "B", "C"] {
            presentation
                .add_slide()
                .unwrap()
                .add_text_box(text, 0, 0, 914_400, 914_400);
        }
    }
    let package = builder.build().unwrap();
    package.to_bytes().unwrap()
}

fn open(bytes: &[u8]) -> SourceBackedPresentationEditor {
    SourceBackedPresentationEditor::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec())))
        .unwrap()
}

fn write_removal(bytes: &[u8], position: usize) -> crate::Result<Vec<u8>> {
    let editor = open(bytes);
    let plan = editor.plan_slide_removal(Position::new(position))?;
    let mut output = Vec::new();
    plan.write_to(&mut output)?;
    Ok(output)
}

fn member(bytes: &[u8], name: &str) -> Vec<u8> {
    ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())
        .unwrap()
        .read(name)
        .unwrap()
}

fn has_member(bytes: &[u8], name: &str) -> bool {
    ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())
        .unwrap()
        .file_names()
        .any(|candidate| candidate == name)
}

fn replace_once(bytes: &[u8], before: &[u8], after: &[u8]) -> Vec<u8> {
    let start = bytes
        .windows(before.len())
        .position(|window| window == before)
        .expect("fixture fragment must exist");
    let mut result = Vec::with_capacity(bytes.len() - before.len() + after.len());
    result.extend_from_slice(&bytes[..start]);
    result.extend_from_slice(after);
    result.extend_from_slice(&bytes[start + before.len()..]);
    result
}

fn rewrite_archive(
    source: &[u8],
    replacements: &[(&str, &[u8])],
    additions: &[(&str, &[u8])],
) -> Vec<u8> {
    let names = ArchiveReader::new(source, soapberry_zip::office::ArchiveLimits::default())
        .unwrap()
        .file_names()
        .map(str::to_owned)
        .collect::<Vec<_>>();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in names {
        let bytes = replacements
            .iter()
            .find(|(candidate, _)| *candidate == name)
            .map_or_else(|| member(source, &name), |(_, bytes)| bytes.to_vec());
        writer.write_stored(&name, &bytes).unwrap();
    }
    for (name, bytes) in additions {
        writer.write_stored(name, bytes).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn rewrite_package(source: &[u8], edit: impl FnOnce(&mut OpcPackage)) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(source, SourceLimits::default()).unwrap();
    edit(&mut package);
    PackageWriter::to_bytes(&package).unwrap()
}

fn with_large_untouched(source: &[u8]) -> (Vec<u8>, Vec<u8>) {
    let payload = (0..256 * 1024)
        .map(|index| (index % 251) as u8)
        .collect::<Vec<_>>();
    let result = rewrite_package(source, |package| {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/ppt/media/untouched.bin").unwrap(),
                "application/octet-stream".to_owned(),
                payload.clone(),
            )))
            .unwrap();
    });
    (result, payload)
}

fn with_shared_incoming(source: &[u8]) -> Vec<u8> {
    rewrite_package(source, |package| {
        package.relate_to("ppt/slides/slide2.xml", rt::SLIDE);
    })
}

fn with_hidden_incoming_relationship(source: &[u8]) -> Vec<u8> {
    let original = member(source, RETAINED_SLIDE_RELS);
    let hidden = format!(
        r#"<x:hidden xmlns:x="urn:litchi:hidden"><Relationship Id="rIdHidden" Type="{}" Target="slide2.xml"/></x:hidden></Relationships>"#,
        rt::SLIDE
    );
    let updated = replace_once(&original, b"</Relationships>", hidden.as_bytes());
    rewrite_archive(source, &[(RETAINED_SLIDE_RELS, &updated)], &[])
}

fn with_slide_dependency(source: &[u8]) -> Vec<u8> {
    rewrite_package(source, |package| {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/ppt/media/dependency.bin").unwrap(),
                "application/octet-stream".to_owned(),
                b"slide dependency".to_vec(),
            )))
            .unwrap();
        package
            .get_part_mut(&PackURI::new(REMOVED_SLIDE).unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "../media/dependency.bin".to_owned(),
                "rIdDependency".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    })
}

fn with_unknown_presentation_metadata(source: &[u8]) -> Vec<u8> {
    let original = member(source, MAIN);
    let rooted = replace_once(
        &original,
        format!(r#"xmlns:r="{REL}""#).as_bytes(),
        format!(r#"xmlns:r="{REL}" xmlns:x="urn:litchi:unknown""#).as_bytes(),
    );
    let updated = replace_once(
        &rooted,
        b"</p:presentation>",
        br#"<x:unknownMetadata/></p:presentation>"#,
    );
    rewrite_archive(source, &[("ppt/presentation.xml", &updated)], &[])
}

fn with_custom_show_reference(source: &[u8]) -> Vec<u8> {
    let main_uri = PackURI::new(MAIN).unwrap();
    let package = OpcPackage::from_bytes(source, SourceLimits::default()).unwrap();
    let presentation = package.get_part(&main_uri).unwrap();
    let relationship_id = presentation
        .rels()
        .iter()
        .find(|relationship| {
            relationship
                .target_partname()
                .is_ok_and(|target| target.as_str() == REMOVED_SLIDE)
        })
        .expect("the fixture must bind slide 2 from the presentation")
        .r_id()
        .to_owned();
    let original = presentation.blob().to_vec();
    let suffix = format!(
        r#"<p:custShowLst><p:custShow name="Opening" id="7"><p:sldLst><p:sld r:id="{relationship_id}"/></p:sldLst></p:custShow></p:custShowLst></p:presentation>"#
    );
    let updated = replace_once(&original, b"</p:presentation>", suffix.as_bytes());
    rewrite_archive(source, &[("ppt/presentation.xml", &updated)], &[])
}

fn with_presentation_properties(source: &[u8], xml: &str) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(source, SourceLimits::default()).unwrap();
    let properties_uri = PackURI::new("/ppt/presProps.xml").unwrap();
    package
        .get_part_mut(&properties_uri)
        .expect("the native fixture includes presProps.xml")
        .set_blob(xml.as_bytes().to_vec());
    PackageWriter::to_bytes(&package).unwrap()
}

fn selected_binding_fragment(xml: &[u8], relationship_id: &str) -> Vec<u8> {
    let marker = format!(r#"r:id="{relationship_id}""#);
    let marker_start = xml
        .windows(marker.len())
        .position(|window| window == marker.as_bytes())
        .expect("selected relationship binding must exist");
    let start = xml[..marker_start]
        .windows(b"<p:sldId".len())
        .rposition(|window| window == b"<p:sldId")
        .expect("selected slide binding start must exist");
    let end_offset = xml[marker_start..]
        .windows(b"/>".len())
        .position(|window| window == b"/>")
        .expect("selected slide binding must be empty-element XML");
    xml[start..marker_start + end_offset + 2].to_vec()
}

fn with_lexical_presentation_root(source: &[u8]) -> (Vec<u8>, Vec<u8>) {
    let main = member(source, MAIN);
    let package = OpcPackage::from_bytes(source, SourceLimits::default()).unwrap();
    let main_part = package.get_part(&PackURI::new(MAIN).unwrap()).unwrap();
    let relationship_id = main_part
        .rels()
        .iter()
        .filter(|relationship| {
            relationship
                .target_partname()
                .is_ok_and(|target| target.as_str() == REMOVED_SLIDE)
        })
        .map(|relationship| relationship.r_id().to_owned())
        .next()
        .expect("the fixture must bind slide 2 from the presentation");
    let selected = selected_binding_fragment(&main, &relationship_id);
    let mut lexical = replace_once(
        &main,
        b"<p:presentation",
        b"<!-- before-root -->\n<p:presentation",
    );
    lexical = replace_once(
        &lexical,
        b"<p:sldIdLst>",
        b"<p:sldIdLst>\n    <!-- list-start -->\n      ",
    );
    let mut selected_with_comments = b"<!-- selected-before -->\n      ".to_vec();
    selected_with_comments.extend_from_slice(&selected);
    selected_with_comments.extend_from_slice(b"\n      <!-- selected-after -->");
    lexical = replace_once(&lexical, &selected, &selected_with_comments);
    lexical = replace_once(
        &lexical,
        b"</p:sldIdLst>",
        b"\n    <!-- list-end -->\n  </p:sldIdLst>",
    );
    lexical = replace_once(
        &lexical,
        b"</p:presentation>",
        b"\n<!-- after-root -->\n</p:presentation>",
    );
    let expected = replace_once(&lexical, &selected, b"");
    (
        rewrite_archive(source, &[("ppt/presentation.xml", &lexical)], &[]),
        expected,
    )
}

fn with_macro_infrastructure(source: &[u8]) -> Vec<u8> {
    rewrite_package(source, |package| {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/ppt/vbaProject.bin").unwrap(),
                ct::OFC_VBA_PROJECT.to_owned(),
                b"macro".to_vec(),
            )))
            .unwrap();
    })
}

fn signed_fixture(source: &[u8]) -> Vec<u8> {
    let content_types = replace_once(
        &member(source, "[Content_Types].xml"),
        b"</Types>",
        format!(
            r#"<Override PartName="/_xmlsignatures/origin.sigs" ContentType="{}"/></Types>"#,
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN
        )
        .as_bytes(),
    );
    let root_relationships = replace_once(
        &member(source, "_rels/.rels"),
        b"</Relationships>",
        format!(
            r#"<Relationship Id="rIdSignatureOrigin" Type="{}" Target="_xmlsignatures/origin.sigs"/></Relationships>"#,
            rt::DIGITAL_SIGNATURE_ORIGIN
        )
        .as_bytes(),
    );
    rewrite_archive(
        source,
        &[
            ("[Content_Types].xml", &content_types),
            ("_rels/.rels", &root_relationships),
        ],
        &[("_xmlsignatures/origin.sigs", b"<origin/>")],
    )
}

fn assert_refusal(source: &[u8], expected: SlideRemovalRefusal) {
    let error = open(source)
        .plan_slide_removal(Position::new(1))
        .err()
        .unwrap();
    assert!(
        matches!(error, Error::Refused(refusal) if refusal.kind() == expected.kind()),
        "unexpected slide-removal refusal: {error:?}"
    );
}

fn stored_payload_range(bytes: &[u8], name: &str) -> Range<u64> {
    let signature = 0x0403_4b50_u32.to_le_bytes();
    for offset in bytes
        .windows(signature.len())
        .enumerate()
        .filter_map(|(offset, window)| (window == signature).then_some(offset))
    {
        let name_len = u16::from_le_bytes(bytes[offset + 26..offset + 28].try_into().unwrap());
        let extra_len = u16::from_le_bytes(bytes[offset + 28..offset + 30].try_into().unwrap());
        let name_start = offset + 30;
        let name_end = name_start + usize::from(name_len);
        if bytes.get(name_start..name_end) != Some(name.as_bytes()) {
            continue;
        }
        let data_start = name_end + usize::from(extra_len);
        let compressed_size =
            u32::from_le_bytes(bytes[offset + 18..offset + 22].try_into().unwrap()) as usize;
        return (data_start as u64)..((data_start + compressed_size) as u64);
    }
    panic!("stored ZIP member not found: {name}");
}

#[derive(Debug)]
struct ToggleSource {
    bytes: Arc<Vec<u8>>,
    denied: AtomicBool,
    denied_range: Range<u64>,
    denied_reads: AtomicU64,
}

impl ToggleSource {
    fn new(bytes: Vec<u8>, denied_range: Range<u64>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            denied: AtomicBool::new(false),
            denied_range,
            denied_reads: AtomicU64::new(0),
        }
    }
}

impl ReadAt for ToggleSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let end = offset.saturating_add(output.len() as u64);
        if self.denied.load(Ordering::SeqCst)
            && offset < self.denied_range.end
            && end > self.denied_range.start
        {
            self.denied_reads.fetch_add(1, Ordering::SeqCst);
            return Err(io::Error::other("untouched payload denied"));
        }
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let end = start.saturating_add(output.len()).min(self.bytes.len());
        output[..end - start].copy_from_slice(&self.bytes[start..end]);
        Ok(end - start)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(17, 0))
    }
}

#[derive(Debug)]
struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let end = start.saturating_add(output.len()).min(self.bytes.len());
        output[..end - start].copy_from_slice(&self.bytes[start..end]);
        Ok(end - start)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(19, self.revision.load(Ordering::SeqCst)))
    }
}

fn managed_context() -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "pptx-source-slide-removal",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(u64::MAX).unwrap(),
        0,
    )
    .unwrap();
    (source, ExecutionContext::new(budget, cancellation, limits))
}

#[test]
fn removes_one_slide_and_its_owned_relationship_member() {
    let source = fixture();
    let output = write_removal(&source, 1).unwrap();
    let package = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    let main = package.get_part(&PackURI::new(MAIN).unwrap()).unwrap();
    assert_eq!(
        SourceBackedPresentationEditor::from_read_at(Arc::new(OwnedSource::new(output.clone())))
            .unwrap()
            .slide_count(),
        2
    );
    assert!(
        package
            .get_part(&PackURI::new(REMOVED_SLIDE).unwrap())
            .is_err()
    );
    assert!(!has_member(&output, REMOVED_SLIDE_RELS));
    assert!(
        main.rels()
            .iter()
            .all(|relationship| relationship.target_partname().unwrap().as_str() != REMOVED_SLIDE)
    );
}

#[test]
fn untouched_payload_is_exact_and_slide_planning_does_not_read_it() {
    let (source, payload) = with_large_untouched(&fixture());
    let denied_range = stored_payload_range(&source, LARGE_MEMBER);
    let handle = Arc::new(ToggleSource::new(source.clone(), denied_range));
    let editor = SourceBackedPresentationEditor::from_read_at(handle.clone()).unwrap();
    handle.denied.store(true, Ordering::SeqCst);
    let plan = editor.plan_slide_removal(Position::new(1)).unwrap();
    assert_eq!(handle.denied_reads.load(Ordering::SeqCst), 0);
    handle.denied.store(false, Ordering::SeqCst);
    let mut output = Vec::new();
    plan.write_to(&mut output).unwrap();
    assert_eq!(member(&output, LARGE_MEMBER), payload);
}

#[test]
fn final_slide_shared_owner_dependency_unknown_macro_and_signed_sources_refuse_atomically() {
    let one_slide = {
        let mut builder = Builder::new().unwrap();
        builder
            .presentation()
            .unwrap()
            .add_slide()
            .unwrap()
            .add_text_box("only", 0, 0, 914_400, 914_400);
        let package = builder.build().unwrap();
        package.to_bytes().unwrap()
    };
    let final_error = open(&one_slide)
        .plan_slide_removal(Position::new(0))
        .err()
        .unwrap();
    assert!(matches!(
        final_error,
        Error::Refused(refusal) if refusal.kind() == SlideRemovalRefusal::FinalSlide.kind()
    ));

    assert_refusal(
        &with_shared_incoming(&fixture()),
        SlideRemovalRefusal::SharedOwner,
    );
    assert_refusal(
        &with_hidden_incoming_relationship(&fixture()),
        SlideRemovalRefusal::UnknownSemanticSurface,
    );
    // A Part only the removed slide targets leaves with it (record 0303).
    let dependent = with_slide_dependency(&fixture());
    let mut output = Vec::new();
    open(&dependent)
        .plan_slide_removal(Position::new(1))
        .unwrap()
        .write_to(&mut output)
        .unwrap();
    let removed = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    assert!(
        removed
            .get_part(&PackURI::new("/ppt/media/dependency.bin").unwrap())
            .is_err()
    );
    assert_refusal(
        &with_unknown_presentation_metadata(&fixture()),
        SlideRemovalRefusal::UnknownSemanticSurface,
    );
    assert_refusal(
        &with_macro_infrastructure(&fixture()),
        SlideRemovalRefusal::MacroEnabledPackage,
    );
    assert_refusal(
        &signed_fixture(&fixture()),
        SlideRemovalRefusal::SignedPackage,
    );
}

#[test]
fn stale_source_and_cancellation_are_rejected_before_output() {
    let source = Arc::new(VersionedSource::new(fixture()));
    let editor = SourceBackedPresentationEditor::from_read_at(source.clone()).unwrap();
    let plan = editor.plan_slide_removal(Position::new(1)).unwrap();
    source.changed();
    let mut output = Vec::new();
    assert!(matches!(
        plan.write_to(&mut output),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());

    let (cancellation_source, context) = managed_context();
    let editor = SourceBackedPresentationEditor::from_read_at_with_options(
        Arc::new(OwnedSource::new(fixture())),
        crate::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let plan = editor.plan_slide_removal(Position::new(1)).unwrap();
    cancellation_source.cancel();
    output.clear();
    assert!(matches!(
        plan.write_to(&mut output),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert!(output.is_empty());
}

#[test]
fn formatted_source_root_preserves_every_byte_outside_selected_binding() {
    let (source, expected_root) = with_lexical_presentation_root(&fixture());
    let output = write_removal(&source, 1).unwrap();
    assert_eq!(member(&output, MAIN), expected_root);
    let root = String::from_utf8(member(&output, MAIN)).unwrap();
    assert!(root.contains("<!-- before-root -->"));
    assert!(root.contains("<!-- selected-before -->"));
    assert!(root.contains("<!-- selected-after -->"));
    assert!(root.contains("<!-- list-start -->"));
    assert!(root.contains("<!-- list-end -->"));
    assert!(root.contains("r:id=\"rId"));
}

#[test]
fn numeric_presentation_properties_range_and_custom_show_reference_refuse_selected_slide() {
    let properties = r#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:showPr><p:sldRg st="1" end="2"/></p:showPr></p:presentationPr>"#;
    assert_refusal(
        &with_presentation_properties(&fixture(), properties),
        SlideRemovalRefusal::SharedOwner,
    );
    assert_refusal(
        &with_custom_show_reference(&fixture()),
        SlideRemovalRefusal::SharedOwner,
    );
}
