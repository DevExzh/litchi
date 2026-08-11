use std::io;
use std::io::Write;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter};
use litchi_pptx::{Error, ReadLimits, SourceBackedPresentation, SourceBackedPresentationEditor};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const DRAWINGML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN: &str = "/ppt/presentation.xml";
const SLIDE: &str = "/ppt/slides/slide1.xml";
const UNUSED: &str = "/ppt/media/unused.bin";

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
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(11, self.revision.load(Ordering::SeqCst)))
    }
}

struct FailingSink {
    accepted: usize,
    limit: usize,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.limit {
            return Err(io::Error::other("injected sink failure"));
        }
        let written = bytes.len().min(self.limit - self.accepted);
        self.accepted += written;
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn shape_xml(pml: &str, text: &str) -> String {
    format!(
        r#"<p:sld xmlns:p="{pml}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Other"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>stable</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr/></p:sld>"#
    )
}

fn mce_slide_xml() -> String {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>before</a:t></a:r></a:p></p:txBody></p:sp>"#;
    format!(
        r#"<p:sld xmlns:p="{PML}" xmlns:a="{DRAWINGML}" xmlns:r="{REL}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><mc:AlternateContent><mc:Choice Requires="x">{shape}</mc:Choice><mc:Fallback>{shape}</mc:Fallback></mc:AlternateContent></p:spTree></p:cSld><p:clrMapOvr/></p:sld>"#
    )
}

fn fixture(presentation_suffix: &str, slide_xml: String, signed: bool) -> Vec<u8> {
    fixture_with_presentation_namespace(PML, presentation_suffix, slide_xml, signed)
}

fn fixture_with_presentation_namespace(
    presentation_namespace: &str,
    presentation_suffix: &str,
    slide_xml: String,
    signed: bool,
) -> Vec<u8> {
    let presentation_xml = format!(
        r#"<p:presentation xmlns:p="{presentation_namespace}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst>{presentation_suffix}</p:presentation>"#
    );
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(MAIN).unwrap(),
            ct::PML_PRESENTATION_MAIN.to_string(),
            presentation_xml.into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(SLIDE).unwrap(),
            ct::PML_SLIDE.to_string(),
            slide_xml.into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(UNUSED).unwrap(),
            "application/octet-stream".to_string(),
            (0..4096).map(|value| (value % 251) as u8).collect(),
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::SLIDE.to_string(),
            "slides/slide1.xml".to_string(),
            "rIdSlide".to_string(),
            litchi_opc::TargetMode::Internal,
        )
        .unwrap();
    package.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
    if signed {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_string(),
                b"<origin/>".to_vec(),
            )))
            .unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    }
    PackageWriter::to_bytes(&package).unwrap()
}

fn edit_commit(editor: &SourceBackedPresentationEditor) -> litchi_pptx::SourceBackedSlideCommit {
    let mut edit = editor.edit_slide(0).unwrap();
    assert!(edit.set_shape_text("Title", "after").unwrap());
    let commit = edit.commit();
    assert!(commit.is_changed());
    assert!(commit.patch().inverse().apply(commit.snapshot()).is_ok());
    commit
}

fn relationship_signatures(relationships: &litchi_opc::Relationships) -> Vec<String> {
    let mut signatures = relationships
        .iter()
        .map(|relationship| {
            format!(
                "{}\u{1f}{}\u{1f}{}\u{1f}{:?}",
                relationship.r_id(),
                relationship.reltype(),
                relationship.target_ref(),
                relationship.target_mode()
            )
        })
        .collect::<Vec<_>>();
    signatures.sort_unstable();
    signatures
}

#[test]
fn changed_edit_reopens_and_changes_only_the_selected_logical_part() {
    let source_bytes = fixture("", shape_xml(PML, "before"), false);
    let source: Arc<dyn ReadAt> = Arc::new(VersionedSource::new(source_bytes.clone()));
    let editor = SourceBackedPresentationEditor::from_read_at(source).unwrap();
    assert_eq!(editor.cache_diagnostics().successful_loads, 1);
    let commit = edit_commit(&editor);
    assert_eq!(editor.cache_diagnostics().successful_loads, 2);
    let mut output = Vec::new();
    editor
        .publish_slide_commit_to_stream(&mut output, &commit)
        .unwrap();

    let source = OpcPackage::from_bytes(&source_bytes).unwrap();
    let candidate = OpcPackage::from_bytes(&output).unwrap();
    assert_eq!(source.part_count(), candidate.part_count());
    assert_eq!(
        relationship_signatures(source.rels()),
        relationship_signatures(candidate.rels())
    );
    for part in source.iter_parts() {
        let output_part = candidate.get_part(part.partname()).unwrap();
        assert_eq!(part.content_type(), output_part.content_type());
        assert_eq!(
            relationship_signatures(part.rels()),
            relationship_signatures(output_part.rels())
        );
        if part.partname().as_str() == SLIDE {
            assert_ne!(part.blob(), output_part.blob());
            let output_xml = std::str::from_utf8(output_part.blob()).unwrap();
            assert!(output_xml.contains("after"));
            assert!(output_xml.contains("name=\"Other\""));
            assert!(output_xml.contains("stable"));
            assert_eq!(
                output_xml.replacen("after", "before", 1).as_bytes(),
                part.blob(),
                "only the selected shape text may differ"
            );
            assert!(
                !litchi_pptx::shape::read(output_part.blob())
                    .unwrap()
                    .is_rewritten()
            );
        } else {
            assert_eq!(part.blob(), output_part.blob());
        }
    }
    let reopened =
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(output))).unwrap();
    assert_eq!(reopened.slide(0).unwrap().text().unwrap(), "after\nstable");
}

#[test]
fn noop_is_byte_exact_and_signed_change_is_refused() {
    let signed = fixture("", shape_xml(PML, "before"), true);
    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        signed.clone(),
    )))
    .unwrap();
    let noop = editor.edit_slide(0).unwrap().commit();
    let mut output = Vec::new();
    editor
        .publish_slide_commit_to_stream(&mut output, &noop)
        .unwrap();
    assert_eq!(output, signed);

    let editor =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(signed)))
            .unwrap();
    let commit = edit_commit(&editor);
    output.clear();
    assert!(matches!(
        editor.publish_slide_commit_to_stream(&mut output, &commit),
        Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty());
}

#[test]
fn stale_foreign_closure_and_changed_source_refuse_before_output() {
    let bytes = fixture("", shape_xml(PML, "before"), false);
    let editor =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone())))
            .unwrap();
    let commit = edit_commit(&editor);
    let foreign = fixture("<p:extLst/>", shape_xml(PML, "before"), false);
    let foreign_editor =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(foreign)))
            .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        foreign_editor.publish_slide_commit_to_stream(&mut output, &commit),
        Err(Error::StaleSource)
    ));
    assert!(output.is_empty());

    let source = Arc::new(VersionedSource::new(bytes));
    let editor = SourceBackedPresentationEditor::from_read_at(source.clone()).unwrap();
    let commit = edit_commit(&editor);
    source.changed();
    assert!(matches!(
        editor.publish_slide_commit_to_stream(&mut output, &commit),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());
}

#[test]
fn mce_limits_strict_and_partial_sink_are_checked() {
    let mce = fixture("", mce_slide_xml(), false);
    let editor =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(mce))).unwrap();
    assert!(matches!(
        editor.edit_slide(0),
        Err(Error::UnsafeEdit { .. })
    ));

    let strict =
        fixture_with_presentation_namespace(STRICT_PML, "", shape_xml(STRICT_PML, "before"), false);
    let editor =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(strict)))
            .unwrap();
    let commit = edit_commit(&editor);
    let mut output = Vec::new();
    editor
        .publish_slide_commit_to_stream(&mut output, &commit)
        .unwrap();

    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        fixture("", shape_xml(PML, "before"), false),
    )))
    .unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    assert!(edit.set_shape_text("Title", "after").unwrap());
    assert!(matches!(
        edit.set_shape_text("Other", "changed"),
        Err(Error::UnsafeEdit { .. })
    ));

    let limits = ReadLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedPresentationEditor::from_read_at_with_limits(
            Arc::new(VersionedSource::new(fixture(
                "",
                shape_xml(PML, "before"),
                false
            ))),
            limits,
        ),
        Err(Error::Opc(OpcError::ReadLimit { .. }))
    ));

    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        fixture("", shape_xml(PML, "before"), false),
    )))
    .unwrap();
    let commit = edit_commit(&editor);
    let mut sink = FailingSink {
        accepted: 0,
        limit: 128,
    };
    assert!(matches!(
        editor.publish_slide_commit_to_stream(&mut sink, &commit),
        Err(Error::Opc(OpcError::IncompleteOutput { .. }))
    ));
    assert_eq!(sink.accepted, 128);
}
