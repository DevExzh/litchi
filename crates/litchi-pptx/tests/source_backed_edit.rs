use std::collections::HashMap;
use std::io;
use std::io::Write;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter};
use litchi_pptx::transition::{Kind as TransitionKind, Ms, Side, Speed, Transition};
use litchi_pptx::{
    Error, MAX_SHAPE_TEXT_REPLACEMENTS, MAX_SOURCE_BACKED_SLIDE_BATCH, ReadLimits,
    ShapeTextReplacement, SourceBackedPresentation, SourceBackedPresentationEditor,
};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const DRAWINGML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN: &str = "/ppt/presentation.xml";
const SLIDE: &str = "/ppt/slides/slide1.xml";
const SLIDE_THREE: &str = "/ppt/slides/slide3.xml";
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

fn slide_with_tail(pml: &str, text: &str, tail: &str) -> String {
    shape_xml(pml, text).replacen("</p:sld>", &format!("{tail}</p:sld>"), 1)
}

fn direct_transition(kind: TransitionKind) -> Transition {
    Transition::new(kind)
        .with_speed(Speed::Fast)
        .with_click(false)
        .with_after(Ms::new(750).unwrap())
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

fn unchecked_slide_fixture(slide_xml: &[u8]) -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer
        .write_stored(
            "[Content_Types].xml",
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
        )
        .unwrap();
    writer
        .write_stored(
            "_rels/.rels",
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        )
        .unwrap();
    writer
        .write_stored(
            "ppt/presentation.xml",
            format!(r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst></p:presentation>"#).as_bytes(),
        )
        .unwrap();
    writer
        .write_stored(
            "ppt/_rels/presentation.xml.rels",
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
        )
        .unwrap();
    writer
        .write_stored("ppt/slides/slide1.xml", slide_xml)
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn multi_fixture(slide_count: usize, signed: bool) -> Vec<u8> {
    let slide_ids = (0..slide_count)
        .map(|index| {
            format!(
                r#"<p:sldId id="{}" r:id="rIdSlide{}"/>"#,
                256 + index,
                index + 1
            )
        })
        .collect::<String>();
    let presentation_xml = format!(
        r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst>{slide_ids}</p:sldIdLst></p:presentation>"#
    );
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(MAIN).unwrap(),
            ct::PML_PRESENTATION_MAIN.to_string(),
            presentation_xml.into_bytes(),
        )))
        .unwrap();
    for index in 0..slide_count {
        let uri = PackURI::new(format!("/ppt/slides/slide{}.xml", index + 1)).unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                uri,
                ct::PML_SLIDE.to_string(),
                shape_xml(PML, &format!("before-{index}")).into_bytes(),
            )))
            .unwrap();
        package
            .get_part_mut(&PackURI::new(MAIN).unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::SLIDE.to_string(),
                format!("slides/slide{}.xml", index + 1),
                format!("rIdSlide{}", index + 1),
                litchi_opc::TargetMode::Internal,
            )
            .unwrap();
    }
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(UNUSED).unwrap(),
            "application/octet-stream".to_string(),
            (0..4096).map(|value| (value % 251) as u8).collect(),
        )))
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

#[derive(Debug)]
struct RawRecord {
    local: Vec<u8>,
    central: Vec<u8>,
}

fn raw_records(data: &[u8]) -> HashMap<Vec<u8>, RawRecord> {
    let archive = soapberry_zip::ZipArchive::from_slice(data)
        .unwrap()
        .into_zip_archive();
    let mut scratch = vec![0; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = soapberry_zip::PreservationIndex::new(&archive, &mut scratch).unwrap();
    index
        .entries()
        .iter()
        .map(|entry| {
            let local = entry.local_span();
            let central = entry.central_record();
            (
                entry.raw_name_bytes().to_vec(),
                RawRecord {
                    local: data[local.start as usize..local.end as usize].to_vec(),
                    central: data[central.start as usize..central.end as usize].to_vec(),
                },
            )
        })
        .collect()
}

fn central_without_local_offset(bytes: &[u8]) -> Vec<u8> {
    let mut bytes = bytes.to_vec();
    bytes[42..46].fill(0);
    bytes
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
fn multi_slide_batch_is_atomic_sorted_and_raw_copies_unselected_members() {
    let source_bytes = multi_fixture(3, false);
    let source_raw = raw_records(&source_bytes);
    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        source_bytes.clone(),
    )))
    .unwrap();
    assert_eq!(editor.cache_diagnostics().successful_loads, 1);
    let mut edit = editor.edit_slides();
    assert_eq!(
        edit.set_shape_texts(2, &[ShapeTextReplacement::named("Title", "after-2")])
            .unwrap(),
        1
    );
    assert_eq!(
        edit.set_shape_texts(0, &[ShapeTextReplacement::named("Title", "after-0")])
            .unwrap(),
        1
    );
    let commit = edit.commit().unwrap();
    assert!(commit.is_changed());
    assert_eq!(
        commit
            .snapshot()
            .slides()
            .map(|slide| slide.position())
            .collect::<Vec<_>>(),
        vec![0, 2]
    );
    let replayed = commit.patch().apply(commit.patch().source()).unwrap();
    assert!(commit.patch().inverse().apply(&replayed).is_ok());
    assert_eq!(editor.cache_diagnostics().successful_loads, 3);

    let mut output = Vec::new();
    let published = editor
        .publish_slide_batch_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert!(commit.patch().inverse().apply(&published).is_ok());
    assert!(commit.patch().apply(&published).is_err());

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
        if matches!(part.partname().as_str(), SLIDE | SLIDE_THREE) {
            assert_ne!(part.blob(), output_part.blob());
        } else {
            assert_eq!(part.blob(), output_part.blob());
        }
    }
    let output_raw = raw_records(&output);
    for (name, source_record) in source_raw {
        if matches!(
            name.as_slice(),
            b"ppt/slides/slide1.xml" | b"ppt/slides/slide3.xml"
        ) {
            assert_ne!(output_raw[&name].local, source_record.local, "{name:?}");
        } else {
            assert_eq!(output_raw[&name].local, source_record.local, "{name:?}");
            assert_eq!(
                central_without_local_offset(&output_raw[&name].central),
                central_without_local_offset(&source_record.central),
                "{name:?}"
            );
        }
    }
    let reopened =
        SourceBackedPresentation::from_read_at(Arc::new(VersionedSource::new(output))).unwrap();
    assert_eq!(
        reopened.slide(0).unwrap().text().unwrap(),
        "after-0\nstable"
    );
    assert_eq!(
        reopened.slide(1).unwrap().text().unwrap(),
        "before-1\nstable"
    );
    assert_eq!(
        reopened.slide(2).unwrap().text().unwrap(),
        "after-2\nstable"
    );
}

#[test]
fn multi_slide_batch_noop_signed_source_is_exact_and_changed_signed_is_refused() {
    let signed = multi_fixture(3, true);
    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        signed.clone(),
    )))
    .unwrap();
    let mut edit = editor.edit_slides();
    assert_eq!(
        edit.set_shape_texts(0, &[ShapeTextReplacement::named("Title", "before-0")])
            .unwrap(),
        0
    );
    assert_eq!(
        edit.set_shape_texts(2, &[ShapeTextReplacement::named("Title", "before-2")])
            .unwrap(),
        0
    );
    let commit = edit.commit().unwrap();
    assert!(!commit.is_changed());
    let mut output = Vec::new();
    editor
        .publish_slide_batch_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, signed);

    let editor =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(signed)))
            .unwrap();
    let mut edit = editor.edit_slides();
    edit.set_shape_texts(0, &[ShapeTextReplacement::named("Title", "changed")])
        .unwrap();
    edit.set_shape_texts(2, &[ShapeTextReplacement::named("Title", "before-2")])
        .unwrap();
    let commit = edit.commit().unwrap();
    output.clear();
    assert!(matches!(
        editor.publish_slide_batch_commit_to_stream(&mut output, &commit),
        Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty());
}

#[test]
fn multi_slide_batch_rejects_empty_duplicate_over_bound_and_stale_sets() {
    let source = multi_fixture(MAX_SOURCE_BACKED_SLIDE_BATCH + 1, false);
    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        source.clone(),
    )))
    .unwrap();
    assert!(matches!(
        editor.edit_slides().commit(),
        Err(Error::UnsafeEdit { .. })
    ));
    let mut edit = editor.edit_slides();
    for position in 0..MAX_SOURCE_BACKED_SLIDE_BATCH {
        let text = format!("before-{position}");
        edit.set_shape_texts(position, &[ShapeTextReplacement::named("Title", &text)])
            .unwrap();
    }
    assert!(matches!(
        edit.set_shape_texts(
            MAX_SOURCE_BACKED_SLIDE_BATCH,
            &[ShapeTextReplacement::named("Title", "bounded")],
        ),
        Err(Error::Limit { .. })
    ));
    assert!(matches!(
        edit.set_shape_texts(0, &[ShapeTextReplacement::named("Title", "duplicate")]),
        Err(Error::UnsafeEdit { .. })
    ));
    let commit = edit.commit().unwrap();

    let foreign = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        multi_fixture(MAX_SOURCE_BACKED_SLIDE_BATCH + 2, false),
    )))
    .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        foreign.publish_slide_batch_commit_to_stream(&mut output, &commit),
        Err(Error::StaleSource)
    ));
    assert!(output.is_empty());

    let versioned = Arc::new(VersionedSource::new(source));
    let editor = SourceBackedPresentationEditor::from_read_at(versioned.clone()).unwrap();
    let mut edit = editor.edit_slides();
    edit.set_shape_texts(0, &[ShapeTextReplacement::named("Title", "changed")])
        .unwrap();
    let commit = edit.commit().unwrap();
    versioned.changed();
    assert!(matches!(
        editor.publish_slide_batch_commit_to_stream(&mut output, &commit),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());
}

#[test]
fn multi_slide_batch_reports_partial_sink_failure() {
    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        multi_fixture(3, false),
    )))
    .unwrap();
    let mut edit = editor.edit_slides();
    edit.set_shape_texts(0, &[ShapeTextReplacement::named("Title", "after-0")])
        .unwrap();
    edit.set_shape_texts(2, &[ShapeTextReplacement::named("Title", "after-2")])
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut sink = FailingSink {
        accepted: 0,
        limit: 128,
    };
    assert!(matches!(
        editor.publish_slide_batch_commit_to_stream(&mut sink, &commit),
        Err(Error::Opc(OpcError::IncompleteOutput { .. }))
    ));
    assert_eq!(sink.accepted, 128);
}

fn publish_batch(
    source: &[u8],
    replacements: &[ShapeTextReplacement<'_>],
) -> (Vec<u8>, usize, u64) {
    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        source.to_vec(),
    )))
    .unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    let changed = edit.set_shape_texts(replacements).unwrap();
    let commit = edit.commit();
    let materializations = editor.cache_diagnostics().successful_loads;
    let mut output = Vec::new();
    editor
        .publish_slide_commit_to_stream(&mut output, &commit)
        .unwrap();
    (output, changed, materializations)
}

#[test]
fn atomic_batch_is_order_independent_and_changes_two_shapes_in_one_part() {
    let source = fixture("", shape_xml(PML, "before"), false);
    let forward = [
        ShapeTextReplacement::named("Title", "after & <title>"),
        ShapeTextReplacement::at(1, "after other"),
    ];
    let reverse = [forward[1], forward[0]];
    let (first, changed, materializations) = publish_batch(&source, &forward);
    let (second, reverse_changed, reverse_materializations) = publish_batch(&source, &reverse);
    assert_eq!(changed, 2);
    assert_eq!(reverse_changed, 2);
    assert_eq!(materializations, 2);
    assert_eq!(reverse_materializations, 2);
    assert_eq!(first, second);

    let before = OpcPackage::from_bytes(&source).unwrap();
    let after = OpcPackage::from_bytes(&first).unwrap();
    for part in before.iter_parts() {
        let candidate = after.get_part(part.partname()).unwrap();
        assert_eq!(part.content_type(), candidate.content_type());
        assert_eq!(
            relationship_signatures(part.rels()),
            relationship_signatures(candidate.rels())
        );
        if part.partname().as_str() == SLIDE {
            let scene = litchi_pptx::shape::read(candidate.blob()).unwrap();
            assert_eq!(
                scene.shape("Title").unwrap().common().text(),
                Some("after & <title>")
            );
            assert_eq!(
                scene.shape("Other").unwrap().common().text(),
                Some("after other")
            );
            let restored = std::str::from_utf8(candidate.blob())
                .unwrap()
                .replacen("after &amp; &lt;title&gt;", "before", 1)
                .replacen("after other", "stable", 1);
            assert_eq!(restored.as_bytes(), part.blob());
        } else {
            assert_eq!(part.blob(), candidate.blob());
        }
    }
}

#[test]
fn batch_preflight_is_atomic_and_duplicate_identity_is_typed() {
    let source = fixture("", shape_xml(PML, "before"), false);
    let editor =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(source)))
            .unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    let over_limit =
        vec![ShapeTextReplacement::named("Title", "bounded"); MAX_SHAPE_TEXT_REPLACEMENTS + 1];
    assert!(matches!(
        edit.set_shape_texts(&over_limit),
        Err(Error::Limit { .. })
    ));
    assert!(matches!(
        edit.set_shape_texts(&[
            ShapeTextReplacement::named("Title", "first"),
            ShapeTextReplacement::at(0, "second"),
        ]),
        Err(Error::DuplicateShapeTextSelection { index: 0 })
    ));
    assert!(
        edit.set_shape_texts(&[
            ShapeTextReplacement::named("Title", "valid"),
            ShapeTextReplacement::named("Other", "bad\0text"),
        ])
        .is_err()
    );
    assert_eq!(
        edit.set_shape_texts(&[
            ShapeTextReplacement::named("Title", "valid"),
            ShapeTextReplacement::named("Other", "also valid"),
        ])
        .unwrap(),
        2
    );
    assert!(matches!(
        edit.set_shape_texts(&[ShapeTextReplacement::at(0, "again")]),
        Err(Error::UnsafeEdit { .. })
    ));
}

#[test]
fn batch_expansion_honors_the_source_part_limit_before_allocation() {
    let source = fixture("", shape_xml(PML, "before"), false);
    let limits = ReadLimits::builder()
        .max_part_bytes(4096)
        .unwrap()
        .build()
        .unwrap();
    let editor = SourceBackedPresentationEditor::from_read_at_with_limits(
        Arc::new(VersionedSource::new(source)),
        limits,
    )
    .unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    let oversized = "x".repeat(5000);
    assert!(matches!(
        edit.set_shape_texts(&[
            ShapeTextReplacement::named("Title", &oversized),
            ShapeTextReplacement::named("Other", "another replacement"),
        ]),
        Err(Error::Limit { .. })
    ));
    assert_eq!(
        edit.set_shape_texts(&[
            ShapeTextReplacement::named("Title", "short"),
            ShapeTextReplacement::named("Other", "tiny"),
        ])
        .unwrap(),
        2
    );
}

#[test]
fn empty_and_all_equal_batch_preserve_signed_source_exactly() {
    let source = fixture("", shape_xml(PML, "before"), true);
    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        source.clone(),
    )))
    .unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    assert_eq!(edit.set_shape_texts(&[]).unwrap(), 0);
    assert_eq!(
        edit.set_shape_texts(&[
            ShapeTextReplacement::named("Title", "before"),
            ShapeTextReplacement::named("Other", "stable"),
        ])
        .unwrap(),
        0
    );
    assert!(matches!(
        edit.set_shape_text("Title", "changed"),
        Err(Error::UnsafeEdit { .. })
    ));
    let commit = edit.commit();
    assert!(!commit.is_changed());
    let mut output = Vec::new();
    editor
        .publish_slide_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source);
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

#[test]
fn direct_transition_set_replaces_clears_and_reopens_with_one_part_overlay() {
    let slide = slide_with_tail(PML, "before", "<p:timing/><p:extLst/>");
    let source_bytes = fixture("", slide.clone(), false);
    let source_raw = raw_records(&source_bytes);
    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        source_bytes.clone(),
    )))
    .unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    assert!(edit.source().transition().unwrap().is_none());
    let requested = direct_transition(TransitionKind::Push(Side::Left));
    assert!(edit.set_transition(&requested).unwrap());
    let commit = edit.commit();
    assert!(commit.is_changed());
    let replayed = commit.patch().apply(commit.patch().source()).unwrap();
    assert!(commit.patch().inverse().apply(&replayed).is_ok());

    let mut output = Vec::new();
    let published = editor
        .publish_slide_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert!(commit.patch().inverse().apply(&published).is_ok());
    assert!(commit.patch().apply(&published).is_err());
    let output_raw = raw_records(&output);
    for (name, source_record) in source_raw {
        if name.as_slice() == b"ppt/slides/slide1.xml" {
            assert_ne!(output_raw[&name].local, source_record.local);
        } else {
            assert_eq!(output_raw[&name].local, source_record.local, "{name:?}");
            assert_eq!(
                central_without_local_offset(&output_raw[&name].central),
                central_without_local_offset(&source_record.central),
                "{name:?}"
            );
        }
    }
    let package = OpcPackage::from_bytes(&output).unwrap();
    let changed_slide = package.get_part(&PackURI::new(SLIDE).unwrap()).unwrap();
    let changed_xml = std::str::from_utf8(changed_slide.blob()).unwrap();
    let fragment = litchi_pptx::transition::write(&requested).unwrap();
    assert_eq!(
        changed_xml,
        slide.replacen("<p:timing/>", &format!("{fragment}<p:timing/>"), 1)
    );
    let reopened =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(output)))
            .unwrap();
    let readback = reopened
        .slide_snapshot(0)
        .unwrap()
        .transition()
        .unwrap()
        .unwrap();
    assert!(readback.same_semantics(&requested));

    let mut replace = reopened.edit_slide(0).unwrap();
    let mut replacement = readback;
    replacement.set_speed(Speed::Slow);
    assert!(replace.set_transition(&replacement).unwrap());
    let replacement_commit = replace.commit();
    let mut replaced = Vec::new();
    reopened
        .publish_slide_commit_to_stream(&mut replaced, &replacement_commit)
        .unwrap();
    let reopened =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(replaced)))
            .unwrap();
    assert!(
        reopened
            .slide_snapshot(0)
            .unwrap()
            .transition()
            .unwrap()
            .unwrap()
            .same_semantics(&replacement)
    );
    let mut clear = reopened.edit_slide(0).unwrap();
    assert!(clear.clear_transition().unwrap());
    let clear_commit = clear.commit();
    let mut cleared = Vec::new();
    reopened
        .publish_slide_commit_to_stream(&mut cleared, &clear_commit)
        .unwrap();
    let reopened =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(cleared)))
            .unwrap();
    assert!(
        reopened
            .slide_snapshot(0)
            .unwrap()
            .transition()
            .unwrap()
            .is_none()
    );
}

#[test]
fn direct_transition_supports_strict_and_noncanonical_prefixes() {
    let strict_slide = shape_xml(STRICT_PML, "before")
        .replace("<p:", "<q:")
        .replace("</p:", "</q:");
    let strict_slide = strict_slide.replace("xmlns:p=", "xmlns:q=");
    let source = fixture_with_presentation_namespace(STRICT_PML, "", strict_slide, false);
    let editor =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(source)))
            .unwrap();
    let requested = direct_transition(TransitionKind::Wipe(Side::Down));
    let mut edit = editor.edit_slide(0).unwrap();
    assert!(edit.set_transition(&requested).unwrap());
    let commit = edit.commit();
    let mut output = Vec::new();
    editor
        .publish_slide_commit_to_stream(&mut output, &commit)
        .unwrap();
    let package = OpcPackage::from_bytes(&output).unwrap();
    let xml = std::str::from_utf8(
        package
            .get_part(&PackURI::new(SLIDE).unwrap())
            .unwrap()
            .blob(),
    )
    .unwrap();
    assert!(xml.contains("<q:transition"));
    assert!(!xml.contains("<p:transition"));
    let reopened =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(output)))
            .unwrap();
    assert!(
        reopened
            .slide_snapshot(0)
            .unwrap()
            .transition()
            .unwrap()
            .unwrap()
            .same_semantics(&requested)
    );
}

#[test]
fn direct_transition_semantic_noop_shares_signed_source_exactly() {
    let slide = slide_with_tail(
        PML,
        "before",
        "<p:transition spd=\"fast\" advClick=\"0\" advTm=\"750\"><p:push dir=\"l\"/></p:transition>",
    );
    let source = fixture("", slide, true);
    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        source.clone(),
    )))
    .unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    let requested = direct_transition(TransitionKind::Push(Side::Left));
    assert!(!edit.set_transition(&requested).unwrap());
    assert!(matches!(
        edit.clear_transition(),
        Err(Error::UnsafeEdit { .. })
    ));
    let commit = edit.commit();
    assert!(!commit.is_changed());
    let replayed = commit.patch().apply(commit.patch().source()).unwrap();
    assert!(!commit.patch().inverse().is_changed());
    assert!(
        commit
            .patch()
            .inverse()
            .apply(&replayed)
            .unwrap()
            .transition()
            .unwrap()
            .is_some()
    );
    let mut output = Vec::new();
    editor
        .publish_slide_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source);

    let editor =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(source)))
            .unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    assert!(
        edit.set_transition(&Transition::new(TransitionKind::Fade { black: None }))
            .unwrap()
    );
    let commit = edit.commit();
    output.clear();
    assert!(matches!(
        editor.publish_slide_commit_to_stream(&mut output, &commit),
        Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty());
}

#[test]
fn direct_transition_refuses_extensions_sound_protection_and_unsafe_targets_atomically() {
    let cases = [
        "<p:transition><p:fade vendor=\"1\"/></p:transition>",
        "<p:transition><p:sndAc><p:stSnd r:embed=\"rIdSound\"/></p:sndAc></p:transition>",
        "<p:transition><p:extLst><p:ext uri=\"urn:vendor\"/></p:extLst></p:transition>",
        "<p:transition><p14:ripple xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\"/></p:transition>",
    ];
    for transition in cases {
        let source = fixture("", slide_with_tail(PML, "before", transition), false);
        let editor =
            SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(source)))
                .unwrap();
        let mut edit = editor.edit_slide(0).unwrap();
        assert!(matches!(
            edit.set_transition(&Transition::new(TransitionKind::Fade { black: None })),
            Err(Error::UnsafeEdit { .. })
        ));
        assert!(matches!(
            edit.clear_transition(),
            Err(Error::UnsafeEdit { .. })
        ));
    }

    let protected = r#"<p:modifyVerifier cryptAlgorithmSid="14" spinCount="1" saltData="AA==" hashData="AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=="/>"#;
    let source = fixture(protected, shape_xml(PML, "before"), false);
    let editor =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(source)))
            .unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    assert!(matches!(
        edit.set_transition(&Transition::new(TransitionKind::Fade { black: None })),
        Err(Error::UnsafeEdit { .. })
    ));
    assert!(!edit.clear_transition().unwrap());

    let source = fixture("", shape_xml(PML, "before"), false);
    let editor =
        SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(source)))
            .unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    let duration =
        Transition::new(TransitionKind::Fade { black: None }).with_duration(Ms::new(900).unwrap());
    assert!(matches!(
        edit.set_transition(&duration),
        Err(Error::UnsafeEdit { .. })
    ));
    let ripple = Transition::new(TransitionKind::Ripple(
        litchi_pptx::transition::Ripple::Center,
    ));
    assert!(matches!(
        edit.set_transition(&ripple),
        Err(Error::UnsafeEdit { .. })
    ));
    assert!(
        edit.set_transition(&Transition::new(TransitionKind::Fade { black: None }))
            .unwrap()
    );
}

#[test]
fn direct_transition_rejects_malformed_dtd_limits_stale_source_and_partial_sink() {
    let duplicate = slide_with_tail(
        PML,
        "before",
        "<p:transition/><p:transition><p:fade/></p:transition>",
    );
    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        fixture("", duplicate, false),
    )))
    .unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    assert!(matches!(edit.clear_transition(), Err(Error::Invalid(_))));

    let dtd = format!("<!DOCTYPE p:sld>{}", shape_xml(PML, "before"));
    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        unchecked_slide_fixture(dtd.as_bytes()),
    )))
    .unwrap();
    assert!(editor.edit_slide(0).is_err());

    let slide = shape_xml(PML, "before").replacen(
        "<p:cSld>",
        &format!("<p:cSld><!--{}-->", "x".repeat(5_000)),
        1,
    );
    let limits = ReadLimits::builder()
        .max_part_bytes(slide.len() as u64)
        .unwrap()
        .build()
        .unwrap();
    let editor = SourceBackedPresentationEditor::from_read_at_with_limits(
        Arc::new(VersionedSource::new(fixture("", slide, false))),
        limits,
    )
    .unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    assert!(matches!(
        edit.set_transition(&Transition::new(TransitionKind::Fade { black: None })),
        Err(Error::Limit { .. })
    ));

    let source_bytes = fixture("", shape_xml(PML, "before"), false);
    let source = Arc::new(VersionedSource::new(source_bytes.clone()));
    let editor = SourceBackedPresentationEditor::from_read_at(source.clone()).unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    edit.set_transition(&Transition::new(TransitionKind::Fade { black: None }))
        .unwrap();
    let commit = edit.commit();
    let foreign = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        fixture("<p:extLst/>", shape_xml(PML, "before"), false),
    )))
    .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        foreign.publish_slide_commit_to_stream(&mut output, &commit),
        Err(Error::StaleSource)
    ));
    assert!(output.is_empty());

    let source = Arc::new(VersionedSource::new(source_bytes));
    let editor = SourceBackedPresentationEditor::from_read_at(source.clone()).unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    edit.set_transition(&Transition::new(TransitionKind::Fade { black: None }))
        .unwrap();
    let commit = edit.commit();
    source.changed();
    assert!(matches!(
        editor.publish_slide_commit_to_stream(&mut output, &commit),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());

    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(VersionedSource::new(
        fixture("", shape_xml(PML, "before"), false),
    )))
    .unwrap();
    let mut edit = editor.edit_slide(0).unwrap();
    edit.set_transition(&Transition::new(TransitionKind::Fade { black: None }))
        .unwrap();
    let commit = edit.commit();
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
