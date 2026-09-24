use std::io::Cursor;

use litchi_core::Position;
use litchi_docx::ink::{ContextKind, InkEffect, Limits};
use litchi_docx::package::story::StoryKind;
use litchi_docx::{Error, Package};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};

#[path = "ink_inventory/authoring.rs"]
mod authoring;

#[path = "ink_inventory/removal.rs"]
mod removal;

#[path = "ink_inventory/source_guards.rs"]
mod source_guards;

#[path = "ink_inventory/durable.rs"]
mod durable;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SW: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const SR: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const INK_TYPE: &str = "application/inkml+xml";
const INK: &[u8] = br##"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:m="http://schemas.microsoft.com/ink/2010/main" xmlns:e="http://www.w3.org/2003/04/emma"><i:definitions><i:context xml:id="ctx0"/><i:brush xml:id="br0"><i:brushProperty name="inkEffects" value="pencil"/></i:brush></i:definitions><i:traceGroup><i:annotationXML><e:emma><e:interpretation e:mode="ink"><m:context type="writingRegion" semanticType="comment"/></e:interpretation></e:emma></i:annotationXML><i:trace contextRef="#ctx0" brushRef="#br0">1 2, 3 4</i:trace></i:traceGroup></i:ink>"##;

fn part(opc: &mut OpcPackage, name: &str, content_type: &str, data: &[u8]) {
    opc.add_part(Box::new(BlobPart::new(
        PackURI::new(name).unwrap(),
        content_type.to_owned(),
        data.to_vec(),
    )));
}

fn edge(opc: &mut OpcPackage, owner: &str, id: &str, target: &str, kind: &str) {
    opc.get_part_mut(&PackURI::new(owner).unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(kind.to_owned(), target.to_owned(), id.to_owned(), false);
}

fn source(strict: bool, all_stories: bool) -> OpcPackage {
    let (word, rel) = if strict { (SW, SR) } else { (W, R) };
    let mut opc = OpcPackage::new();
    let paragraph = "<u:p><u:r><u:contentPart z:id=\"ink\"/></u:r></u:p>";
    part(&mut opc, "/word/document.xml", ct::WML_DOCUMENT_MAIN,
        format!("<u:document xmlns:u=\"{word}\" xmlns:z=\"{rel}\"><u:body>{paragraph}</u:body></u:document>").as_bytes());
    opc.relate_to("word/document.xml", &format!("{rel}/officeDocument"));
    part(&mut opc, "/payload/handwriting.xml", INK_TYPE, INK);
    edge(
        &mut opc,
        "/word/document.xml",
        "ink",
        "../payload/handwriting.xml",
        &format!("{rel}/customXml"),
    );
    if all_stories {
        for (name, typ, root, contents, relationship) in [
            (
                "a-header",
                ct::WML_HEADER,
                "hdr",
                paragraph.to_owned(),
                "header",
            ),
            (
                "b-header",
                ct::WML_HEADER,
                "hdr",
                paragraph.to_owned(),
                "header",
            ),
            (
                "footer",
                ct::WML_FOOTER,
                "ftr",
                paragraph.to_owned(),
                "footer",
            ),
            (
                "footnotes",
                ct::WML_FOOTNOTES,
                "footnotes",
                format!("<u:footnote u:id=\"1\">{paragraph}</u:footnote>"),
                "footnotes",
            ),
            (
                "endnotes",
                ct::WML_ENDNOTES,
                "endnotes",
                format!("<u:endnote u:id=\"1\">{paragraph}</u:endnote>"),
                "endnotes",
            ),
            (
                "comments",
                ct::WML_COMMENTS,
                "comments",
                format!("<u:comment u:id=\"1\" u:author=\"A\">{paragraph}</u:comment>"),
                "comments",
            ),
            (
                "glossary",
                ct::WML_DOCUMENT_GLOSSARY,
                "glossaryDocument",
                format!(
                    "<u:docParts><u:docPart><u:docPartBody>{paragraph}</u:docPartBody></u:docPart></u:docParts>"
                ),
                "glossaryDocument",
            ),
        ] {
            let uri = format!("/word/{name}.xml");
            part(
                &mut opc,
                &uri,
                typ,
                format!("<u:{root} xmlns:u=\"{word}\" xmlns:z=\"{rel}\">{contents}</u:{root}>")
                    .as_bytes(),
            );
            edge(
                &mut opc,
                "/word/document.xml",
                name,
                &format!("{name}.xml"),
                &format!("{rel}/{relationship}"),
            );
            // Nonstandard names and differently cased target references must
            // resolve through the owning story and canonical package identity.
            edge(
                &mut opc,
                &uri,
                "ink",
                "../PAYLOAD/handwriting.xml",
                &format!("{rel}/customXml"),
            );
        }
    }
    opc
}

fn open(opc: &OpcPackage) -> Package {
    Package::from_reader(Cursor::new(PackageWriter::to_bytes(opc).unwrap())).unwrap()
}

fn save(package: &mut Package) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    package.to_stream(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn all_story_owners_share_one_payload_and_keep_signed_source_exact() {
    for strict in [false, true] {
        let mut opc = source(strict, true);
        part(
            &mut opc,
            "/_xmlsignatures/origin.sigs",
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
            b"",
        );
        opc.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        let original = PackageWriter::to_bytes(&opc).unwrap();
        let mut package = Package::from_reader(Cursor::new(original.as_slice())).unwrap();
        assert_eq!(
            package
                .story_inventory()
                .unwrap()
                .stories()
                .iter()
                .map(|story| story.part().as_str())
                .collect::<Vec<_>>(),
            [
                "/word/document.xml",
                "/word/a-header.xml",
                "/word/b-header.xml",
                "/word/comments.xml",
                "/word/endnotes.xml",
                "/word/footer.xml",
                "/word/footnotes.xml",
                "/word/glossary.xml",
            ]
        );
        let limits = Limits {
            max_total_payload_bytes: INK.len(),
            ..Limits::default()
        };
        let snapshot = package.ink_with_limits(limits).unwrap();
        assert_eq!(snapshot.annotations().len(), 8);
        assert_eq!(snapshot.distinct_payload_count(), 1);
        assert!(snapshot.get(Position::new(8)).is_none());
        let main = snapshot.get(Position::new(0)).unwrap();
        assert_eq!(main.location().kind(), StoryKind::Main);
        assert_eq!(main.trace_count(), 1);
        assert!(matches!(
            main.contexts().next().unwrap().kind(),
            ContextKind::WritingRegion
        ));
        assert!(matches!(
            main.brush_properties().next().unwrap().ink_effect(),
            Some(InkEffect::Pencil)
        ));
        let headers = snapshot
            .annotations()
            .iter()
            .filter(|item| item.location().kind() == StoryKind::Header)
            .collect::<Vec<_>>();
        assert_eq!(headers.len(), 2);
        assert_eq!(headers[0].location().position(), Position::new(0));
        assert_eq!(headers[1].location().position(), Position::new(1));
        assert!(package.is_signed());
        assert_eq!(save(&mut package), original);
        let reopened = Package::from_reader(Cursor::new(original)).unwrap();
        assert_eq!(reopened.ink().unwrap().annotations().len(), 8);
        drop(package);
        assert_eq!(
            snapshot
                .clone()
                .get(Position::new(0))
                .unwrap()
                .trace_count(),
            1
        );
    }
}

#[test]
fn base_text_xml_ink_is_supported_and_generic_content_is_not_projected() {
    let mut opc = source(false, false);
    part(&mut opc, "/payload/handwriting.xml", "text/xml", INK);
    assert_eq!(open(&opc).ink().unwrap().annotations().len(), 1);
    part(
        &mut opc,
        "/payload/handwriting.xml",
        "text/xml",
        b"<math xmlns=\"urn:other\"/>",
    );
    assert!(open(&opc).ink().unwrap().annotations().is_empty());
    part(
        &mut opc,
        "/payload/handwriting.xml",
        "application/mathml+xml",
        b"<math xmlns=\"http://www.w3.org/1998/Math/MathML\"/>",
    );
    assert!(open(&opc).ink().unwrap().annotations().is_empty());
}

#[test]
fn drawing_ink_uses_its_explicit_content_type_and_preserves_vml_fallback() {
    let mut opc = source(false, false);
    let main = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" mc:Ignorable="wi w14"><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="wi"><w:drawing><wp:inline><wp:extent cx="127000" cy="127000"/><wp:docPr id="1" name="Ink"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"><w14:contentPart r:id="ink"/></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="fallback" style="width:10pt;height:10pt"><v:imagedata r:id="fallbackImage"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p></w:body></w:document>"#
    );
    opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap()
        .set_blob(main.as_bytes().to_vec());
    part(
        &mut opc,
        "/payload/fallback.png",
        "image/png",
        b"opaque fallback image",
    );
    edge(
        &mut opc,
        "/word/document.xml",
        "fallbackImage",
        "../payload/fallback.png",
        rt::IMAGE,
    );
    let original = PackageWriter::to_bytes(&opc).unwrap();
    let mut package = Package::from_reader(Cursor::new(original.as_slice())).unwrap();
    assert_eq!(package.ink().unwrap().annotations().len(), 1);
    assert_eq!(save(&mut package), original);
    part(&mut opc, "/payload/handwriting.xml", "text/xml", INK);
    assert!(open(&opc).ink().is_err());

    let canvas = main.replace("wordprocessingInk", "wordprocessingCanvas").replace(
        "<w14:contentPart r:id=\"ink\"/>",
        "<c:wpc xmlns:c=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\"><w14:contentPart r:id=\"ink\"/></c:wpc>",
    );
    opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap()
        .set_blob(canvas.into_bytes());
    part(
        &mut opc,
        "/payload/handwriting.xml",
        "application/mathml+xml",
        b"<math xmlns=\"http://www.w3.org/1998/Math/MathML\"/>",
    );
    assert!(open(&opc).ink().unwrap().annotations().is_empty());
    part(&mut opc, "/payload/handwriting.xml", INK_TYPE, INK);
    assert_eq!(open(&opc).ink().unwrap().annotations().len(), 1);
}

#[test]
fn malformed_ink_and_relationships_fail_without_changing_source() {
    for payload in [
        b"<ink/>".as_slice(),
        b"<inkml:ink/>".as_slice(),
        b"<i:ink xmlns:i=\"urn:foreign\"/>".as_slice(),
        b"<i:ink xmlns:i=\"http://www.w3.org/2003/InkML\">&unknown;</i:ink>".as_slice(),
    ] {
        let mut opc = source(false, false);
        part(&mut opc, "/payload/handwriting.xml", INK_TYPE, payload);
        // Preserve malformed/unsupported input by building its archive directly
        // through an existing source fixture wrapper below.
        let archive = raw_archive(&opc);
        let mut package = Package::from_reader(Cursor::new(archive.as_slice())).unwrap();
        assert!(package.ink().is_err());
        assert_eq!(save(&mut package), archive);
    }
    let mut opc = source(false, false);
    opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap()
        .rels_mut()
        .remove("ink");
    assert!(open(&opc).ink().is_err());
    edge(
        &mut opc,
        "/word/document.xml",
        "ink",
        "../payload/handwriting.xml",
        rt::IMAGE,
    );
    assert!(open(&opc).ink().is_err());
    let mut opc = source(false, false);
    part(&mut opc, "/payload/missing.xml", "text/xml", b"<other/>");
    edge(
        &mut opc,
        "/payload/handwriting.xml",
        "sidecar",
        "missing.xml",
        rt::CUSTOM_XML,
    );
    assert!(open(&opc).ink().is_err());
}

#[test]
fn distinct_payload_budget_and_external_targets_are_checked() {
    let mut opc = source(false, false);
    let main = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:r><w:contentPart r:id="ink"/><w:contentPart r:id="second"/></w:r></w:p></w:body></w:document>"#
    );
    opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap()
        .set_blob(main.into_bytes());
    part(&mut opc, "/payload/second.xml", INK_TYPE, INK);
    edge(
        &mut opc,
        "/word/document.xml",
        "second",
        "../payload/second.xml",
        rt::CUSTOM_XML,
    );
    let package = open(&opc);
    assert_eq!(package.ink().unwrap().distinct_payload_count(), 2);
    assert!(matches!(
        package.ink_with_limits(Limits {
            max_total_payload_bytes: INK.len(),
            ..Limits::default()
        }),
        Err(Error::InkLimit {
            resource: "total payload bytes",
            ..
        })
    ));

    let mut opc = source(false, false);
    let rels = opc
        .get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap()
        .rels_mut();
    rels.remove("ink");
    rels.add_relationship(
        rt::CUSTOM_XML.into(),
        "https://never-fetch.invalid/ink.xml".into(),
        "ink".into(),
        true,
    );
    assert!(matches!(
        open(&opc).ink(),
        Err(Error::InvalidRelationship(_))
    ));
}

#[test]
fn inactive_unknown_choice_is_not_an_annotation_or_semantic_error() {
    let mut opc = source(false, false);
    let main = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:f="urn:future"><w:body><w:p><w:r><mc:AlternateContent><mc:Choice Requires="f"><w:contentPart/></mc:Choice><mc:Fallback><w:contentPart r:id="ink"/></mc:Fallback></mc:AlternateContent></w:r></w:p></w:body></w:document>"#
    );
    opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap()
        .set_blob(main.into_bytes());
    let mut package = open(&opc);
    let before = save(&mut package);
    assert_eq!(
        package
            .ink_with_limits(Limits {
                max_annotations: 1,
                ..Limits::default()
            })
            .unwrap()
            .annotations()
            .len(),
        1
    );
    assert_eq!(save(&mut package), before);
}

fn raw_archive(opc: &OpcPackage) -> Vec<u8> {
    use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
    let valid = PackageWriter::to_bytes(&source(false, false)).unwrap();
    let archive = ArchiveReader::new(&valid).unwrap();
    let payload = opc
        .get_part(&PackURI::new("/payload/handwriting.xml").unwrap())
        .unwrap()
        .blob();
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        if name == "payload/handwriting.xml" {
            writer.write_stored(name, payload).unwrap();
        } else {
            writer
                .write_stored(name, &archive.read(name).unwrap())
                .unwrap();
        }
    }
    writer.finish_to_bytes().unwrap()
}

#[test]
fn payload_annotation_relationship_and_dirty_state_limits_are_enforced() {
    let package = open(&source(false, true));
    for limits in [
        Limits {
            max_annotations: 7,
            ..Limits::default()
        },
        Limits {
            max_payload_bytes: INK.len() - 1,
            ..Limits::default()
        },
        Limits {
            max_total_payload_bytes: INK.len() - 1,
            ..Limits::default()
        },
        Limits {
            max_relationships: 1,
            ..Limits::default()
        },
        Limits {
            max_xml_depth: 2,
            ..Limits::default()
        },
        Limits {
            max_annotations: 0,
            ..Limits::default()
        },
    ] {
        assert!(package.ink_with_limits(limits).is_err());
    }
    let mut package = Package::new().unwrap();
    package
        .document_mut()
        .unwrap()
        .add_paragraph()
        .add_run_with_text("pending");
    assert!(matches!(package.ink(), Err(Error::UnsafeEdit { .. })));
}

fn managed_source(
    bytes: Vec<u8>,
) -> (
    litchi_core::Budget,
    litchi_core::CancellationSource,
    litchi_docx::source_backed::Package,
) {
    let memory = 1024 * 1024;
    let budget = litchi_core::Budget::root(
        "docx-ink-pinned",
        litchi_core::Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation, token) = litchi_core::CancellationSource::pair();
    let execution = litchi_core::ExecutionContext::new(
        budget.clone(),
        token,
        litchi_core::ExecutionLimits::new(
            std::num::NonZeroUsize::MIN,
            std::num::NonZeroUsize::MIN,
            std::num::NonZeroU64::new(memory).unwrap(),
            0,
        )
        .unwrap(),
    );
    let package =
        litchi_docx::source_backed::Package::from_read_at_with_cache_limits_and_execution_context(
            std::sync::Arc::new(litchi_core::OwnedSource::new(bytes)),
            litchi_opc::SourceCacheLimits::new(1, 1).unwrap(),
            execution,
        )
        .unwrap();
    (budget, cancellation, package)
}

#[test]
fn source_backed_inventory_matches_owned_and_pins_managed_payload_once() {
    for strict in [false, true] {
        let mut opc = source(strict, true);
        // No Ink query should materialize either this unrelated large payload
        // or the orphan declared Ink part. These are not active annotations.
        part(
            &mut opc,
            "/unrelated.bin",
            "application/octet-stream",
            &[7; 64 * 1024],
        );
        part(&mut opc, "/orphan.xml", INK_TYPE, INK);
        // A glossary may itself own a header; the owner is resolved through
        // relationships rather than hard-coded main-document paths.
        opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .rels_mut()
            .remove("b-header")
            .unwrap();
        edge(
            &mut opc,
            "/word/glossary.xml",
            "header",
            "b-header.xml",
            &format!("{}/header", if strict { SR } else { R }),
        );
        let expected = open(&opc).ink().unwrap();
        let (budget, _cancellation, package) =
            managed_source(PackageWriter::to_bytes(&opc).unwrap());
        assert_eq!(package.cache_diagnostics().cold_loads, 0);
        assert_eq!(budget.used(litchi_core::Resource::Memory), 0);
        let snapshot = package.ink().unwrap();
        assert_eq!(snapshot.distinct_payload_count(), 1);
        assert_eq!(snapshot.annotations().len(), 8);
        assert_eq!(package.cache_diagnostics().successful_loads, 9);
        assert_eq!(package.cache_diagnostics().retained_entries, 0);
        assert_eq!(budget.used(litchi_core::Resource::Memory), INK.len() as u64);
        for (actual, expected) in snapshot.annotations().iter().zip(expected.annotations()) {
            assert_eq!(actual.location(), expected.location());
            assert_eq!(actual.trace_count(), expected.trace_count());
            assert_eq!(
                actual.contexts().next().unwrap().kind(),
                expected.contexts().next().unwrap().kind()
            );
            assert_eq!(
                actual.brush_properties().next().unwrap().ink_effect(),
                Some(InkEffect::Pencil)
            );
        }
        let retained = snapshot.annotations()[0].clone();
        drop(snapshot);
        drop(package);
        assert_eq!(retained.trace_count(), 1);
        assert_eq!(budget.used(litchi_core::Resource::Memory), INK.len() as u64);
        drop(retained);
        assert_eq!(budget.used(litchi_core::Resource::Memory), 0);
    }
}

#[test]
fn source_backed_payload_limit_refuses_before_decompression_and_releases_story_pins() {
    let opc = source(false, true);
    let (budget, cancellation, package) = managed_source(PackageWriter::to_bytes(&opc).unwrap());
    assert!(matches!(
        package.ink_with_limits(Limits {
            max_payload_bytes: INK.len() - 1,
            ..Limits::default()
        }),
        Err(Error::InkLimit {
            resource: "payload bytes",
            ..
        })
    ));
    assert_eq!(package.cache_diagnostics().successful_loads, 8);
    assert_eq!(budget.used(litchi_core::Resource::Memory), 0);
    let snapshot = package.ink().unwrap();
    cancellation.cancel();
    assert!(package.ink().is_err());
    assert_eq!(snapshot.annotations()[0].trace_count(), 1);
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(litchi_core::Resource::Memory), 0);
}

#[test]
fn source_backed_failed_payload_parse_does_not_retain_managed_bytes() {
    let mut opc = source(false, false);
    part(&mut opc, "/payload/handwriting.xml", INK_TYPE, b"<wrong/>");
    let (budget, _cancellation, package) = managed_source(PackageWriter::to_bytes(&opc).unwrap());
    assert!(package.ink().is_err());
    assert_eq!(package.cache_diagnostics().successful_loads, 2);
    assert_eq!(budget.used(litchi_core::Resource::Memory), 0);
}

struct ChangingInkSource {
    bytes: litchi_core::OwnedSource,
    revision: std::sync::atomic::AtomicU64,
    change_on_read: std::sync::atomic::AtomicBool,
}

impl litchi_core::ReadAt for ChangingInkSource {
    fn len(&self) -> std::io::Result<u64> {
        litchi_core::ReadAt::len(&self.bytes)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let read = litchi_core::ReadAt::read_at(&self.bytes, offset, output)?;
        if self
            .change_on_read
            .swap(false, std::sync::atomic::Ordering::SeqCst)
        {
            self.revision
                .fetch_add(1, std::sync::atomic::Ordering::SeqCst);
        }
        Ok(read)
    }

    fn version(&self) -> std::io::Result<litchi_core::SourceVersion> {
        Ok(litchi_core::SourceVersion::new(
            119,
            self.revision.load(std::sync::atomic::Ordering::SeqCst),
        ))
    }
}

#[test]
fn source_backed_inventory_checks_freshness_on_cached_and_cold_reads() {
    for during_read in [false, true] {
        let versioned = std::sync::Arc::new(ChangingInkSource {
            bytes: litchi_core::OwnedSource::new(
                PackageWriter::to_bytes(&source(false, false)).unwrap(),
            ),
            revision: std::sync::atomic::AtomicU64::new(0),
            change_on_read: std::sync::atomic::AtomicBool::new(false),
        });
        let package = litchi_docx::source_backed::Package::from_read_at(versioned.clone()).unwrap();
        let retained = if during_read {
            versioned
                .change_on_read
                .store(true, std::sync::atomic::Ordering::SeqCst);
            None
        } else {
            let snapshot = package.ink().unwrap();
            versioned
                .revision
                .fetch_add(1, std::sync::atomic::Ordering::SeqCst);
            Some(snapshot)
        };
        assert!(matches!(
            package.ink(),
            Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
        ));
        if let Some(snapshot) = retained {
            assert_eq!(snapshot.annotations()[0].trace_count(), 1);
        }
    }
}

#[test]
fn source_backed_story_graph_and_aggregate_limits_fail_before_payload_reads() {
    for scenario in 0..6 {
        let mut opc = source(false, true);
        match scenario {
            0 => {
                opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())
                    .unwrap()
                    .rels_mut()
                    .remove("a-header")
                    .unwrap();
            },
            1 => edge(
                &mut opc,
                "/word/document.xml",
                "ambiguous",
                "a-header.xml",
                rt::HEADER,
            ),
            2 => edge(
                &mut opc,
                "/word/a-header.xml",
                "nested",
                "footer.xml",
                rt::FOOTER,
            ),
            _ => {},
        }
        let (budget, _cancellation, package) =
            managed_source(PackageWriter::to_bytes(&opc).unwrap());
        let mut limits = Limits::default();
        match scenario {
            3 => limits.stories.max_total_story_bytes = 1,
            4 => limits.stories.max_topology_bytes = 1,
            5 => limits.max_relationships = 1,
            _ => {},
        }
        assert!(
            package.ink_with_limits(limits).is_err(),
            "scenario {scenario}"
        );
        assert_eq!(
            package.cache_diagnostics().cold_loads,
            0,
            "scenario {scenario}"
        );
        assert_eq!(budget.used(litchi_core::Resource::Memory), 0);
    }
}

#[test]
fn source_session_keeps_distinct_story_payloads_correct_across_cold_and_cached_reads() {
    for strict in [false, true] {
        let mut opc = source(strict, true);
        let owners = [
            "/word/document.xml",
            "/word/a-header.xml",
            "/word/b-header.xml",
            "/word/comments.xml",
            "/word/endnotes.xml",
            "/word/footer.xml",
            "/word/footnotes.xml",
            "/word/glossary.xml",
        ];
        for (index, owner) in owners.iter().enumerate() {
            let traces =
                "<i:trace contextRef=\"#ctx0\" brushRef=\"#br0\">1 2</i:trace>".repeat(index + 1);
            let payload = format!(
                "<i:ink xmlns:i=\"http://www.w3.org/2003/InkML\"><i:definitions><i:context xml:id=\"ctx0\"/><i:brush xml:id=\"br0\"/></i:definitions>{traces}</i:ink>"
            );
            let target = format!("/payload/ink{index}.xml");
            part(&mut opc, &target, INK_TYPE, payload.as_bytes());
            opc.get_part_mut(&PackURI::new(*owner).unwrap())
                .unwrap()
                .rels_mut()
                .remove("ink")
                .unwrap();
            edge(
                &mut opc,
                owner,
                "ink",
                &format!("../payload/ink{index}.xml"),
                &format!("{}/customXml", if strict { SR } else { R }),
            );
        }
        let expected = open(&opc).ink().unwrap();
        let package = litchi_docx::source_backed::Package::from_read_at(std::sync::Arc::new(
            litchi_core::OwnedSource::new(PackageWriter::to_bytes(&opc).unwrap()),
        ))
        .unwrap();
        let first = package.ink().unwrap();
        assert_eq!(package.cache_diagnostics().successful_loads, 16);
        let second = package.ink().unwrap();
        assert_eq!(package.cache_diagnostics().successful_loads, 16);
        drop(package);
        for snapshot in [first, second] {
            assert_eq!(snapshot.distinct_payload_count(), 8);
            assert_eq!(snapshot.annotations().len(), 8);
            for (actual, expected) in snapshot.annotations().iter().zip(expected.annotations()) {
                assert_eq!(actual.location(), expected.location());
                assert_eq!(actual.trace_count(), expected.trace_count());
            }
        }
    }
}
