// Diagnostic only: include as a temporary child module of opened::patch.
// The assertions confirm a known loss, not the desired library contract.
#[test]
fn metadata_replay_probe() {
    use super::{Patch, apply};
    use crate::opened::Limits;
    use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};
    use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

    let initial = crate::Package::new().unwrap().to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&initial, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        let mut bytes = archive.read(name).unwrap();
        if name == "[Content_Types].xml" {
            let mut xml = String::from_utf8(bytes).unwrap();
            let end = xml.rfind("</").unwrap();
            xml.insert_str(end, "<!--CT-SENTINEL-->");
            bytes = xml.into_bytes();
        } else if name == "ppt/_rels/presentation.xml.rels" {
            let mut xml = String::from_utf8(bytes).unwrap();
            let end = xml.rfind("</").unwrap();
            xml.insert_str(end, "<!--RELS-SENTINEL-->");
            bytes = xml.into_bytes();
        }
        writer.write_stored(name, &bytes).unwrap();
    }
    let original = writer.finish().unwrap().into_inner();
    let source = OpcPackage::from_bytes(&original).unwrap();
    let main = PackURI::new("/ppt/presentation.xml").unwrap();
    let added = PackURI::new("/custom/probe.bin").unwrap();
    let content_types = source.source_content_types().unwrap();
    let relationships = source.source_relationships(&main).unwrap();
    assert!(
        content_types
            .bytes()
            .windows(11)
            .any(|w| w == b"CT-SENTINEL")
    );
    assert!(
        relationships
            .bytes()
            .windows(13)
            .any(|w| w == b"RELS-SENTINEL")
    );
    let next_types = content_types
        .with_part_overrides(&[(&added, "application/octet-stream")], 1 << 20)
        .unwrap();
    let next_rels = relationships
        .with_relationship(
            "urn:litchi:probe",
            "../custom/probe.bin",
            "rIdProbe",
            TargetMode::Internal,
            1 << 20,
        )
        .unwrap();
    let mut direct = source.clone();
    direct
        .try_add_parts_with_source_tokens(
            content_types.bytes(),
            &next_types,
            &relationships,
            &next_rels,
            vec![Box::new(BlobPart::new(
                added,
                "application/octet-stream".into(),
                vec![1, 2, 3],
            ))],
        )
        .unwrap();
    let patch = Patch::capture(&source, &direct, main.clone(), Limits::default()).unwrap();
    let patch = Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap();
    let mut replayed = source.clone();
    apply(&mut replayed, &patch, true).unwrap();
    assert_ne!(
        direct.source_content_types().unwrap().bytes(),
        replayed.source_content_types().unwrap().bytes()
    );
    assert_ne!(
        direct.source_relationships(&main).unwrap().bytes(),
        replayed.source_relationships(&main).unwrap().bytes()
    );
    assert_ne!(
        PackageWriter::to_bytes(&direct).unwrap(),
        PackageWriter::to_bytes(&replayed).unwrap()
    );
    apply(&mut replayed, &patch.inverse(), true).unwrap();
    assert_eq!(
        source.source_content_types().unwrap().bytes(),
        replayed.source_content_types().unwrap().bytes()
    );
    assert_eq!(
        source.source_relationships(&main).unwrap().bytes(),
        replayed.source_relationships(&main).unwrap().bytes()
    );
    println!("confirmed: durable forward replay loses metadata XML; inverse restores this fixture");
}
