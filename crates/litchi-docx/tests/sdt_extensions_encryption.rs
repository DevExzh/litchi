#![allow(
    clippy::redundant_closure_call,
    reason = "raw OPC test edits stage one closure over the edit's graph"
)]
#![cfg(feature = "encryption")]

use litchi_docx::content_control::{Checksum, Limits, PackageChecksumStatus};
use litchi_docx::custom_xml::NewStore;
use litchi_docx::encryption::Mode;
use litchi_docx::{Error, Package};
use litchi_ooxml_common::custom_xml::Conformance;
use litchi_opc::PackURI;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const HASH: &str = "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash";
const ITEM: &str = "{11111111-1111-4111-8111-111111111111}";

fn store(xml: &[u8]) -> NewStore {
    NewStore {
        xml: xml.to_vec(),
        content_type: "application/xml".to_owned(),
        id: ITEM.to_owned(),
        schemas: vec!["urn:test".to_owned()],
        conformance: Conformance::Transitional,
    }
}

fn encrypted_source(payload: &[u8]) -> Package {
    let checksum = Checksum::compute(payload, &Limits::default())
        .unwrap()
        .to_base64();
    let source = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:h="{HASH}" mc:Ignorable="h"><w:body><w:sdtPr><w:id w:val="1"/><w:dataBinding w:xpath="/x" w:storeItemID="{ITEM}" h:storeItemChecksum="{checksum}"/></w:sdtPr></w:body></w:document>"#
    )
    .into_bytes();
    let mut package = Package::new().unwrap();
    package = {
        let mut edit = package.edit_custom_xml();
        edit.add(store(payload)).unwrap();
        edit.commit().unwrap().into_package()
    };
    let main = PackURI::new("/word/document.xml").unwrap();
    package = {
        let mut edit = package.edit_opc();
        (|opc: &mut litchi_opc::OpcPackage| -> litchi_docx::Result<_> {
            {
                opc.get_part_mut(&main)?.set_blob(source);
                Ok::<_, Error>(())
            }
        })(edit.package_mut())
        .unwrap();
        edit.commit().unwrap().into_package()
    };
    package
}

fn assert_matching(package: &Package, expected: &[u8]) {
    assert_eq!(
        package.custom_xml_by_id(ITEM).unwrap().unwrap().xml(),
        expected
    );
    let reports = package.verify_content_control_checksums().unwrap();
    assert_eq!(reports.len(), 1);
    assert!(matches!(
        reports[0].status(),
        PackageChecksumStatus::Matches
    ));
}

fn assert_encrypted_payload(package: &Package, expected: &[u8]) {
    assert_eq!(
        package.custom_xml_by_id(ITEM).unwrap().unwrap().xml(),
        expected
    );
    assert!(matches!(
            package.verify_content_control_checksums(),
            Err(Error::Refused(refusal)) if refusal.operation() == "verify_content_control_checksums"
    ));
}

#[test]
fn encrypted_set_and_replace_reencrypt_with_current_checksums() {
    let directory = tempfile::tempdir().unwrap();
    let source_path = directory.path().join("source.docx");
    let set_plain_path = directory.path().join("set-plain.docx");
    let set_path = directory.path().join("set.docx");
    let replace_plain_path = directory.path().join("replace-plain.docx");
    let replaced_path = directory.path().join("replaced.docx");
    let final_plain_path = directory.path().join("final-plain.docx");
    let initial = b"<root>encrypted-initial</root>";
    let set_payload = b"<root>encrypted-set</root>";
    let replacement_payload = b"<root>encrypted-replacement</root>";

    let package = encrypted_source(initial);
    package
        .save_encrypted(&source_path, "initial-password", Mode::Agile)
        .unwrap();
    let encrypted_source = std::fs::read(&source_path).unwrap();

    let opened = Package::open_with_password(&source_path, "initial-password").unwrap();
    assert!(matches!(
            {
                let mut edit = opened.edit_custom_xml();
                edit.set(ITEM, set_payload.to_vec()).map(|_| ()).and_then(|()| edit.commit().map(|_| ()))
            },
            Err(Error::Refused(refusal)) if refusal.operation() == "set_custom_xml"
    ));
    assert_encrypted_payload(&opened, initial);
    assert_eq!(std::fs::read(&source_path).unwrap(), encrypted_source);
    opened.save_plain(&set_plain_path).unwrap();

    let mut plaintext = Package::open(&set_plain_path).unwrap();
    plaintext = {
        let mut edit = plaintext.edit_custom_xml();
        edit.set(ITEM, set_payload.to_vec()).unwrap();
        edit.commit().unwrap().into_package()
    };
    assert_matching(&plaintext, set_payload);
    plaintext
        .save_encrypted(&set_path, "set-password", Mode::Agile)
        .unwrap();
    let reopened = Package::open_with_password(&set_path, "set-password").unwrap();
    assert_encrypted_payload(&reopened, set_payload);

    let mut refused_replacement = store(replacement_payload);
    refused_replacement.content_type = "application/vnd.litchi.encrypted+xml".to_owned();
    assert!(matches!(
            {
                let mut edit = reopened.edit_custom_xml();
                edit.replace(ITEM, refused_replacement).map(|_| ()).and_then(|()| edit.commit().map(|_| ()))
            },
            Err(Error::Refused(refusal)) if refusal.operation() == "replace_custom_xml"
    ));
    assert_encrypted_payload(&reopened, set_payload);
    reopened.save_plain(&replace_plain_path).unwrap();

    let mut plaintext = Package::open(&replace_plain_path).unwrap();
    assert_matching(&plaintext, set_payload);
    let mut replacement = store(replacement_payload);
    replacement.content_type = "application/vnd.litchi.encrypted+xml".to_owned();
    replacement
        .schemas
        .push("urn:encrypted:replacement".to_owned());
    plaintext = {
        let mut edit = plaintext.edit_custom_xml();
        edit.replace(ITEM, replacement).unwrap();
        edit.commit().unwrap().into_package()
    };
    assert_matching(&plaintext, replacement_payload);
    plaintext
        .save_encrypted(&replaced_path, "replacement-password", Mode::Agile)
        .unwrap();
    let final_package =
        Package::open_with_password(&replaced_path, "replacement-password").unwrap();
    assert_encrypted_payload(&final_package, replacement_payload);
    final_package.save_plain(&final_plain_path).unwrap();
    let final_plain = Package::open(&final_plain_path).unwrap();
    assert_matching(&final_plain, replacement_payload);
}
