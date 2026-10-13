#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::too_many_lines,
    reason = "integration tests use panic-on-failure extraction"
)]

//! The package-level families refuse a change to a signed document instead of
//! removing its signature (ADR 0006; unified-ops record 0014b).
//!
//! Each family is checked three ways: a change to a signed document is
//! refused with `OpcError::SignedSourceRequiresExplicitPolicy` and leaves the
//! document as it was (it still publishes its signed source byte for byte);
//! an exact no-op keeps the signature byte for byte; and after the explicit
//! removal (`Package::unsign` or `OpcPackage::unsign`) the same change
//! publishes a document without signature parts. Changes to the
//! producer-signed fixtures are stores; removals run against a document that
//! already holds the family's content and is signed in the test.

use std::io::Cursor;
use std::path::PathBuf;

use litchi_docx::metadata::Error as MetadataError;
use litchi_docx::{Error, Package};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part as _, TargetMode};

/// The producer-signed fixtures: signed once and twice by `xmlsec`, and once
/// by Office 2010.
const SIGNED: [&str; 3] = [
    "hello-world-signed.docx",
    "hello-world-signed-twice.docx",
    "ms-office-2010-signed.docx",
];

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/xmldsign")
            .join(name),
    )
    .expect("signed DOCX fixture")
}

fn open(bytes: &[u8]) -> Package {
    Package::from_reader(Cursor::new(bytes)).unwrap()
}

fn bytes(package: &mut Package) -> Vec<u8> {
    package.to_bytes().unwrap()
}

/// `plain` with an OPC signature graph (origin part, one XML signature and
/// their relationships), written out so the signature is an ingressed source
/// signature, as in a file a producer signed.
fn sign(plain: &[u8]) -> Vec<u8> {
    let source = OpcPackage::from_vec(plain.to_vec()).unwrap();
    let mut package = OpcPackage::new();
    for relationship in source.rels().iter() {
        package
            .rels_mut()
            .try_add_relationship(
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )
            .unwrap();
    }
    let names: Vec<_> = source
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    for name in names {
        let part = source.get_part(&name).unwrap();
        let mut copy = BlobPart::new(
            part.partname().clone(),
            part.content_type().to_owned(),
            part.blob().to_vec(),
        );
        for relationship in part.rels().iter() {
            copy.rels_mut()
                .try_add_relationship(
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.r_id().to_owned(),
                    relationship.target_mode(),
                )
                .unwrap();
        }
        package.try_add_part(Box::new(copy)).unwrap();
    }
    let origin = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
    let signature = PackURI::new("/_xmlsignatures/sig1.xml").unwrap();
    let mut origin_part = BlobPart::new(
        origin.clone(),
        litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        Vec::new(),
    );
    origin_part
        .rels_mut()
        .try_add_relationship(
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
                .to_owned(),
            signature.relative_ref(origin.base_uri()),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(origin_part)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            signature,
            litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
            b"<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>".to_vec(),
        )))
        .unwrap();
    package
        .rels_mut()
        .try_add_relationship(
            rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let bytes = PackageWriter::to_bytes(&package).unwrap();
    assert!(open(&bytes).is_signed());
    bytes
}

/// The first fixture with its signature removed explicitly, edited by `edit`,
/// and signed again: a signed document that holds the family's content.
fn signed_with(edit: impl FnOnce(&mut Package)) -> Vec<u8> {
    let mut plain = open(&fixture(SIGNED[0]));
    plain.unsign();
    edit(&mut plain);
    sign(&bytes(&mut plain))
}

/// Whether `bytes` still carries any OPC signature infrastructure.
fn carries_signature(bytes: &[u8]) -> bool {
    let package = OpcPackage::from_vec(bytes.to_vec()).unwrap();
    package.is_signed()
        || package
            .iter_parts()
            .any(|part| part.partname().as_str().starts_with("/_xmlsignatures/"))
}

#[track_caller]
fn assert_signed<T>(result: litchi_docx::Result<T>) {
    let Err(error) = result else {
        panic!("a change to a signed document must be refused");
    };
    assert!(
        matches!(
            error,
            Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)
        ),
        "{error:?}"
    );
}

mod web_settings {
    use super::*;
    use litchi_docx::web::{self, Screen};

    #[test]
    fn a_signed_document_refuses_changes_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = fixture(name);
            let mut package = open(&original);
            let (settings, conformance) = package.web().unwrap().expect("web settings");
            assert!(!package.put_web(settings.clone(), conformance).unwrap());
            assert_eq!(bytes(&mut package), original, "{name}: an equal store");

            let mut changed = settings.clone();
            changed.set_target_screen_size(Screen::Pixels1024x768);
            assert_signed(package.put_web(changed.clone(), conformance));
            assert_signed(package.remove_web());
            assert!(package.is_signed());
            assert_eq!(bytes(&mut package), original, "{name}: refusals");

            package.unsign();
            assert!(package.put_web(changed.clone(), conformance).unwrap());
            let output = bytes(&mut package);
            assert!(!carries_signature(&output), "{name}");
            assert_eq!(open(&output).web().unwrap().unwrap().0, changed);
        }
    }

    #[test]
    fn a_web_settings_patch_refuses_a_signed_package() {
        for name in SIGNED {
            let original = fixture(name);
            let mut package = OpcPackage::from_vec(original.clone()).unwrap();
            let snapshot = web::load_snapshot(&package).unwrap().unwrap();
            let empty = snapshot.edit().commit().unwrap().into_patch();
            web::apply_patch(&mut package, &empty).unwrap();
            let mut edit = snapshot.edit();
            edit.set_target_screen_size(Some(Screen::Pixels800x600));
            let patch = edit.commit().unwrap().into_patch();
            assert_signed(web::apply_patch(&mut package, &patch));
            assert_eq!(
                PackageWriter::to_bytes(&package).unwrap(),
                original,
                "{name}"
            );

            package.unsign();
            web::apply_patch(&mut package, &patch).unwrap();
            let output = PackageWriter::to_bytes(&package).unwrap();
            assert!(!carries_signature(&output), "{name}");
        }
    }
}

mod font_table {
    use super::*;
    use litchi_docx::font::{Conformance, Font};

    #[test]
    fn a_signed_document_refuses_changes_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = fixture(name);
            let mut package = open(&original);
            let table = package.fonts().unwrap().expect("font table");
            assert!(
                !package
                    .put_fonts(table.clone(), Conformance::Transitional)
                    .unwrap()
            );
            assert_eq!(bytes(&mut package), original, "{name}: an equal table");

            let mut changed = table.clone();
            changed
                .add(Font::new("Litchi Signed Sans").unwrap())
                .unwrap();
            assert_signed(package.put_fonts(changed.clone(), Conformance::Transitional));
            assert_signed(package.remove_fonts());
            assert!(package.is_signed());
            assert_eq!(bytes(&mut package), original, "{name}: refusals");

            package.unsign();
            assert!(
                package
                    .put_fonts(changed, Conformance::Transitional)
                    .unwrap()
            );
            let output = bytes(&mut package);
            assert!(!carries_signature(&output), "{name}");
            let reopened = open(&output).fonts().unwrap().unwrap();
            assert!(reopened.get("Litchi Signed Sans").unwrap().is_some());
        }
    }
}

mod glossary_family {
    use super::*;
    use litchi_docx::glossary::{self, Catalog, Conformance, Entry, Transaction};

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn entry(name: &str) -> Entry {
        Entry::new(
            name,
            format!(r#"<w:docPartBody xmlns:w="{W}"><w:p><w:r><w:t>{name}</w:t></w:r></w:p></w:docPartBody>"#)
                .into_bytes(),
        )
        .unwrap()
    }

    fn catalog(names: &[&str]) -> Catalog {
        let mut catalog = Catalog::new();
        for name in names {
            catalog.add(entry(name)).unwrap();
        }
        catalog
    }

    #[test]
    fn a_signed_fixture_refuses_a_new_glossary_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = fixture(name);
            let mut package = open(&original);
            assert!(!package.remove_glossary().unwrap());
            assert!(package.take_glossary_graph().unwrap().is_none());
            assert_signed(package.put_glossary(catalog(&["Keep"]), Conformance::Transitional));
            assert_eq!(bytes(&mut package), original, "{name}");

            package.unsign();
            assert!(
                package
                    .put_glossary(catalog(&["Keep"]), Conformance::Transitional)
                    .unwrap()
            );
            let output = bytes(&mut package);
            assert!(!carries_signature(&output), "{name}");
            assert!(open(&output).glossary().unwrap().is_some());
        }
    }

    #[test]
    fn a_signed_glossary_refuses_changes_removal_and_patches() {
        let signed = signed_with(|plain| {
            assert!(
                plain
                    .put_glossary(catalog(&["Keep"]), Conformance::Transitional)
                    .unwrap()
            );
        });
        let mut package = open(&signed);
        let graph = package.glossary_graph().unwrap().unwrap();
        assert!(!package.put_glossary_graph(&graph).unwrap());
        assert_eq!(bytes(&mut package), signed, "an identical graph");

        let mut changed = graph.clone();
        changed.catalog.add(entry("Changed")).unwrap();
        assert_signed(package.put_glossary_graph(&changed));
        assert_signed(package.take_glossary_graph());
        assert_signed(package.remove_glossary());
        assert_eq!(bytes(&mut package), signed, "refused changes");

        // The transaction and its patch over the OPC graph.
        let mut opc = OpcPackage::from_vec(signed.clone()).unwrap();
        let commit = Transaction::new(&mut opc).unwrap().commit().unwrap();
        assert!(!commit.changed());
        let mut transaction = Transaction::new(&mut opc).unwrap();
        transaction.add_entry(entry("Staged")).unwrap();
        assert_signed(transaction.commit());
        let patch = {
            let mut plain = OpcPackage::from_vec(signed.clone()).unwrap();
            plain.unsign();
            let mut transaction = Transaction::new(&mut plain).unwrap();
            transaction.add_entry(entry("Patched")).unwrap();
            transaction.commit().unwrap().patch().clone()
        };
        assert_signed(patch.apply(&mut opc));
        assert_eq!(PackageWriter::to_bytes(&opc).unwrap(), signed);

        opc.unsign();
        patch.apply(&mut opc).unwrap();
        let output = PackageWriter::to_bytes(&opc).unwrap();
        assert!(!carries_signature(&output));
        let (catalog, _) = glossary::load(&OpcPackage::from_vec(output).unwrap())
            .unwrap()
            .unwrap();
        assert!(catalog.get("Patched").unwrap().is_some());

        package.unsign();
        assert!(package.remove_glossary().unwrap());
        assert!(!carries_signature(&bytes(&mut package)));
    }
}

mod mail_merge_family {
    use super::*;
    use litchi_docx::mail_merge::{Conformance, DataType, Destination, Settings, Source};

    fn settings() -> Settings {
        let mut settings = Settings::new();
        settings
            .set_data_type(Some(DataType::Spreadsheet))
            .set_query(Some("SELECT * FROM [Sheet1$]".into()))
            .set_destination(Destination::Email);
        settings
    }

    fn source() -> Option<Source> {
        Some(Source::External(
            "https://example.invalid/never-fetch.csv".into(),
        ))
    }

    #[test]
    fn a_signed_document_refuses_mail_merge_changes_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = fixture(name);
            let mut package = open(&original);
            assert!(!package.clear_mail_merge().unwrap());
            assert_signed(package.set_mail_merge(
                settings(),
                source(),
                None,
                None,
                Conformance::Transitional,
            ));
            assert_eq!(bytes(&mut package), original, "{name}");

            package.unsign();
            package
                .set_mail_merge(settings(), source(), None, None, Conformance::Transitional)
                .unwrap();
            let output = bytes(&mut package);
            assert!(!carries_signature(&output), "{name}");
            assert!(open(&output).mail_merge_settings().unwrap().is_some());
        }

        let signed = signed_with(|plain| {
            plain
                .set_mail_merge(settings(), source(), None, None, Conformance::Transitional)
                .unwrap();
        });
        let mut package = open(&signed);
        assert_signed(package.clear_mail_merge());
        assert_signed(package.update_mail_merge(
            settings(),
            None,
            None,
            None,
            Conformance::Transitional,
        ));
        assert_eq!(bytes(&mut package), signed);
        package.unsign();
        assert!(package.clear_mail_merge().unwrap());
        assert!(!carries_signature(&bytes(&mut package)));
    }
}

mod ribbon_family {
    use super::*;
    use litchi_ooxml_common::ribbon::{Family, Version};

    const UI_2007: &[u8] =
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#;
    const UI_2007_CHANGED: &[u8] =
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon/></customUI>"#;

    #[test]
    fn a_signed_document_refuses_ribbon_changes_and_keeps_exact_noops() {
        for name in SIGNED {
            let original = fixture(name);
            let mut package = open(&original);
            assert!(!package.remove_ribbon(Family::Legacy).unwrap());
            assert_signed(package.put_ribbon(Version::V2007, UI_2007.to_vec()));
            assert_eq!(bytes(&mut package), original, "{name}");

            package.unsign();
            package
                .put_ribbon(Version::V2007, UI_2007.to_vec())
                .unwrap();
            let output = bytes(&mut package);
            assert!(!carries_signature(&output), "{name}");
            assert_eq!(
                open(&output).ribbon().unwrap().legacy().unwrap().xml(),
                UI_2007
            );
        }

        let signed = signed_with(|plain| {
            plain.put_ribbon(Version::V2007, UI_2007.to_vec()).unwrap();
        });
        let mut package = open(&signed);
        package
            .put_ribbon(Version::V2007, UI_2007.to_vec())
            .unwrap();
        assert!(!package.remove_ribbon(Family::Modern).unwrap());
        assert_eq!(bytes(&mut package), signed, "exact no-ops");
        assert_signed(package.put_ribbon(Version::V2007, UI_2007_CHANGED.to_vec()));
        assert_signed(package.remove_ribbon(Family::Legacy));
        assert_eq!(bytes(&mut package), signed, "refusals");

        package.unsign();
        assert!(package.remove_ribbon(Family::Legacy).unwrap());
        assert!(!carries_signature(&bytes(&mut package)));
    }
}

/// The shared core- and custom-properties owners of `litchi-ooxml-common`
/// refuse a change to a signed document instead of removing the signature
/// (unified-ops record 0014a, F1b): the core-properties transaction refuses
/// at commit with `metadata::Error::Signed`, and a custom-property edit is
/// refused at a managed save with this crate's typed signed error, staying
/// pending so the save succeeds once the signature is removed explicitly.
mod shared_properties {
    use super::*;

    #[track_caller]
    fn assert_save_refused(package: &mut Package) {
        let mut sink = Cursor::new(Vec::new());
        assert_signed(package.write_to(&mut sink));
        assert_signed(package.to_bytes());
        assert_signed(package.to_plain_bytes());
        assert!(package.is_signed());
    }

    #[test]
    fn edited_properties_refuse_a_signed_document_at_save() {
        for name in SIGNED {
            let original = fixture(name);

            // Values equal to the stored ones publish nothing.
            let package = open(&original);
            let props = package.metadata().cloned().expect("core properties");
            let mut edit = package.edit_metadata();
            edit.set(props);
            let commit = edit.commit().expect("exact no-op");
            assert!(!commit.diagnostics().changed());
            let mut package = commit.into_package();
            let _ = package.custom_props_mut();
            assert_eq!(bytes(&mut package), original, "{name}: exact no-ops");

            // A changed title is refused at commit; a new custom property
            // is refused at save.
            let mut package = open(&original);
            let mut edit = package.edit_metadata();
            edit.update(|props| props.title = Some("Signed".into()));
            assert!(
                matches!(edit.commit(), Err(MetadataError::Signed)),
                "{name}"
            );
            assert_eq!(bytes(&mut package), original, "{name}: refusal");
            package.unsign();
            let mut edit = package.edit_metadata();
            edit.update(|props| props.title = Some("Signed".into()));
            let output = bytes(&mut edit.commit().expect("unsigned").into_package());
            assert!(!carries_signature(&output), "{name}");
            assert_eq!(
                open(&output)
                    .metadata()
                    .and_then(|props| props.title.clone()),
                Some("Signed".into())
            );

            let mut package = open(&original);
            package
                .custom_props_mut()
                .insert("Project", "Litchi")
                .unwrap();
            assert_save_refused(&mut package);
            package.unsign();
            let output = bytes(&mut package);
            assert!(!carries_signature(&output), "{name}");
            assert!(open(&output).custom_props().get("Project").is_some());
        }
    }

    #[test]
    fn removing_properties_refuses_a_signed_document_at_save() {
        let signed = signed_with(|plain| {
            plain
                .custom_props_mut()
                .insert("Project", "Litchi")
                .unwrap();
        });
        let mut package = open(&signed);
        let _ = package.custom_props_mut();
        assert_eq!(bytes(&mut package), signed, "an unchanged custom set");

        let mut package = open(&signed);
        assert!(package.custom_props_mut().remove("Project").is_some());
        assert_save_refused(&mut package);
        let mut package = open(&signed);
        let mut edit = package.edit_metadata();
        edit.clear();
        assert!(matches!(edit.commit(), Err(MetadataError::Signed)));
        assert_eq!(bytes(&mut package), signed, "refusal");

        package.unsign();
        let mut edit = package.edit_metadata();
        edit.clear();
        let output = bytes(&mut edit.commit().expect("unsigned").into_package());
        assert!(!carries_signature(&output));
        assert!(open(&output).metadata().is_none());
    }
}
