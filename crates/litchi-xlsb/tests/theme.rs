#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "theme fixtures deliberately fail fast while asserting package contracts"
)]

//! Integration coverage for the XLSB Theme owner ([MS-XLSB] 2.1.7.52).
//!
//! The package graph is owned by XLSB, while the typed `a:theme` payload is
//! decoded by the shared DrawingML codec.  These tests keep those boundaries
//! visible: a real XLSB fixture must resolve one internal workbook-owned
//! theme, typed scheme replacement must retain untouched XML, and malformed or
//! oversized payloads must be refused before they can be published.

use litchi_drawingml::theme::{Color, Face, FontSet, Palette, Slot, codec};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, PackURI, TargetMode};
use litchi_xlsb::Workbook;
use litchi_xlsb::theme::{
    Commit as ThemeCommit, Limits as ThemeLimits, Snapshot as ThemeSnapshot, Theme,
};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
use std::io::Cursor;
use std::path::PathBuf;

const FIXTURE: &str = "test-data/poi/test-data/spreadsheet/testVarious.xlsb";
const THEME_NAME: &str = "/xl/theme/theme1.xml";
const WORKBOOK_NAME: &str = "/xl/workbook.bin";
const CORE_PROPERTIES_NAME: &str = "/docProps/core.xml";
const STRICT_THEME_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/theme";

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn native_workbook() -> Workbook {
    Workbook::new(Cursor::new(
        std::fs::read(fixture(FIXTURE)).expect("native XLSB fixture"),
    ))
    .expect("native XLSB package")
}

fn theme_uri() -> PackURI {
    PackURI::new(THEME_NAME).expect("theme URI")
}

fn workbook_uri() -> PackURI {
    PackURI::new(WORKBOOK_NAME).expect("workbook URI")
}

fn theme_bytes(workbook: &Workbook) -> Vec<u8> {
    workbook
        .opc_package()
        .get_part(&theme_uri())
        .expect("theme part")
        .blob()
        .to_vec()
}

fn relationship_fingerprint(
    workbook: &Workbook,
    owner: &PackURI,
) -> Vec<(String, String, String, TargetMode)> {
    let mut relationships = workbook
        .opc_package()
        .get_part(owner)
        .expect("relationship owner part")
        .rels()
        .iter()
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.target_mode(),
            )
        })
        .collect::<Vec<_>>();
    relationships.sort_by(|left, right| left.0.cmp(&right.0));
    relationships
}

fn relationship_xml(workbook: &Workbook, owner: &PackURI) -> Vec<u8> {
    workbook
        .opc_package()
        .source_relationships(owner)
        .expect("capture relationship XML")
        .bytes()
        .to_vec()
}

fn replace_theme_relationship_source_lexically(package: &mut litchi_opc::OpcPackage) {
    let owner = theme_uri();
    let source = package
        .source_relationships(&owner)
        .expect("capture Theme relationship source");
    let replacement = ["rIdInternalThemeImage", "rIdExternalThemeImage"]
        .into_iter()
        .find_map(|id| {
            let (reltype, target, mode) = {
                let relationship = package
                    .get_part(&owner)
                    .expect("Theme part")
                    .rels()
                    .get(id)?;
                (
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.target_mode(),
                )
            };
            let candidate = source
                .without_relationship(id, usize::MAX)
                .ok()?
                .with_relationship(&reltype, &target, id, mode, usize::MAX)
                .ok()?;
            (candidate.bytes() != source.bytes()).then_some(candidate)
        })
        .expect("reinsert one Theme edge with changed lexical bytes");
    package
        .try_replace_relationships(&source, &replacement)
        .expect("replace Theme relationship source token");
}

fn all_rgb_palette(name: &str, value: &str) -> Palette {
    Slot::ALL
        .into_iter()
        .fold(Palette::new(name), |palette, slot| {
            palette.with(slot, Color::rgb(value).expect("valid RGB color"))
        })
}

fn aptos_fonts() -> FontSet {
    FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"))
}

fn authored_theme_xml() -> Vec<u8> {
    codec::encode_part(
        "Office",
        &all_rgb_palette("Office", "4472C4"),
        &aptos_fonts(),
    )
    .expect("authored theme XML")
}

fn custom_theme() -> Theme {
    Theme {
        name: "Litchi Custom Theme".to_owned(),
        colors: all_rgb_palette("Litchi Custom", "112233"),
        fonts: FontSet::new(
            "Litchi Custom",
            Face::new("Aptos Display"),
            Face::new("Aptos"),
        ),
    }
}

fn replace_theme_blob(workbook: &mut Workbook, replacement: Vec<u8>) {
    workbook
        .edit_opc(|opc| {
            opc.get_part_mut(&theme_uri())
                .expect("theme part")
                .set_blob(replacement);
            Ok(())
        })
        .expect("reopen workbook with theme");
}

fn saved_bytes(workbook: &Workbook) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).expect("save workbook");
    output.into_inner()
}

fn typed_theme(workbook: &Workbook) -> ThemeSnapshot {
    workbook
        .theme()
        .expect("read workbook-owned theme")
        .expect("native workbook theme")
}

fn authored_palette(theme: &ThemeSnapshot, value: &str) -> Palette {
    let mut palette = theme.theme().colors.clone();
    for slot in Slot::ALL {
        palette = palette.with(slot, Color::rgb(value).expect("valid RGB color"));
    }
    palette
}

fn changed_commit(theme: &ThemeSnapshot, value: &str) -> ThemeCommit {
    let mut edit = theme.edit();
    edit.set_palette(authored_palette(theme, value))
        .expect("stage typed palette");
    edit.commit().expect("commit typed palette")
}

fn package_with(edit: impl FnOnce(&mut litchi_opc::OpcPackage)) -> Workbook {
    let mut package = native_workbook().opc_package().clone();
    edit(&mut package);
    Workbook::from_opc_package(package).expect("host package graph remains readable")
}

fn image_bearing_theme_workbook() -> (Workbook, PackURI) {
    let image_uri = PackURI::new("/xl/media/theme-owner.png").expect("image URI");
    let workbook = package_with(|package| {
        package.add_part(Box::new(BlobPart::new(
            image_uri.clone(),
            ct::PNG.to_owned(),
            vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A],
        )));
        package
            .get_part_mut(&theme_uri())
            .expect("theme part")
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "../media/theme-owner.png".to_owned(),
                "rIdInternalThemeImage".to_owned(),
                TargetMode::Internal,
            )
            .expect("internal Theme image relationship");
        package
            .get_part_mut(&theme_uri())
            .expect("theme part")
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "https://example.invalid/theme-owner.png".to_owned(),
                "rIdExternalThemeImage".to_owned(),
                TargetMode::External,
            )
            .expect("external Theme image relationship");
    });
    (workbook, image_uri)
}

fn remove_workbook_theme(mut package: litchi_opc::OpcPackage) -> litchi_opc::OpcPackage {
    let relationship_id = package
        .get_part(&workbook_uri())
        .expect("workbook part")
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::THEME)
        .expect("workbook theme relationship")
        .r_id()
        .to_owned();
    package
        .get_part_mut(&workbook_uri())
        .expect("workbook part")
        .rels_mut()
        .remove(&relationship_id);
    assert!(package.remove_part(&theme_uri()));
    package
}

fn replace_workbook_theme_relationship(
    package: &mut litchi_opc::OpcPackage,
    relationship_type: &str,
) {
    let (relationship_id, target_ref) = package
        .get_part(&workbook_uri())
        .expect("workbook part")
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::THEME)
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.target_ref().to_owned(),
            )
        })
        .expect("workbook theme relationship");
    let workbook = package
        .get_part_mut(&workbook_uri())
        .expect("workbook part");
    workbook.rels_mut().remove(&relationship_id);
    workbook
        .rels_mut()
        .try_add_relationship(
            relationship_type.to_owned(),
            target_ref,
            relationship_id,
            TargetMode::Internal,
        )
        .expect("replacement theme relationship");
}

fn strict_theme_xml() -> Vec<u8> {
    let transitional = authored_theme_xml();
    String::from_utf8(transitional)
        .expect("authored UTF-8")
        .replace(codec::NAMESPACE, codec::STRICT_NAMESPACE)
        .into_bytes()
}

fn append_opaque_extension(source: &[u8]) -> Vec<u8> {
    let root_start = source
        .windows(b"<a:theme".len())
        .position(|window| window == b"<a:theme")
        .expect("theme root opening tag");
    let root_end = root_start
        + source[root_start..]
            .iter()
            .position(|byte| *byte == b'>')
            .expect("theme root opening tag");
    let closing = b"</a:theme>";
    let closing_start = source
        .windows(closing.len())
        .rposition(|window| window == closing)
        .expect("theme root closing tag");
    let mut output = Vec::with_capacity(source.len() + 80);
    output.extend_from_slice(&source[..root_end]);
    output.extend_from_slice(br#" xmlns:x="urn:litchi:theme-test">"#);
    output.extend_from_slice(&source[root_end + 1..closing_start]);
    output.extend_from_slice(br#"<x:opaque marker="keep"/>"#);
    output.extend_from_slice(&source[closing_start..]);
    let scheme_start = output
        .windows(b"<a:clrScheme".len())
        .position(|window| window == b"<a:clrScheme")
        .expect("theme color scheme");
    let scheme_end = scheme_start + b"<a:clrScheme".len();
    output.splice(
        scheme_end..scheme_end,
        br#" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#
            .iter()
            .copied(),
    );
    output
}

#[test]
fn native_fixture_theme_is_typed_and_workbook_owned() {
    let workbook = native_workbook();
    let theme = workbook
        .opc_package()
        .get_part(&theme_uri())
        .expect("native theme part");
    assert_eq!(theme.content_type(), ct::OFC_THEME);
    assert!(theme.rels().is_empty(), "theme is a leaf part");

    let workbook_part = workbook
        .opc_package()
        .get_part(&workbook_uri())
        .expect("native workbook part");
    let theme_relationships = workbook_part
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == rt::THEME)
        .collect::<Vec<_>>();
    assert_eq!(theme_relationships.len(), 1);
    let relationship = theme_relationships[0];
    assert_eq!(relationship.target_mode(), TargetMode::Internal);
    assert!(!relationship.is_external());
    assert_eq!(
        relationship.target_partname().expect("theme target"),
        theme_uri()
    );

    let parsed = codec::read(theme.blob()).expect("typed native theme");
    assert!(!parsed.name.is_empty());
    assert_eq!(
        parsed.colors.color(Slot::Accent1),
        Some(&Color::Rgb("5B9BD5".to_owned()))
    );
    assert_eq!(parsed.fonts.major().latin, "Calibri Light");
    assert_eq!(parsed.fonts.minor().latin, "Calibri");

    let snapshot = typed_theme(&workbook);
    assert_eq!(snapshot.source_xml(), theme.blob());
    assert_eq!(snapshot.theme().name, "Office Theme");
}

#[test]
fn authored_typed_theme_survives_save_and_reopen() {
    let mut workbook = native_workbook();
    replace_theme_blob(&mut workbook, authored_theme_xml());
    let before = typed_theme(&workbook);
    let mut edit = workbook.edit_theme().expect("start workbook theme edit");
    edit.set_name("Litchi Authored Theme")
        .expect("stage authored display name");
    edit.set_palette(authored_palette(&before, "4472C4"))
        .expect("stage authored palette");
    let commit = edit.commit().expect("commit authored theme");
    assert!(commit.changed());
    let expected = commit.snapshot().source_xml().to_vec();
    workbook
        .apply_theme(&commit)
        .expect("publish authored theme");

    let saved = saved_bytes(&workbook);
    let reopened = Workbook::new(Cursor::new(saved)).expect("reopen saved package");
    assert_eq!(theme_bytes(&reopened), expected);
    let reopened_theme = typed_theme(&reopened);
    assert_eq!(reopened_theme.theme().name, "Litchi Authored Theme");
    assert_eq!(
        reopened_theme.theme().colors.color(Slot::Accent1),
        Some(&Color::Rgb("4472C4".to_owned()))
    );
}

#[test]
fn typed_change_retains_opaque_xml_and_inverse_is_exact() {
    let mut workbook = native_workbook();
    let original = authored_theme_xml();
    let extended = append_opaque_extension(&original);
    replace_theme_blob(&mut workbook, extended.clone());
    let before = typed_theme(&workbook);
    assert!(
        before
            .source_xml()
            .windows(b"x:opaque marker=\"keep\"".len())
            .any(|window| window == b"x:opaque marker=\"keep\"")
    );

    let commit = changed_commit(&before, "FF0000");
    workbook
        .apply_theme(&commit)
        .expect("publish typed theme change");
    let changed = theme_bytes(&workbook);
    assert_ne!(changed, extended);
    assert!(
        changed
            .windows(b"x:opaque marker=\"keep\"".len())
            .any(|window| { window == b"x:opaque marker=\"keep\"" })
    );
    assert_eq!(
        typed_theme(&workbook).theme().colors.color(Slot::Accent1),
        Some(&Color::Rgb("FF0000".to_owned()))
    );

    assert_eq!(commit.patch().inverse().after().source_xml(), extended);
    let changed_snapshot = typed_theme(&workbook);
    let mut inverse_edit = changed_snapshot.edit();
    inverse_edit
        .replace(before.theme().clone())
        .expect("stage original typed theme");
    let inverse_commit = inverse_edit.commit().expect("commit exact inverse");
    workbook
        .apply_theme(&inverse_commit)
        .expect("publish exact inverse");
    assert_eq!(theme_bytes(&workbook), extended);
}

#[test]
fn typed_noop_is_exact_and_stale_publication_is_atomic() {
    let mut workbook = native_workbook();
    let original = theme_bytes(&workbook);
    let before = typed_theme(&workbook);
    let noop = before.edit().commit().expect("commit no-op");
    assert!(!noop.changed());
    assert!(noop.patch().is_empty());
    workbook.apply_theme(&noop).expect("publish no-op");
    assert_eq!(theme_bytes(&workbook), original);

    let stale_commit = changed_commit(&before, "00B050");
    replace_theme_blob(&mut workbook, authored_theme_xml());
    let current = theme_bytes(&workbook);
    assert!(workbook.apply_theme(&stale_commit).is_err());
    assert_eq!(theme_bytes(&workbook), current);
}

#[test]
fn malformed_and_oversized_theme_payloads_are_refused() {
    let original_workbook = native_workbook();
    let original = theme_bytes(&original_workbook);
    for malformed in [
        b"<a:not-a-theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>".to_vec(),
        b"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:themeElements/></a:theme>".to_vec(),
        b"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:themeElements>".to_vec(),
    ] {
        assert!(codec::read(&malformed).is_err());
        let mut workbook = native_workbook();
        replace_theme_blob(&mut workbook, malformed);
        assert!(workbook.theme().is_err());
        assert_eq!(theme_bytes(&original_workbook), original);
    }

    let oversized = vec![b' '; codec::MAX_XML_BYTES + 1];
    assert!(codec::read(&oversized).is_err());
    assert!(codec::replace_scheme(&oversized, b"clrScheme", b"<a:clrScheme/>").is_err());

    let workbook = native_workbook();
    let exact = theme_bytes(&workbook).len();
    assert!(
        workbook
            .theme_with_limits(ThemeLimits::new(exact - 1, 100_000, 128))
            .is_err()
    );
    assert!(
        workbook
            .theme_with_limits(ThemeLimits::new(exact, 100_000, 128))
            .expect("exact eager theme limit")
            .is_some()
    );
}

#[test]
fn strict_theme_namespace_is_accepted_but_wrong_namespace_is_not() {
    let transitional = authored_theme_xml();
    let strict = transitional
        .windows(codec::NAMESPACE.len())
        .enumerate()
        .find_map(|(index, window)| (window == codec::NAMESPACE.as_bytes()).then_some(index))
        .map(|index| {
            let mut bytes = transitional.clone();
            bytes.splice(
                index..index + codec::NAMESPACE.len(),
                codec::STRICT_NAMESPACE.as_bytes().iter().copied(),
            );
            bytes
        })
        .expect("transitional namespace");
    assert!(codec::read(&strict).is_ok());

    let wrong_namespace = String::from_utf8(transitional)
        .expect("authored UTF-8")
        .replace(codec::NAMESPACE, "urn:example:not-drawingml")
        .into_bytes();
    assert!(codec::read(&wrong_namespace).is_err());

    let mut workbook = native_workbook();
    replace_theme_blob(&mut workbook, wrong_namespace);
    assert!(workbook.theme().is_err());
}

#[test]
fn theme_graph_ownership_rejects_bad_edges_and_content_types() {
    let wrong_content_type = package_with(|package| {
        package
            .get_part_mut(&theme_uri())
            .expect("theme part")
            .set_content_type("application/octet-stream".to_owned())
            .expect("blob content type is mutable");
    });
    assert!(wrong_content_type.theme().is_err());

    let duplicate_theme = package_with(|package| {
        package
            .get_part_mut(&workbook_uri())
            .expect("workbook part")
            .rels_mut()
            .try_add_relationship(
                rt::THEME.to_owned(),
                "theme/theme1.xml".to_owned(),
                "rIdThemeDuplicate".to_owned(),
                TargetMode::Internal,
            )
            .expect("second relationship id");
    });
    assert!(duplicate_theme.theme().is_err());

    let external_theme = package_with(|package| {
        let workbook = package
            .get_part_mut(&workbook_uri())
            .expect("workbook part");
        let id = workbook
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == rt::THEME)
            .expect("theme relationship")
            .r_id()
            .to_owned();
        workbook.rels_mut().remove(&id);
        workbook
            .rels_mut()
            .try_add_relationship(
                rt::THEME.to_owned(),
                "https://example.invalid/theme.xml".to_owned(),
                id,
                TargetMode::External,
            )
            .expect("external replacement relationship");
    });
    assert!(external_theme.theme().is_err());

    let unsupported_leaf_relationship = package_with(|package| {
        package
            .get_part_mut(&theme_uri())
            .expect("theme part")
            .rels_mut()
            .try_add_relationship(
                rt::CHART.to_owned(),
                "../charts/chart1.xml".to_owned(),
                "rIdUnexpected".to_owned(),
                TargetMode::Internal,
            )
            .expect("unsupported relationship");
    });
    assert!(unsupported_leaf_relationship.theme().is_err());
}

#[test]
fn theme_image_relationship_is_owned_and_preserved_by_typed_edit() {
    let image_uri = PackURI::new("/xl/media/theme-test.png").expect("image URI");
    let mut workbook = package_with(|package| {
        package.add_part(Box::new(BlobPart::new(
            image_uri.clone(),
            ct::PNG.to_owned(),
            vec![0x89, b'P', b'N', b'G'],
        )));
        package
            .get_part_mut(&theme_uri())
            .expect("theme part")
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "../media/theme-test.png".to_owned(),
                "rIdThemeImage".to_owned(),
                TargetMode::Internal,
            )
            .expect("theme image relationship");
    });
    let before = typed_theme(&workbook);
    let workbook_stream = workbook
        .opc_package()
        .get_part(&workbook_uri())
        .expect("workbook part before edit")
        .blob()
        .to_vec();
    assert!(before.source_xml().starts_with(b"<?xml"));
    let commit = changed_commit(&before, "7030A0");
    workbook
        .apply_theme(&commit)
        .expect("publish theme with image relationship");

    let theme = workbook
        .opc_package()
        .get_part(&theme_uri())
        .expect("theme part after edit");
    assert!(theme.rels().get("rIdThemeImage").is_some());
    assert_eq!(
        workbook
            .opc_package()
            .get_part(&image_uri)
            .expect("image part after edit")
            .blob(),
        [0x89, b'P', b'N', b'G']
    );
    assert_eq!(
        workbook
            .opc_package()
            .get_part(&workbook_uri())
            .expect("workbook part after edit")
            .blob(),
        workbook_stream
    );
}

#[test]
fn writer_custom_theme_is_created_and_round_trips_as_typed_metadata() {
    let custom = custom_theme();
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Sheet1"));
    writer
        .set_theme(custom.clone())
        .expect("validate custom writer theme");
    let mut bytes = Cursor::new(Vec::new());
    writer
        .save(&mut bytes)
        .expect("write custom themed workbook");

    let workbook = Workbook::new(Cursor::new(bytes.into_inner())).expect("open custom workbook");
    let snapshot = typed_theme(&workbook);
    assert_eq!(snapshot.theme(), &custom);
    assert_eq!(
        snapshot.source_xml(),
        codec::encode_part(&custom.name, &custom.colors, &custom.fonts)
            .expect("encode custom theme")
            .as_slice()
    );
    assert_eq!(
        workbook
            .opc_package()
            .get_part(&theme_uri())
            .expect("writer theme part")
            .content_type(),
        ct::OFC_THEME
    );
}

#[test]
fn typed_theme_owner_remove_inverse_restores_xml_graph_and_resources() {
    let (mut workbook, image_uri) = image_bearing_theme_workbook();

    let owner = workbook
        .theme_owner()
        .expect("read image-bearing Theme owner");
    assert!(owner.is_present());
    let before_xml = owner.source_xml().expect("present Theme XML").to_vec();
    let before_theme_relationships = relationship_fingerprint(&workbook, &theme_uri());
    let before_theme_relationship_xml = relationship_xml(&workbook, &theme_uri());
    let before_workbook_relationships = relationship_fingerprint(&workbook, &workbook_uri());
    let before_workbook_relationship_xml = relationship_xml(&workbook, &workbook_uri());
    let before_image = workbook
        .opc_package()
        .get_part(&image_uri)
        .expect("Theme image payload")
        .blob()
        .to_vec();
    let before_content_types = workbook
        .opc_package()
        .source_content_types()
        .expect("capture content-types XML")
        .bytes()
        .to_vec();

    let mut removal = owner.edit();
    assert!(removal.remove().expect("stage Theme owner removal"));
    let removal = removal.commit().expect("commit Theme owner removal");
    assert!(removal.changed());
    let removed = workbook
        .apply_theme_owner(&removal)
        .expect("publish Theme owner removal");
    assert!(!removed.is_present());
    assert!(workbook.opc_package().get_part(&theme_uri()).is_err());

    let restored = workbook
        .apply_theme_owner_patch(&removal.patch().inverse())
        .expect("publish exact inverse Theme owner patch");
    assert!(restored.is_present());
    assert_eq!(restored.source_xml(), Some(before_xml.as_slice()));
    assert_eq!(
        relationship_fingerprint(&workbook, &theme_uri()),
        before_theme_relationships
    );
    assert_eq!(
        relationship_xml(&workbook, &theme_uri()),
        before_theme_relationship_xml
    );
    assert_eq!(
        relationship_fingerprint(&workbook, &workbook_uri()),
        before_workbook_relationships
    );
    assert_eq!(
        relationship_xml(&workbook, &workbook_uri()),
        before_workbook_relationship_xml
    );
    assert_eq!(
        workbook
            .opc_package()
            .get_part(&image_uri)
            .expect("restored Theme image payload")
            .blob(),
        before_image
    );
    assert_eq!(
        workbook
            .opc_package()
            .source_content_types()
            .expect("capture restored content-types XML")
            .bytes(),
        before_content_types
    );
}

#[test]
fn typed_theme_owner_rejects_lexically_stale_theme_relationships_atomically() {
    let (mut workbook, _) = image_bearing_theme_workbook();
    let owner = workbook
        .theme_owner()
        .expect("read image-bearing Theme owner");
    let before_relationships = relationship_fingerprint(&workbook, &theme_uri());
    let mut edit = owner.edit();
    assert!(edit.set_name("Stale Theme Edit").expect("stage Theme edit"));
    let commit = edit.commit().expect("commit Theme edit");

    workbook
        .edit_opc(|package| {
            replace_theme_relationship_source_lexically(package);
            Ok(())
        })
        .expect("publish semantically unchanged Theme relationship spelling");
    let stale_relationship_xml = relationship_xml(&workbook, &theme_uri());
    assert_eq!(
        relationship_fingerprint(&workbook, &theme_uri()),
        before_relationships
    );
    let stale_theme_xml = theme_bytes(&workbook);

    assert!(workbook.apply_theme_owner(&commit).is_err());
    assert_eq!(
        relationship_xml(&workbook, &theme_uri()),
        stale_relationship_xml
    );
    assert_eq!(theme_bytes(&workbook), stale_theme_xml);
}

#[test]
fn typed_theme_owner_uses_theme2_when_theme1_is_occupied() {
    let absent_package = remove_workbook_theme(native_workbook().opc_package().clone());
    let mut absent = Workbook::from_opc_package(absent_package).expect("open absent-theme host");
    absent
        .edit_opc(|package| {
            package.add_part(Box::new(BlobPart::new(
                theme_uri(),
                ct::PNG.to_owned(),
                vec![0x89, b'P', b'N', b'G'],
            )));
            Ok(())
        })
        .expect("occupy theme1 with unrelated part");

    let owner = absent.theme_owner().expect("read absent Theme owner");
    assert!(!owner.is_present());
    assert_eq!(owner.part_name(), "/xl/theme/theme2.xml");
    let custom = custom_theme();
    let mut create = owner.edit();
    assert!(
        create
            .set_theme(custom.clone())
            .expect("stage Theme creation")
    );
    let create = create.commit().expect("commit Theme creation");
    let created = absent
        .apply_theme_owner(&create)
        .expect("publish Theme creation at theme2");
    assert!(created.is_present());
    assert_eq!(created.part_name(), "/xl/theme/theme2.xml");
    assert_eq!(created.theme(), Some(&custom));
    assert_eq!(
        absent
            .opc_package()
            .get_part(&theme_uri())
            .expect("unrelated theme1 part")
            .content_type(),
        ct::PNG
    );

    let mut remove = created.edit();
    assert!(remove.remove().expect("stage Theme2 removal"));
    let remove = remove.commit().expect("commit Theme2 removal");
    let absent_again = absent
        .apply_theme_owner(&remove)
        .expect("publish Theme2 removal");
    assert!(!absent_again.is_present());
    assert!(
        absent
            .opc_package()
            .get_part(&PackURI::new("/xl/theme/theme2.xml").unwrap())
            .is_err()
    );

    let mut recreate = absent_again.edit();
    assert!(
        recreate
            .set_theme(custom.clone())
            .expect("stage Theme2 recreation")
    );
    let recreate = recreate.commit().expect("commit Theme2 recreation");
    let recreated = absent
        .apply_theme_owner(&recreate)
        .expect("publish Theme2 recreation");
    assert_eq!(recreated.part_name(), "/xl/theme/theme2.xml");
    assert_eq!(recreated.theme(), Some(&custom));
}

#[test]
fn newly_created_theme_inverse_rejects_changed_empty_relationship_member() {
    let absent_package = remove_workbook_theme(native_workbook().opc_package().clone());
    let mut workbook = Workbook::from_opc_package(absent_package).expect("open absent-theme host");
    let owner = workbook.theme_owner().expect("read absent Theme owner");
    assert!(!owner.is_present());

    let custom = custom_theme();
    let mut create = owner.edit();
    assert!(create.set_theme(custom).expect("stage Theme creation"));
    let create = create.commit().expect("commit Theme creation");
    let created = workbook
        .apply_theme_owner(&create)
        .expect("publish Theme creation");
    assert!(created.is_present());
    let created_theme_uri = PackURI::new(created.part_name()).expect("created Theme part URI");
    let image_uri = PackURI::new("/xl/media/theme-empty-rels.png").expect("image URI");
    workbook
        .edit_opc(|package| {
            package.add_part(Box::new(BlobPart::new(
                image_uri.clone(),
                ct::PNG.to_owned(),
                vec![0x89, b'P', b'N', b'G'],
            )));
            package
                .get_part_mut(&created_theme_uri)
                .expect("created Theme part")
                .rels_mut()
                .try_add_relationship(
                    rt::IMAGE.to_owned(),
                    "../media/theme-empty-rels.png".to_owned(),
                    "rIdTemporaryImage".to_owned(),
                    TargetMode::Internal,
                )
                .expect("temporary Theme image relationship");
            let source = package
                .source_relationships(&created_theme_uri)
                .expect("capture non-empty Theme relationship source");
            let empty = source
                .without_relationship("rIdTemporaryImage", usize::MAX)
                .expect("remove temporary Theme image relationship");
            assert!(empty.member_present(), "retain physical empty .rels member");
            package
                .try_replace_relationships(&source, &empty)
                .expect("publish physical empty Theme .rels member");
            Ok(())
        })
        .expect("publish changed empty Theme2 relationship member");
    let empty_relationship_xml = relationship_xml(&workbook, &created_theme_uri);
    assert!(!empty_relationship_xml.is_empty());
    let theme_xml = theme_bytes(&workbook);

    assert!(
        workbook
            .apply_theme_owner_patch(&create.patch().inverse())
            .is_err(),
        "inverse must reject a newly authored physical empty .rels change"
    );
    assert!(
        workbook
            .theme_owner()
            .expect("read unchanged owner")
            .is_present()
    );
    assert_eq!(
        relationship_xml(&workbook, &created_theme_uri),
        empty_relationship_xml
    );
    assert_eq!(theme_bytes(&workbook), theme_xml);
}

#[test]
fn absent_theme_owner_remove_noop_preserves_source_bytes() {
    let absent_package = remove_workbook_theme(native_workbook().opc_package().clone());
    let mut workbook = Workbook::from_opc_package(absent_package).expect("open absent-theme host");
    let before_workbook_relationships = relationship_xml(&workbook, &workbook_uri());
    let before_content_types = workbook
        .opc_package()
        .source_content_types()
        .expect("capture absent content-types XML")
        .bytes()
        .to_vec();
    let before_package = saved_bytes(&workbook);

    let owner = workbook.theme_owner().expect("read absent Theme owner");
    let mut remove = owner.edit();
    assert!(!remove.remove().expect("stage absent Theme no-op"));
    let noop = remove.commit().expect("commit absent Theme no-op");
    assert!(!noop.changed());
    assert!(noop.patch().is_empty());
    let applied = workbook
        .apply_theme_owner(&noop)
        .expect("publish absent Theme no-op");
    assert!(!applied.is_present());
    assert_eq!(
        relationship_xml(&workbook, &workbook_uri()),
        before_workbook_relationships
    );
    assert_eq!(
        workbook
            .opc_package()
            .source_content_types()
            .expect("capture preserved content-types XML")
            .bytes(),
        before_content_types
    );
    assert_eq!(saved_bytes(&workbook), before_package);
}

#[test]
fn theme_graph_rejects_orphans_multiple_parts_and_foreign_inbound_edges() {
    let orphan = package_with(|package| {
        let relationship_id = package
            .get_part(&workbook_uri())
            .expect("workbook part")
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == rt::THEME)
            .expect("workbook theme relationship")
            .r_id()
            .to_owned();
        package
            .get_part_mut(&workbook_uri())
            .expect("workbook part")
            .rels_mut()
            .remove(&relationship_id);
    });
    assert!(orphan.theme().is_err());

    let multiple = package_with(|package| {
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/theme/theme2.xml").expect("second theme URI"),
            ct::OFC_THEME.to_owned(),
            authored_theme_xml(),
        )));
    });
    assert!(multiple.theme().is_err());

    let inbound = package_with(|package| {
        package
            .get_part_mut(&PackURI::new(CORE_PROPERTIES_NAME).expect("core properties URI"))
            .expect("core properties part")
            .rels_mut()
            .try_add_relationship(
                rt::THEME.to_owned(),
                "../xl/theme/theme1.xml".to_owned(),
                "rIdForeignTheme".to_owned(),
                TargetMode::Internal,
            )
            .expect("foreign inbound Theme edge");
    });
    assert!(inbound.theme().is_err());

    let package_root_inbound = package_with(|package| {
        package
            .rels_mut()
            .try_add_relationship(
                rt::THEME.to_owned(),
                "xl/theme/theme1.xml".to_owned(),
                "rIdPackageRootTheme".to_owned(),
                TargetMode::Internal,
            )
            .expect("package-root inbound Theme edge");
    });
    assert!(package_root_inbound.theme().is_err());
}

#[test]
fn theme_graph_validates_image_targets_and_transitional_strict_conformance() {
    let bad_type = package_with(|package| {
        let image_uri = PackURI::new("/xl/media/theme-invalid.bin").expect("image URI");
        package.add_part(Box::new(BlobPart::new(
            image_uri,
            ct::OFC_THEME.to_owned(),
            authored_theme_xml(),
        )));
        package
            .get_part_mut(&theme_uri())
            .expect("theme part")
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "../media/theme-invalid.bin".to_owned(),
                "rIdInvalidImage".to_owned(),
                TargetMode::Internal,
            )
            .expect("invalid image target relationship");
    });
    assert!(bad_type.theme().is_err());

    let non_leaf = package_with(|package| {
        let image_uri = PackURI::new("/xl/media/theme-non-leaf.png").expect("image URI");
        package.add_part(Box::new(BlobPart::new(
            image_uri.clone(),
            ct::PNG.to_owned(),
            vec![0x89, b'P', b'N', b'G'],
        )));
        package
            .get_part_mut(&image_uri)
            .expect("image part")
            .rels_mut()
            .try_add_relationship(
                rt::CUSTOM_XML.to_owned(),
                "../xl/media/theme-non-leaf.png".to_owned(),
                "rIdImageChild".to_owned(),
                TargetMode::Internal,
            )
            .expect("image child relationship");
        package
            .get_part_mut(&theme_uri())
            .expect("theme part")
            .rels_mut()
            .try_add_relationship(
                rt::IMAGE.to_owned(),
                "../media/theme-non-leaf.png".to_owned(),
                "rIdNonLeafImage".to_owned(),
                TargetMode::Internal,
            )
            .expect("non-leaf image relationship");
    });
    assert!(non_leaf.theme().is_err());

    let mismatch = package_with(|package| {
        replace_workbook_theme_relationship(package, STRICT_THEME_RELATIONSHIP);
    });
    assert!(mismatch.theme().is_err());

    let strict = package_with(|package| {
        package
            .get_part_mut(&theme_uri())
            .expect("theme part")
            .set_blob(strict_theme_xml());
        replace_workbook_theme_relationship(package, STRICT_THEME_RELATIONSHIP);
    });
    assert_eq!(typed_theme(&strict).theme().name, "Office");
}
