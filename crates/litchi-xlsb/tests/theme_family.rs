#![allow(
    clippy::expect_used,
    clippy::pedantic,
    clippy::unwrap_used,
    reason = "native package fixtures are checked and fail fast"
)]

//! Independent XLSB Theme-part integration coverage for DrawingML 2012
//! `themeFamily` ownership.
//!
//! These tests deliberately exercise the package facade rather than private
//! scanner helpers.  The fixture is a real XLSB Theme part, while variants
//! are made by changing only the XML source needed to test ownership and
//! source-preserving publication.

use std::io::{self, Cursor};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_core::{ReadAt, SourceVersion};
use litchi_drawingml::theme::{codec, family};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use litchi_xlsb::theme::{
    Family, Limits, THEME_FAMILY_EXTENSION_URI, THEME_FAMILY_NAMESPACE_EXTENSION_URI,
};
use litchi_xlsb::{SourceBackedWorkbook, Workbook};

const NATIVE_FAMILY_FIXTURE: &str = "test-data/ooxml/xlsb/hyperlink.xlsb";
const FAMILY_FREE_FIXTURE: &str = "test-data/ooxml/xlsb/62815.xlsb";
const THEME_NAME: &str = "/xl/theme/theme1.xml";
const WORKBOOK_NAME: &str = "/xl/workbook.bin";
const FAMILY_NAMESPACE: &str = family::NAMESPACE;
const OFFICE_ID: &str = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}";
const OFFICE_VID: &str = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}";
const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const STRICT_THEME_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/theme";

fn fixture(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn workbook_from_fixture(relative: &str) -> Workbook {
    Workbook::new(Cursor::new(
        std::fs::read(fixture(relative)).expect("native XLSB fixture"),
    ))
    .expect("native XLSB package")
}

fn native_workbook() -> Workbook {
    workbook_from_fixture(NATIVE_FAMILY_FIXTURE)
}

fn theme_uri() -> PackURI {
    PackURI::new(THEME_NAME).expect("Theme URI")
}

fn workbook_uri() -> PackURI {
    PackURI::new(WORKBOOK_NAME).expect("Workbook URI")
}

fn theme_bytes(workbook: &Workbook) -> Vec<u8> {
    workbook
        .opc_package()
        .get_part(&theme_uri())
        .expect("Theme part")
        .blob()
        .to_vec()
}

fn theme_snapshot(workbook: &Workbook) -> litchi_xlsb::theme::Snapshot {
    workbook
        .theme()
        .expect("read Theme")
        .expect("Workbook Theme")
}

fn family_value(name: &str) -> Family {
    Family::new(name, OFFICE_ID, OFFICE_VID).expect("valid Theme family")
}

fn replace_theme_blob(workbook: &mut Workbook, xml: Vec<u8>) {
    workbook
        .edit_opc(|package| {
            package
                .get_part_mut(&theme_uri())
                .expect("Theme part")
                .set_blob(xml);
            Ok(())
        })
        .expect("replace Theme source");
}

fn replace_bytes(source: &[u8], old: &[u8], new: &[u8]) -> Vec<u8> {
    let start = source
        .windows(old.len())
        .position(|window| window == old)
        .expect("source marker");
    let mut output = Vec::with_capacity(source.len() + new.len() - old.len());
    output.extend_from_slice(&source[..start]);
    output.extend_from_slice(new);
    output.extend_from_slice(&source[start + old.len()..]);
    output
}

fn compact_xml(source: Vec<u8>) -> Vec<u8> {
    if source
        .windows(b"?>\r\n".len())
        .any(|window| window == b"?>\r\n")
    {
        replace_bytes(&source, b"?>\r\n", b"?>")
    } else {
        source
    }
}

fn theme_family_fragment(source: &[u8]) -> &[u8] {
    let start = source
        .windows(b"<thm15:themeFamily".len())
        .position(|window| window == b"<thm15:themeFamily")
        .expect("family element");
    let end = source[start..]
        .windows(2)
        .position(|window| window == b"/>")
        .map(|offset| start + offset + 2)
        .expect("family close");
    &source[start..end]
}

fn familyless_source() -> Vec<u8> {
    let source = workbook_from_fixture(FAMILY_FREE_FIXTURE)
        .opc_package()
        .get_part(&theme_uri())
        .expect("family-free Theme part")
        .blob()
        .to_vec();
    // The native member has a declaration line break that the package XML
    // publication validator intentionally rejects after a source splice.
    replace_bytes(&source, b"?>\r\n", b"?>")
}

fn with_root_namespace(source: &[u8], declaration: &str) -> Vec<u8> {
    let marker = b"<a:theme ";
    let insertion = marker.len();
    let start = source
        .windows(marker.len())
        .position(|window| window == marker)
        .expect("Theme root");
    let mut output = Vec::with_capacity(source.len() + declaration.len());
    output.extend_from_slice(&source[..start + insertion]);
    output.extend_from_slice(declaration.as_bytes());
    output.extend_from_slice(&source[start + insertion..]);
    output
}

fn with_unknown_extension_sibling(source: &[u8]) -> Vec<u8> {
    let source = with_root_namespace(source, "xmlns:x=\"urn:litchi:theme-family\" ");
    replace_bytes(&source, b"</a:ext>", b"<x:future marker=\"keep\"/></a:ext>")
}

fn with_foreign_wrapper(source: &[u8]) -> Vec<u8> {
    let source = with_root_namespace(source, "xmlns:x=\"urn:litchi:theme-family-wrapper\" ");
    let old = theme_family_fragment(&source).to_vec();
    let mut replacement = b"<x:wrapper>".to_vec();
    replacement.extend_from_slice(&old);
    replacement.extend_from_slice(b"</x:wrapper>");
    replace_bytes(&source, &old, &replacement)
}

fn with_mce_wrapper(source: &[u8]) -> Vec<u8> {
    let source = with_root_namespace(
        source,
        &format!("xmlns:mc=\"{MC_NAMESPACE}\" xmlns:x=\"urn:litchi:unsupported\" "),
    );
    let old = theme_family_fragment(&source).to_vec();
    let mut replacement = b"<mc:AlternateContent><mc:Choice Requires=\"x\">".to_vec();
    replacement.extend_from_slice(&old);
    replacement.extend_from_slice(b"</mc:Choice><mc:Fallback/></mc:AlternateContent>");
    replace_bytes(&source, &old, &replacement)
}

fn with_mce_hidden_extlst(source: &[u8]) -> Vec<u8> {
    let list_start = source
        .windows(b"<a:extLst>".len())
        .position(|window| window == b"<a:extLst>")
        .expect("Theme extLst");
    let list_end = source[list_start..]
        .windows(b"</a:extLst>".len())
        .position(|window| window == b"</a:extLst>")
        .map(|offset| list_start + offset + b"</a:extLst>".len())
        .expect("Theme extLst close");
    let old = source[list_start..list_end].to_vec();
    let family = std::str::from_utf8(theme_family_fragment(source)).expect("family UTF-8");
    let replacement = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="x"><a:extLst><a:ext uri="{uri}">{family}</a:ext></a:extLst></mc:Choice><mc:Fallback/></mc:AlternateContent>"#,
        uri = THEME_FAMILY_EXTENSION_URI,
    );
    let wrapped = replace_bytes(source, &old, replacement.as_bytes());
    with_root_namespace(
        &wrapped,
        &format!("xmlns:mc=\"{MC_NAMESPACE}\" xmlns:x=\"urn:litchi:mce-hidden\" "),
    )
}

fn with_inherited_family_namespace(source: &[u8]) -> Vec<u8> {
    let without_local = replace_bytes(
        source,
        format!(" xmlns:thm15=\"{FAMILY_NAMESPACE}\"").as_bytes(),
        b"",
    );
    let value = with_root_namespace(
        &without_local,
        &format!("xmlns:thm15=\"{FAMILY_NAMESPACE}\" "),
    );
    assert!(
        value
            .windows(b"<thm15:themeFamily".len())
            .any(|window| window == b"<thm15:themeFamily")
    );
    value
}

fn with_prefix_collision(source: &[u8]) -> Vec<u8> {
    with_root_namespace(source, "xmlns:thm15=\"urn:litchi:colliding-prefix\" ")
}

fn with_alternate_outer_uri(source: &[u8], uri: &str) -> Vec<u8> {
    replace_bytes(
        source,
        THEME_FAMILY_EXTENSION_URI.as_bytes(),
        uri.as_bytes(),
    )
}

fn strict_theme_xml(source: &[u8]) -> Vec<u8> {
    String::from_utf8(source.to_vec())
        .expect("Theme XML is UTF-8")
        .replace(codec::NAMESPACE, codec::STRICT_NAMESPACE)
        .into_bytes()
}

fn replace_workbook_theme_relationship(package: &mut OpcPackage, relationship_type: &str) {
    let (id, target) = package
        .get_part(&workbook_uri())
        .expect("Workbook part")
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::THEME)
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.target_ref().to_owned(),
            )
        })
        .expect("Workbook Theme relationship");
    let workbook = package
        .get_part_mut(&workbook_uri())
        .expect("Workbook part");
    workbook.rels_mut().remove(&id);
    workbook
        .rels_mut()
        .try_add_relationship(
            relationship_type.to_owned(),
            target,
            id,
            TargetMode::Internal,
        )
        .expect("strict Theme relationship");
}

fn signed_workbook() -> Workbook {
    let mut package = native_workbook().opc_package().clone();
    let origin = PackURI::new("/_xmlsignatures/origin.sigs").expect("signature origin");
    let signature = PackURI::new("/_xmlsignatures/sig1.xml").expect("signature");
    let mut origin_part = BlobPart::new(
        origin.clone(),
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
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
        .expect("signature relationship");
    package.add_part(Box::new(origin_part));
    package.add_part(Box::new(BlobPart::new(
        signature,
        ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
        b"<Signature/>".to_vec(),
    )));
    package
        .rels_mut()
        .try_add_relationship(
            rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            origin.as_str().to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature origin relationship");
    Workbook::from_opc_package(package).expect("signed workbook")
}

#[derive(Debug)]
struct CountingSource {
    bytes: Vec<u8>,
    reads: AtomicUsize,
}

impl ReadAt for CountingSource {
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
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x5448_454d, 0))
    }
}

#[test]
fn native_xlsb_profiles_project_the_direct_family_owner() {
    for relative in [
        "test-data/ooxml/xlsb/hyperlink.xlsb",
        "test-data/ooxml/xlsb/bug66682.xlsb",
        "test-data/ooxml/xlsb/date.xlsb",
        "test-data/ooxml/xlsb/cond_format.xlsb",
    ] {
        let workbook = workbook_from_fixture(relative);
        let snapshot = theme_snapshot(&workbook);
        let family = snapshot.family().expect("native Theme family");
        assert_eq!(family.name(), "Office Theme");
        assert_eq!(family.id().as_str(), OFFICE_ID);
        assert_eq!(family.variant_id().as_str(), OFFICE_VID);
        assert!(
            snapshot
                .source_xml()
                .windows(b"themeFamily".len())
                .any(|window| { window == b"themeFamily" })
        );
    }
}

#[test]
fn strict_theme_profile_keeps_family_ownership() {
    let mut package = native_workbook().opc_package().clone();
    let strict = strict_theme_xml(package.get_part(&theme_uri()).expect("Theme part").blob());
    package
        .get_part_mut(&theme_uri())
        .expect("Theme part")
        .set_blob(strict);
    replace_workbook_theme_relationship(&mut package, STRICT_THEME_RELATIONSHIP);
    let workbook = Workbook::from_opc_package(package).expect("strict package");
    assert_eq!(
        theme_snapshot(&workbook)
            .family()
            .expect("strict family")
            .name(),
        "Office Theme"
    );
}

#[test]
fn outer_uri_namespace_uri_and_foreign_uri_are_distinguished() {
    let source = theme_bytes(&native_workbook());

    let namespace_uri = with_alternate_outer_uri(&source, THEME_FAMILY_NAMESPACE_EXTENSION_URI);
    let mut workbook = native_workbook();
    replace_theme_blob(&mut workbook, namespace_uri);
    assert!(theme_snapshot(&workbook).family().is_some());

    let foreign_uri = with_alternate_outer_uri(&source, "urn:litchi:unrelated-extension");
    replace_theme_blob(&mut workbook, foreign_uri);
    assert!(theme_snapshot(&workbook).family().is_none());
}

#[test]
fn supported_outer_uri_token_whitespace_resolves_and_keeps_lexical_uri() {
    let source = theme_bytes(&native_workbook());
    for (uri, lexical) in [
        (
            THEME_FAMILY_EXTENSION_URI,
            format!("  {THEME_FAMILY_EXTENSION_URI} \t"),
        ),
        (
            THEME_FAMILY_NAMESPACE_EXTENSION_URI,
            format!("&#x20;{THEME_FAMILY_NAMESPACE_EXTENSION_URI}&#9;"),
        ),
    ] {
        let old = format!("uri=\"{THEME_FAMILY_EXTENSION_URI}\"");
        let replacement = format!("uri=\"{lexical}\"");
        let candidate = compact_xml(replace_bytes(
            &source,
            old.as_bytes(),
            replacement.as_bytes(),
        ));
        let mut workbook = native_workbook();
        replace_theme_blob(&mut workbook, candidate);
        assert_eq!(
            theme_snapshot(&workbook)
                .family()
                .expect("token-whitespace family")
                .name(),
            "Office Theme"
        );

        let before = theme_snapshot(&workbook);
        let mut edit = before.edit();
        edit.set_family(family_value("Lexical URI Family"))
            .expect("stage URI-preserving edit");
        let commit = edit.commit().expect("commit URI-preserving edit");
        workbook
            .apply_theme(&commit)
            .expect("publish URI-preserving edit");
        let changed = theme_bytes(&workbook);
        assert!(
            changed
                .windows(replacement.len())
                .any(|window| window == replacement.as_bytes()),
            "supported URI {uri} must retain its lexical source"
        );
    }
}

#[test]
fn family_ownership_does_not_cross_foreign_or_ignored_mce_ancestry() {
    let source = theme_bytes(&native_workbook());
    let mut workbook = native_workbook();
    replace_theme_blob(&mut workbook, with_foreign_wrapper(&source));
    assert!(theme_snapshot(&workbook).family().is_none());

    replace_theme_blob(&mut workbook, with_mce_wrapper(&source));
    assert!(theme_snapshot(&workbook).family().is_none());
}

#[test]
fn adding_family_refuses_an_ignored_mce_extlst_branch_without_source_mutation() {
    let source = theme_bytes(&native_workbook());
    let hidden = compact_xml(with_mce_hidden_extlst(&source));
    let hidden_family = theme_family_fragment(&hidden).to_vec();
    let mut workbook = native_workbook();
    replace_theme_blob(&mut workbook, hidden);
    let before_bytes = theme_bytes(&workbook);
    let before = theme_snapshot(&workbook);
    assert!(before.family().is_none());

    let mut edit = before.edit();
    edit.set_family(family_value("Fresh Direct Family"))
        .expect("stage direct family");
    assert!(
        edit.commit().is_err(),
        "a hidden MCE family must refuse the package mutation"
    );
    let changed = theme_bytes(&workbook);
    assert_eq!(
        changed, before_bytes,
        "failed mutation changed Theme source"
    );
    assert!(
        changed
            .windows(hidden_family.len())
            .any(|window| window == hidden_family.as_slice()),
        "the hidden branch must retain its original family bytes"
    );
}

#[test]
fn inherited_namespace_and_prefix_collision_remain_resolved() {
    let source = theme_bytes(&native_workbook());
    let mut workbook = native_workbook();
    replace_theme_blob(&mut workbook, with_inherited_family_namespace(&source));
    assert_eq!(
        theme_snapshot(&workbook)
            .family()
            .expect("inherited family namespace")
            .name(),
        "Office Theme"
    );

    replace_theme_blob(&mut workbook, with_prefix_collision(&source));
    assert_eq!(
        theme_snapshot(&workbook)
            .family()
            .expect("local family prefix wins over inherited collision")
            .name(),
        "Office Theme"
    );
}

#[test]
fn fresh_family_replacement_keeps_unknown_sibling_bytes() {
    let source = compact_xml(with_unknown_extension_sibling(&theme_bytes(
        &native_workbook(),
    )));
    let mut workbook = native_workbook();
    replace_theme_blob(&mut workbook, source.clone());
    let snapshot = theme_snapshot(&workbook);
    let mut edit = snapshot.edit();
    assert!(
        edit.set_family(family_value("Fresh Family"))
            .expect("stage fresh family")
    );
    let commit = edit.commit().expect("commit fresh family");
    workbook.apply_theme(&commit).expect("publish fresh family");
    let changed = theme_bytes(&workbook);
    assert!(
        changed
            .windows(b"x:future marker=\"keep\"".len())
            .any(|window| window == b"x:future marker=\"keep\"")
    );
    assert_eq!(
        theme_snapshot(&workbook)
            .family()
            .expect("fresh family")
            .name(),
        "Fresh Family"
    );
}

#[test]
fn family_edit_preserves_opaque_family_markup_and_inverse_is_exact() {
    let source = theme_bytes(&native_workbook());
    let old = theme_family_fragment(&source).to_vec();
    let replacement = String::from_utf8(old.clone())
        .expect("family UTF-8")
        .replace(
            &format!(" vid=\"{OFFICE_VID}\"/>") ,
            &format!(
                " vid=\"{OFFICE_VID}\" xmlns:x=\"urn:litchi:opaque-family\"><x:child marker=\"keep\"/></thm15:themeFamily>"
            ),
        )
        .into_bytes();
    let source = compact_xml(replace_bytes(&source, &old, &replacement));
    let mut workbook = native_workbook();
    replace_theme_blob(&mut workbook, source.clone());
    let snapshot = theme_snapshot(&workbook);
    let mut family = snapshot.family().expect("opaque family").clone();
    family.set_name("Edited Family").expect("edit family name");
    let mut edit = snapshot.edit();
    edit.set_family(family).expect("stage opaque family edit");
    let commit = edit.commit().expect("commit opaque family edit");
    workbook
        .apply_theme(&commit)
        .expect("publish opaque family edit");
    let changed = theme_bytes(&workbook);
    assert!(
        changed
            .windows(b"x:child marker=\"keep\"".len())
            .any(|window| window == b"x:child marker=\"keep\"")
    );

    let after = theme_snapshot(&workbook);
    let inverse = commit.patch().inverse();
    workbook
        .apply_theme_patch(&inverse)
        .expect("publish family inverse");
    assert_eq!(theme_bytes(&workbook), source);
    assert_eq!(
        after.family().expect("after family").name(),
        "Edited Family"
    );
}

#[test]
fn absent_family_add_remove_noop_inverse_and_reopen_are_exact() {
    let original = familyless_source();
    let mut workbook = workbook_from_fixture(FAMILY_FREE_FIXTURE);
    replace_theme_blob(&mut workbook, original.clone());
    assert!(theme_snapshot(&workbook).family().is_none());

    let before = theme_snapshot(&workbook);
    let mut noop_edit = before.edit();
    assert!(
        !noop_edit
            .remove_family()
            .expect("absent family remove no-op")
    );
    let noop = noop_edit.commit().expect("absent family no-op");
    assert!(!noop.changed());
    workbook.apply_theme(&noop).expect("publish absent no-op");
    assert_eq!(theme_bytes(&workbook), original);

    let mut add_edit = before.edit();
    assert!(
        add_edit
            .set_family(family_value("Added Family"))
            .expect("stage family add")
    );
    let add = add_edit.commit().expect("commit family add");
    workbook.apply_theme(&add).expect("publish family add");
    assert!(theme_snapshot(&workbook).family().is_some());
    assert!(
        theme_bytes(&workbook)
            .windows(b"<a:extLst>".len())
            .any(|window| { window == b"<a:extLst>" })
    );

    let reopened = Workbook::new(Cursor::new(saved_bytes(&workbook))).expect("reopen family add");
    assert_eq!(
        theme_snapshot(&reopened)
            .family()
            .expect("reopened family")
            .name(),
        "Added Family"
    );

    let current = theme_snapshot(&workbook);
    let mut remove_edit = current.edit();
    assert!(remove_edit.remove_family().expect("stage family remove"));
    let remove = remove_edit.commit().expect("commit family remove");
    workbook
        .apply_theme(&remove)
        .expect("publish family remove");
    assert!(theme_snapshot(&workbook).family().is_none());
    workbook
        .apply_theme_patch(&remove.patch().inverse())
        .expect("reapply family through inverse remove");
    workbook
        .apply_theme_patch(&add.patch().inverse())
        .expect("publish family inverse");
    assert!(theme_snapshot(&workbook).family().is_none());
    assert_eq!(theme_bytes(&workbook), original);
}

#[test]
fn stale_family_patch_is_atomic_and_limits_are_bounded() {
    let mut workbook = native_workbook();
    let before = theme_snapshot(&workbook);
    let mut edit = before.edit();
    edit.set_family(family_value("Stale Family"))
        .expect("stage stale family");
    let commit = edit.commit().expect("commit stale family");

    let replacement = familyless_source();
    replace_theme_blob(&mut workbook, replacement.clone());
    assert!(workbook.apply_theme(&commit).is_err());
    assert_eq!(theme_bytes(&workbook), replacement);

    let exact = theme_bytes(&native_workbook()).len();
    let exact_limits = Limits::new(exact, 100_000, 128);
    assert!(
        workbook_from_fixture(NATIVE_FAMILY_FIXTURE)
            .theme_with_limits(exact_limits)
            .expect("exact Theme source limit")
            .is_some()
    );
    assert!(
        workbook_from_fixture(NATIVE_FAMILY_FIXTURE)
            .theme_with_limits(Limits::new(exact - 1, 100_000, 128))
            .is_err()
    );
    assert!(
        workbook_from_fixture(NATIVE_FAMILY_FIXTURE)
            .theme_with_limits(Limits::new(exact, 1, 128))
            .is_err()
    );
    assert!(
        workbook_from_fixture(NATIVE_FAMILY_FIXTURE)
            .theme_with_limits(Limits::new(exact, 100_000, 1))
            .is_err()
    );
}

#[test]
fn malformed_recognized_family_content_fails_closed() {
    let source = theme_bytes(&native_workbook());
    let family = theme_family_fragment(&source).to_vec();
    let missing_vid = replace_bytes(&source, &format!(" vid=\"{OFFICE_VID}\"").into_bytes(), b"");
    assert!(
        Workbook::from_opc_package({
            let mut package = native_workbook().opc_package().clone();
            package
                .get_part_mut(&theme_uri())
                .expect("Theme part")
                .set_blob(missing_vid);
            package
        })
        .expect("package graph")
        .theme()
        .is_err()
    );

    let malformed_reference = replace_bytes(
        &source,
        b"name=\"Office Theme\"",
        b"name=\"Office &unknown;\"",
    );
    let mut package = native_workbook().opc_package().clone();
    package
        .get_part_mut(&theme_uri())
        .expect("Theme part")
        .set_blob(malformed_reference);
    assert!(
        Workbook::from_opc_package(package)
            .expect("package graph")
            .theme()
            .is_err()
    );

    assert!(
        family::read(
            &family
                .iter()
                .copied()
                .chain(Some(b'\0'))
                .collect::<Vec<_>>(),
        )
        .is_err()
    );
}

#[test]
fn signed_theme_family_noop_preserves_signature_but_change_requires_policy() {
    let mut signed = signed_workbook();
    assert!(signed.is_signed());
    let before_theme = theme_bytes(&signed);
    let snapshot = theme_snapshot(&signed);
    let noop = snapshot.edit().commit().expect("signed family no-op");
    assert!(!noop.changed());
    signed.apply_theme(&noop).expect("signed no-op");
    assert!(signed.is_signed());
    assert_eq!(theme_bytes(&signed), before_theme);

    let mut edit = snapshot.edit();
    edit.set_family(family_value("Signed Change"))
        .expect("stage signed family");
    let changed = edit.commit().expect("commit signed family");
    assert!(signed.apply_theme(&changed).is_err());
    signed.unsign();
    signed
        .apply_theme(&changed)
        .expect("explicitly unsign before family edit");
    assert!(!signed.is_signed());
}

#[test]
fn source_backed_family_projection_does_not_materialize_theme_media() {
    let mut package = native_workbook().opc_package().clone();
    let image_uri = PackURI::new("/xl/media/theme-family-test.png").expect("image URI");
    package.add_part(Box::new(BlobPart::new(
        image_uri,
        ct::PNG.to_owned(),
        vec![0xA5; 1 << 20],
    )));
    package
        .get_part_mut(&theme_uri())
        .expect("Theme part")
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/theme-family-test.png".to_owned(),
            "rIdFamilyImage".to_owned(),
            TargetMode::Internal,
        )
        .expect("Theme image edge");
    let mut bytes = Vec::new();
    package
        .to_stream(&mut bytes)
        .expect("serialized source package");
    let source = Arc::new(CountingSource {
        bytes,
        reads: AtomicUsize::new(0),
    });
    let workbook = SourceBackedWorkbook::from_read_at(source.clone()).expect("source workbook");
    let before = workbook.cache_diagnostics();
    let view = workbook
        .theme()
        .expect("source Theme")
        .expect("source Theme owner");
    assert_eq!(view.family().expect("source family").name(), "Office Theme");
    let after = workbook.cache_diagnostics();
    assert_eq!(after.cold_loads - before.cold_loads, 1);
    assert_eq!(after.retained_entries - before.retained_entries, 1);
    assert_eq!(after.retained_bytes - before.retained_bytes, 8_390);
    let reads = source.reads.load(Ordering::SeqCst);
    assert_eq!(
        view.family().expect("cached source family").name(),
        "Office Theme"
    );
    assert_eq!(source.reads.load(Ordering::SeqCst), reads);
}

fn saved_bytes(workbook: &Workbook) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).expect("save workbook");
    output.into_inner()
}
