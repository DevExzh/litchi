#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "native package fixtures are checked and fail fast"
)]

//! XLSB host coverage for source-preserving Theme Family removal.

use std::io::Cursor;
use std::path::Path;

use litchi_opc::PackURI;
use litchi_xlsb::Workbook;

const FIXTURE: &str = "test-data/ooxml/xlsb/hyperlink.xlsb";
const THEME_NAME: &str = "/xl/theme/theme1.xml";
const FAMILY_URI: &str = "{05A4C25C-085E-4340-85A3-A5531E510DB2}";
const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn fixture() -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(FIXTURE)
}

fn workbook() -> Workbook {
    Workbook::new(Cursor::new(
        std::fs::read(fixture()).expect("native XLSB fixture"),
    ))
    .expect("native XLSB workbook")
}

fn theme_uri() -> PackURI {
    PackURI::new(THEME_NAME).expect("Theme URI")
}

fn theme_bytes(workbook: &Workbook) -> Vec<u8> {
    workbook
        .opc_package()
        .get_part(&theme_uri())
        .expect("Theme part")
        .blob()
        .to_vec()
}

fn saved_bytes(workbook: &Workbook) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).expect("save XLSB workbook");
    output.into_inner()
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

fn theme_extension_fragment(source: &[u8]) -> &[u8] {
    let start = source
        .windows(b"<a:ext uri=".len())
        .position(|window| window == b"<a:ext uri=")
        .expect("extension element");
    let end = source[start..]
        .windows(b"</a:ext>".len())
        .position(|window| window == b"</a:ext>")
        .map(|offset| start + offset + b"</a:ext>".len())
        .expect("extension close");
    &source[start..end]
}

fn with_root_namespace(source: &[u8], declaration: &str) -> Vec<u8> {
    let marker = b"<a:theme ";
    let start = source
        .windows(marker.len())
        .position(|window| window == marker)
        .expect("Theme root");
    let insertion = start + marker.len();
    let mut output = Vec::with_capacity(source.len() + declaration.len());
    output.extend_from_slice(&source[..insertion]);
    output.extend_from_slice(declaration.as_bytes());
    output.extend_from_slice(&source[insertion..]);
    output
}

fn direct_extension_mce(source: &[u8]) -> Vec<u8> {
    let family = std::str::from_utf8(theme_family_fragment(source)).expect("family UTF-8");
    let replacement = format!(
        r#"<a:ext uri="{FAMILY_URI}"><mc:AlternateContent><mc:Choice Requires="a">{family}</mc:Choice><mc:Fallback/></mc:AlternateContent></a:ext>"#
    );
    let replaced = replace_bytes(
        source,
        theme_extension_fragment(source),
        replacement.as_bytes(),
    );
    with_root_namespace(&replaced, &format!(r#"xmlns:mc="{MC_NAMESPACE}" "#))
}

#[test]
fn native_family_removal_closes_wrappers_and_save_reopen_inverse_is_exact() {
    let mut workbook = workbook();
    let original = theme_bytes(&workbook);
    let snapshot = workbook.theme().expect("read Theme").expect("native Theme");
    let mut edit = snapshot.edit();
    assert!(edit.remove_family().expect("stage native family removal"));
    let removal = edit.commit().expect("commit native family removal");
    workbook
        .apply_theme(&removal)
        .expect("publish native family removal");

    let removed = theme_bytes(&workbook);
    let removed_text = std::str::from_utf8(&removed).expect("Theme UTF-8");
    assert!(!removed_text.contains("<a:extLst>"));
    assert!(!removed_text.contains(FAMILY_URI));
    assert!(
        workbook
            .theme()
            .expect("read removed Theme")
            .expect("removed Theme")
            .family()
            .is_none()
    );

    let reopened_removed =
        Workbook::new(Cursor::new(saved_bytes(&workbook))).expect("reopen removed XLSB workbook");
    assert_eq!(theme_bytes(&reopened_removed), removed);
    assert!(
        reopened_removed
            .theme()
            .expect("read reopened removed Theme")
            .expect("reopened removed Theme")
            .family()
            .is_none()
    );

    workbook
        .apply_theme_patch(&removal.patch().inverse())
        .expect("publish exact family removal inverse");
    assert_eq!(theme_bytes(&workbook), original);
    let reopened_restored =
        Workbook::new(Cursor::new(saved_bytes(&workbook))).expect("reopen restored XLSB workbook");
    assert_eq!(theme_bytes(&reopened_restored), original);
    assert_eq!(
        reopened_restored
            .theme()
            .expect("read restored Theme")
            .expect("restored Theme")
            .family()
            .expect("restored family")
            .name(),
        "Office Theme"
    );
}

#[test]
fn xlsb_removal_retains_comment_and_foreign_payload_wrappers() {
    let mut workbook = workbook();
    let source = theme_bytes(&workbook);
    let source = String::from_utf8(source)
        .expect("Theme UTF-8")
        .replacen(
            "</a:ext>",
            "<x:payload xmlns:x=\"urn:litchi:xlsb-removal\" marker=\"keep\"/><!-- keep wrapper --></a:ext>",
            1,
        )
        .replace("?>\r\n", "?>")
        .into_bytes();
    workbook
        .edit_opc(|package| {
            package
                .get_part_mut(&theme_uri())
                .expect("Theme part")
                .set_blob(source);
            Ok(())
        })
        .expect("replace Theme source");

    let snapshot = workbook
        .theme()
        .expect("read payload Theme")
        .expect("payload Theme");
    let mut edit = snapshot.edit();
    assert!(edit.remove_family().expect("stage payload family removal"));
    let removal = edit.commit().expect("commit payload family removal");
    workbook
        .apply_theme(&removal)
        .expect("publish payload family removal");

    let removed = theme_bytes(&workbook);
    let removed_text = std::str::from_utf8(&removed).expect("Theme UTF-8");
    assert!(removed_text.contains("<a:extLst>"));
    assert!(removed_text.contains("<a:ext "));
    assert!(removed_text.contains("x:payload"));
    assert!(removed_text.contains("<!-- keep wrapper -->"));
    assert!(!removed_text.contains("themeFamily"));
}

#[test]
fn xlsb_add_refuses_direct_extension_mce_child_without_source_mutation() {
    let mut workbook = workbook();
    let source = direct_extension_mce(&theme_bytes(&workbook));
    let source = String::from_utf8(source)
        .expect("Theme UTF-8")
        .replace("?>\r\n", "?>")
        .into_bytes();
    workbook
        .edit_opc(|package| {
            package
                .get_part_mut(&theme_uri())
                .expect("Theme part")
                .set_blob(source);
            Ok(())
        })
        .expect("replace Theme source");

    let before_bytes = theme_bytes(&workbook);
    let before = workbook
        .theme()
        .expect("read MCE Theme")
        .expect("MCE Theme");
    assert!(before.family().is_none());

    let mut edit = before.edit();
    edit.set_family(
        litchi_xlsb::theme::Family::new(
            "Fresh Direct Family",
            "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}",
            "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}",
        )
        .expect("valid Theme family"),
    )
    .expect("stage direct family");
    assert!(
        edit.commit().is_err(),
        "an MCE child under a recognized ext must refuse mutation"
    );
    assert_eq!(
        theme_bytes(&workbook),
        before_bytes,
        "failed mutation changed source"
    );
}

#[test]
fn native_family_name_edit_preserves_declaration_crlf_through_save_and_inverse() {
    let mut workbook = workbook();
    let original = theme_bytes(&workbook);
    assert!(original.starts_with(b"<?xml "));
    assert!(original.windows(4).any(|bytes| bytes == b"?>\r\n"));
    let snapshot = workbook.theme().expect("read Theme").expect("native Theme");
    let mut family = snapshot.family().expect("native family").clone();
    family.set_name("Saved Native Family").expect("name");
    let old_fragment = theme_family_fragment(&original);
    let expected_fragment = replace_bytes(
        old_fragment,
        b"name=\"Office Theme\"",
        b"name=\"Saved Native Family\"",
    );
    let expected = replace_bytes(&original, old_fragment, &expected_fragment);
    let mut edit = snapshot.edit();
    edit.set_family(family).expect("stage name edit");
    let commit = edit.commit().expect("commit name edit");
    workbook.apply_theme(&commit).expect("publish edit");
    assert_eq!(theme_bytes(&workbook), expected);
    let reopened = Workbook::new(Cursor::new(saved_bytes(&workbook))).expect("reopen edit");
    assert_eq!(theme_bytes(&reopened), expected);
    workbook
        .apply_theme_patch(&commit.patch().inverse())
        .expect("publish inverse");
    let restored = Workbook::new(Cursor::new(saved_bytes(&workbook))).expect("reopen inverse");
    assert_eq!(theme_bytes(&restored), original);
}
