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
