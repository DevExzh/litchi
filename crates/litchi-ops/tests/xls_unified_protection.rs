//! Every unified XLS P0 write refuses signed and encrypted workbooks with a
//! typed error and leaves the owner unchanged.
#![cfg(feature = "xls")]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "integration fixtures fail fast"
)]

use litchi_ops::{Code, File, NewNode, Placement, props};

fn builder() -> litchi_xls::Builder {
    let mut builder = litchi_xls::Builder::new();
    let data = builder.add_worksheet("Data").unwrap();
    builder.add_worksheet("Notes").unwrap();
    builder.write_number(data, 0, 0, 3.0).unwrap();
    builder.write_number(data, 1, 0, 4.0).unwrap();
    builder
}

fn signed() -> File {
    let bytes = builder().build().unwrap().to_bytes().unwrap();
    let mut ole = litchi_cfb::OleFile::open(
        std::io::Cursor::new(bytes),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let mut writer = litchi_cfb::OleWriter::new();
    for path in ole.list_streams() {
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        let data = ole.open_stream(&refs).unwrap();
        writer.create_stream(&refs, &data).unwrap();
    }
    writer
        .create_stream(&["_signatures"], b"synthetic")
        .unwrap();
    let mut signed = std::io::Cursor::new(Vec::new());
    writer.write_to(&mut signed).unwrap();
    let file = litchi_ops::open_bytes(signed.into_inner()).unwrap();
    assert!(file.is_signed());
    file
}

#[cfg(feature = "encryption")]
fn unlocked_encrypted() -> File {
    let mut builder = builder();
    builder
        .set_password("secret", litchi_xls::EncryptionProfile::OfficeBinaryRc4)
        .unwrap();
    builder.set_encryption_entropy(litchi_xls::EncryptionEntropy::generate().unwrap());
    let bytes = builder.build().unwrap().to_encrypted_bytes().unwrap();
    let File::Locked(locked) = litchi_ops::open_bytes(bytes).unwrap() else {
        panic!("encrypted XLS must open locked");
    };
    locked
        .unlock(litchi_ops::Password::new("secret".to_owned()))
        .unwrap()
}

fn refusals(file: &File) -> Vec<(&'static str, litchi_ops::Error)> {
    vec![
        (
            "add sheet",
            file.add("/", NewNode::new("sheet")).unwrap_err(),
        ),
        ("remove sheet", file.remove("/Notes").unwrap_err()),
        (
            "move sheet",
            file.move_to("/Notes", Placement::Start).unwrap_err(),
        ),
        ("swap sheets", file.swap("/Data", "/Notes").unwrap_err()),
        (
            "rename sheet",
            file.set("/Notes", props! { "name" => "Later" })
                .unwrap_err(),
        ),
        (
            "formula",
            file.set("/Data/B1", props! { "formula" => "A1*2" })
                .unwrap_err(),
        ),
        (
            "row height",
            file.set("/Data/row[1]", props! { "height" => "24pt" })
                .unwrap_err(),
        ),
        (
            "row hidden",
            file.set("/Data/row[2]", props! { "hidden" => true })
                .unwrap_err(),
        ),
        (
            "column width",
            file.set("/Data/col[A]", props! { "width" => 12 })
                .unwrap_err(),
        ),
        (
            "cell format",
            file.set("/Data/A1", props! { "alignment.horizontal" => "right" })
                .unwrap_err(),
        ),
        (
            "row insertion",
            file.add("/Data", NewNode::new("row").at(Placement::Start))
                .unwrap_err(),
        ),
    ]
}

#[test]
fn signed_workbooks_refuse_every_p0_write() {
    let file = signed();
    let before = file.to_bytes().unwrap();
    for (operation, error) in refusals(&file) {
        assert_eq!(
            error.code(),
            Code::DocumentProtected,
            "{operation}: {error}"
        );
    }
    assert_eq!(file.to_bytes().unwrap(), before);
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_workbooks_refuse_every_p0_write() {
    let file = unlocked_encrypted();
    for (operation, error) in refusals(&file) {
        assert_eq!(
            error.code(),
            Code::DocumentProtected,
            "{operation}: {error}"
        );
    }
    let text = file
        .get("/Data/A1")
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(std::borrow::Cow::into_owned);
    assert_eq!(text.as_deref(), Some("3"));
}
