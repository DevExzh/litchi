//! Unified XLS cell formatting reads and writes through the public [`File`]
//! API on both owner routes.
#![cfg(feature = "xls")]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "integration fixtures fail fast"
)]

use std::ops::Deref;

use litchi_ops::{Code, File, OpenOptions, Props, SourceMode, props};

fn fixture() -> Vec<u8> {
    let mut builder = litchi_xls::Builder::new();
    let data = builder.add_worksheet("Data").unwrap();
    builder.write_number(data, 0, 0, 1.5).unwrap();
    builder.write_string(data, 1, 0, "text").unwrap();
    builder.write_formula(data, 2, 0, "A1*2").unwrap();
    builder.build().unwrap().to_bytes().unwrap()
}

struct Owner {
    file: File,
    _guard: Option<tempfile::TempDir>,
}

impl Deref for Owner {
    type Target = File;

    fn deref(&self) -> &File {
        &self.file
    }
}

fn open_owner(bytes: &[u8], eager: bool) -> Owner {
    if eager {
        return Owner {
            file: litchi_ops::open_bytes_with(
                bytes.to_vec(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )
            .unwrap(),
            _guard: None,
        };
    }
    let guard = tempfile::tempdir().unwrap();
    let path = guard.path().join("source.xls");
    std::fs::write(&path, bytes).unwrap();
    Owner {
        file: litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
            .unwrap(),
        _guard: Some(guard),
    }
}

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(std::borrow::Cow::into_owned)
        .unwrap_or_default()
}

fn styled(file: &File) -> File {
    file.set(
        "/Data/A1",
        props! {
            "alignment.horizontal" => "center",
            "alignment.vertical" => "top",
            "alignment.wrapText" => true,
            "alignment.indent" => 2,
            "border.bottom" => "thin",
            "border.left" => "double",
            "fill" => "#FF0000",
        },
    )
    .unwrap()
}

#[test]
fn formatting_writes_read_back_on_both_routes() {
    let bytes = fixture();
    for eager in [false, true] {
        let owner = open_owner(&bytes, eager);
        assert_eq!(property(&owner, "/Data/A1", "alignment.horizontal"), None);
        let written = styled(&owner)
            .set("/Data/E5", props! { "border.top" => "medium" })
            .unwrap();
        let reopened_eager = litchi_ops::open_bytes(written.to_bytes().unwrap()).unwrap();
        let guard = tempfile::tempdir().unwrap();
        let path = guard.path().join("styled.xls");
        std::fs::write(&path, written.to_bytes().unwrap()).unwrap();
        let reopened_source = litchi_ops::open(&path).unwrap();
        for file in [&written, &reopened_eager, &reopened_source] {
            assert_eq!(
                property(file, "/Data/A1", "alignment.horizontal").as_deref(),
                Some("center")
            );
            assert_eq!(
                property(file, "/Data/A1", "alignment.vertical").as_deref(),
                Some("top")
            );
            assert_eq!(
                property(file, "/Data/A1", "alignment.wrapText").as_deref(),
                Some("true")
            );
            assert_eq!(
                property(file, "/Data/A1", "alignment.indent").as_deref(),
                Some("2")
            );
            assert_eq!(
                property(file, "/Data/A1", "border.bottom").as_deref(),
                Some("thin")
            );
            assert_eq!(
                property(file, "/Data/A1", "border.left").as_deref(),
                Some("double")
            );
            assert_eq!(
                property(file, "/Data/A1", "fill").as_deref(),
                Some("#FF0000")
            );
            assert_eq!(property(file, "/Data/A1", "fillPattern"), None);
            assert_eq!(text(file, "/Data/A1"), "1.5");
            // An absent coordinate became a formatting-only blank cell.
            assert_eq!(
                property(file, "/Data/E5", "border.top").as_deref(),
                Some("medium")
            );
            assert_eq!(property(file, "/Data/E5", "empty").as_deref(), Some("true"));
        }
        assert_eq!(owner.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn formatting_ranges_formula_cells_and_inverse() {
    let bytes = fixture();
    let owner = open_owner(&bytes, true);
    let commit = owner
        .edit(|tx| {
            tx.set("/Data/A1:A3", props! { "alignment.horizontal" => "right" })?;
            tx.set(
                "/Data/A1",
                props! { "border.bottom" => litchi_ops::PropValue::Null },
            )?;
            Ok::<_, litchi_ops::Error>(())
        })
        .unwrap();
    for cell in ["A1", "A2", "A3"] {
        assert_eq!(
            property(
                commit.file(),
                &format!("/Data/{cell}"),
                "alignment.horizontal"
            )
            .as_deref(),
            Some("right"),
            "{cell}"
        );
    }
    // The formula cell keeps its formula.
    assert_eq!(
        property(commit.file(), "/Data/A3", "formula").as_deref(),
        Some("(A1*2)")
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        bytes
    );
    // Re-applying the current alignment is an exact no-op.
    let same = commit
        .file()
        .set("/Data/A2", props! { "alignment.horizontal" => "right" })
        .unwrap();
    assert_eq!(same.to_bytes().unwrap(), commit.file().to_bytes().unwrap());
}

#[test]
fn unsupported_and_invalid_formatting_is_typed_and_atomic() {
    let bytes = fixture();
    let owner = open_owner(&bytes, true);
    let cases: Vec<(Props<'static>, Code)> = vec![
        (
            Props::from_props(props! { "fill" => "#123456" }).unwrap(),
            Code::InvalidValue,
        ),
        (
            Props::from_props(props! { "alignment.horizontal" => "sideways" }).unwrap(),
            Code::InvalidValue,
        ),
        (
            Props::from_props(props! { "value" => 2, "alignment.horizontal" => "left" }).unwrap(),
            Code::InvalidArgument,
        ),
    ];
    for (properties, code) in cases {
        let error = owner.set("/Data/A1", properties).unwrap_err();
        assert_eq!(error.code(), code, "{error}");
    }
    for key in ["font.bold", "protection.locked", "numberformat"] {
        let error = owner
            .set(
                "/Data/A1",
                Props::from_props(props! { key => true }).unwrap(),
            )
            .unwrap_err();
        assert!(
            !matches!(error.code(), Code::InternalError),
            "{key}: {error}"
        );
    }
    assert_eq!(owner.to_bytes().unwrap(), bytes);
}
