//! Unified XLS row and column layout reads and writes through the public
//! [`File`] API on both owner routes.
#![cfg(feature = "xls")]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "integration fixtures fail fast"
)]

use std::ops::Deref;

use litchi_ops::{Code, File, OpenOptions, SourceMode, props};
use litchi_xls::writer::{Column, Row};

fn fixture() -> Vec<u8> {
    let mut builder = litchi_xls::Builder::new();
    let data = builder.add_worksheet("Data").unwrap();
    builder.write_number(data, 0, 0, 1.0).unwrap();
    builder.write_string(data, 2, 1, "two").unwrap();
    builder
        .set_row_height(data, Row::new(2).unwrap(), 30.0)
        .unwrap();
    builder
        .set_column_width(data, Column::new(1).unwrap(), 20.0)
        .unwrap();
    builder.hide_column(data, Column::new(4).unwrap()).unwrap();
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

#[test]
fn layout_reads_on_both_routes() {
    let bytes = fixture();
    for eager in [false, true] {
        let owner = open_owner(&bytes, eager);
        assert_eq!(
            property(&owner, "/Data/row[3]", "height").as_deref(),
            Some("30pt")
        );
        assert_eq!(
            property(&owner, "/Data/row[3]", "customHeight").as_deref(),
            Some("true")
        );
        assert_eq!(property(&owner, "/Data/row[3]", "hidden"), None);
        assert_eq!(
            property(&owner, "/Data/col[B]", "width").as_deref(),
            Some("20")
        );
        assert_eq!(
            property(&owner, "/Data/col[E]", "hidden").as_deref(),
            Some("true")
        );
        assert_eq!(property(&owner, "/Data/col[C]", "width"), None);
    }
}

#[test]
fn row_and_column_layout_writes_round_trip() {
    let bytes = fixture();
    for eager in [false, true] {
        let owner = open_owner(&bytes, eager);
        let written = owner
            .set("/Data/row[1]", props! { "height" => "24pt" })
            .unwrap()
            .set("/Data/row[10]", props! { "hidden" => true })
            .unwrap()
            .set("/Data/col[C]", props! { "width" => 12.5 })
            .unwrap()
            .set("/Data/col[E]", props! { "hidden" => false })
            .unwrap();
        let reopened = litchi_ops::open_bytes(written.to_bytes().unwrap()).unwrap();
        for file in [&written, &reopened] {
            assert_eq!(
                property(file, "/Data/row[1]", "height").as_deref(),
                Some("24pt")
            );
            assert_eq!(
                property(file, "/Data/row[1]", "customHeight").as_deref(),
                Some("true")
            );
            assert_eq!(
                property(file, "/Data/row[10]", "hidden").as_deref(),
                Some("true")
            );
            assert_eq!(
                property(file, "/Data/col[C]", "width").as_deref(),
                Some("12.5")
            );
            assert_eq!(
                property(file, "/Data/col[C]", "customWidth").as_deref(),
                Some("true")
            );
            assert_eq!(property(file, "/Data/col[E]", "hidden"), None);
            assert_eq!(
                property(file, "/Data/col[B]", "width").as_deref(),
                Some("20")
            );
        }
        assert_eq!(owner.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn layout_transactions_invert_exactly() {
    let bytes = fixture();
    let owner = open_owner(&bytes, true);
    let commit = owner
        .edit(|tx| {
            tx.set("/Data/A1", props! { "value" => 5 })?;
            tx.set(
                "/Data/row[1]",
                props! { "height" => "40pt", "hidden" => true },
            )?;
            tx.set("/Data/col[A]", props! { "width" => 30 })?;
            Ok::<_, litchi_ops::Error>(())
        })
        .unwrap();
    assert_eq!(
        property(commit.file(), "/Data/row[1]", "hidden").as_deref(),
        Some("true")
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
    // Empty property sets are exact no-ops.
    let same = owner.set("/Data/row[2]", props! {}).unwrap();
    assert_eq!(same.to_bytes().unwrap(), bytes);
}

#[test]
fn invalid_layout_values_are_typed_and_atomic() {
    let bytes = fixture();
    let owner = open_owner(&bytes, true);
    for (path, properties) in [
        (
            "/Data/row[1]",
            litchi_ops::Props::from_props(props! { "height" => "0pt" }).unwrap(),
        ),
        (
            "/Data/row[1]",
            litchi_ops::Props::from_props(props! { "height" => "500pt" }).unwrap(),
        ),
        (
            "/Data/col[A]",
            litchi_ops::Props::from_props(props! { "width" => 300 }).unwrap(),
        ),
        (
            "/Data/col[A]",
            litchi_ops::Props::from_props(props! { "outline" => 1 }).unwrap(),
        ),
    ] {
        let error = owner.set(path, properties).unwrap_err();
        assert_ne!(error.code(), Code::InternalError, "{path}: {error}");
        assert!(
            matches!(
                error.code(),
                Code::InvalidValue | Code::OutOfRange | Code::UnsupportedProperty
            ),
            "{path}: {error}"
        );
    }
    assert_eq!(owner.to_bytes().unwrap(), bytes);
}

/// Record 0310: a layout write to a row or column named by its number
/// resolves the coordinates without materializing the worksheet grid. The
/// published bytes are still exactly the native edit's on both owner
/// routes, and `last()` selectors, which do read the grid, still resolve.
#[test]
fn layout_path_writes_publish_the_native_edit_bytes() {
    let bytes = fixture();
    let native = |edit: &dyn Fn(&mut litchi_xls::Edit)| {
        let package = litchi_xls::Package::from_bytes(&bytes).unwrap();
        let mut native = package.edit().unwrap();
        edit(&mut native);
        native.commit().unwrap().into_package().to_bytes().unwrap()
    };
    let row = native(&|edit| edit.set_row_height("Data", 0, 480).unwrap());
    let column = native(&|edit| edit.set_column_width("Data", 2, 3200).unwrap());
    for eager in [false, true] {
        let owner = open_owner(&bytes, eager);
        let written = owner
            .set("/Data/row[1]", props! { "height" => "24pt" })
            .unwrap();
        assert_eq!(written.to_bytes().unwrap(), row, "eager={eager}");
        let written = owner
            .set("/Data/col[C]", props! { "width" => 12.5 })
            .unwrap();
        assert_eq!(written.to_bytes().unwrap(), column, "eager={eager}");
        let last = owner
            .set("/Data/row[last()]", props! { "height" => "24pt" })
            .unwrap();
        assert_eq!(
            property(&last, "/Data/row[3]", "height").as_deref(),
            Some("24pt"),
            "eager={eager}"
        );
        assert_eq!(owner.to_bytes().unwrap(), bytes);
    }
}
