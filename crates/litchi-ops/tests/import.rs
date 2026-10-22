//! CSV/TSV import through `File::import` (record 0266): one atomic commit per
//! import into XLSX, XLSB and XLS, exact no-ops, typed refusals and limits.
#![cfg(all(feature = "xlsx", feature = "xlsb", feature = "xls", feature = "docx"))]

use litchi_core::FileFormat;
use litchi_ops::import::{Encoding, Options};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{Code, File, props};

const WORKBOOKS: [FileFormat; 3] = [FileFormat::Xlsx, FileFormat::Xlsb, FileFormat::Xls];

fn text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .text()
        .unwrap_or_else(|error| panic!("read {path}: {error}"))
        .map(std::borrow::Cow::into_owned)
}

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path}: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

/// No value: absent, or the projection's empty-cell rendering.
fn blank(file: &File, path: &str) -> bool {
    text(file, path).is_none_or(|value| value.is_empty() || value == "(empty)")
}

fn reopen(file: &File) -> File {
    litchi_ops::open_bytes(file.to_bytes().expect("publish")).expect("reopen")
}

const SOURCE: &[u8] =
    b"name,score,active,joined\nAda,36,TRUE,2026-10-13\n\"Lovelace, A.\",-2.5,false,1815-12-10\n";

#[test]
fn imports_typed_values_into_every_workbook_format() {
    for format in WORKBOOKS {
        let original = litchi_ops::create(format).expect("create");
        let before = original.to_bytes().expect("publish original");
        let commit = original
            .import("/sheet[1]", SOURCE, &Options::csv().header(true))
            .unwrap_or_else(|error| panic!("{format:?} import: {error}"));
        let summary = commit.value();
        assert_eq!((summary.rows(), summary.columns()), (3, 4), "{format:?}");
        assert_eq!(summary.cells(), 12, "{format:?}");
        assert_eq!(summary.numbers(), 2, "{format:?}");
        assert_eq!(summary.booleans(), 2, "{format:?}");
        assert!(commit.diagnostics().changed());
        assert_eq!(
            commit.diagnostics().commits(),
            1,
            "{format:?}: one native commit"
        );
        // The source handle is untouched.
        assert_eq!(original.to_bytes().expect("publish original"), before);

        for file in [commit.file().clone(), reopen(commit.file())] {
            assert_eq!(text(&file, "/sheet[1]/A1").as_deref(), Some("name"));
            assert_eq!(text(&file, "/sheet[1]/A3").as_deref(), Some("Lovelace, A."));
            assert_eq!(text(&file, "/sheet[1]/B2").as_deref(), Some("36"));
            assert_eq!(text(&file, "/sheet[1]/B3").as_deref(), Some("-2.5"));
            assert_eq!(
                property(&file, "/sheet[1]/B2", "type").as_deref(),
                Some("Number")
            );
            assert_eq!(
                property(&file, "/sheet[1]/C2", "type").as_deref(),
                Some("Boolean"),
                "{format:?}"
            );
            // 1815 predates every date system: it stays text everywhere.
            assert_eq!(text(&file, "/sheet[1]/D3").as_deref(), Some("1815-12-10"));
        }
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse");
        assert_eq!(restored.to_bytes().expect("publish restored"), before);
    }
}

#[test]
fn dates_become_serials_with_a_date_format_where_the_format_can_write_one() {
    for format in [FileFormat::Xlsx, FileFormat::Xlsb] {
        let file = litchi_ops::create(format).expect("create");
        let commit = file
            .import(
                "/sheet[1]",
                b"2026-10-13,2026-10-13T12:00:00\n",
                &Options::csv(),
            )
            .expect("import");
        assert_eq!(commit.value().dates(), 2);
        let file = reopen(commit.file());
        assert_eq!(
            property(&file, "/sheet[1]/A1", "type").as_deref(),
            Some("Number")
        );
        let format_code = property(&file, "/sheet[1]/A1", "numberformat")
            .or_else(|| property(&file, "/sheet[1]/A1", "numFmt"));
        assert_eq!(format_code.as_deref(), Some("yyyy-mm-dd"), "{format:?}");
        let raw = file
            .workbook()
            .expect("workbook")
            .worksheet_names()
            .expect("names");
        assert_eq!(raw.len(), 1);
    }
    // XLS keeps the ISO text.
    let xls = litchi_ops::create(FileFormat::Xls).expect("create XLS");
    let commit = xls
        .import("/sheet[1]", b"2026-10-13\n", &Options::csv())
        .expect("import XLS");
    assert_eq!(commit.value().dates_as_text(), 1);
    assert_eq!(
        text(commit.file(), "/sheet[1]/A1").as_deref(),
        Some("2026-10-13")
    );
}

#[test]
fn origin_cell_offsets_the_rectangle_and_the_grid_is_checked() {
    for format in WORKBOOKS {
        let file = litchi_ops::create(format).expect("create");
        let commit = file
            .import("/sheet[1]/B2", b"1,2\n3,4\n", &Options::csv())
            .expect("import at B2");
        assert_eq!(
            (commit.value().row(), commit.value().column()),
            (1, 1),
            "{format:?}"
        );
        assert_eq!(text(commit.file(), "/sheet[1]/C3").as_deref(), Some("4"));
        assert!(blank(commit.file(), "/sheet[1]/A1"), "{format:?}");
    }
    let xls = litchi_ops::create(FileFormat::Xls).expect("create XLS");
    let error = xls
        .import("/sheet[1]/IV1", b"1,2\n", &Options::csv())
        .expect_err("past column IV");
    assert_eq!(error.code(), Code::OutOfRange);
    let xlsx = litchi_ops::create(FileFormat::Xlsx).expect("create XLSX");
    let error = xlsx
        .import("/sheet[1]/XFD1048576", b"1\n2\n", &Options::csv())
        .expect_err("past the last row");
    assert_eq!(error.code(), Code::OutOfRange);
}

#[test]
fn empty_input_is_an_exact_noop_that_still_validates_the_target() {
    for format in WORKBOOKS {
        let file = litchi_ops::create(format).expect("create");
        let before = file.to_bytes().expect("publish");
        for source in [&b""[..], b"\xEF\xBB\xBF", b"sep=;\n"] {
            let commit = file
                .import("/sheet[1]", source, &Options::csv())
                .expect("empty import");
            assert!(commit.patch().is_noop(), "{format:?}");
            assert!(!commit.diagnostics().changed());
            assert_eq!(commit.diagnostics().commits(), 0);
            assert_eq!(commit.value().cells(), 0);
            assert_eq!(commit.file().to_bytes().expect("publish"), before);
        }
        let error = file
            .import("/sheet[9]", b"", &Options::csv())
            .expect_err("missing sheet");
        assert_eq!(error.code(), Code::PathNotFound, "{format:?}");
    }
}

#[test]
fn empty_fields_clear_existing_values_and_skip_absent_cells() {
    for format in WORKBOOKS {
        let file = litchi_ops::create(format)
            .expect("create")
            .set("/sheet[1]/A1", props! { "value" => "old" })
            .expect("seed");
        let commit = file
            .import("/sheet[1]", b",new\n,\n", &Options::csv())
            .expect("import");
        assert_eq!(commit.value().empty(), 3, "{format:?}");
        assert_eq!(commit.value().cells(), 1);
        let reopened = reopen(commit.file());
        assert_eq!(text(&reopened, "/sheet[1]/B1").as_deref(), Some("new"));
        assert!(
            blank(&reopened, "/sheet[1]/A1"),
            "{format:?}: the empty field clears A1"
        );
    }
}

#[test]
fn formulas_are_opt_in() {
    for format in WORKBOOKS {
        let file = litchi_ops::create(format).expect("create");
        let plain = file
            .import("/sheet[1]", b"1,=A1+1\n", &Options::csv())
            .expect("plain import");
        assert_eq!(text(plain.file(), "/sheet[1]/B1").as_deref(), Some("=A1+1"));
        assert_eq!(property(plain.file(), "/sheet[1]/B1", "formula"), None);
        let formulas = file
            .import("/sheet[1]", b"1,=A1+1\n", &Options::csv().formulas(true))
            .expect("formula import");
        assert_eq!(formulas.value().formulas(), 1);
        // XLSB reads formulas back by decompiling its tokens, which may add
        // redundant parentheses.
        let formula = property(formulas.file(), "/sheet[1]/B1", "formula");
        assert!(
            matches!(formula.as_deref(), Some("A1+1" | "(A1+1)")),
            "{format:?}: {formula:?}"
        );
    }
}

#[test]
fn tsv_utf16_and_disabled_inference() {
    for format in WORKBOOKS {
        let file = litchi_ops::create(format).expect("create");
        let source: Vec<u8> = "\u{FEFF}α\t007\r\nβ\tTRUE\r\n"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let commit = file
            .import("/sheet[1]", &source, &Options::tsv().infer(false))
            .expect("import");
        assert_eq!(commit.value().texts(), 4);
        assert_eq!(text(commit.file(), "/sheet[1]/B1").as_deref(), Some("007"));
        assert_eq!(text(commit.file(), "/sheet[1]/A2").as_deref(), Some("β"));
        let explicit = file
            .import(
                "/sheet[1]",
                &source,
                &Options::tsv().encoding(Encoding::Utf16Le),
            )
            .expect("explicit encoding");
        assert_eq!(explicit.value().numbers(), 1, "{format:?}");
    }
}

#[test]
fn failures_are_typed_and_leave_the_input_unchanged() {
    let file = litchi_ops::create(FileFormat::Xlsx).expect("create");
    let before = file.to_bytes().expect("publish");
    for (source, options, code) in [
        (&b"\"open"[..], Options::csv(), Code::InvalidInput),
        (b"a\xFFb", Options::csv(), Code::InvalidInput),
        (b"a", Options::csv().delimiter('"'), Code::InvalidArgument),
    ] {
        assert_eq!(
            file.import("/sheet[1]", source, &options)
                .expect_err("refused")
                .code(),
            code
        );
    }
    assert_eq!(
        file.import("/sheet[1]/A1:B2", b"1", &Options::csv())
            .expect_err("a range is not a target")
            .code(),
        Code::InvalidArgument
    );
    let limits = SemanticLimits::builder()
        .max_expanded_targets(2)
        .expect("limit")
        .build()
        .expect("limits");
    assert_eq!(
        file.import_with_options(
            "/sheet[1]",
            b"1,2,3",
            &Options::csv(),
            litchi_ops::EditOptions::new().semantic_limits(limits)
        )
        .expect_err("over the target bound")
        .code(),
        Code::LimitExceeded
    );
    assert_eq!(file.to_bytes().expect("publish"), before);

    let document = litchi_ops::create(FileFormat::Docx).expect("create DOCX");
    assert_eq!(
        document
            .import("/body", b"1", &Options::csv())
            .expect_err("not a workbook")
            .code(),
        Code::UnsupportedCapability
    );
}

#[test]
fn xls_refuses_to_overwrite_formula_cells_atomically() {
    let file = litchi_ops::create(FileFormat::Xls)
        .expect("create")
        .set("/sheet[1]/B1", props! { "formula" => "1+1" })
        .expect("formula");
    let before = file.to_bytes().expect("publish");
    let error = file
        .import("/sheet[1]", b"x,y\n", &Options::csv())
        .expect_err("over a formula");
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(file.to_bytes().expect("publish"), before);
}

#[test]
fn import_composes_inside_a_closure_transaction() {
    let file = litchi_ops::create(FileFormat::Xlsx).expect("create");
    let commit = file
        .edit(|edit| {
            let first = edit.import("/sheet[1]", b"1,2\n", &Options::csv())?;
            edit.set("/sheet[1]/C1", props! { "value" => 3 })?;
            let second = edit.import("/sheet[1]/A2", b"4\n", &Options::csv())?;
            Ok(first.cells() + second.cells())
        })
        .expect("edit");
    assert_eq!(*commit.value(), 3);
    assert_eq!(text(commit.file(), "/sheet[1]/C1").as_deref(), Some("3"));
    assert_eq!(text(commit.file(), "/sheet[1]/A2").as_deref(), Some("4"));
}

#[test]
fn large_imports_fit_every_format_in_one_unified_commit() {
    // XLS stages at most 1,024 cells per native edit (its editor admits
    // 4,096 staged changes), so 1,200 cells chain two native commits; BIFF8
    // insertion grows with the sheet, so XLS takes the smaller table.
    let limits = SemanticLimits::builder()
        .max_expanded_targets(10_000)
        .expect("limit")
        .build()
        .expect("limits");
    for (format, rows) in [
        (FileFormat::Xlsx, 3_000),
        (FileFormat::Xlsb, 3_000),
        (FileFormat::Xls, 400),
    ] {
        let mut source = String::new();
        for row in 0..rows {
            source.push_str(&format!("{row},name {row},{}.5\n", row % 7));
        }
        let file = litchi_ops::create(format).expect("create");
        let commit = file
            .import_with_options(
                "/sheet[1]",
                source.as_bytes(),
                &Options::csv(),
                litchi_ops::EditOptions::new().semantic_limits(limits),
            )
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        assert_eq!(commit.value().cells(), rows * 3);
        assert_eq!(commit.diagnostics().operations(), 1);
        if format == FileFormat::Xls {
            assert_eq!(commit.diagnostics().commits(), 2);
        }
        let last = rows - 1;
        let reopened = reopen(commit.file());
        assert_eq!(
            text(&reopened, &format!("/sheet[1]/B{rows}")),
            Some(format!("name {last}"))
        );
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.file())
            .expect("inverse");
        assert_eq!(
            restored.to_bytes().expect("publish"),
            file.to_bytes().expect("publish")
        );
    }
}
