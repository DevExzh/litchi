//! Formula evaluation through `File::evaluate` and `File::recalculate`
//! (record 0266), against values the producing applications cached.
#![cfg(all(
    feature = "xlsx",
    feature = "xlsb",
    feature = "xls",
    feature = "docx",
    feature = "eval"
))]

use std::path::PathBuf;

use litchi_core::FileFormat;
use litchi_ops::calc::{ErrorValue, Options, Origin, Status, Value};
use litchi_ops::{Code, File, props};

fn corpus(relative: &str) -> File {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative);
    litchi_ops::open(&path).unwrap_or_else(|error| panic!("open {relative}: {error}"))
}

#[test]
fn computes_uncached_formulas_and_reads_literals() {
    for format in [FileFormat::Xlsx, FileFormat::Xlsb, FileFormat::Xls] {
        let file = litchi_ops::create(format)
            .expect("create")
            .set("/sheet[1]/A1", props! { "value" => 20 })
            .expect("value")
            .set("/sheet[1]/A2", props! { "formula" => "SUM(A1,22)" })
            .expect("formula")
            .set("/sheet[1]/A3", props! { "formula" => "A2/0" })
            .expect("division");
        // The XLSB writer stores an empty string cache with a new formula, so
        // ask for recalculation in every format.
        let recalculate = Options::new().recalculate(true);
        let evaluated = file
            .evaluate("/sheet[1]/A2", &recalculate)
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        assert_eq!(evaluated.value(), &Value::Number(42.0), "{format:?}");
        assert_eq!(evaluated.origin(), Origin::Computed);
        let literal = file
            .evaluate("/sheet[1]/A1", &Options::new())
            .expect("literal");
        assert_eq!(
            (literal.value(), literal.origin()),
            (&Value::Number(20.0), Origin::Literal)
        );
        assert_eq!(
            file.evaluate("/sheet[1]/A3", &recalculate)
                .expect("error value")
                .value(),
            &Value::Error(ErrorValue::Div0)
        );
        // A blank cell is an empty literal.
        assert_eq!(
            file.evaluate("/sheet[1]/Z9", &Options::new())
                .expect("blank")
                .value(),
            &Value::Empty
        );
    }
}

#[test]
fn recalculation_matches_cached_results_in_each_format() {
    // The XLS workbook stores no cached value for five string results.
    for (relative, formulas, uncached) in [
        ("ooxml/xlsx/shared_formulas.xlsx", 40, 0),
        (
            "libreoffice-core/sc/qa/unit/data/xlsb/shared_formula.xlsb",
            30,
            0,
        ),
        ("ole/xls/ComplexFunctionTestCaseData.xls", 12, 5),
    ] {
        let file = corpus(relative);
        let before = file.to_bytes().expect("publish");
        let report = file.recalculate(&Options::new()).expect("recalculate");
        assert_eq!(report.formulas().len(), formulas, "{relative}");
        assert_eq!(report.uncached(), uncached, "{relative}: {report:?}");
        assert_eq!(
            report.matches(),
            formulas - uncached,
            "{relative}: {report:?}"
        );
        for formula in report.formulas() {
            assert!(!formula.text().is_empty());
            if formula.status() == Status::Match {
                assert_eq!(formula.cached(), formula.computed());
            }
        }
        // Evaluation is a read.
        assert_eq!(file.to_bytes().expect("publish"), before);
    }
}

#[test]
fn cached_values_are_preferred_unless_recalculation_is_asked_for() {
    let file = corpus("ooxml/xlsx/shared_formulas.xlsx");
    let report = file.recalculate(&Options::new()).expect("recalculate");
    let first = &report.formulas()[0];
    let path = format!(
        "/sheet[{}]/{}{}",
        first.sheet() + 1,
        column_label(first.column()),
        first.row() + 1
    );
    let cached = file
        .evaluate(path.as_str(), &Options::new())
        .expect("cached");
    assert_eq!(cached.origin(), Origin::Cached);
    assert_eq!(Some(cached.value()), first.cached());
    let computed = file
        .evaluate(path.as_str(), &Options::new().recalculate(true))
        .expect("computed");
    assert_eq!(computed.origin(), Origin::Computed);
    assert_eq!(Some(computed.value()), first.computed());
}

fn column_label(column: u32) -> String {
    let mut label = String::new();
    let mut column = column + 1;
    while column > 0 {
        label.insert(
            0,
            char::from(b'A' + u8::try_from((column - 1) % 26).unwrap()),
        );
        column = (column - 1) / 26;
    }
    label
}

#[test]
fn volatile_formulas_keep_their_cache() {
    let file = corpus("ole/xls/WithCustomViews.xls");
    let report = file.recalculate(&Options::new()).expect("recalculate");
    assert!(report.volatile() > 0);
    for formula in report.formulas() {
        if formula.status() == Status::Volatile {
            assert_eq!(formula.computed(), formula.cached());
        }
    }
    let uncached = litchi_ops::create(FileFormat::Xlsx)
        .expect("create")
        .set("/Sheet1/A1", props! { "formula" => "NOW()" })
        .expect("formula");
    let error = uncached
        .evaluate("/Sheet1/A1", &Options::new())
        .expect_err("volatile without a cache");
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let report = uncached.recalculate(&Options::new()).expect("recalculate");
    assert_eq!(report.volatile(), 1);
}

#[test]
fn unsupported_syntax_is_reported_not_guessed() {
    let file = litchi_ops::create(FileFormat::Xlsx)
        .expect("create")
        .set("/Sheet1/A1", props! { "value" => "a" })
        .expect("value")
        .set("/Sheet1/A2", props! { "formula" => "A1&\"b\"" })
        .expect("formula");
    assert_eq!(
        file.evaluate("/Sheet1/A2", &Options::new())
            .expect_err("concatenation operator")
            .code(),
        Code::UnsupportedCapability
    );
    let report = file.recalculate(&Options::new()).expect("recalculate");
    assert_eq!(report.unsupported(), 1);
    assert_eq!(report.formulas()[0].computed(), None);
}

#[test]
fn chains_are_bounded_and_forward_chains_resolve_in_order() {
    // A1 = 1 and each later row adds one: a 399-formula forward chain.
    let file = litchi_ops::create(FileFormat::Xlsx)
        .expect("create")
        .edit(|edit| {
            edit.set("/Sheet1/A1", props! { "value" => 1 })?;
            for row in 2..=400 {
                edit.set(
                    format!("/Sheet1/A{row}").as_str(),
                    props! { "formula" => format!("A{}+1", row - 1) },
                )?;
            }
            Ok(())
        })
        .expect("chain")
        .into_file();
    let deep = file
        .evaluate("/Sheet1/A400", &Options::new())
        .expect("computed in sheet order");
    assert_eq!(deep.value(), &Value::Number(400.0));
    let report = file.recalculate(&Options::new()).expect("in order");
    assert_eq!(report.uncached(), 399);
    assert_eq!(
        report
            .formulas()
            .last()
            .and_then(|formula| formula.computed()),
        Some(&Value::Number(400.0))
    );

    // A backward chain (each row reads the next) cannot be ordered away.
    let backward = litchi_ops::create(FileFormat::Xlsx)
        .expect("create")
        .edit(|edit| {
            for row in 1..=12 {
                edit.set(
                    format!("/Sheet1/A{row}").as_str(),
                    props! { "formula" => format!("A{}+1", row + 1) },
                )?;
            }
            Ok(())
        })
        .expect("backward chain")
        .into_file();
    let bounded = Options::new().max_depth(8);
    assert_eq!(
        backward
            .evaluate("/Sheet1/A1", &bounded)
            .expect_err("twelve deep")
            .code(),
        Code::LimitExceeded
    );
    assert_eq!(
        backward
            .evaluate("/Sheet1/A6", &bounded)
            .expect("seven deep")
            .value(),
        &Value::Number(7.0)
    );
    assert_eq!(
        backward
            .recalculate(&bounded)
            .expect_err("the first row is twelve deep")
            .code(),
        Code::LimitExceeded
    );
}

#[test]
fn cell_bound_refuses_large_reads() {
    let file = corpus("ooxml/xlsx/shared_formulas.xlsx");
    assert_eq!(
        file.recalculate(&Options::new().max_cells(5))
            .expect_err("bounded")
            .code(),
        Code::LimitExceeded
    );
}

#[test]
fn targets_and_formats_are_checked() {
    let file = litchi_ops::create(FileFormat::Xlsx).expect("create");
    assert_eq!(
        file.evaluate("/Sheet1", &Options::new())
            .expect_err("a sheet")
            .code(),
        Code::InvalidArgument
    );
    let document = litchi_ops::create(FileFormat::Docx).expect("create DOCX");
    assert_eq!(
        document
            .recalculate(&Options::new())
            .expect_err("not a workbook")
            .code(),
        Code::UnsupportedCapability
    );
}
