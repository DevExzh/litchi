//! Native Excel cached values versus the bounded, read-only evaluator.
#![cfg(all(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "eval"))]
use litchi_ops::calc::{Options, Origin, Status, Value};
use litchi_ops::{Code, OpenOptions, SourceMode, props};
use std::path::Path;
fn native(base: &str, extension: &str, mode: SourceMode) -> litchi_ops::File {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/excel")
        .join(format!("{base}.{extension}"));
    let options = OpenOptions::new().source_mode(mode);
    if mode == SourceMode::Auto {
        litchi_ops::open_with(&path, &options).unwrap()
    } else {
        litchi_ops::open_bytes_with(std::fs::read(path).unwrap(), &options).unwrap()
    }
}
#[test]
fn native_1904_date_formulas_compute_in_the_workbook_date_system() {
    for extension in ["xls", "xlsx", "xlsb"] {
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let f = native("excel-date1904-breaks", extension, mode);
            let source = f.to_bytes().unwrap();
            for (cell, expected) in [
                ("B5", 43889.0),
                ("B6", 44_848.551_562_5),
                ("B8", 0.270_833_333_333_333_3),
            ] {
                let path = format!("/sheet[1]/{cell}");
                let cached = f.evaluate(path.as_str(), &Options::new()).unwrap();
                assert_eq!(cached.origin(), Origin::Cached);
                let computed = f
                    .evaluate(path.as_str(), &Options::new().recalculate(true))
                    .unwrap();
                assert_eq!(computed.origin(), Origin::Computed);
                assert_eq!(cached.value(), &Value::Number(expected));
                let Value::Number(value) = computed.value() else {
                    panic!("unexpected result {computed:?}")
                };
                assert!(
                    (value - expected).abs() < 1e-9,
                    "{extension} {mode:?} {cell}: {value} != {expected}"
                );
            }
            let report = f.recalculate(&Options::new()).unwrap();
            assert_eq!(report.formulas().len(), 3);
            assert_eq!(report.matches(), 3);
            assert!(f.to_bytes().unwrap() == source);
        }
    }
}
fn unsupported(f: &litchi_ops::File, path: &str) {
    let error = f
        .evaluate(path, &Options::new().recalculate(true))
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert!(matches!(
        error.leaf(),
        litchi_ops::error::Source::Calc(litchi_ops::calc::Error::Unsupported)
    ));
}

#[test]
fn native_scalar_edits_leave_caches_distinct_from_computed_values() {
    for extension in ["xls", "xlsx", "xlsb"] {
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let f = native("excel-values-formulas-chart", extension, mode);
            let source = f.to_bytes().unwrap();
            for (path, value) in [
                ("Sheet1!D2", 3.75),
                ("Sheet1!D6", 49.75),
                ("Sheet1!B6", 14.0),
                ("Sheet1!B23", 3.5),
            ] {
                let cached = f.evaluate(path, &Options::new()).unwrap();
                let computed = f.evaluate(path, &Options::new().recalculate(true)).unwrap();
                assert_eq!(cached.origin(), Origin::Cached);
                assert_eq!(computed.origin(), Origin::Computed);
                assert_eq!(cached.value(), &Value::Number(value));
                assert_eq!(computed.value(), cached.value());
            }
            let changed = f.set("Sheet1!B2", props! {"value" => 4}).unwrap();
            let before_eval = changed.to_bytes().unwrap();
            let cached = changed.evaluate("Sheet1!D2", &Options::new()).unwrap();
            assert_eq!(
                (cached.value(), cached.origin()),
                (&Value::Number(3.75), Origin::Cached)
            );
            let computed = changed
                .evaluate("Sheet1!D2", &Options::new().recalculate(true))
                .unwrap();
            assert_eq!(
                (computed.value(), computed.origin()),
                (&Value::Number(5.0), Origin::Computed)
            );
            assert_eq!(
                changed
                    .evaluate("Sheet1!D6", &Options::new().recalculate(true))
                    .unwrap()
                    .value(),
                &Value::Number(51.0)
            );
            assert!(changed.to_bytes().unwrap() == before_eval);
            let directory = tempfile::tempdir().unwrap();
            let path = directory.path().join(format!("derivative.{extension}"));
            std::fs::write(&path, &before_eval).unwrap();
            let reopened =
                litchi_ops::open_with(&path, &OpenOptions::new().source_mode(mode)).unwrap();
            assert_eq!(
                reopened
                    .evaluate("Sheet1!D2", &Options::new())
                    .unwrap()
                    .value(),
                &Value::Number(3.75)
            );
            assert_eq!(
                reopened
                    .evaluate("Sheet1!D2", &Options::new().recalculate(true))
                    .unwrap()
                    .value(),
                &Value::Number(5.0)
            );
            assert!(reopened.to_bytes().unwrap() == before_eval);
            assert!(f.to_bytes().unwrap() == source);
        }
    }
}

#[test]
fn native_shared_formulas_and_error_caches_keep_exact_result_kinds() {
    use litchi_ops::calc::ErrorValue;
    for extension in ["xls", "xlsx", "xlsb"] {
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let shared = native("excel-formula-owners", extension, mode);
            let source = shared.to_bytes().unwrap();
            for (cell, expected) in [("D4", 8.0), ("D53", 106.0), ("D103", 206.0)] {
                let path = format!("Sheet1!{cell}");
                for recalculate in [false, true] {
                    let result = shared
                        .evaluate(path.as_str(), &Options::new().recalculate(recalculate))
                        .unwrap();
                    assert_eq!(result.value(), &Value::Number(expected));
                    assert_eq!(
                        result.origin(),
                        if recalculate {
                            Origin::Computed
                        } else {
                            Origin::Cached
                        }
                    );
                }
            }
            assert!(shared.to_bytes().unwrap() == source);
            let metadata = native("excel-metadata", extension, mode);
            let source = metadata.to_bytes().unwrap();
            for (path, expected) in [
                ("Sheet1!A4", Value::Error(ErrorValue::Value)),
                ("Sheet1!A5", Value::Error(ErrorValue::Div0)),
                ("Sheet1!A7", Value::Number(30.0)),
            ] {
                assert_eq!(
                    metadata.evaluate(path, &Options::new()).unwrap().value(),
                    &expected
                );
                assert_eq!(
                    metadata
                        .evaluate(path, &Options::new().recalculate(true))
                        .unwrap()
                        .value(),
                    &expected
                );
            }
            let report = metadata.recalculate(&Options::new()).unwrap();
            assert_eq!((report.formulas().len(), report.matches()), (9, 9));
            assert!(metadata.to_bytes().unwrap() == source);
        }
    }
}

#[test]
fn native_external_and_unregistered_table_formulas_refuse_instead_of_fabricating_values() {
    for extension in ["xls", "xlsx", "xlsb"] {
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let links = native("excel-links-objects", extension, mode);
            let source = links.to_bytes().unwrap();
            for (cell, cached) in [("B2", 49.75), ("B3", 14.0), ("B4", 63.75)] {
                let path = format!("/sheet[1]/{cell}");
                let read = links.evaluate(path.as_str(), &Options::new()).unwrap();
                assert_eq!(
                    (read.value(), read.origin()),
                    (&Value::Number(cached), Origin::Cached)
                );
                unsupported(&links, path.as_str());
            }
            let report = links.recalculate(&Options::new()).unwrap();
            assert_eq!((report.formulas().len(), report.unsupported()), (3, 3));
            for f in report.formulas() {
                assert_eq!(f.status(), Status::Unsupported);
                assert!(f.computed().is_none());
            }
            assert!(links.to_bytes().unwrap() == source);
            let values = native("excel-values-formulas-chart", extension, mode);
            let source = values.to_bytes().unwrap();
            assert_eq!(
                values
                    .evaluate("Sheet1!B22", &Options::new())
                    .unwrap()
                    .value(),
                &Value::Text("native Excel".into())
            );
            unsupported(&values, "Sheet1!B22");
            if extension == "xlsx" {
                assert_eq!(
                    values
                        .evaluate("Sheet2!B4", &Options::new())
                        .unwrap()
                        .value(),
                    &Value::Number(14.0)
                );
                unsupported(&values, "Sheet2!B4");
            }
            assert!(values.to_bytes().unwrap() == source);
        }
    }
}

#[test]
fn native_evaluation_limits_and_volatile_derivative_boundaries_are_explicit() {
    use litchi_ops::calc::{Error, LimitKind};
    for extension in ["xls", "xlsx", "xlsb"] {
        let f = native("excel-metadata", extension, SourceMode::Auto);
        let before = f.to_bytes().unwrap();
        let limit = f
            .evaluate("Sheet1!A7", &Options::new().max_cells(1))
            .unwrap_err();
        assert_eq!(limit.code(), Code::LimitExceeded);
        assert!(matches!(
            limit.leaf(),
            litchi_ops::error::Source::Calc(Error::Limit {
                kind: LimitKind::Cells,
                observed: 2,
                maximum: 1
            })
        ));
        assert_eq!(
            f.evaluate("Sheet1!A7", &Options::new().recalculate(true))
                .unwrap()
                .value(),
            &Value::Number(30.0)
        );
        assert!(f.to_bytes().unwrap() == before);
        // The XLS cell writer deliberately uses its 14-function tokenizer;
        // NOW is outside that authoring profile, unlike the read evaluator.
        if extension == "xls" {
            let error = f
                .set("Sheet1!Z100", props! {"formula" => "NOW()"})
                .unwrap_err();
            assert_eq!(error.code(), Code::InvalidValue);
            assert_eq!(error.property(), Some("formula"));
            assert_eq!(
                error.message(),
                "the XLS formula expression cannot be encoded as BIFF8 tokens"
            );
            assert!(f.to_bytes().unwrap() == before);
            continue;
        }
        // An explicitly library-authored derivative, not a native NOW specimen.
        // XLSB keeps the existing formula owner and its cache because its
        // native calculation chain forbids introducing a new formula cell.
        let target = if extension == "xlsb" {
            "Sheet1!A7"
        } else {
            "Sheet1!Z100"
        };
        let derivative = f
            .set(target, props! {"formula" => "NOW()"})
            .unwrap_or_else(|e| panic!("{extension}: {e}; {:?}", e.leaf()));
        let source = derivative.to_bytes().unwrap();
        let result = derivative.evaluate(target, &Options::new().recalculate(true));
        // The retained XLSB cache is preferred even when recalculation is
        // requested. The uncached XLSX derivative refuses a clock read.
        if extension == "xlsb" {
            let result = result.unwrap();
            assert_eq!(result.origin(), Origin::Cached);
            assert_eq!(result.value(), &Value::Number(30.0));
        } else {
            let e = result.unwrap_err();
            assert_eq!(e.code(), Code::UnsupportedCapability);
            assert!(matches!(
                e.leaf(),
                litchi_ops::error::Source::Calc(Error::Volatile)
            ));
        }
        let report = derivative.recalculate(&Options::new()).unwrap();
        assert_eq!(report.volatile(), 1);
        assert!(derivative.to_bytes().unwrap() == source);
        assert!(f.to_bytes().unwrap() == before);
    }
}
