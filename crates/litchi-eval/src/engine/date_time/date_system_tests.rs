use crate::engine::{evaluate_expression, test_helpers::TestEngine, to_number};
use crate::parser::parse_expression;
use litchi_core::sheet::CellValue;

fn value(engine: &TestEngine, formula: &str) -> f64 {
    let expr = parse_expression("Sheet1", formula).unwrap();
    let result = evaluate_expression(engine.ctx(), "Sheet1", &expr).unwrap();
    to_number(&result).unwrap_or_else(|| panic!("{formula}: {result:?}"))
}

#[test]
fn date_systems_shift_serials_without_shifting_calendar_results() {
    for date1904 in [false, true] {
        let offset = if date1904 { 0.0 } else { 1462.0 };
        let engine = TestEngine::new().with_date_system(date1904);
        engine.set_cell("Sheet1", 1, 1, CellValue::DateTime(43889.0 + offset));
        for (formula, serial) in [
            ("DATE(2024,2,29)", 43889.0),
            ("DATEVALUE(\"2024-02-29\")", 43889.0),
            ("EDATE(A1,1)", 43918.0),
            ("EOMONTH(A1,1)", 43920.0),
            ("WORKDAY(A1,1)", 43890.0),
            ("WORKDAY.INTL(A1,1,1)", 43890.0),
        ] {
            assert_eq!(
                value(&engine, formula),
                serial + offset,
                "{formula} date1904={date1904}"
            );
        }
        for (formula, expected) in [
            ("YEAR(A1)", 2024.0),
            ("MONTH(A1)", 2.0),
            ("DAY(A1)", 29.0),
            ("WEEKDAY(A1)", 5.0),
            ("WEEKNUM(A1,2)", 9.0),
            ("ISOWEEKNUM(A1)", 9.0),
            ("DAYS(DATE(2024,3,1),A1)", 1.0),
            ("DATEDIF(A1,DATE(2024,3,29),\"m\")", 1.0),
            ("DAYS360(DATE(2024,1,15),DATE(2024,4,15))", 90.0),
            ("YEARFRAC(DATE(2024,1,15),DATE(2024,4,15),0)", 0.25),
            ("HOUR(A1+0.5)", 12.0),
            ("TIME(6,30,0)", 6.5 / 24.0),
        ] {
            assert_eq!(
                value(&engine, formula),
                expected,
                "{formula} date1904={date1904}"
            );
        }
    }
}

#[test]
fn holiday_ranges_use_the_same_epoch_as_workday_endpoints() {
    for date1904 in [false, true] {
        let offset = if date1904 { 0.0 } else { 1462.0 };
        let engine = TestEngine::new().with_date_system(date1904);
        engine.set_cell("Sheet1", 1, 1, CellValue::DateTime(43889.0 + offset));
        engine.set_cell("Sheet1", 1, 2, CellValue::DateTime(43890.0 + offset));
        for formula in ["WORKDAY(A1,1,B1)", "WORKDAY.INTL(A1,1,1,B1)"] {
            assert_eq!(value(&engine, formula), 43893.0 + offset, "{formula}");
        }
        for formula in [
            "NETWORKDAYS(A1,DATE(2024,3,4),B1)",
            "NETWORKDAYS.INTL(A1,DATE(2024,3,4),1,B1)",
        ] {
            assert_eq!(value(&engine, formula), 2.0, "{formula}");
        }
    }
}
