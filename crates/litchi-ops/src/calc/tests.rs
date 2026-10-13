//! Formula-evaluation helper tests.

use litchi_core::sheet::CellValue;

use super::{ErrorValue, Value, convert, is_volatile, normalize};

#[test]
fn volatile_functions_are_found_outside_strings() {
    for formula in [
        "NOW()",
        "1+today()",
        "IF(A1, RAND(), 0)",
        "_xlfn.RANDARRAY(2)",
        "RANDBETWEEN(1,6)*2",
    ] {
        assert!(is_volatile(formula), "{formula}");
    }
    for formula in [
        "\"NOW()\"&A1",
        "NOWHERE(1)",
        "SUM(A1:A3)",
        "NOW",
        "MY.NOW(1)",
    ] {
        assert!(!is_volatile(formula), "{formula}");
    }
}

#[test]
fn modern_function_prefixes_are_removed_outside_strings() {
    assert_eq!(normalize("SUM(A1)"), "SUM(A1)");
    assert_eq!(
        normalize("_xlfn.XLOOKUP(1,A:A,B:B)+_xlws.FILTER(C1)"),
        "XLOOKUP(1,A:A,B:B)+FILTER(C1)"
    );
    assert_eq!(
        normalize("\"_xlfn.\"&_XLFN.CONCAT(A1)"),
        "\"_xlfn.\"&CONCAT(A1)"
    );
}

#[test]
fn engine_values_convert_and_unknown_errors_are_unsupported() {
    assert_eq!(convert(&CellValue::Int(3)), Some(Value::Number(3.0)));
    assert_eq!(convert(&CellValue::DateTime(2.5)), Some(Value::Number(2.5)));
    assert_eq!(
        convert(&CellValue::Error("#DIV/0!".to_owned())),
        Some(Value::Error(ErrorValue::Div0))
    );
    assert_eq!(
        convert(&CellValue::Error("Unsupported function: X".to_owned())),
        None
    );
    let cached = CellValue::Formula {
        formula: "A1".to_owned(),
        cached_value: Some(Box::new(CellValue::Bool(true))),
        is_array: false,
        array_range: None,
    };
    assert_eq!(convert(&cached), Some(Value::Bool(true)));
}

#[test]
fn agreement_uses_a_relative_tolerance() {
    let tolerance = 1e-9;
    assert!(Value::Number(1.0).agrees(&Value::Number(1.0 + 1e-12), tolerance));
    assert!(Value::Number(1e12).agrees(&Value::Number(1e12 + 1e2), tolerance));
    assert!(!Value::Number(1.0).agrees(&Value::Number(1.0001), tolerance));
    assert!(Value::Empty.agrees(&Value::Text(String::new()), tolerance));
    assert!(!Value::Text("a".to_owned()).agrees(&Value::Text("A".to_owned()), tolerance));
    assert!(!Value::Bool(true).agrees(&Value::Number(1.0), tolerance));
}

#[test]
fn error_values_round_trip() {
    for (value, text) in ErrorValue::ALL {
        assert_eq!(value.as_str(), text);
        assert_eq!(ErrorValue::parse(&text.to_ascii_lowercase()), Some(value));
    }
    assert_eq!(ErrorValue::parse("#BOGUS!"), None);
}
