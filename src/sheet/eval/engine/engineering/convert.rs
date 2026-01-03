use crate::sheet::eval::engine::{EvalCtx, evaluate_expression, to_number, to_text};
use crate::sheet::eval::parser::Expr;
use crate::sheet::{CellValue, Result};
use once_cell::sync::Lazy;
use std::collections::HashMap;

struct UnitInfo {
    category: &'static str,
    factor: f64,
}

static UNITS: Lazy<HashMap<&'static str, UnitInfo>> = Lazy::new(|| {
    let mut m = HashMap::new();
    // Weight and mass
    m.insert(
        "g",
        UnitInfo {
            category: "weight",
            factor: 1.0,
        },
    );
    m.insert(
        "kg",
        UnitInfo {
            category: "weight",
            factor: 1000.0,
        },
    );
    m.insert(
        "mg",
        UnitInfo {
            category: "weight",
            factor: 0.001,
        },
    );
    m.insert(
        "lbm",
        UnitInfo {
            category: "weight",
            factor: 453.59237,
        },
    );
    m.insert(
        "ozm",
        UnitInfo {
            category: "weight",
            factor: 28.349523125,
        },
    );

    // Distance
    m.insert(
        "m",
        UnitInfo {
            category: "distance",
            factor: 1.0,
        },
    );
    m.insert(
        "km",
        UnitInfo {
            category: "distance",
            factor: 1000.0,
        },
    );
    m.insert(
        "cm",
        UnitInfo {
            category: "distance",
            factor: 0.01,
        },
    );
    m.insert(
        "mm",
        UnitInfo {
            category: "distance",
            factor: 0.001,
        },
    );
    m.insert(
        "in",
        UnitInfo {
            category: "distance",
            factor: 0.0254,
        },
    );
    m.insert(
        "ft",
        UnitInfo {
            category: "distance",
            factor: 0.3048,
        },
    );
    m.insert(
        "yd",
        UnitInfo {
            category: "distance",
            factor: 0.9144,
        },
    );
    m.insert(
        "mi",
        UnitInfo {
            category: "distance",
            factor: 1609.344,
        },
    );

    // Time
    m.insert(
        "yr",
        UnitInfo {
            category: "time",
            factor: 31536000.0,
        },
    );
    m.insert(
        "day",
        UnitInfo {
            category: "time",
            factor: 86400.0,
        },
    );
    m.insert(
        "hr",
        UnitInfo {
            category: "time",
            factor: 3600.0,
        },
    );
    m.insert(
        "mn",
        UnitInfo {
            category: "time",
            factor: 60.0,
        },
    );
    m.insert(
        "sec",
        UnitInfo {
            category: "time",
            factor: 1.0,
        },
    );

    // Pressure
    m.insert(
        "Pa",
        UnitInfo {
            category: "pressure",
            factor: 1.0,
        },
    );
    m.insert(
        "atm",
        UnitInfo {
            category: "pressure",
            factor: 101325.0,
        },
    );
    m.insert(
        "mmHg",
        UnitInfo {
            category: "pressure",
            factor: 133.322368,
        },
    );

    // Force
    m.insert(
        "N",
        UnitInfo {
            category: "force",
            factor: 1.0,
        },
    );
    m.insert(
        "dyn",
        UnitInfo {
            category: "force",
            factor: 0.00001,
        },
    );
    m.insert(
        "lbf",
        UnitInfo {
            category: "force",
            factor: 4.4482216152605,
        },
    );

    // Energy
    m.insert(
        "J",
        UnitInfo {
            category: "energy",
            factor: 1.0,
        },
    );
    m.insert(
        "e",
        UnitInfo {
            category: "energy",
            factor: 1e-7,
        },
    );
    m.insert(
        "cal",
        UnitInfo {
            category: "energy",
            factor: 4.1868,
        },
    );
    m.insert(
        "BTU",
        UnitInfo {
            category: "energy",
            factor: 1055.05585,
        },
    );

    // Power
    m.insert(
        "W",
        UnitInfo {
            category: "power",
            factor: 1.0,
        },
    );
    m.insert(
        "HP",
        UnitInfo {
            category: "power",
            factor: 745.69987158227,
        },
    );

    // Magnetism
    m.insert(
        "T",
        UnitInfo {
            category: "magnetism",
            factor: 1.0,
        },
    );
    m.insert(
        "ga",
        UnitInfo {
            category: "magnetism",
            factor: 0.0001,
        },
    );

    // Temperature (Special handling)
    m.insert(
        "C",
        UnitInfo {
            category: "temp",
            factor: 1.0,
        },
    );
    m.insert(
        "F",
        UnitInfo {
            category: "temp",
            factor: 1.0,
        },
    );
    m.insert(
        "K",
        UnitInfo {
            category: "temp",
            factor: 1.0,
        },
    );

    // Volume
    m.insert(
        "l",
        UnitInfo {
            category: "volume",
            factor: 0.001,
        },
    );
    m.insert(
        "L",
        UnitInfo {
            category: "volume",
            factor: 0.001,
        },
    );
    m.insert(
        "gal",
        UnitInfo {
            category: "volume",
            factor: 0.003785411784,
        },
    );
    m.insert(
        "qt",
        UnitInfo {
            category: "volume",
            factor: 0.000946352946,
        },
    );
    m.insert(
        "pt",
        UnitInfo {
            category: "volume",
            factor: 0.000473176473,
        },
    );

    m
});

pub(crate) async fn eval_convert(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 3 {
        return Ok(CellValue::Error("CONVERT expects 3 arguments".to_string()));
    }

    let number = match to_number(&evaluate_expression(ctx, current_sheet, &args[0]).await?) {
        Some(n) => n,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let from_unit = to_text(&evaluate_expression(ctx, current_sheet, &args[1]).await?);
    let to_unit = to_text(&evaluate_expression(ctx, current_sheet, &args[2]).await?);

    let from_info = match UNITS.get(from_unit.as_str()) {
        Some(i) => i,
        None => return Ok(CellValue::Error("#N/A".to_string())),
    };

    let to_info = match UNITS.get(to_unit.as_str()) {
        Some(i) => i,
        None => return Ok(CellValue::Error("#N/A".to_string())),
    };

    if from_info.category != to_info.category {
        return Ok(CellValue::Error("#N/A".to_string()));
    }

    if from_info.category == "temp" {
        let result = convert_temp(number, &from_unit, &to_unit);
        return Ok(CellValue::Float(result));
    }

    let result = number * (from_info.factor / to_info.factor);
    Ok(CellValue::Float(result))
}

fn convert_temp(val: f64, from: &str, to: &str) -> f64 {
    let kelvin = match from {
        "C" => val + 273.15,
        "F" => (val + 459.67) * 5.0 / 9.0,
        "K" => val,
        _ => val,
    };
    match to {
        "C" => kelvin - 273.15,
        "F" => kelvin * 9.0 / 5.0 - 459.67,
        "K" => kelvin,
        _ => kelvin,
    }
}
