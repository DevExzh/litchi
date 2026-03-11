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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::eval::parser::Expr;

    fn num_expr(n: f64) -> Expr {
        if n == n.floor() {
            Expr::Literal(CellValue::Int(n as i64))
        } else {
            Expr::Literal(CellValue::Float(n))
        }
    }

    fn str_expr(s: &str) -> Expr {
        Expr::Literal(CellValue::String(s.to_string()))
    }

    #[tokio::test]
    async fn test_eval_convert_weight() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Convert 1 kg to grams
        let args = vec![num_expr(1.0), str_expr("kg"), str_expr("g")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1000.0).abs() < 1e-9),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_convert_distance() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Convert 1 mile to meters
        let args = vec![num_expr(1.0), str_expr("mi"), str_expr("m")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 1609.344).abs() < 0.001),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_convert_time() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Convert 1 hour to seconds
        let args = vec![num_expr(1.0), str_expr("hr"), str_expr("sec")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3600.0).abs() < 1e-9),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_convert_temperature_c_to_f() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Convert 0°C to Fahrenheit (32°F)
        let args = vec![num_expr(0.0), str_expr("C"), str_expr("F")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 32.0).abs() < 1e-9),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_convert_temperature_f_to_c() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Convert 212°F to Celsius (100°C)
        let args = vec![num_expr(212.0), str_expr("F"), str_expr("C")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 100.0).abs() < 1e-9),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_convert_temperature_c_to_k() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Convert 0°C to Kelvin (273.15K)
        let args = vec![num_expr(0.0), str_expr("C"), str_expr("K")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 273.15).abs() < 1e-9),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_convert_wrong_args() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0), str_expr("kg")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 3 arguments")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_convert_invalid_unit() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(1.0), str_expr("kg"), str_expr("invalid")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#N/A")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_convert_incompatible_units() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Can't convert weight to distance
        let args = vec![num_expr(1.0), str_expr("kg"), str_expr("m")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#N/A")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_convert_pressure() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Convert 1 atm to Pa
        let args = vec![num_expr(1.0), str_expr("atm"), str_expr("Pa")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 101325.0).abs() < 1.0),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_convert_force() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Convert 1 lbf to Newtons
        let args = vec![num_expr(1.0), str_expr("lbf"), str_expr("N")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 4.448).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_convert_power() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Convert 1 HP to Watts
        let args = vec![num_expr(1.0), str_expr("HP"), str_expr("W")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 745.7).abs() < 1.0),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_convert_volume() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        // Convert 1 gallon to liters
        let args = vec![num_expr(1.0), str_expr("gal"), str_expr("l")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Float(v) => assert!((v - 3.785).abs() < 0.01),
            _ => panic!("Expected Float"),
        }
    }

    #[tokio::test]
    async fn test_eval_convert_non_numeric() {
        let engine = crate::sheet::eval::engine::test_helpers::TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![str_expr("not a number"), str_expr("kg"), str_expr("g")];
        let result = eval_convert(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("#VALUE!")),
            _ => panic!("Expected Error"),
        }
    }
}
