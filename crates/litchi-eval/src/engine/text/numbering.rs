use crate::engine::{EvalCtx, evaluate_expression, to_number, to_text};
use crate::parser::Expr;
use litchi_core::sheet::{CellValue, Result};

const LOWER_DIGITS: [char; 10] = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
const UPPER_DIGITS: [char; 10] = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];
const LOWER_UNITS: [char; 3] = ['十', '百', '千'];
const UPPER_UNITS: [char; 3] = ['拾', '佰', '仟'];
const GROUP_UNITS: [&str; 4] = ["", "万", "亿", "万亿"];

/// Largest integer that f64 can represent exactly; NUMBERSTRING is limited
/// to this range (ECMA-376 §18.17.7.242 does not define behavior beyond it).
const MAX_EXACT_INT: f64 = 9_007_199_254_740_992.0; // 2^53

/// Renders 1..=9999 with 千百十 units, inserting 零 across internal zero
/// gaps and dropping a leading 一 before 十 (12 -> 十二, not 一十二).
fn four_digit_group(n: u16, digits: &[char; 10], units: &[char; 3]) -> String {
    let mut s = String::new();
    let mut started = false;
    let mut zero_pending = false;
    let place_values = [n / 1000, (n / 100) % 10, (n / 10) % 10, n % 10];
    for (i, &place) in place_values.iter().enumerate() {
        let d = place as usize;
        if d == 0 {
            if started {
                zero_pending = true;
            }
            continue;
        }
        if zero_pending {
            s.push(digits[0]);
            zero_pending = false;
        }
        if i == 3 {
            s.push(digits[d]);
        } else {
            if !(i == 2 && d == 1 && !started) {
                s.push(digits[d]);
            }
            s.push(units[2 - i]);
        }
        started = true;
    }
    s
}

/// Positional reading (type 1/2): groups of four decimal digits joined with
/// the 万/亿/万亿 units, with 零 bridging zero gaps between and within groups.
fn positional_reading(mut value: u64, digits: &[char; 10], units: &[char; 3]) -> String {
    if value == 0 {
        return digits[0].to_string();
    }
    let mut groups = Vec::new();
    while value > 0 {
        groups.push((value % 10000) as u16);
        value /= 10000;
    }
    let mut result = String::new();
    let mut zero_between = false;
    for i in (0..groups.len()).rev() {
        let group = groups[i];
        if group == 0 {
            if !result.is_empty() {
                zero_between = true;
            }
            continue;
        }
        if !result.is_empty() && (zero_between || group < 1000) {
            result.push(digits[0]);
        }
        result.push_str(&four_digit_group(group, digits, units));
        result.push_str(GROUP_UNITS[i]);
        zero_between = false;
    }
    result
}

/// NUMBERSTRING(value, type) — East Asian number-to-string conversion
/// (ECMA-376 §18.17.7.242): type 1 = lowercase numerals (一千二百三十四),
/// type 2 = uppercase formal numerals (壹仟贰佰叁拾肆), type 3 = plain
/// digit-by-digit reading (一二三四). Non-integers are truncated; negative
/// values, values beyond 2^53, and unknown types yield #VALUE!.
pub(crate) async fn eval_numberstring(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 2 {
        return Ok(CellValue::Error(
            "NUMBERSTRING expects 2 arguments (value, type)".to_string(),
        ));
    }

    let value_val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let value = match to_number(&value_val) {
        Some(n) => n,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };
    let type_val = evaluate_expression(ctx, current_sheet, &args[1]).await?;
    let num_type = match to_number(&type_val) {
        Some(n) if (1.0..=3.0).contains(&n.trunc()) => n.trunc() as u8,
        _ => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    let truncated = value.trunc();
    if !(0.0..=MAX_EXACT_INT).contains(&truncated) {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    let n = truncated as u64;

    let text = match num_type {
        1 => positional_reading(n, &LOWER_DIGITS, &LOWER_UNITS),
        2 => positional_reading(n, &UPPER_DIGITS, &UPPER_UNITS),
        _ => {
            // Type 3: plain digit-by-digit reading with lowercase numerals.
            if n == 0 {
                LOWER_DIGITS[0].to_string()
            } else {
                n.to_string()
                    .chars()
                    .map(|c| LOWER_DIGITS[c.to_digit(10).unwrap_or(0) as usize])
                    .collect()
            }
        },
    };

    Ok(CellValue::String(text))
}

pub(crate) async fn eval_arabic(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.len() != 1 {
        return Ok(CellValue::Error("ARABIC expects 1 argument".to_string()));
    }
    let text = to_text(&evaluate_expression(ctx, current_sheet, &args[0]).await?)
        .trim()
        .to_uppercase();
    if text.is_empty() {
        return Ok(CellValue::Int(0));
    }

    let mut result = 0i64;
    let mut last_val = 0i64;

    // Handle negative sign
    let (haystack, multiplier) = if let Some(stripped) = text.strip_prefix('-') {
        (stripped, -1)
    } else {
        (text.as_str(), 1)
    };

    for c in haystack.chars().rev() {
        let val = match c {
            'I' => 1,
            'V' => 5,
            'X' => 10,
            'L' => 50,
            'C' => 100,
            'D' => 500,
            'M' => 1000,
            _ => return Ok(CellValue::Error("#VALUE!".to_string())),
        };
        if val < last_val {
            result -= val;
        } else {
            result += val;
        }
        last_val = val;
    }

    Ok(CellValue::Int(result * multiplier))
}

pub(crate) async fn eval_roman(
    ctx: EvalCtx<'_>,
    current_sheet: &str,
    args: &[Expr],
) -> Result<CellValue> {
    if args.is_empty() || args.len() > 2 {
        return Ok(CellValue::Error(
            "ROMAN expects 1 or 2 arguments".to_string(),
        ));
    }
    let val = evaluate_expression(ctx, current_sheet, &args[0]).await?;
    let num = match crate::engine::to_number(&val) {
        Some(n) => n.trunc() as i64,
        None => return Ok(CellValue::Error("#VALUE!".to_string())),
    };

    if !(0..=3999).contains(&num) {
        return Ok(CellValue::Error("#VALUE!".to_string()));
    }
    if num == 0 {
        return Ok(CellValue::String(String::new()));
    }

    // Excel ROMAN function has different forms (0 to 4), but usually 0 (classic) is used.
    // We'll implement classic (form 0) for now.
    let mut n = num;
    let mut result = String::new();
    let mapping = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];

    for (val, sym) in mapping {
        while n >= val {
            result.push_str(sym);
            n -= val;
        }
    }

    Ok(CellValue::String(result))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::engine::test_helpers::TestEngine;
    use crate::parser::Expr;

    #[tokio::test]
    async fn test_eval_arabic_simple() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String("XII".to_string()))];
        let result = eval_arabic(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(12));
    }

    #[tokio::test]
    async fn test_eval_arabic_complex() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String("MCMXCIV".to_string()))];
        let result = eval_arabic(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(1994));
    }

    #[tokio::test]
    async fn test_eval_arabic_lowercase() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String("xiv".to_string()))];
        let result = eval_arabic(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(14));
    }

    #[tokio::test]
    async fn test_eval_arabic_empty() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String("".to_string()))];
        let result = eval_arabic(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Int(0));
    }

    #[tokio::test]
    async fn test_eval_arabic_invalid() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::String("ABC".to_string()))];
        let result = eval_arabic(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Error("#VALUE!".to_string()));
    }

    #[tokio::test]
    async fn test_eval_arabic_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_arabic(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 argument")),
            _ => panic!("Expected Error"),
        }
    }

    #[tokio::test]
    async fn test_eval_roman_simple() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Int(12))];
        let result = eval_roman(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("XII".to_string()));
    }

    #[tokio::test]
    async fn test_eval_roman_complex() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Int(1994))];
        let result = eval_roman(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("MCMXCIV".to_string()));
    }

    #[tokio::test]
    async fn test_eval_roman_zero() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Int(0))];
        let result = eval_roman(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::String("".to_string()));
    }

    #[tokio::test]
    async fn test_eval_roman_too_large() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Int(4000))];
        let result = eval_roman(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Error("#VALUE!".to_string()));
    }

    #[tokio::test]
    async fn test_eval_roman_negative() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Int(-5))];
        let result = eval_roman(ctx, "Sheet1", &args).await.unwrap();
        assert_eq!(result, CellValue::Error("#VALUE!".to_string()));
    }

    #[tokio::test]
    async fn test_eval_roman_wrong_args() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args: Vec<Expr> = vec![];
        let result = eval_roman(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 1 or 2")),
            _ => panic!("Expected Error"),
        }
    }

    fn num_expr(n: f64) -> Expr {
        if n == n.floor() {
            Expr::Literal(CellValue::Int(n as i64))
        } else {
            Expr::Literal(CellValue::Float(n))
        }
    }

    async fn numberstring(value: f64, num_type: i64) -> CellValue {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![num_expr(value), Expr::Literal(CellValue::Int(num_type))];
        eval_numberstring(ctx, "Sheet1", &args).await.unwrap()
    }

    fn expect_string(value: CellValue) -> String {
        match value {
            CellValue::String(s) => s,
            other => panic!("Expected String, got {:?}", other),
        }
    }

    #[tokio::test]
    async fn test_numberstring_type1_ecma_example() {
        // ECMA-376 §18.17.7.242 example.
        assert_eq!(
            expect_string(numberstring(123456789.0, 1).await),
            "一亿二千三百四十五万六千七百八十九"
        );
    }

    #[tokio::test]
    async fn test_numberstring_type2_ecma_example() {
        assert_eq!(
            expect_string(numberstring(123456789.0, 2).await),
            "壹亿贰仟叁佰肆拾伍万陆仟柒佰捌拾玖"
        );
    }

    #[tokio::test]
    async fn test_numberstring_type3_ecma_example() {
        assert_eq!(
            expect_string(numberstring(123456789.0, 3).await),
            "一二三四五六七八九"
        );
    }

    #[tokio::test]
    async fn test_numberstring_zero() {
        assert_eq!(expect_string(numberstring(0.0, 1).await), "零");
        assert_eq!(expect_string(numberstring(0.0, 2).await), "零");
        assert_eq!(expect_string(numberstring(0.0, 3).await), "零");
    }

    #[tokio::test]
    async fn test_numberstring_small_numbers() {
        assert_eq!(expect_string(numberstring(5.0, 1).await), "五");
        assert_eq!(expect_string(numberstring(10.0, 1).await), "十");
        assert_eq!(expect_string(numberstring(12.0, 1).await), "十二");
        assert_eq!(expect_string(numberstring(112.0, 1).await), "一百一十二");
    }

    #[tokio::test]
    async fn test_numberstring_zero_gaps() {
        assert_eq!(expect_string(numberstring(101.0, 1).await), "一百零一");
        assert_eq!(expect_string(numberstring(1001.0, 1).await), "一千零一");
        assert_eq!(expect_string(numberstring(10000.0, 1).await), "一万");
        assert_eq!(expect_string(numberstring(10001.0, 1).await), "一万零一");
        assert_eq!(expect_string(numberstring(11000.0, 1).await), "一万一千");
        assert_eq!(
            expect_string(numberstring(100000001.0, 1).await),
            "一亿零一"
        );
        assert_eq!(
            expect_string(numberstring(100010000.0, 1).await),
            "一亿零一万"
        );
    }

    #[tokio::test]
    async fn test_numberstring_type2_zero_gaps() {
        assert_eq!(expect_string(numberstring(101.0, 2).await), "壹佰零壹");
        assert_eq!(expect_string(numberstring(10000.0, 2).await), "壹万");
    }

    #[tokio::test]
    async fn test_numberstring_truncates_fraction() {
        assert_eq!(expect_string(numberstring(12.9, 1).await), "十二");
    }

    #[tokio::test]
    async fn test_numberstring_negative_value() {
        match numberstring(-1.0, 1).await {
            CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
            other => panic!("Expected #VALUE!, got {:?}", other),
        }
    }

    #[tokio::test]
    async fn test_numberstring_invalid_type() {
        match numberstring(5.0, 4).await {
            CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
            other => panic!("Expected #VALUE!, got {:?}", other),
        }
    }

    #[tokio::test]
    async fn test_numberstring_beyond_exact_f64_range() {
        match numberstring(1.0e16, 1).await {
            CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
            other => panic!("Expected #VALUE!, got {:?}", other),
        }
    }

    #[tokio::test]
    async fn test_numberstring_non_numeric() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![
            Expr::Literal(CellValue::String("abc".to_string())),
            Expr::Literal(CellValue::Int(1)),
        ];
        let result = eval_numberstring(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert_eq!(e, "#VALUE!"),
            other => panic!("Expected #VALUE!, got {:?}", other),
        }
    }

    #[tokio::test]
    async fn test_numberstring_wrong_arg_count() {
        let engine = TestEngine::new();
        let ctx = engine.ctx();
        let args = vec![Expr::Literal(CellValue::Int(1))];
        let result = eval_numberstring(ctx, "Sheet1", &args).await.unwrap();
        match result {
            CellValue::Error(e) => assert!(e.contains("expects 2 arguments")),
            _ => panic!("Expected Error"),
        }
    }
}
