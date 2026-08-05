//! Semantic conversions used by the XLSB worksheet cell reader.

use crate::conditional_formatting::{Color, Scale, Value};
use crate::package::error::{Error, Result};
use litchi_core::sheet::CellValue;

pub(super) fn build_color_scale(
    cfvos: Vec<Value>,
    colors: Vec<Color>,
    record: &'static str,
    extension14: bool,
) -> Result<Scale> {
    if !(cfvos.len() == 2 || cfvos.len() == 3) || colors.len() != cfvos.len() {
        return Err(Error::Unrecognized {
            typ: record.to_string(),
            val: format!("{} thresholds and {} colors", cfvos.len(), colors.len()),
        });
    }
    if (extension14 && cfvos.iter().any(|cfvo| matches!(cfvo.cfvo_type, 8 | 9)))
        || cfvos[0].cfvo_type == 3
        || cfvos[cfvos.len() - 1].cfvo_type == 2
        || (cfvos.len() == 3 && matches!(cfvos[1].cfvo_type, 2 | 3))
    {
        return Err(Error::Unrecognized {
            typ: record.to_string(),
            val: "invalid min/mid/max threshold type".to_string(),
        });
    }
    let has_middle = colors.len() == 3;
    let mut cfvos = cfvos.into_iter();
    let min_cfvo = cfvos.next().expect("validated threshold count");
    let middle_cfvo = if has_middle { cfvos.next() } else { None };
    let max_cfvo = cfvos.next().expect("validated threshold count");
    let mut colors = colors.into_iter();
    let min_color_record = colors.next().expect("validated color count");
    let mid_color_record = if has_middle { colors.next() } else { None };
    let max_color_record = colors.next().expect("validated color count");
    Ok(Scale {
        min_cfvo,
        mid_cfvo: middle_cfvo,
        max_cfvo,
        min_color: min_color_record.argb.unwrap_or(0),
        mid_color: mid_color_record.and_then(|color| color.argb),
        max_color: max_color_record.argb.unwrap_or(0),
        min_color_record,
        mid_color_record,
        max_color_record,
    })
}

pub(super) fn error_text(error_code: u8) -> &'static str {
    match error_code {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0F => "#VALUE!",
        0x17 => "#REF!",
        0x1D => "#NAME?",
        0x24 => "#NUM!",
        0x2A => "#N/A",
        0x2B => "#GETTING_DATA",
        _ => "#ERR!",
    }
}

pub(super) fn cell_value_from_number(value: f64) -> CellValue {
    if value == value.round() && value >= i64::MIN as f64 && value <= i64::MAX as f64 {
        CellValue::Int(value as i64)
    } else {
        CellValue::Float(value)
    }
}
