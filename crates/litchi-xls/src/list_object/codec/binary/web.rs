//! Web-query field-info wire primitives.

use super::primitives::u32_at;
use super::strings::{append_string, parse_string};
use crate::Result;
use crate::list_object::invalid;
use crate::list_object::model::{WebColumnType, WebDefaultValue, WebFieldInfo, WebReadingOrder};

pub(in crate::list_object) fn append_web_info(
    out: &mut Vec<u8>,
    info: &WebFieldInfo,
) -> Result<()> {
    out.extend_from_slice(&info.locale.to_le_bytes());
    out.extend_from_slice(&info.decimal_places.to_le_bytes());
    let flags1 = u32::from(info.percent)
        | (u32::from(info.fixed_decimal) << 1)
        | (u32::from(info.date_only) << 2)
        | (info.reading_order.code() << 3)
        | (u32::from(info.rich_text) << 5)
        | (u32::from(info.unknown_rich_text) << 6)
        | (u32::from(info.alert_unknown_rich_text) << 7)
        | info.ignored_display_flags;
    out.extend_from_slice(&flags1.to_le_bytes());
    let default_type = match info.default_value {
        None => 0,
        Some(WebDefaultValue::String(_)) => 1,
        Some(WebDefaultValue::Boolean(_)) => 2,
        Some(WebDefaultValue::Number(_) | WebDefaultValue::DateTime(_)) => 3,
    };
    let flags2 = u32::from(info.read_only)
        | (u32::from(info.required) << 1)
        | (u32::from(info.minimum_set) << 2)
        | (u32::from(info.maximum_set) << 3)
        | (u32::from(info.default_value.is_some()) << 4)
        | (u32::from(info.default_today) << 5)
        | (u32::from(info.validation_formula.is_some()) << 6)
        | (u32::from(info.allow_fill_in) << 7)
        | (default_type << 8)
        | info.ignored_validation_flags;
    out.extend_from_slice(&flags2.to_le_bytes());
    if let Some(value) = &info.default_value {
        match value {
            WebDefaultValue::String(v) => append_string(out, v),
            WebDefaultValue::Boolean(v) => out.extend_from_slice(&u32::from(*v).to_le_bytes()),
            WebDefaultValue::Number(v) | WebDefaultValue::DateTime(v) => {
                out.extend_from_slice(&v.to_le_bytes());
            },
        }
    }
    if let Some(v) = &info.validation_formula {
        append_string(out, v);
    }
    out.extend_from_slice(&0u32.to_le_bytes());
    Ok(())
}

pub(in crate::list_object) fn parse_web_info(
    data: &[u8],
    offset: &mut usize,
    kind: WebColumnType,
    rt: u16,
) -> Result<WebFieldInfo> {
    let locale = u32_at(data, *offset, rt, "Web LCID")?;
    let decimal_places = u32_at(data, *offset + 4, rt, "Web cDec")?;
    let a = u32_at(data, *offset + 8, rt, "Web display flags")?;
    let b = u32_at(data, *offset + 12, rt, "Web validation flags")?;
    let reading_order = WebReadingOrder::from_code((a >> 3) & 3)?;
    let default_set = b & 0x10 != 0;
    let default_type = ((b >> 8) & 0xff) as u8;
    *offset += 16;
    let default_value = if default_set {
        Some(match (default_type, kind) {
            (1, WebColumnType::Text | WebColumnType::Choice | WebColumnType::MultipleChoices) => {
                let (v, end) = parse_string(data, *offset, rt, "Web default string")?;
                *offset = end;
                WebDefaultValue::String(v)
            },
            (2, WebColumnType::Boolean) => {
                let v = u32_at(data, *offset, rt, "Web default boolean")?;
                if v > 1 {
                    return Err(invalid(rt, "invalid Web default boolean"));
                }
                *offset += 4;
                WebDefaultValue::Boolean(v != 0)
            },
            (3, WebColumnType::Number | WebColumnType::Currency | WebColumnType::DateTime) => {
                let bytes = data
                    .get(*offset..*offset + 8)
                    .ok_or_else(|| invalid(rt, "truncated Web default number"))?;
                *offset += 8;
                let v = f64::from_le_bytes(bytes.try_into().unwrap());
                if kind == WebColumnType::DateTime {
                    WebDefaultValue::DateTime(v)
                } else {
                    WebDefaultValue::Number(v)
                }
            },
            _ => return Err(invalid(rt, "Web default type does not match column type")),
        })
    } else {
        if default_type != 0 {
            return Err(invalid(rt, "Web default type exists without a default"));
        }
        None
    };
    let validation_formula = if b & 0x40 != 0 {
        let (v, end) = parse_string(data, *offset, rt, "Web validation formula")?;
        *offset = end;
        Some(v)
    } else {
        None
    };
    if u32_at(data, *offset, rt, "Web reserved")? != 0 {
        return Err(invalid(rt, "Web field-info reserved value must be zero"));
    }
    *offset += 4;
    let value = WebFieldInfo {
        locale,
        decimal_places,
        percent: a & 1 != 0,
        fixed_decimal: a & 2 != 0,
        date_only: a & 4 != 0,
        reading_order,
        rich_text: a & 0x20 != 0,
        unknown_rich_text: a & 0x40 != 0,
        alert_unknown_rich_text: a & 0x80 != 0,
        read_only: b & 1 != 0,
        required: b & 2 != 0,
        minimum_set: b & 4 != 0,
        maximum_set: b & 8 != 0,
        default_today: b & 0x20 != 0,
        allow_fill_in: b & 0x80 != 0,
        default_value,
        validation_formula,
        ignored_display_flags: a & !0xff,
        ignored_validation_flags: b & 0xffff_0000,
    };
    value.validate(kind)?;
    Ok(value)
}
