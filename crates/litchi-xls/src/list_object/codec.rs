//! Bounded BIFF8 payload codecs for worksheet tables and List12 records.

use super::model::*;
use super::model::{validate_column_name, validate_name, validate_table_name};
use super::{
    AUTO_FILTER12_RECORD_TYPE, CONTINUE_FRT11_RECORD_TYPE, FEATURE11_RECORD_TYPE,
    FEATURE12_RECORD_TYPE, ISF_LIST, LIST12_RECORD_TYPE, MAX_CONTINUE_RGB, MAX_FEATURE_BYTES,
    MAX_PAYLOAD, invalid,
};
use crate::Result;
use crate::autofilter12::write_table_autofilter12;

pub(super) struct PendingFeature {
    pub(super) record_type: u16,
    pub(super) base: Vec<u8>,
    pub(super) continuations: Vec<Vec<u8>>,
    pub(super) combined: Vec<u8>,
}
pub(super) fn u16_at(data: &[u8], offset: usize, rt: u16, field: &str) -> Result<u16> {
    data.get(offset..offset + 2)
        .map(|v| u16::from_le_bytes([v[0], v[1]]))
        .ok_or_else(|| invalid(rt, format!("truncated {field}")))
}
pub(super) fn u32_at(data: &[u8], offset: usize, rt: u16, field: &str) -> Result<u32> {
    data.get(offset..offset + 4)
        .map(|v| u32::from_le_bytes(v.try_into().unwrap()))
        .ok_or_else(|| invalid(rt, format!("truncated {field}")))
}
pub(super) fn append_range(out: &mut Vec<u8>, range: ListObjectRange) {
    out.extend_from_slice(&range.first_row.to_le_bytes());
    out.extend_from_slice(&range.last_row.to_le_bytes());
    out.extend_from_slice(&range.first_column.to_le_bytes());
    out.extend_from_slice(&range.last_column.to_le_bytes());
}
pub(super) fn parse_range(data: &[u8], offset: usize, rt: u16) -> Result<ListObjectRange> {
    ListObjectRange::try_new(
        u16_at(data, offset, rt, "rwFirst")?,
        u16_at(data, offset + 2, rt, "rwLast")?,
        u16_at(data, offset + 4, rt, "colFirst")?,
        u16_at(data, offset + 6, rt, "colLast")?,
    )
}
pub(super) fn append_frt(out: &mut Vec<u8>, rt: u16, range: Option<ListObjectRange>) {
    out.extend_from_slice(&rt.to_le_bytes());
    out.extend_from_slice(&u16::from(range.is_some()).to_le_bytes());
    if let Some(range) = range {
        append_range(out, range);
    } else {
        out.extend_from_slice(&[0; 8]);
    }
}
pub(super) fn validate_frt(data: &[u8], rt: u16, reference: bool) -> Result<()> {
    if u16_at(data, 0, rt, "frt.rt")? != rt
        || u16_at(data, 2, rt, "frt.flags")? != u16::from(reference)
    {
        return Err(invalid(rt, "future-record header is invalid"));
    }
    if !reference && data.get(4..12).is_none_or(|v| v.iter().any(|b| *b != 0)) {
        return Err(invalid(rt, "future-record reserved bytes must be zero"));
    }
    Ok(())
}
pub(super) fn validate_frt_any(data: &[u8], rt: u16) -> Result<()> {
    if u16_at(data, 0, rt, "frt.rt")? != rt {
        return Err(invalid(rt, "future-record type echo is invalid"));
    }
    let flags = u16_at(data, 2, rt, "frt.flags")?;
    if flags & 0x0002 != 0 {
        return Err(invalid(rt, "future-record alert flag must be zero"));
    }
    if flags & 0x0001 == 0 && data.get(4..12).is_none_or(|v| v.iter().any(|b| *b != 0)) {
        return Err(invalid(
            rt,
            "future-record reference is present without fFrtRef",
        ));
    }
    Ok(())
}
pub(super) fn record(rt: u16, payload: Vec<u8>) -> Result<Vec<u8>> {
    let len =
        u16::try_from(payload.len()).map_err(|_| invalid(rt, "payload exceeds BIFF8 length"))?;
    let mut out = Vec::with_capacity(4 + payload.len());
    out.extend_from_slice(&rt.to_le_bytes());
    out.extend_from_slice(&len.to_le_bytes());
    out.extend_from_slice(&payload);
    Ok(out)
}
pub(super) fn append_string(out: &mut Vec<u8>, value: &str) {
    let units = value.encode_utf16().collect::<Vec<_>>();
    out.extend_from_slice(&(units.len() as u16).to_le_bytes());
    if value.is_ascii() {
        out.push(0);
        out.extend_from_slice(value.as_bytes());
    } else {
        out.push(1);
        out.extend(units.into_iter().flat_map(u16::to_le_bytes));
    }
}
pub(super) fn parse_string(
    data: &[u8],
    offset: usize,
    rt: u16,
    field: &str,
) -> Result<(String, usize)> {
    let count = usize::from(u16_at(data, offset, rt, field)?);
    let flags = *data
        .get(offset + 2)
        .ok_or_else(|| invalid(rt, format!("truncated {field} flags")))?;
    if flags & !1 != 0 {
        return Err(invalid(rt, format!("{field} flags are unsupported")));
    }
    let width = if flags == 0 { 1 } else { 2 };
    let end = offset
        .checked_add(3)
        .and_then(|v| count.checked_mul(width).and_then(|n| v.checked_add(n)))
        .ok_or_else(|| invalid(rt, format!("{field} length overflows")))?;
    let bytes = data
        .get(offset + 3..end)
        .ok_or_else(|| invalid(rt, format!("truncated {field}")))?;
    let value = if width == 1 {
        bytes.iter().map(|b| char::from(*b)).collect()
    } else {
        char::decode_utf16(
            bytes
                .chunks_exact(2)
                .map(|c| u16::from_le_bytes([c[0], c[1]])),
        )
        .collect::<Result<String, _>>()
        .map_err(|_| invalid(rt, format!("invalid UTF-16 in {field}")))?
    };
    Ok((value, end))
}

#[derive(Clone, Copy)]
enum FormulaExtraKind {
    Array,
    Memory,
}

fn parse_list_formula_extra_end(
    data: &[u8],
    tokens: &[u8],
    mut offset: usize,
    rt: u16,
) -> Result<usize> {
    let mut extras = Vec::new();
    let mut position = 0usize;
    while position < tokens.len() {
        let opcode = tokens[position];
        let base = if opcode < 0x20 {
            opcode
        } else {
            (opcode & 0x1f) | 0x20
        };
        let size = match base {
            0x03..=0x16 => 1,
            0x17 => {
                let count = usize::from(
                    *tokens
                        .get(position + 1)
                        .ok_or_else(|| invalid(rt, "truncated formula string token"))?,
                );
                let flags = *tokens
                    .get(position + 2)
                    .ok_or_else(|| invalid(rt, "truncated formula string flags"))?;
                if flags & !1 != 0 {
                    return Err(invalid(rt, "unsupported formula string flags"));
                }
                3usize
                    .checked_add(
                        count
                            .checked_mul(if flags == 0 { 1 } else { 2 })
                            .ok_or_else(|| invalid(rt, "formula string length overflows"))?,
                    )
                    .ok_or_else(|| invalid(rt, "formula string length overflows"))?
            },
            0x19 => {
                let header = tokens
                    .get(position..position + 4)
                    .ok_or_else(|| invalid(rt, "truncated Attr token"))?;
                if header[1] & 0x04 != 0 {
                    4usize
                        .checked_add(
                            (usize::from(u16::from_le_bytes([header[2], header[3]])) + 1) * 2,
                        )
                        .ok_or_else(|| invalid(rt, "Attr token length overflows"))?
                } else {
                    4
                }
            },
            0x1c | 0x1d => 2,
            0x1e => 3,
            0x1f => 9,
            0x20 => {
                extras.push(FormulaExtraKind::Array);
                8
            },
            0x21 => 3,
            0x22 => 4,
            0x23 | 0x24 | 0x2a | 0x2c => 5,
            0x25 | 0x2b | 0x2d => 9,
            0x26 => {
                extras.push(FormulaExtraKind::Memory);
                7
            },
            0x27 => 7,
            0x29 => 3,
            0x39 | 0x3a | 0x3c => 7,
            0x3b | 0x3d => 11,
            _ => {
                return Err(invalid(
                    rt,
                    "invalid or forbidden token in list array formula",
                ));
            },
        };
        position = position
            .checked_add(size)
            .ok_or_else(|| invalid(rt, "formula token length overflows"))?;
        if position > tokens.len() {
            return Err(invalid(rt, "truncated formula token"));
        }
    }
    for extra in extras {
        match extra {
            FormulaExtraKind::Memory => {
                let count = usize::from(u16_at(data, offset, rt, "PtgExtraMem count")?);
                offset = offset
                    .checked_add(2)
                    .and_then(|value| value.checked_add(count.checked_mul(8)?))
                    .ok_or_else(|| invalid(rt, "PtgExtraMem length overflows"))?;
                data.get(..offset)
                    .ok_or_else(|| invalid(rt, "truncated PtgExtraMem"))?;
            },
            FormulaExtraKind::Array => {
                let dimensions = data
                    .get(offset..offset + 3)
                    .ok_or_else(|| invalid(rt, "truncated PtgExtraArray dimensions"))?;
                let count = (usize::from(dimensions[0]) + 1)
                    .checked_mul(
                        usize::from(u16::from_le_bytes([dimensions[1], dimensions[2]])) + 1,
                    )
                    .ok_or_else(|| invalid(rt, "PtgExtraArray dimensions overflow"))?;
                offset += 3;
                for _ in 0..count {
                    let kind = *data
                        .get(offset)
                        .ok_or_else(|| invalid(rt, "truncated PtgExtraArray value"))?;
                    offset += 1;
                    match kind {
                        0 | 1 | 4 | 16 => {
                            offset = offset
                                .checked_add(8)
                                .ok_or_else(|| invalid(rt, "PtgExtraArray length overflows"))?;
                            data.get(..offset)
                                .ok_or_else(|| invalid(rt, "truncated PtgExtraArray value"))?;
                        },
                        2 => {
                            offset = parse_string(data, offset, rt, "PtgExtraArray string")?.1;
                        },
                        _ => return Err(invalid(rt, "invalid PtgExtraArray value type")),
                    }
                }
            },
        }
    }
    Ok(offset)
}

fn append_formula(out: &mut Vec<u8>, tokens: &[u8]) -> Result<()> {
    let len = u16::try_from(tokens.len())
        .map_err(|_| invalid(FEATURE11_RECORD_TYPE, "formula token length exceeds 65535"))?;
    out.extend_from_slice(&len.to_le_bytes());
    out.extend_from_slice(tokens);
    Ok(())
}
fn parse_formula(data: &[u8], offset: &mut usize, rt: u16, field: &str) -> Result<Vec<u8>> {
    let len = usize::from(u16_at(data, *offset, rt, field)?);
    if len == 0 {
        return Err(invalid(rt, format!("empty {field}")));
    }
    let end = (*offset)
        .checked_add(2 + len)
        .ok_or_else(|| invalid(rt, format!("{field} length overflows")))?;
    let value = data
        .get(*offset + 2..end)
        .ok_or_else(|| invalid(rt, format!("truncated {field}")))?
        .to_vec();
    *offset = end;
    Ok(value)
}
fn append_web_info(out: &mut Vec<u8>, info: &WebFieldInfo) -> Result<()> {
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
                out.extend_from_slice(&v.to_le_bytes())
            },
        }
    }
    if let Some(v) = &info.validation_formula {
        append_string(out, v)
    }
    out.extend_from_slice(&0u32.to_le_bytes());
    Ok(())
}
fn parse_web_info(
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

impl ListObject {
    pub(crate) fn to_feature_record_bytes(&self) -> Result<Vec<Vec<u8>>> {
        self.validate()?;
        if let Some(opaque) = &self.opaque_feature {
            let mut records = vec![record(opaque.record_type, opaque.base_payload.clone())?];
            for payload in &opaque.continuation_payloads {
                records.push(record(CONTINUE_FRT11_RECORD_TYPE, payload.clone())?);
            }
            return Ok(records);
        }
        if let Some(metadata) = &self.external_metadata {
            return self.to_external_feature_record_bytes(metadata);
        }
        if let Some(metadata) = &self.source_metadata {
            return self.to_source_feature_record_bytes(metadata);
        }
        let rt = match self.feature_version {
            ListObjectFeatureVersion::Feature11 => FEATURE11_RECORD_TYPE,
            ListObjectFeatureVersion::Feature12 => FEATURE12_RECORD_TYPE,
        };
        let mut feature = Vec::new();
        feature.extend_from_slice(&0u32.to_le_bytes());
        feature.extend_from_slice(&self.id.value().to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_header).to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_totals).to_le_bytes());
        let next = self
            .columns
            .iter()
            .map(|c| c.id.value())
            .max()
            .unwrap()
            .checked_add(1)
            .ok_or_else(|| invalid(FEATURE11_RECORD_TYPE, "column id overflows"))?;
        feature.extend_from_slice(&next.to_le_bytes());
        feature.extend_from_slice(&64u32.to_le_bytes());
        feature.extend_from_slice(&[0; 4]);
        let mut flags = 0x001B_0000u32;
        if self.autofilter {
            flags |= 0x806;
        }
        if self.has_totals {
            flags |= 0x40;
        }
        feature.extend_from_slice(&flags.to_le_bytes());
        feature.extend_from_slice(&[0; 32]);
        append_string(&mut feature, &self.name);
        feature.extend_from_slice(&(self.columns.len() as u16).to_le_bytes());
        append_string(&mut feature, &self.id.value().to_string());
        for column in &self.columns {
            feature.extend_from_slice(&column.id.value().to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&column.aggregation.code().to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&u32::MAX.to_le_bytes());
            let cflags = u32::from(self.autofilter)
                | (u32::from(column.total_formula.is_some()) << 7)
                | (u32::from(column.total_string.is_some()) << 10);
            feature.extend_from_slice(&cflags.to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&u32::MAX.to_le_bytes());
            append_string(&mut feature, &column.name);
            append_string(&mut feature, &column.name);
            if self.autofilter {
                feature.extend_from_slice(&0u32.to_le_bytes());
                feature.extend_from_slice(&(column.id.value() as u16).to_le_bytes());
            }
            if let Some(tokens) = &column.total_formula {
                feature.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
                feature.extend_from_slice(tokens);
            }
            if let Some(value) = &column.total_string {
                append_string(&mut feature, value);
            }
        }
        let mut payload = Vec::new();
        append_frt(&mut payload, rt, Some(self.range));
        payload.extend_from_slice(&ISF_LIST.to_le_bytes());
        payload.push(0);
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        append_range(&mut payload, self.range);
        payload.extend_from_slice(&feature);
        if payload.len() > MAX_FEATURE_BYTES {
            return Err(invalid(
                rt,
                "table feature exceeds aggregate resource bound",
            ));
        }
        let first_len = payload.len().min(MAX_PAYLOAD);
        let mut records = vec![record(rt, payload[..first_len].to_vec())?];
        for chunk in payload[first_len..].chunks(MAX_CONTINUE_RGB) {
            let mut continuation = Vec::with_capacity(12 + chunk.len());
            append_frt(&mut continuation, CONTINUE_FRT11_RECORD_TYPE, None);
            continuation.extend_from_slice(chunk);
            records.push(record(CONTINUE_FRT11_RECORD_TYPE, continuation)?);
        }
        Ok(records)
    }
    fn to_source_feature_record_bytes(
        &self,
        source: &ListObjectSourceMetadata,
    ) -> Result<Vec<Vec<u8>>> {
        let rt = match self.feature_version {
            ListObjectFeatureVersion::Feature11 => FEATURE11_RECORD_TYPE,
            ListObjectFeatureVersion::Feature12 => FEATURE12_RECORD_TYPE,
        };
        let (lt, version, build, single, fields_len): (u32, _, _, _, _) = match source {
            ListObjectSourceMetadata::Web(v) => {
                (1u32, v.version, v.build_number, false, v.fields.len())
            },
            ListObjectSourceMetadata::Xml(v) => (
                2u32,
                v.version,
                v.build_number,
                v.single_cell,
                v.fields.len(),
            ),
        };
        let mut feature = Vec::new();
        feature.extend_from_slice(&lt.to_le_bytes());
        feature.extend_from_slice(&self.id.value().to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_header).to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_totals).to_le_bytes());
        let next = self
            .columns
            .iter()
            .map(|c| c.id.value())
            .max()
            .unwrap()
            .checked_add(1)
            .ok_or_else(|| invalid(FEATURE11_RECORD_TYPE, "column id overflows"))?;
        feature.extend_from_slice(&next.to_le_bytes());
        feature.extend_from_slice(&64u32.to_le_bytes());
        feature.extend_from_slice(&build.to_le_bytes());
        let ignored_fixed_word = match source {
            ListObjectSourceMetadata::Web(v) => v.ignored_fixed_word,
            ListObjectSourceMetadata::Xml(v) => v.ignored_fixed_word,
        };
        feature.extend_from_slice(&ignored_fixed_word.to_le_bytes());
        let mut flags = (version.code() << 16)
            | (u32::from(self.autofilter) * 0x806)
            | (u32::from(self.has_totals) * 0x40)
            | (u32::from(single) << 9)
            | 0x0040_0000;
        match source {
            ListObjectSourceMetadata::Web(v) => {
                flags |= u32::from(!v.deleted_row_ids.is_empty()) << 5
                    | u32::from(v.needs_commit) << 8
                    | u32::from(v.compressed_cache) << 13
                    | u32::from(v.provider_name.is_some()) << 14
                    | u32::from(!v.changed_row_ids.is_empty()) << 15
                    | u32::from(v.entry_id.is_some()) << 20
                    | u32::from(!v.invalid_cells.is_empty()) << 21
                    | v.ignored_flags
            },
            ListObjectSourceMetadata::Xml(v) => {
                flags |= u32::from(v.entry_id.is_some()) << 20 | v.ignored_flags
            },
        };
        feature.extend_from_slice(&flags.to_le_bytes());
        match source {
            ListObjectSourceMetadata::Web(v) => {
                feature.extend_from_slice(&v.cache_position.to_le_bytes());
                feature.extend_from_slice(&v.cache_size.to_le_bytes());
                feature.extend_from_slice(&v.cache_characters.to_le_bytes());
                feature.extend_from_slice(&v.edit_mode.code().to_le_bytes());
                feature.extend_from_slice(&v.hash_parameters)
            },
            ListObjectSourceMetadata::Xml(v) => feature.extend_from_slice(&v.ignored_fixed_tail),
        };
        append_string(&mut feature, &self.name);
        feature.extend_from_slice(&(fields_len as u16).to_le_bytes());
        match source {
            ListObjectSourceMetadata::Web(v) => {
                if let Some(name) = &v.provider_name {
                    append_string(&mut feature, name)
                }
                if let Some(entry) = &v.entry_id {
                    append_string(&mut feature, entry)
                }
            },
            ListObjectSourceMetadata::Xml(v) => {
                if let Some(entry) = &v.entry_id {
                    append_string(&mut feature, entry)
                }
            },
        }
        for (index, column) in self.columns.iter().enumerate() {
            let (web, xml) = match source {
                ListObjectSourceMetadata::Web(v) => (Some(&v.fields[index]), None),
                ListObjectSourceMetadata::Xml(v) => (None, Some(&v.fields[index])),
            };
            let (
                source_name,
                web_type,
                xml_type,
                mapped,
                calc,
                auto_filter,
                agg_fmt,
                insert_fmt,
                total_extra,
                ignored_flags,
            ) = if let Some(v) = web {
                (
                    &v.source_name,
                    v.data_type.code(),
                    0,
                    false,
                    v.calculated_formula.as_deref(),
                    v.auto_filter.as_slice(),
                    v.aggregate_format.as_slice(),
                    v.insert_row_format.as_slice(),
                    v.total_formula_extra.as_slice(),
                    v.ignored_flags,
                )
            } else {
                let v = xml.unwrap();
                (
                    &v.source_name,
                    0,
                    v.data_type.value(),
                    v.mapping.is_some(),
                    None,
                    v.auto_filter.as_slice(),
                    v.aggregate_format.as_slice(),
                    v.insert_row_format.as_slice(),
                    v.total_formula_extra.as_slice(),
                    v.ignored_flags,
                )
            };
            feature.extend_from_slice(&column.id.value().to_le_bytes());
            feature.extend_from_slice(&web_type.to_le_bytes());
            feature.extend_from_slice(&xml_type.to_le_bytes());
            feature.extend_from_slice(&column.aggregation.code().to_le_bytes());
            feature.extend_from_slice(&(agg_fmt.len() as u32).to_le_bytes());
            feature.extend_from_slice(&u32::MAX.to_le_bytes());
            let ff = u32::from(self.autofilter)
                | (u32::from(mapped) << 2)
                | (u32::from(calc.is_some()) << 3)
                | (u32::from(column.total_formula.is_some()) << 7)
                | (u32::from(!total_extra.is_empty()) << 8)
                | (u32::from(column.total_string.is_some()) << 10)
                | ignored_flags;
            feature.extend_from_slice(&ff.to_le_bytes());
            feature.extend_from_slice(&(insert_fmt.len() as u32).to_le_bytes());
            feature.extend_from_slice(&u32::MAX.to_le_bytes());
            append_string(&mut feature, source_name);
            if !single {
                append_string(&mut feature, &column.name)
            }
            feature.extend_from_slice(agg_fmt);
            feature.extend_from_slice(insert_fmt);
            if self.autofilter {
                feature.extend_from_slice(auto_filter)
            }
            if let Some(v) = xml.and_then(|v| v.mapping.as_ref()) {
                feature.extend_from_slice(&1u16.to_le_bytes());
                feature.extend_from_slice(&(2u32 | u32::from(v.can_be_single) << 2).to_le_bytes());
                feature.extend_from_slice(&v.map_id.to_le_bytes());
                append_string(&mut feature, &v.xpath)
            }
            if let Some(tokens) = calc {
                append_formula(&mut feature, tokens)?
            }
            if let Some(tokens) = &column.total_formula {
                append_formula(&mut feature, tokens)?;
                feature.extend_from_slice(total_extra)
            }
            if let Some(value) = &column.total_string {
                append_string(&mut feature, value)
            }
            if let Some(v) = web {
                append_web_info(&mut feature, &v.info)?
            }
        }
        if let ListObjectSourceMetadata::Web(v) = source {
            if !v.deleted_row_ids.is_empty() {
                feature.extend_from_slice(&(v.deleted_row_ids.len() as u16).to_le_bytes());
                for id in &v.deleted_row_ids {
                    feature.extend_from_slice(&id.to_le_bytes())
                }
            }
            if !v.changed_row_ids.is_empty() {
                feature.extend_from_slice(&(v.changed_row_ids.len() as u16).to_le_bytes());
                for id in &v.changed_row_ids {
                    feature.extend_from_slice(&id.to_le_bytes())
                }
            }
            if !v.invalid_cells.is_empty() {
                feature.extend_from_slice(&(v.invalid_cells.len() as u16).to_le_bytes());
                for cell in &v.invalid_cells {
                    feature.extend_from_slice(&cell.row_id.to_le_bytes());
                    feature.extend_from_slice(&cell.column_id.value().to_le_bytes())
                }
            }
        }
        let mut payload = Vec::new();
        append_frt(&mut payload, rt, Some(self.range));
        payload.extend_from_slice(&ISF_LIST.to_le_bytes());
        payload.push(0);
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&(feature.len() as u32).to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        append_range(&mut payload, self.range);
        payload.extend_from_slice(&feature);
        if payload.len() > MAX_FEATURE_BYTES {
            return Err(invalid(
                rt,
                "Web/XML table feature exceeds aggregate resource bound",
            ));
        }
        let first = payload.len().min(MAX_PAYLOAD);
        let mut records = vec![record(rt, payload[..first].to_vec())?];
        for chunk in payload[first..].chunks(MAX_CONTINUE_RGB) {
            let mut continuation = Vec::with_capacity(12 + chunk.len());
            append_frt(&mut continuation, CONTINUE_FRT11_RECORD_TYPE, None);
            continuation.extend_from_slice(chunk);
            records.push(record(CONTINUE_FRT11_RECORD_TYPE, continuation)?)
        }
        Ok(records)
    }
    fn to_external_feature_record_bytes(
        &self,
        metadata: &ExternalTableMetadata,
    ) -> Result<Vec<Vec<u8>>> {
        let mut feature = Vec::new();
        feature.extend_from_slice(&3u32.to_le_bytes());
        feature.extend_from_slice(&self.id.value().to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_header).to_le_bytes());
        feature.extend_from_slice(&u32::from(self.has_totals).to_le_bytes());
        let next = self
            .columns
            .iter()
            .map(|column| column.id.value())
            .max()
            .unwrap()
            .checked_add(1)
            .ok_or_else(|| invalid(FEATURE12_RECORD_TYPE, "column id overflows"))?;
        feature.extend_from_slice(&next.to_le_bytes());
        feature.extend_from_slice(&64u32.to_le_bytes());
        feature.extend_from_slice(&metadata.build_number.to_le_bytes());
        feature.extend_from_slice(&0u16.to_le_bytes());
        let mut flags = metadata.version.code() << 16 | 0x0010_0000;
        if self.autofilter {
            flags |= 0x0000_0806;
        }
        if self.has_totals {
            flags |= 0x40;
        }
        feature.extend_from_slice(&flags.to_le_bytes());
        feature.extend_from_slice(&[0; 12]);
        feature.extend_from_slice(&0u32.to_le_bytes());
        feature.extend_from_slice(&[0; 16]);
        append_string(&mut feature, &self.name);
        feature.extend_from_slice(&(self.columns.len() as u16).to_le_bytes());
        append_string(&mut feature, &self.id.value().to_string());
        for (column, field) in self.columns.iter().zip(&metadata.fields) {
            feature.extend_from_slice(&column.id.value().to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&0u32.to_le_bytes());
            feature.extend_from_slice(&column.aggregation.code().to_le_bytes());
            feature.extend_from_slice(&(field.aggregate_format.len() as u32).to_le_bytes());
            feature.extend_from_slice(&field.aggregate_style.to_le_bytes());
            let mut field_flags = u32::from(self.autofilter)
                | (u32::from(field.filter_hidden) << 1)
                | (u32::from(column.total_formula.is_some()) << 7)
                | (u32::from(field.total_array_formula) << 8)
                | (u32::from(column.total_string.is_some()) << 10)
                | (u32::from(field.auto_create_calculated_column) << 11);
            if !self.has_header && field.header_cache.style_name().is_some() {
                field_flags |= 0x200;
            }
            feature.extend_from_slice(&field_flags.to_le_bytes());
            feature.extend_from_slice(&(field.insert_row_format.len() as u32).to_le_bytes());
            feature.extend_from_slice(&field.insert_row_style.to_le_bytes());
            append_string(&mut feature, &field.source_name);
            append_string(&mut feature, &column.name);
            feature.extend_from_slice(&field.aggregate_format);
            feature.extend_from_slice(&field.insert_row_format);
            if self.autofilter {
                feature.extend_from_slice(&field.auto_filter);
            }
            if let Some(tokens) = &column.total_formula {
                feature.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
                feature.extend_from_slice(tokens);
                feature.extend_from_slice(&field.formula_extra);
            }
            if let Some(value) = &column.total_string {
                append_string(&mut feature, value);
            }
            feature.extend_from_slice(&field.query_field_id.to_le_bytes());
            if !self.has_header {
                feature.extend_from_slice(field.header_cache.as_bytes());
            }
        }
        let mut payload = Vec::new();
        append_frt(&mut payload, FEATURE12_RECORD_TYPE, Some(self.range));
        payload.extend_from_slice(&ISF_LIST.to_le_bytes());
        payload.push(0);
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&(feature.len() as u32).to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        append_range(&mut payload, self.range);
        payload.extend_from_slice(&feature);
        if payload.len() > MAX_FEATURE_BYTES {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "external table feature exceeds aggregate resource bound",
            ));
        }
        let first_len = payload.len().min(MAX_PAYLOAD);
        let mut records = vec![record(
            FEATURE12_RECORD_TYPE,
            payload[..first_len].to_vec(),
        )?];
        for chunk in payload[first_len..].chunks(MAX_CONTINUE_RGB) {
            let mut continuation = Vec::with_capacity(12 + chunk.len());
            append_frt(&mut continuation, CONTINUE_FRT11_RECORD_TYPE, None);
            continuation.extend_from_slice(chunk);
            records.push(record(CONTINUE_FRT11_RECORD_TYPE, continuation)?);
        }
        Ok(records)
    }
    pub(crate) fn to_list12_record_bytes(&self) -> Result<Vec<Vec<u8>>> {
        let mut block = Vec::new();
        append_frt(&mut block, LIST12_RECORD_TYPE, None);
        block.extend_from_slice(&0u16.to_le_bytes());
        block.extend_from_slice(&self.id.value().to_le_bytes());
        for v in [0i32, -1, 0, -1, 0, -1, 0, 0, 0] {
            block.extend_from_slice(&v.to_le_bytes());
        }
        let style = self.style.as_ref().unwrap();
        let mut styled = Vec::new();
        append_frt(&mut styled, LIST12_RECORD_TYPE, None);
        styled.extend_from_slice(&1u16.to_le_bytes());
        styled.extend_from_slice(&self.id.value().to_le_bytes());
        let bits = u16::from(style.first)
            | u16::from(style.last) << 1
            | u16::from(style.row_stripes) << 2
            | u16::from(style.column_stripes) << 3
            | u16::from(style.default_style) << 6;
        styled.extend_from_slice(&bits.to_le_bytes());
        append_string(&mut styled, &style.name);
        let mut display = Vec::new();
        append_frt(&mut display, LIST12_RECORD_TYPE, None);
        display.extend_from_slice(&2u16.to_le_bytes());
        display.extend_from_slice(&self.id.value().to_le_bytes());
        append_string(&mut display, &self.name);
        append_string(&mut display, &self.comment);
        Ok(vec![
            record(LIST12_RECORD_TYPE, block)?,
            record(LIST12_RECORD_TYPE, styled)?,
            record(LIST12_RECORD_TYPE, display)?,
        ])
    }
    pub(crate) fn to_following_record_bytes(&self) -> Result<Vec<Vec<u8>>> {
        let list12 = self.to_list12_record_bytes()?;
        let mut output = Vec::new();
        for (index, item) in list12.into_iter().enumerate() {
            output.push(item);
            if index == 0
                && let Some(filter) = &self.autofilter12_criteria
            {
                output.extend(write_table_autofilter12(filter, self.range, self.id)?);
            }
            for future in self
                .opaque_future_records
                .iter()
                .filter(|v| v.after_list12_count == index + 1)
            {
                output.push(record(future.record_type, future.payload.clone())?);
                for payload in &future.continuation_payloads {
                    output.push(record(
                        crate::sort_data::CONTINUE_FRT12_RECORD_TYPE,
                        payload.clone(),
                    )?);
                }
            }
        }
        if self
            .opaque_future_records
            .iter()
            .any(|v| v.after_list12_count == 0 || v.after_list12_count > 3)
        {
            return Err(invalid(
                AUTO_FILTER12_RECORD_TYPE,
                "opaque table future-record insertion point is invalid",
            ));
        }
        Ok(output)
    }
    pub(super) fn parse_opaque_feature12(pending: PendingFeature) -> Result<Self> {
        let data = &pending.combined;
        if data.len() < 108 {
            return Err(invalid(FEATURE12_RECORD_TYPE, "truncated opaque Feature12"));
        }
        validate_frt(data, FEATURE12_RECORD_TYPE, true)?;
        if u16_at(data, 12, FEATURE12_RECORD_TYPE, "isf")? != ISF_LIST
            || data[14] != 0
            || u32_at(data, 55, FEATURE12_RECORD_TYPE, "cbFSData")? != 64
        {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "invalid opaque Feature12 fixed fields",
            ));
        }
        let range = parse_range(data, 4, FEATURE12_RECORD_TYPE)?;
        let id = ListObjectId::try_new(u32_at(data, 39, FEATURE12_RECORD_TYPE, "idList")?)?;
        let header = u32_at(data, 43, FEATURE12_RECORD_TYPE, "crwHeader")?;
        let totals = u32_at(data, 47, FEATURE12_RECORD_TYPE, "crwTotals")?;
        if header > 1 || totals > 1 {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "invalid opaque Feature12 row flags",
            ));
        }
        let flags = u32_at(data, 63, FEATURE12_RECORD_TYPE, "flags")?;
        let (name, _) = parse_string(data, 99, FEATURE12_RECORD_TYPE, "rgbName")?;
        let opaque_feature = OpaqueListObjectFeature {
            record_type: pending.record_type,
            base_payload: pending.base,
            continuation_payloads: pending.continuations,
        };
        Ok(Self {
            id,
            name,
            range,
            columns: Vec::new(),
            style: None,
            has_header: header != 0,
            has_totals: totals != 0,
            autofilter: flags & 2 != 0,
            comment: String::new(),
            feature_version: ListObjectFeatureVersion::Feature12,
            opaque_feature: Some(opaque_feature),
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: None,
            source_metadata: None,
        })
    }
    pub(super) fn parse_external_feature12(pending: PendingFeature) -> Result<Self> {
        let data = &pending.combined;
        let rt = FEATURE12_RECORD_TYPE;
        if !(99..=MAX_FEATURE_BYTES).contains(&data.len()) {
            return Err(invalid(rt, "invalid external table feature length"));
        }
        validate_frt(data, rt, true)?;
        let range = parse_range(data, 4, rt)?;
        if u16_at(data, 12, rt, "isf")? != ISF_LIST
            || data[14] != 0
            || u32_at(data, 15, rt, "reserved2")? != 0
            || u16_at(data, 19, rt, "cref2")? != 1
            || u16_at(data, 25, rt, "reserved3")? != 0
            || parse_range(data, 27, rt)? != range
        {
            return Err(invalid(rt, "invalid Feature12 fixed fields"));
        }
        let declared = usize::try_from(u32_at(data, 21, rt, "cbFeatData")?)
            .map_err(|_| invalid(rt, "cbFeatData overflows"))?;
        if declared != 0 && declared != data.len() - 35 {
            return Err(invalid(
                rt,
                "cbFeatData does not match external feature size",
            ));
        }
        let base = 35;
        if u32_at(data, base, rt, "lt")? != 3 {
            return Err(invalid(rt, "external parser requires LTEXTERNALDATA"));
        }
        let id = ListObjectId::try_new(u32_at(data, base + 4, rt, "idList")?)?;
        let header = u32_at(data, base + 8, rt, "crwHeader")?;
        let totals = u32_at(data, base + 12, rt, "crwTotals")?;
        if header > 1
            || totals > 1
            || u32_at(data, base + 20, rt, "cbFSData")? != 64
            || u32_at(data, base + 44, rt, "lem")? != 0
        {
            return Err(invalid(
                rt,
                "invalid external TableFeatureType fixed fields",
            ));
        }
        let flags = u32_at(data, base + 28, rt, "flags")?;
        let version = ExternalTableVersion::from_code((flags >> 16) & 0xF)?;
        if flags & 0x0020_E7A0 != 0
            || flags & 0x4 != 0 && flags & 0x2 == 0
            || flags & 0x10 != 0 && flags & 0x8 == 0
            || flags & 0x2 != 0 && header == 0
        {
            return Err(invalid(rt, "invalid external table flags"));
        }
        let (name, mut offset) = parse_string(data, base + 64, rt, "rgbName")?;
        validate_table_name(&name)?;
        let count = usize::from(u16_at(data, offset, rt, "cFieldData")?);
        offset += 2;
        if !(1..=256).contains(&count) || count != range.column_count() {
            return Err(invalid(rt, "external field count must match table range"));
        }
        if flags & 0x0000_4000 != 0 {
            return Err(invalid(
                rt,
                "external table cannot load a SharePoint CSP name",
            ));
        }
        if flags & 0x0010_0000 != 0 {
            offset = parse_string(data, offset, rt, "entryId")?.1;
        }
        let mut columns = Vec::with_capacity(count);
        let mut fields = Vec::with_capacity(count);
        for _ in 0..count {
            let start = offset;
            let cid = ListColumnId::try_new(u32_at(data, start, rt, "idField")?)?;
            if u32_at(data, start + 4, rt, "lfdt")? != 0
                || u32_at(data, start + 8, rt, "lfxidt")? != 0
            {
                return Err(invalid(rt, "external field data types must be zero"));
            }
            let aggregation =
                ListTotalAggregation::from_code(u32_at(data, start + 12, rt, "ilta")?)?;
            let aggregate_len = usize::try_from(u32_at(data, start + 16, rt, "cbFmtAgg")?)
                .map_err(|_| invalid(rt, "aggregate format length overflows"))?;
            let aggregate_style = u32_at(data, start + 20, rt, "istnAgg")?;
            let field_flags = u32_at(data, start + 24, rt, "field flags")?;
            let insert_len = usize::try_from(u32_at(data, start + 28, rt, "cbFmtInsertRow")?)
                .map_err(|_| invalid(rt, "insert format length overflows"))?;
            let insert_row_style = u32_at(data, start + 32, rt, "istnInsertRow")?;
            if field_flags & 0x4C != 0
                || field_flags & 0x40 != 0
                || field_flags & 0x100 != 0 && field_flags & 0x80 == 0
                || field_flags & 0x80 != 0 && aggregation != ListTotalAggregation::Custom
                || field_flags & 0x400 != 0 && aggregation != ListTotalAggregation::None
                || field_flags & 2 != 0 && field_flags & 1 == 0
                || (field_flags & 1 != 0) != (flags & 2 != 0)
            {
                return Err(invalid(rt, "invalid external field flags"));
            }
            let (source_name, after_source) = parse_string(data, start + 36, rt, "strFieldName")?;
            validate_name(&source_name, "external source field name")?;
            let (caption, after_caption) = parse_string(data, after_source, rt, "strCaption")?;
            validate_column_name(&caption)?;
            let aggregate_end = after_caption
                .checked_add(aggregate_len)
                .ok_or_else(|| invalid(rt, "aggregate format length overflows"))?;
            let aggregate_format = data
                .get(after_caption..aggregate_end)
                .ok_or_else(|| invalid(rt, "truncated aggregate format"))?
                .to_vec();
            let insert_end = aggregate_end
                .checked_add(insert_len)
                .ok_or_else(|| invalid(rt, "insert format length overflows"))?;
            let insert_row_format = data
                .get(aggregate_end..insert_end)
                .ok_or_else(|| invalid(rt, "truncated insert-row format"))?
                .to_vec();
            offset = insert_end;
            let auto_filter = if field_flags & 1 != 0 {
                let size = usize::try_from(u32_at(data, offset, rt, "cbAutoFilter")?)
                    .map_err(|_| invalid(rt, "AutoFilter length overflows"))?;
                if size > 2080 {
                    return Err(invalid(rt, "AutoFilter exceeds 2080 bytes"));
                }
                let end = offset
                    .checked_add(6 + size)
                    .ok_or_else(|| invalid(rt, "AutoFilter length overflows"))?;
                let value = data
                    .get(offset..end)
                    .ok_or_else(|| invalid(rt, "truncated AutoFilter"))?
                    .to_vec();
                offset = end;
                value
            } else {
                vec![0; 6]
            };
            let (total_formula, formula_extra) = if field_flags & 0x80 != 0 {
                let size = usize::from(u16_at(data, offset, rt, "total formula length")?);
                if size == 0 {
                    return Err(invalid(rt, "empty total formula"));
                }
                let token_end = offset
                    .checked_add(2 + size)
                    .ok_or_else(|| invalid(rt, "total formula length overflows"))?;
                let tokens = data
                    .get(offset + 2..token_end)
                    .ok_or_else(|| invalid(rt, "truncated total formula"))?
                    .to_vec();
                offset = token_end;
                let extra_end = if field_flags & 0x100 != 0 {
                    parse_list_formula_extra_end(data, &tokens, offset, rt)?
                } else {
                    offset
                };
                let extra = data
                    .get(offset..extra_end)
                    .ok_or_else(|| invalid(rt, "truncated formula extra data"))?
                    .to_vec();
                offset = extra_end;
                (Some(tokens), extra)
            } else {
                (None, Vec::new())
            };
            let total_string = if field_flags & 0x400 != 0 {
                let (value, end) = parse_string(data, offset, rt, "strTotal")?;
                offset = end;
                Some(value)
            } else {
                None
            };
            let query_field_id = u32_at(data, offset, rt, "qsif")?;
            if query_field_id == 0 {
                return Err(invalid(rt, "external qsif must be nonzero"));
            }
            offset += 4;
            let header_cache = if header == 0 {
                let size = usize::try_from(u32_at(data, offset, rt, "cbdxfHdrDisk")?)
                    .map_err(|_| invalid(rt, "header cache length overflows"))?;
                let format_end = offset
                    .checked_add(4 + size)
                    .ok_or_else(|| invalid(rt, "header cache length overflows"))?;
                data.get(offset..format_end)
                    .ok_or_else(|| invalid(rt, "truncated header cache"))?;
                let end = if field_flags & 0x200 != 0 {
                    parse_string(data, format_end, rt, "header style name")?.1
                } else {
                    format_end
                };
                let value = CachedDiskHeader::parse(
                    data[offset..end].to_vec(),
                    field_flags & 0x200 != 0,
                    rt,
                )?;
                offset = end;
                value
            } else {
                if field_flags & 0x200 != 0 {
                    return Err(invalid(rt, "header style name requires a CachedDiskHeader"));
                }
                CachedDiskHeader::empty()
            };
            let column = ListObjectColumn {
                id: cid,
                name: caption,
                aggregation,
                total_formula,
                total_string,
            };
            column.validate_totals()?;
            columns.push(column);
            fields.push(ExternalTableField {
                column_id: cid,
                source_name,
                query_field_id,
                aggregate_format,
                insert_row_format,
                auto_filter,
                formula_extra,
                header_cache,
                aggregate_style,
                insert_row_style,
                filter_hidden: field_flags & 2 != 0,
                total_array_formula: field_flags & 0x100 != 0,
                auto_create_calculated_column: field_flags & 0x800 != 0,
            });
        }
        if offset != data.len() {
            return Err(invalid(rt, "trailing external Feature12 data"));
        }
        let metadata = ExternalTableMetadata {
            version,
            build_number: u16_at(data, base + 24, rt, "rupBuild")?,
            fields,
        };
        metadata.validate()?;
        let opaque_feature = OpaqueListObjectFeature {
            record_type: pending.record_type,
            base_payload: pending.base,
            continuation_payloads: pending.continuations,
        };
        Ok(Self {
            id,
            name,
            range,
            columns,
            style: None,
            has_header: header != 0,
            has_totals: totals != 0,
            autofilter: flags & 2 != 0,
            comment: String::new(),
            feature_version: ListObjectFeatureVersion::Feature12,
            opaque_feature: Some(opaque_feature),
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: Some(metadata),
            source_metadata: None,
        })
    }
    fn parse_source_feature(data: &[u8], rt: u16, lt: u32) -> Result<Self> {
        if !matches!(rt, FEATURE11_RECORD_TYPE | FEATURE12_RECORD_TYPE) {
            return Err(invalid(
                rt,
                "Web/XML source table requires Feature11 or Feature12",
            ));
        }
        let range = parse_range(data, 4, rt)?;
        let base = 35;
        let id = ListObjectId::try_new(u32_at(data, base + 4, rt, "idList")?)?;
        let header = u32_at(data, base + 8, rt, "crwHeader")?;
        let totals = u32_at(data, base + 12, rt, "crwTotals")?;
        if header > 1 || totals > 1 || u32_at(data, base + 20, rt, "cbFSData")? != 64 {
            return Err(invalid(rt, "invalid Web/XML TableFeatureType fixed fields"));
        }
        let build = u16_at(data, base + 24, rt, "rupBuild")?;
        let ignored_fixed_word = u16_at(data, base + 26, rt, "unused1")?;
        let flags = u32_at(data, base + 28, rt, "flags")?;
        let version = ExternalTableVersion::from_code((flags >> 16) & 0xf)?;
        if flags & 0x0000_0480 != 0
            || flags & 4 != 0 && flags & 2 == 0
            || flags & 0x10 != 0 && flags & 8 == 0
            || flags & 2 != 0 && header == 0
        {
            return Err(invalid(rt, "invalid Web/XML table flags"));
        }
        let single = flags & 0x200 != 0;
        if single && (lt != 2 || header != 0 || totals != 0 || range.column_count() != 1) {
            return Err(invalid(rt, "invalid single-cell XML table"));
        }
        if lt == 2 && flags & 0x0020_e120 != 0 {
            return Err(invalid(rt, "Web-only flags occur on XML table"));
        }
        let mut ignored_fixed_tail = [0; 32];
        let (cache_position, cache_size, cache_characters, edit_mode, hash) = if lt == 1 {
            let mut hash = [0; 16];
            hash.copy_from_slice(
                data.get(base + 48..base + 64)
                    .ok_or_else(|| invalid(rt, "truncated Web hash parameters"))?,
            );
            (
                u32_at(data, base + 32, rt, "cache position")?,
                u32_at(data, base + 36, rt, "cache size")?,
                u32_at(data, base + 40, rt, "cache characters")?,
                WebEditMode::from_code(u32_at(data, base + 44, rt, "edit mode")?)?,
                hash,
            )
        } else {
            ignored_fixed_tail.copy_from_slice(
                data.get(base + 32..base + 64)
                    .ok_or_else(|| invalid(rt, "truncated XML fixed tail"))?,
            );
            if ignored_fixed_tail[12..16].iter().any(|byte| *byte != 0) {
                return Err(invalid(rt, "XML edit mode must be zero"));
            }
            (0, 0, 0, WebEditMode::Normal, [0; 16])
        };
        let (name, mut offset) = parse_string(data, base + 64, rt, "rgbName")?;
        validate_table_name(&name)?;
        let count = usize::from(u16_at(data, offset, rt, "cFieldData")?);
        offset += 2;
        if !(1..=256).contains(&count) || count != range.column_count() {
            return Err(invalid(rt, "Web/XML field count must match table range"));
        }
        let provider_name = if flags & 0x4000 != 0 {
            let (v, end) = parse_string(data, offset, rt, "cSPName")?;
            offset = end;
            Some(v)
        } else {
            None
        };
        let entry_id = if flags & 0x0010_0000 != 0 {
            let (v, end) = parse_string(data, offset, rt, "entryId")?;
            offset = end;
            if lt == 2 && v != id.value().to_string() {
                return Err(invalid(rt, "XML entryId does not match table id"));
            }
            Some(v)
        } else {
            None
        };
        let mut columns = Vec::with_capacity(count);
        let mut web_fields = Vec::with_capacity(count);
        let mut xml_fields = Vec::with_capacity(count);
        for _ in 0..count {
            let start = offset;
            let cid = ListColumnId::try_new(u32_at(data, start, rt, "idField")?)?;
            let web_type = u32_at(data, start + 4, rt, "lfdt")?;
            let xml_type = u32_at(data, start + 8, rt, "lfxidt")?;
            if (lt == 1 && xml_type != 0) || (lt == 2 && web_type != 0) {
                return Err(invalid(rt, "field data type does not match table source"));
            }
            let aggregation =
                ListTotalAggregation::from_code(u32_at(data, start + 12, rt, "ilta")?)?;
            let agg_len = usize::try_from(u32_at(data, start + 16, rt, "cbFmtAgg")?)
                .map_err(|_| invalid(rt, "aggregate format length overflows"))?;
            let field_flags = u32_at(data, start + 24, rt, "field flags")?;
            let insert_len = usize::try_from(u32_at(data, start + 28, rt, "cbFmtInsertRow")?)
                .map_err(|_| invalid(rt, "insert-row format length overflows"))?;
            if field_flags & 0x0000_0040 != 0
                || field_flags & 2 != 0 && field_flags & 1 == 0
                || field_flags & 0x100 != 0 && field_flags & 0x80 == 0
                || field_flags & 0x80 != 0 && aggregation != ListTotalAggregation::Custom
                || field_flags & 0x400 != 0 && aggregation != ListTotalAggregation::None
                || (field_flags & 1 != 0) != (flags & 2 != 0)
                || (lt == 1 && field_flags & 0x804 != 0)
                || (lt == 2 && field_flags & 8 != 0)
                || (rt == FEATURE11_RECORD_TYPE && field_flags & 0x480 != 0)
            {
                return Err(invalid(rt, "invalid source field condition flags"));
            }
            let (source_name, after_source) = parse_string(data, start + 36, rt, "strFieldName")?;
            let (caption, after_caption) = if single {
                (source_name.clone(), after_source)
            } else {
                parse_string(data, after_source, rt, "strCaption")?
            };
            validate_column_name(&caption)?;
            let agg_end = after_caption
                .checked_add(agg_len)
                .ok_or_else(|| invalid(rt, "aggregate format length overflows"))?;
            let aggregate_format = data
                .get(after_caption..agg_end)
                .ok_or_else(|| invalid(rt, "truncated aggregate format"))?
                .to_vec();
            let insert_end = agg_end
                .checked_add(insert_len)
                .ok_or_else(|| invalid(rt, "insert format length overflows"))?;
            let insert_row_format = data
                .get(agg_end..insert_end)
                .ok_or_else(|| invalid(rt, "truncated insert-row format"))?
                .to_vec();
            offset = insert_end;
            let auto_filter = if field_flags & 1 != 0 {
                let n = usize::try_from(u32_at(data, offset, rt, "cbAutoFilter")?)
                    .map_err(|_| invalid(rt, "AutoFilter size overflows"))?;
                if n > 2080 {
                    return Err(invalid(rt, "AutoFilter exceeds 2080 bytes"));
                }
                let end = offset
                    .checked_add(6 + n)
                    .ok_or_else(|| invalid(rt, "AutoFilter size overflows"))?;
                let v = data
                    .get(offset..end)
                    .ok_or_else(|| invalid(rt, "truncated AutoFilter"))?
                    .to_vec();
                offset = end;
                v
            } else {
                vec![0; 6]
            };
            let mapping = if field_flags & 4 != 0 {
                if u16_at(data, offset, rt, "iXmapMac")? != 1 {
                    return Err(invalid(rt, "XML mapped field must contain one map entry"));
                }
                let map_flags = u32_at(data, offset + 2, rt, "XMap flags")?;
                if map_flags & !6 != 0 || map_flags & 2 == 0 {
                    return Err(invalid(rt, "invalid XML map flags"));
                }
                let map_id = u32_at(data, offset + 6, rt, "XML map id")?;
                let (xpath, end) = parse_string(data, offset + 10, rt, "XPath")?;
                offset = end;
                Some(XmlColumnMapping::try_new(
                    map_id,
                    xpath,
                    map_flags & 4 != 0,
                )?)
            } else {
                None
            };
            let calculated_formula = if field_flags & 8 != 0 {
                Some(parse_formula(data, &mut offset, rt, "calculated formula")?)
            } else {
                None
            };
            let (total_formula, total_extra) = if field_flags & 0x80 != 0 {
                let tokens = parse_formula(data, &mut offset, rt, "total formula")?;
                let extra_end = if field_flags & 0x100 != 0 {
                    parse_list_formula_extra_end(data, &tokens, offset, rt)?
                } else {
                    offset
                };
                let extra = data
                    .get(offset..extra_end)
                    .ok_or_else(|| invalid(rt, "truncated total formula extra data"))?
                    .to_vec();
                offset = extra_end;
                (Some(tokens), extra)
            } else {
                (None, Vec::new())
            };
            let total_string = if field_flags & 0x400 != 0 {
                let (v, end) = parse_string(data, offset, rt, "strTotal")?;
                offset = end;
                Some(v)
            } else {
                None
            };
            let web_kind = if lt == 1 {
                Some(WebColumnType::from_code(web_type)?)
            } else {
                None
            };
            let web_info = if let Some(kind) = web_kind {
                Some(parse_web_info(data, &mut offset, kind, rt)?)
            } else {
                None
            };
            if header == 0 && !single {
                let n = usize::try_from(u32_at(data, offset, rt, "cached header format size")?)
                    .map_err(|_| invalid(rt, "cached header size overflows"))?;
                let end = offset
                    .checked_add(4 + n)
                    .ok_or_else(|| invalid(rt, "cached header size overflows"))?;
                data.get(offset..end)
                    .ok_or_else(|| invalid(rt, "truncated cached header"))?;
                offset = if field_flags & 0x200 != 0 {
                    parse_string(data, end, rt, "cached header style")?.1
                } else {
                    end
                }
            } else if field_flags & 0x200 != 0 {
                return Err(invalid(rt, "cached header style lacks cached header"));
            }
            let column = ListObjectColumn {
                id: cid,
                name: caption,
                aggregation,
                total_formula,
                total_string,
            };
            column.validate_totals()?;
            columns.push(column);
            if let (Some(kind), Some(info)) = (web_kind, web_info) {
                web_fields.push(WebTableField {
                    column_id: cid,
                    source_name,
                    data_type: kind,
                    info,
                    calculated_formula,
                    auto_filter,
                    aggregate_format,
                    insert_row_format,
                    total_formula_extra: total_extra,
                    header_cache: vec![0; 4],
                    ignored_flags: field_flags & 0xffff_f030,
                })
            } else {
                xml_fields.push(XmlTableField {
                    column_id: cid,
                    source_name,
                    data_type: XmlDataType::try_new(xml_type)?,
                    mapping,
                    auto_filter,
                    aggregate_format,
                    insert_row_format,
                    total_formula_extra: total_extra,
                    header_cache: vec![0; 4],
                    ignored_flags: field_flags & 0xffff_f030,
                })
            }
        }
        let source_metadata = if lt == 1 {
            let parse_ids = |data: &[u8], offset: &mut usize, label: &str| -> Result<Vec<u32>> {
                let count = usize::from(u16_at(data, *offset, rt, label)?);
                *offset += 2;
                let end = (*offset)
                    .checked_add(
                        count
                            .checked_mul(4)
                            .ok_or_else(|| invalid(rt, "source id count overflows"))?,
                    )
                    .ok_or_else(|| invalid(rt, "source id count overflows"))?;
                let bytes = data
                    .get(*offset..end)
                    .ok_or_else(|| invalid(rt, format!("truncated {label}")))?;
                let ids = bytes
                    .chunks_exact(4)
                    .map(|v| u32::from_le_bytes(v.try_into().unwrap()))
                    .collect();
                *offset = end;
                Ok(ids)
            };
            let deleted_row_ids = if flags & 0x20 != 0 {
                parse_ids(data, &mut offset, "deleted row ids")?
            } else {
                Vec::new()
            };
            let changed_row_ids = if flags & 0x8000 != 0 {
                parse_ids(data, &mut offset, "changed row ids")?
            } else {
                Vec::new()
            };
            let invalid_cells = if flags & 0x0020_0000 != 0 {
                let n = usize::from(u16_at(data, offset, rt, "invalid cell count")?);
                offset += 2;
                let mut out = Vec::with_capacity(n);
                for _ in 0..n {
                    let row = u32_at(data, offset, rt, "invalid cell row")?;
                    let column =
                        ListColumnId::try_new(u32_at(data, offset + 4, rt, "invalid cell field")?)?;
                    offset += 8;
                    out.push(WebInvalidCell::new(row, column))
                }
                out
            } else {
                Vec::new()
            };
            ListObjectSourceMetadata::Web(WebTableMetadata {
                version,
                build_number: build,
                fields: web_fields,
                edit_mode,
                cache_position,
                cache_size,
                cache_characters,
                hash_parameters: hash,
                provider_name,
                entry_id,
                deleted_row_ids,
                changed_row_ids,
                invalid_cells,
                needs_commit: flags & 0x100 != 0,
                compressed_cache: flags & 0x2000 != 0,
                ignored_fixed_word,
                ignored_flags: flags & 0xfe80_0001,
            })
        } else {
            ListObjectSourceMetadata::Xml(XmlTableMetadata {
                version,
                build_number: build,
                fields: xml_fields,
                entry_id,
                single_cell: single,
                ignored_fixed_word,
                ignored_flags: flags & 0xfe80_0001,
                ignored_fixed_tail,
            })
        };
        let has_feature12_field = columns
            .iter()
            .any(|column| column.total_formula.is_some() || column.total_string.is_some());
        if rt == FEATURE12_RECORD_TYPE && header != 0 && !has_feature12_field {
            return Err(invalid(
                rt,
                "Feature12 Web/XML source lacks a Feature12-only property",
            ));
        }
        if offset != data.len() {
            return Err(invalid(rt, "trailing Web/XML feature data"));
        }
        Ok(Self {
            id,
            name,
            range,
            columns,
            style: None,
            has_header: header != 0,
            has_totals: totals != 0,
            autofilter: flags & 2 != 0,
            comment: String::new(),
            feature_version: if rt == FEATURE12_RECORD_TYPE {
                ListObjectFeatureVersion::Feature12
            } else {
                ListObjectFeatureVersion::Feature11
            },
            opaque_feature: None,
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: None,
            source_metadata: Some(source_metadata),
        })
    }
    pub(super) fn parse_feature(data: &[u8], rt: u16) -> Result<Self> {
        if !(99..=MAX_FEATURE_BYTES).contains(&data.len()) {
            return Err(invalid(rt, "invalid table feature length"));
        }
        validate_frt(data, rt, true)?;
        let range = parse_range(data, 4, rt)?;
        if u16_at(data, 12, FEATURE11_RECORD_TYPE, "isf")? != ISF_LIST
            || data[14] != 0
            || u32_at(data, 15, FEATURE11_RECORD_TYPE, "reserved2")? != 0
            || u16_at(data, 19, FEATURE11_RECORD_TYPE, "cref2")? != 1
            || u16_at(data, 25, FEATURE11_RECORD_TYPE, "reserved3")? != 0
            || parse_range(data, 27, FEATURE11_RECORD_TYPE)? != range
        {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "invalid Feature11 fixed fields",
            ));
        }
        let base = 35;
        let source_type = u32_at(data, base, FEATURE11_RECORD_TYPE, "lt")?;
        if matches!(source_type, 1 | 2) {
            return Self::parse_source_feature(data, rt, source_type);
        }
        if source_type != 0 {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "unsupported table source type",
            ));
        }
        let id = ListObjectId::try_new(u32_at(data, base + 4, FEATURE11_RECORD_TYPE, "idList")?)?;
        let header = u32_at(data, base + 8, FEATURE11_RECORD_TYPE, "crwHeader")?;
        let totals = u32_at(data, base + 12, FEATURE11_RECORD_TYPE, "crwTotals")?;
        if header > 1
            || totals > 1
            || u32_at(data, base + 20, FEATURE11_RECORD_TYPE, "cbFSData")? != 64
        {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "invalid TableFeatureType fixed fields",
            ));
        }
        let flags = u32_at(data, base + 28, FEATURE11_RECORD_TYPE, "flags")?;
        if flags & 0x0020_E320 != 0 {
            return Err(invalid(
                FEATURE11_RECORD_TYPE,
                "unsupported external table flags",
            ));
        }
        let (name, mut offset) = parse_string(data, base + 64, FEATURE11_RECORD_TYPE, "rgbName")?;
        let count = usize::from(u16_at(data, offset, FEATURE11_RECORD_TYPE, "cFieldData")?);
        offset += 2;
        if !(1..=256).contains(&count) {
            return Err(invalid(FEATURE11_RECORD_TYPE, "invalid table column count"));
        }
        if flags & 0x0010_0000 != 0 {
            let (entry, next) = parse_string(data, offset, FEATURE11_RECORD_TYPE, "entryId")?;
            if entry != id.value().to_string() {
                return Err(invalid(FEATURE11_RECORD_TYPE, "entryId mismatch"));
            }
            offset = next;
        }
        let mut columns = Vec::with_capacity(count);
        for _ in 0..count {
            let start = offset;
            let cid =
                ListColumnId::try_new(u32_at(data, start, FEATURE11_RECORD_TYPE, "idField")?)?;
            if u32_at(data, start + 4, FEATURE11_RECORD_TYPE, "lfdt")? != 0
                || u32_at(data, start + 8, FEATURE11_RECORD_TYPE, "lfxidt")? != 0
            {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "external column data is unsupported",
                ));
            }
            let agg = u32_at(data, start + 16, FEATURE11_RECORD_TYPE, "cbFmtAgg")? as usize;
            let cflags = u32_at(data, start + 24, FEATURE11_RECORD_TYPE, "column flags")?;
            let insert = u32_at(data, start + 28, FEATURE11_RECORD_TYPE, "cbFmtInsert")? as usize;
            let aggregation = ListTotalAggregation::from_code(u32_at(
                data,
                start + 12,
                FEATURE11_RECORD_TYPE,
                "ilta",
            )?)?;
            // Feature11 permits calculated columns (fAutoCreateCalcCol) but not
            // Feature12-only XML/Web mappings or loaded total formulas/strings.
            let forbidden = 0x4c
                | if rt == FEATURE11_RECORD_TYPE {
                    0x580
                } else {
                    0
                };
            if cflags & forbidden != 0 || cflags & 0x100 != 0 {
                return Err(invalid(
                    FEATURE11_RECORD_TYPE,
                    "external, array, or reserved column data is unsupported",
                ));
            }
            let (_, after_field) =
                parse_string(data, start + 36, FEATURE11_RECORD_TYPE, "field name")?;
            let (caption, after_caption) =
                parse_string(data, after_field, FEATURE11_RECORD_TYPE, "caption")?;
            offset = after_caption
                .checked_add(agg + insert)
                .ok_or_else(|| invalid(FEATURE11_RECORD_TYPE, "column size overflow"))?;
            if cflags & 1 != 0 {
                let n = u32_at(data, offset, FEATURE11_RECORD_TYPE, "AutoFilter size")? as usize;
                offset = offset
                    .checked_add(6 + n)
                    .ok_or_else(|| invalid(FEATURE11_RECORD_TYPE, "AutoFilter size overflow"))?;
            }
            let total_formula = if cflags & 0x80 != 0 {
                let n = usize::from(u16_at(data, offset, rt, "total formula length")?);
                if n == 0 {
                    return Err(invalid(rt, "empty total formula"));
                }
                let end = offset + 2 + n;
                let value = data
                    .get(offset + 2..end)
                    .ok_or_else(|| invalid(rt, "truncated total formula"))?
                    .to_vec();
                offset = end;
                Some(value)
            } else {
                None
            };
            let total_string = if cflags & 0x400 != 0 {
                let (value, end) = parse_string(data, offset, rt, "total string")?;
                offset = end;
                Some(value)
            } else {
                None
            };
            if cflags & 0x200 != 0 {
                offset = parse_string(data, offset, FEATURE11_RECORD_TYPE, "cached style")?.1;
            }
            if offset > data.len() {
                return Err(invalid(FEATURE11_RECORD_TYPE, "truncated column data"));
            }
            let column = ListObjectColumn {
                id: cid,
                name: caption,
                aggregation,
                total_formula,
                total_string,
            };
            column.validate_totals()?;
            columns.push(column);
        }
        if offset != data.len() {
            return Err(invalid(FEATURE11_RECORD_TYPE, "trailing Feature11 data"));
        }
        Ok(Self {
            id,
            name,
            range,
            columns,
            style: None,
            has_header: header != 0,
            has_totals: totals != 0,
            autofilter: flags & 2 != 0,
            comment: String::new(),
            feature_version: if rt == FEATURE12_RECORD_TYPE {
                ListObjectFeatureVersion::Feature12
            } else {
                ListObjectFeatureVersion::Feature11
            },
            opaque_feature: None,
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: None,
            source_metadata: None,
        })
    }
    pub(super) fn apply_list12(&mut self, data: &[u8]) -> Result<u16> {
        if data.len() < 18 {
            return Err(invalid(LIST12_RECORD_TYPE, "truncated List12"));
        }
        validate_frt(data, LIST12_RECORD_TYPE, false)?;
        let kind = u16_at(data, 12, LIST12_RECORD_TYPE, "lsd")?;
        if u32_at(data, 14, LIST12_RECORD_TYPE, "idList")? != self.id.value() {
            return Err(invalid(LIST12_RECORD_TYPE, "List12 id mismatch"));
        }
        match kind {
            0 => {
                if data.len() < 54 {
                    return Err(invalid(LIST12_RECORD_TYPE, "truncated block-level List12"));
                }
            },
            1 => {
                let bits = u16_at(data, 18, LIST12_RECORD_TYPE, "style flags")?;
                let (name, end) = parse_string(data, 20, LIST12_RECORD_TYPE, "style name")?;
                if end != data.len() {
                    return Err(invalid(LIST12_RECORD_TYPE, "trailing style List12 data"));
                }
                self.style = Some(ListObjectStyleOptions {
                    name,
                    first: bits & 1 != 0,
                    last: bits & 2 != 0,
                    row_stripes: bits & 4 != 0,
                    column_stripes: bits & 8 != 0,
                    default_style: bits & 0x40 != 0,
                });
            },
            2 => {
                let (name, next) = parse_string(data, 18, LIST12_RECORD_TYPE, "display name")?;
                let (comment, end) = parse_string(data, next, LIST12_RECORD_TYPE, "comment")?;
                if end != data.len() || (!name.is_empty() && name != self.name) {
                    return Err(invalid(LIST12_RECORD_TYPE, "inconsistent display List12"));
                }
                self.comment = comment;
            },
            _ => return Err(invalid(LIST12_RECORD_TYPE, "reserved List12 type")),
        }
        Ok(kind)
    }
}
