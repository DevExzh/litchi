//! Classic text-import metadata (MS-XLSB 2.4.65/67/69, 2.5.40/41).
use super::model::{TextCodePage, TextField, TextFieldType, TextImportProperties, TextQualifier};
use crate::package::error::Result;
use crate::package::walker::{RecordWalker, malformed};
use crate::raw::{Cursor, kind as rt};

pub(super) const MAX_FIELDS: usize = 16_384;
const IGNORED_FLAGS: u32 = 0xff90_0000;
const CONTEXT: &str = "BrtBeginECTxtWiz";

pub(super) fn validate(value: &TextImportProperties) -> Result<()> {
    if value.application_code_page > 1023 || value.ignored_flags & !IGNORED_FLAGS != 0 {
        return Err(malformed(
            CONTEXT,
            "text import flags are outside their wire fields",
        ));
    }
    if value.first_row == 0 || value.first_row >= 100_000_000 {
        return Err(malformed(
            CONTEXT,
            "text import first row must be in 1..100000000",
        ));
    }
    if value.source_file.encode_utf16().count() >= 219
        || (value.source_file.is_empty() && !value.prompt_for_file)
    {
        return Err(malformed(
            CONTEXT,
            "invalid text import source path or prompt flag",
        ));
    }
    if value
        .fields
        .as_ref()
        .is_some_and(|fields| fields.len() > MAX_FIELDS)
    {
        return Err(malformed(CONTEXT, "text import field count limit exceeded"));
    }
    Ok(())
}

pub(super) fn parse(
    walker: &mut RecordWalker<'_, '_>,
    payload: &[u8],
) -> Result<TextImportProperties> {
    let mut cursor = Cursor::new(payload, CONTEXT, walker.limits());
    let flags = cursor.read_u32()?;
    let code_page = match flags & 3 {
        0 => TextCodePage::Macintosh,
        1 => TextCodePage::WindowsAnsi,
        2 => TextCodePage::Dos,
        _ => return Err(malformed(CONTEXT, "invalid text character-set kind")),
    };
    let qualifier = match (flags >> 18) & 3 {
        0 => TextQualifier::DoubleQuote,
        1 => TextQualifier::SingleQuote,
        2 => TextQualifier::None,
        _ => TextQualifier::NoneAlternate,
    };
    let mut value = TextImportProperties {
        code_page,
        application_code_page: ((flags >> 2) & 1023) as u16,
        delimited: flags & (1 << 12) != 0,
        tab: flags & (1 << 13) != 0,
        space: flags & (1 << 14) != 0,
        comma: flags & (1 << 15) != 0,
        semicolon: flags & (1 << 16) != 0,
        consecutive: flags & (1 << 17) != 0,
        qualifier,
        prompt_for_file: flags & (1 << 21) != 0,
        custom_delimiter: flags & (1 << 22) != 0,
        delimiter: cursor.read_u16()?,
        first_row: cursor.read_u32()?,
        decimal_separator: cursor.read_u8()?,
        thousands_separator: cursor.read_u8()?,
        source_file: cursor.read_wide_string()?,
        fields: None,
        ignored_flags: flags & IGNORED_FLAGS,
    };
    walker.finish(&cursor)?;
    loop {
        let record = walker.required(CONTEXT)?;
        match record.kind() {
            rt::BEGIN_EC_TW_FLD_INFO_LIST => {
                if value.fields.is_some() {
                    return Err(malformed(CONTEXT, "duplicate text field lists"));
                }
                let mut cursor = Cursor::new(record.payload(), CONTEXT, walker.limits());
                let count = cursor.read_u32()? as usize;
                walker.finish(&cursor)?;
                if count > MAX_FIELDS {
                    return Err(malformed(CONTEXT, "text import field count limit exceeded"));
                }
                let mut fields = Vec::new();
                loop {
                    let record = walker.required(CONTEXT)?;
                    match record.kind() {
                        rt::BEGIN_EC_TW_FLD_INFO => {
                            if fields.len() >= count {
                                return Err(malformed(
                                    CONTEXT,
                                    "text field count exceeds declaration",
                                ));
                            }
                            let mut cursor = Cursor::new(
                                record.payload(),
                                "BrtBeginECTwFldInfo",
                                walker.limits(),
                            );
                            let field_type = match cursor.read_u32()? {
                                0 => TextFieldType::General,
                                1 => TextFieldType::Text,
                                2 => TextFieldType::Mdy,
                                3 => TextFieldType::Dmy,
                                4 => TextFieldType::Ymd,
                                5 => TextFieldType::Myd,
                                6 => TextFieldType::Dym,
                                7 => TextFieldType::Ydm,
                                8 => TextFieldType::Skip,
                                9 => TextFieldType::Emd,
                                _ => return Err(malformed(CONTEXT, "invalid text field type")),
                            };
                            fields.push(TextField {
                                field_type,
                                start: cursor.read_u32()?,
                            });
                            walker.finish(&cursor)?;
                        },
                        rt::END_EC_TW_FLD_INFO_LIST => {
                            let cursor = Cursor::new(
                                record.payload(),
                                "BrtEndECTWFldInfoLst",
                                walker.limits(),
                            );
                            walker.finish(&cursor)?;
                            if fields.len() != count {
                                return Err(malformed(
                                    CONTEXT,
                                    "text field count differs from declaration",
                                ));
                            }
                            break;
                        },
                        rt::END_EC_TXT_WIZ | rt::END_EXT_CONNECTION | rt::END_EXT_CONNECTIONS => {
                            return Err(malformed(CONTEXT, "text field list is not closed"));
                        },
                        other => walker.skip_unhandled(other, CONTEXT)?,
                    }
                }
                value.fields = Some(fields);
            },
            rt::END_EC_TXT_WIZ => {
                let cursor = Cursor::new(record.payload(), "BrtEndECTxtWiz", walker.limits());
                walker.finish(&cursor)?;
                validate(&value)?;
                return Ok(value);
            },
            rt::BEGIN_EC_TXT_WIZ
            | rt::BEGIN_EXT_CONNECTION
            | rt::BEGIN_EXT_CONNECTIONS
            | rt::BEGIN_EC_TW_FLD_INFO
            | rt::END_EC_TW_FLD_INFO_LIST
            | rt::END_EXT_CONNECTION
            | rt::END_EXT_CONNECTIONS => {
                return Err(malformed(
                    CONTEXT,
                    "unexpected text import collection boundary",
                ));
            },
            other => walker.skip_unhandled(other, CONTEXT)?,
        }
    }
}

pub(super) fn payload(value: &TextImportProperties) -> Result<Vec<u8>> {
    validate(value)?;
    let flags = u32::from(value.code_page as u8)
        | (u32::from(value.application_code_page) << 2)
        | (u32::from(value.delimited) << 12)
        | (u32::from(value.tab) << 13)
        | (u32::from(value.space) << 14)
        | (u32::from(value.comma) << 15)
        | (u32::from(value.semicolon) << 16)
        | (u32::from(value.consecutive) << 17)
        | (u32::from(value.qualifier as u8) << 18)
        | value.ignored_flags
        | (u32::from(value.prompt_for_file) << 21)
        | (u32::from(value.custom_delimiter) << 22);
    let mut output = flags.to_le_bytes().to_vec();
    output.extend_from_slice(&value.delimiter.to_le_bytes());
    output.extend_from_slice(&value.first_row.to_le_bytes());
    output.push(value.decimal_separator);
    output.push(value.thousands_separator);
    let units = value.source_file.encode_utf16().collect::<Vec<_>>();
    output.extend_from_slice(
        &u32::try_from(units.len())
            .map_err(|_error| malformed(CONTEXT, "text path length overflow"))?
            .to_le_bytes(),
    );
    for unit in units {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(output)
}
