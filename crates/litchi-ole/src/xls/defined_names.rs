//! BIFF8 internal defined-name (`Lbl`) parsing and public models.

use super::error::{XlsError, XlsResult};
use super::formula::{FormulaContext, render_formula};

pub(crate) const LBL_RECORD_TYPE: u16 = 0x0018;
pub(crate) const NAME_PUBLISH_RECORD_TYPE: u16 = 0x0893;
pub(crate) const NAME_CMT_RECORD_TYPE: u16 = 0x0894;
pub(crate) const NAME_FN_GRP12_RECORD_TYPE: u16 = 0x0899;
const LBL_HEADER_LEN: usize = 14;
const FLAG_HIDDEN: u16 = 0x0001;
const FLAG_FUNCTION: u16 = 0x0002;
const FLAG_VBA: u16 = 0x0004;
const FLAG_PROCEDURE: u16 = 0x0008;
const FLAG_BUILT_IN: u16 = 0x0020;
const RESERVED_FLAG_MASK: u16 = 0x9000;

/// A non-macro internal defined name from the workbook globals substream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsDefinedName {
    /// One-based position in the complete `Lbl` collection, including macro slots.
    pub record_index: u32,
    pub name: String,
    pub scope: XlsNameScope,
    pub hidden: bool,
    pub function: bool,
    pub vba_procedure: bool,
    pub procedure: bool,
    pub calculated_expression: bool,
    pub function_group: u8,
    pub published: bool,
    pub workbook_parameter: bool,
    pub shortcut_key: Option<u8>,
    pub kind: XlsDefinedNameKind,
    /// Rendered formula using the same leading-`=` convention as cell formulas.
    pub formula: Option<String>,
    /// Original `NameParsedFormula.rgce` bytes.
    pub formula_tokens: Vec<u8>,
    pub formula_extra: Vec<u8>,
    pub continuation_chunks: Vec<Vec<u8>>,
    pub custom_menu: String,
    pub description: String,
    pub help_topic: String,
    pub status_bar: String,
    pub comment: Option<String>,
    pub future_records: XlsDefinedNameFutureRecords,
}

/// Optional BIFF8 future records associated with one immediately preceding `Lbl`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsDefinedNameFutureRecords {
    pub function_group: Option<XlsNameFnGrp12>,
    pub publication: Option<XlsNamePublish>,
}

/// Extended function-category metadata from `NameFnGrp12`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsNameFnGrp12 {
    pub function_name: String,
    /// Raw BIFF category number in the inclusive range 32..=255.
    pub category: u8,
}

impl XlsNameFnGrp12 {
    pub fn category_index(&self) -> usize {
        usize::from(self.category - 32)
    }
}

/// Server publication metadata from `NamePublish`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsNamePublish {
    pub published: bool,
    pub workbook_parameter: bool,
    pub name: String,
}

impl XlsDefinedName {
    pub fn is_macro(&self) -> bool {
        self.function || self.vba_procedure || self.procedure
    }
    /// Whether the rendered definition contains a deleted reference.
    pub fn is_deleted(&self) -> bool {
        self.formula
            .as_deref()
            .is_some_and(|formula| formula.contains("#REF!"))
    }
}

/// Scope of a defined name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsNameScope {
    Workbook,
    Worksheet(usize),
}

/// User-defined or reserved built-in name kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsDefinedNameKind {
    User,
    BuiltIn(XlsBuiltInName),
}

/// Built-in BIFF8 defined-name identifiers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsBuiltInName {
    ConsolidateArea,
    AutoOpen,
    AutoClose,
    Extract,
    Database,
    Criteria,
    PrintArea,
    PrintTitles,
    Recorder,
    DataForm,
    AutoActivate,
    AutoDeactivate,
    SheetTitle,
    FilterDatabase,
}

impl XlsBuiltInName {
    fn from_code(code: u8) -> Option<Self> {
        Some(match code {
            0x00 => Self::ConsolidateArea,
            0x01 => Self::AutoOpen,
            0x02 => Self::AutoClose,
            0x03 => Self::Extract,
            0x04 => Self::Database,
            0x05 => Self::Criteria,
            0x06 => Self::PrintArea,
            0x07 => Self::PrintTitles,
            0x08 => Self::Recorder,
            0x09 => Self::DataForm,
            0x0A => Self::AutoActivate,
            0x0B => Self::AutoDeactivate,
            0x0C => Self::SheetTitle,
            0x0D => Self::FilterDatabase,
            _ => return None,
        })
    }

    pub(crate) fn canonical_name(self) -> &'static str {
        match self {
            Self::ConsolidateArea => "Consolidate_Area",
            Self::AutoOpen => "Auto_Open",
            Self::AutoClose => "Auto_Close",
            Self::Extract => "Extract",
            Self::Database => "Database",
            Self::Criteria => "Criteria",
            Self::PrintArea => "Print_Area",
            Self::PrintTitles => "Print_Titles",
            Self::Recorder => "Recorder",
            Self::DataForm => "Data_Form",
            Self::AutoActivate => "Auto_Activate",
            Self::AutoDeactivate => "Auto_Deactivate",
            Self::SheetTitle => "Sheet_Title",
            Self::FilterDatabase => "_FilterDatabase",
        }
    }

    pub(crate) fn code(self) -> u8 {
        match self {
            Self::ConsolidateArea => 0x00,
            Self::AutoOpen => 0x01,
            Self::AutoClose => 0x02,
            Self::Extract => 0x03,
            Self::Database => 0x04,
            Self::Criteria => 0x05,
            Self::PrintArea => 0x06,
            Self::PrintTitles => 0x07,
            Self::Recorder => 0x08,
            Self::DataForm => 0x09,
            Self::AutoActivate => 0x0a,
            Self::AutoDeactivate => 0x0b,
            Self::SheetTitle => 0x0c,
            Self::FilterDatabase => 0x0d,
        }
    }
}

/// One parsed `Lbl` slot. Macro slots remain here so `PtgName` indices do not shift.
#[derive(Debug, Clone)]
pub(crate) struct DefinedNameSlot {
    record_index: u32,
    name: String,
    itab: u16,
    hidden: bool,
    function: bool,
    vba_procedure: bool,
    procedure: bool,
    calculated_expression: bool,
    function_group: u8,
    published: bool,
    workbook_parameter: bool,
    shortcut_key: Option<u8>,
    kind: XlsDefinedNameKind,
    formula_tokens: Vec<u8>,
    formula_extra: Vec<u8>,
    continuation_chunks: Vec<Vec<u8>>,
    custom_menu: String,
    description: String,
    help_topic: String,
    status_bar: String,
    comment: Option<String>,
    future_records: XlsDefinedNameFutureRecords,
}

impl DefinedNameSlot {
    #[cfg(test)]
    pub(crate) fn parse(data: &[u8], record_index: u32) -> XlsResult<Self> {
        Self::parse_with_continuations(data, record_index, Vec::new())
    }

    pub(crate) fn parse_with_continuations(
        data: &[u8],
        record_index: u32,
        continuation_chunks: Vec<Vec<u8>>,
    ) -> XlsResult<Self> {
        if data.len() < LBL_HEADER_LEN + 1 {
            return invalid("Lbl is missing its fixed header or name flags");
        }
        let flags = u16::from_le_bytes([data[0], data[1]]);
        if flags & RESERVED_FLAG_MASK != 0 {
            return invalid("Lbl contains nonzero reserved option flags");
        }
        let function = flags & FLAG_FUNCTION != 0;
        let vba = flags & FLAG_VBA != 0;
        let procedure = flags & FLAG_PROCEDURE != 0;
        if (function || vba) && !procedure {
            return invalid("Lbl function/VBA flags require the procedure flag");
        }
        let function_group = (flags >> 6) & 0x3f;
        if function_group > 31 {
            return invalid("Lbl function group must be at most 31");
        }
        let shortcut = data[2];
        if (function || !procedure) && shortcut != 0 {
            return invalid("Lbl shortcut key is not valid for these macro flags");
        }
        if procedure && !function && shortcut != 0 && !shortcut.is_ascii_alphabetic() {
            return invalid("Lbl macro shortcut key must be an ASCII letter");
        }

        let character_count = usize::from(data[3]);
        let formula_len = usize::from(u16::from_le_bytes([data[4], data[5]]));
        if data[6] != 0 || data[7] != 0 {
            return invalid("Lbl reserved3 field must be zero");
        }
        let itab = u16::from_le_bytes([data[8], data[9]]);
        let auxiliary_lengths = [data[10], data[11], data[12], data[13]].map(usize::from);
        let string_flags = data[LBL_HEADER_LEN];
        if string_flags & !0x01 != 0 {
            return invalid("Lbl name contains unsupported Unicode flags");
        }
        let char_width = if string_flags & 0x01 == 0 { 1 } else { 2 };
        let name_byte_len = character_count
            .checked_mul(char_width)
            .ok_or_else(|| invalid_error("Lbl name length overflows"))?;
        let name_start = LBL_HEADER_LEN + 1;
        let name_end = name_start
            .checked_add(name_byte_len)
            .ok_or_else(|| invalid_error("Lbl name range overflows"))?;
        let name_bytes = data
            .get(name_start..name_end)
            .ok_or_else(|| invalid_error("Lbl name is truncated"))?;
        let decoded_name = if char_width == 1 {
            name_bytes.iter().map(|byte| char::from(*byte)).collect()
        } else {
            let units: Vec<u16> = name_bytes
                .chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                .collect();
            String::from_utf16(&units)
                .map_err(|_| invalid_error("Lbl name contains invalid UTF-16"))?
        };

        let built_in = flags & FLAG_BUILT_IN != 0;
        let (name, kind) = if built_in {
            if character_count != 1 {
                return invalid("built-in Lbl name must contain exactly one character");
            }
            let code = decoded_name
                .encode_utf16()
                .next()
                .and_then(|value| u8::try_from(value).ok())
                .ok_or_else(|| invalid_error("built-in Lbl identifier is invalid"))?;
            let built_in = XlsBuiltInName::from_code(code)
                .ok_or_else(|| invalid_error("built-in Lbl identifier is out of range"))?;
            (
                built_in.canonical_name().to_string(),
                XlsDefinedNameKind::BuiltIn(built_in),
            )
        } else {
            if decoded_name.is_empty() || decoded_name.contains('\0') {
                return invalid("user-defined Lbl name must be nonempty and contain no NUL");
            }
            (decoded_name, XlsDefinedNameKind::User)
        };

        let formula_end = name_end
            .checked_add(formula_len)
            .ok_or_else(|| invalid_error("Lbl formula range overflows"))?;
        let formula_tokens = data
            .get(name_end..formula_end)
            .ok_or_else(|| invalid_error("Lbl formula is truncated"))?
            .to_vec();
        let auxiliary_len = auxiliary_lengths.iter().sum::<usize>();
        let extra_end = data
            .len()
            .checked_sub(auxiliary_len)
            .ok_or_else(|| invalid_error("Lbl auxiliary strings exceed payload"))?;
        if extra_end < formula_end {
            return invalid("Lbl formula or auxiliary strings are truncated");
        }
        let formula_extra = data[formula_end..extra_end].to_vec();
        let mut auxiliary_offset = extra_end;
        let mut auxiliary_strings = Vec::with_capacity(4);
        for length in auxiliary_lengths {
            let end = auxiliary_offset
                .checked_add(length)
                .ok_or_else(|| invalid_error("Lbl auxiliary string range overflows"))?;
            let bytes = data
                .get(auxiliary_offset..end)
                .ok_or_else(|| invalid_error("Lbl auxiliary string is truncated"))?;
            auxiliary_strings.push(
                bytes
                    .iter()
                    .map(|byte| char::from(*byte))
                    .collect::<String>(),
            );
            auxiliary_offset = end;
        }
        if auxiliary_offset != data.len() {
            return invalid("Lbl payload has unconsumed bytes");
        }

        Ok(Self {
            record_index,
            name,
            itab,
            hidden: flags & FLAG_HIDDEN != 0,
            function,
            vba_procedure: vba,
            procedure,
            calculated_expression: flags & 0x0010 != 0,
            function_group: function_group as u8,
            published: flags & 0x2000 != 0,
            workbook_parameter: flags & 0x4000 != 0,
            shortcut_key: (shortcut != 0).then_some(shortcut),
            kind,
            formula_tokens,
            formula_extra,
            continuation_chunks,
            custom_menu: auxiliary_strings.remove(0),
            description: auxiliary_strings.remove(0),
            help_topic: auxiliary_strings.remove(0),
            status_bar: auxiliary_strings.remove(0),
            comment: None,
            future_records: XlsDefinedNameFutureRecords::default(),
        })
    }

    pub(crate) fn attach_comment(&mut self, data: &[u8]) -> XlsResult<()> {
        if self.comment.is_some() {
            return invalid("Lbl has duplicate NameCmt records");
        }
        let (name, comment) = parse_name_comment(data)?;
        if !unicode_name_eq(&name, &self.name) {
            return optional_invalid(
                NAME_CMT_RECORD_TYPE,
                "NameCmt name does not match preceding Lbl",
            );
        }
        self.comment = Some(comment);
        Ok(())
    }

    pub(crate) fn attach_function_group(&mut self, data: &[u8]) -> XlsResult<()> {
        if self.future_records.function_group.is_some() {
            return optional_invalid(
                NAME_FN_GRP12_RECORD_TYPE,
                "Lbl has duplicate NameFnGrp12 records",
            );
        }
        let value = parse_name_fn_grp12(data)?;
        if !unicode_name_eq(&value.function_name, &self.name) {
            return optional_invalid(
                NAME_FN_GRP12_RECORD_TYPE,
                "NameFnGrp12 name does not match preceding Lbl",
            );
        }
        self.future_records.function_group = Some(value);
        Ok(())
    }

    pub(crate) fn attach_publication(&mut self, data: &[u8]) -> XlsResult<()> {
        if self.future_records.publication.is_some() {
            return optional_invalid(
                NAME_PUBLISH_RECORD_TYPE,
                "Lbl has duplicate NamePublish records",
            );
        }
        let value = parse_name_publish(data)?;
        if !unicode_name_eq(&value.name, &self.name) {
            return optional_invalid(
                NAME_PUBLISH_RECORD_TYPE,
                "NamePublish name does not match preceding Lbl",
            );
        }
        self.future_records.publication = Some(value);
        Ok(())
    }

    pub(crate) fn validate_extended_category(&self, count: usize) -> XlsResult<()> {
        if self
            .future_records
            .function_group
            .as_ref()
            .is_some_and(|value| value.category_index() >= count)
        {
            return optional_invalid(
                NAME_FN_GRP12_RECORD_TYPE,
                "NameFnGrp12 category does not reference an FnGrp12 record",
            );
        }
        Ok(())
    }

    /// Symbol table entry. Macro slots intentionally remain `None`.
    pub(crate) fn symbol(&self) -> Option<String> {
        (!(self.function || self.vba_procedure || self.procedure)).then(|| self.name.clone())
    }

    pub(crate) fn into_public(
        self,
        sheet_count: usize,
        context: &FormulaContext,
    ) -> XlsResult<XlsDefinedName> {
        let scope = if self.itab == 0 {
            XlsNameScope::Workbook
        } else {
            let sheet_index = usize::from(self.itab - 1);
            if sheet_index >= sheet_count {
                return invalid("Lbl.itab is outside the BoundSheet8 collection");
            }
            XlsNameScope::Worksheet(sheet_index)
        };
        let formula = (!self.formula_tokens.is_empty())
            .then(|| render_formula(&self.formula_tokens, Some(context)))
            .flatten();
        Ok(XlsDefinedName {
            record_index: self.record_index,
            name: self.name,
            scope,
            hidden: self.hidden,
            function: self.function,
            vba_procedure: self.vba_procedure,
            procedure: self.procedure,
            calculated_expression: self.calculated_expression,
            function_group: self.function_group,
            published: self.published,
            workbook_parameter: self.workbook_parameter,
            shortcut_key: self.shortcut_key,
            kind: self.kind,
            formula,
            formula_tokens: self.formula_tokens,
            formula_extra: self.formula_extra,
            continuation_chunks: self.continuation_chunks,
            custom_menu: self.custom_menu,
            description: self.description,
            help_topic: self.help_topic,
            status_bar: self.status_bar,
            comment: self.comment,
            future_records: self.future_records,
        })
    }
}

fn parse_name_comment(data: &[u8]) -> XlsResult<(String, String)> {
    if data.len() < 18 {
        return optional_invalid(NAME_CMT_RECORD_TYPE, "NameCmt is truncated");
    }
    if u16::from_le_bytes([data[0], data[1]]) != NAME_CMT_RECORD_TYPE
        || data[2..12].iter().any(|byte| *byte != 0)
    {
        return optional_invalid(
            NAME_CMT_RECORD_TYPE,
            "NameCmt future-record header is invalid",
        );
    }
    let name_len = usize::from(u16::from_le_bytes([data[12], data[13]]));
    let comment_len = usize::from(u16::from_le_bytes([data[14], data[15]]));
    if name_len > 255 || comment_len > 255 {
        return optional_invalid(
            NAME_CMT_RECORD_TYPE,
            "NameCmt strings exceed 255 characters",
        );
    }
    let (name, offset) = parse_no_cch_string(data, 16, name_len)?;
    let (comment, offset) = parse_no_cch_string(data, offset, comment_len)?;
    if offset != data.len() {
        return optional_invalid(
            NAME_CMT_RECORD_TYPE,
            "NameCmt strings do not consume payload",
        );
    }
    Ok((name, comment))
}

fn parse_name_fn_grp12(data: &[u8]) -> XlsResult<XlsNameFnGrp12> {
    validate_frt_header(data, NAME_FN_GRP12_RECORD_TYPE)?;
    if data.len() < 20 {
        return optional_invalid(NAME_FN_GRP12_RECORD_TYPE, "NameFnGrp12 is truncated");
    }
    let cached = usize::from(u16::from_le_bytes([data[12], data[13]]));
    if !(1..=255).contains(&cached) {
        return optional_invalid(
            NAME_FN_GRP12_RECORD_TYPE,
            "NameFnGrp12 cached name length is outside 1..=255",
        );
    }
    let category = u16::from_le_bytes([data[14], data[15]]);
    if !(32..=255).contains(&category) {
        return optional_invalid(
            NAME_FN_GRP12_RECORD_TYPE,
            "NameFnGrp12 category is outside 32..=255",
        );
    }
    let (name, count, end) = parse_xl_name_unicode(data, 16, NAME_FN_GRP12_RECORD_TYPE)?;
    if count != cached {
        return optional_invalid(
            NAME_FN_GRP12_RECORD_TYPE,
            "NameFnGrp12 cached and embedded name lengths differ",
        );
    }
    if end != data.len() {
        return optional_invalid(
            NAME_FN_GRP12_RECORD_TYPE,
            "NameFnGrp12 string does not consume payload",
        );
    }
    Ok(XlsNameFnGrp12 {
        function_name: name,
        category: category as u8,
    })
}

fn parse_name_publish(data: &[u8]) -> XlsResult<XlsNamePublish> {
    validate_frt_header(data, NAME_PUBLISH_RECORD_TYPE)?;
    if data.len() < 16 {
        return optional_invalid(NAME_PUBLISH_RECORD_TYPE, "NamePublish is truncated");
    }
    let flags = u16::from_le_bytes([data[12], data[13]]);
    let (name, _, end) = parse_xl_name_unicode(data, 14, NAME_PUBLISH_RECORD_TYPE)?;
    if end != data.len() {
        return optional_invalid(
            NAME_PUBLISH_RECORD_TYPE,
            "NamePublish string does not consume payload",
        );
    }
    Ok(XlsNamePublish {
        published: flags & 1 != 0,
        workbook_parameter: flags & 2 != 0,
        name,
    })
}

fn validate_frt_header(data: &[u8], record_type: u16) -> XlsResult<()> {
    if data.len() < 12 {
        return optional_invalid(record_type, "future-record header is truncated");
    }
    if u16::from_le_bytes([data[0], data[1]]) != record_type
        || data[2..12].iter().any(|byte| *byte != 0)
    {
        return optional_invalid(
            record_type,
            "future-record header flags or reserved fields are invalid",
        );
    }
    Ok(())
}

fn parse_xl_name_unicode(
    data: &[u8],
    offset: usize,
    record_type: u16,
) -> XlsResult<(String, usize, usize)> {
    let header = data.get(offset..offset + 3).ok_or_else(|| {
        optional_invalid_error(record_type, "XLNameUnicodeString header is truncated")
    })?;
    let count = usize::from(u16::from_le_bytes([header[0], header[1]]));
    if !(1..=255).contains(&count) {
        return optional_invalid(record_type, "XLNameUnicodeString length is outside 1..=255");
    }
    if header[2] & !1 != 0 {
        return optional_invalid(record_type, "XLNameUnicodeString flags are invalid");
    }
    let width = if header[2] == 0 { 1usize } else { 2 };
    let start = offset + 3;
    let end = start
        .checked_add(count.checked_mul(width).ok_or_else(|| {
            optional_invalid_error(record_type, "XLNameUnicodeString size overflows")
        })?)
        .ok_or_else(|| optional_invalid_error(record_type, "XLNameUnicodeString end overflows"))?;
    let bytes = data
        .get(start..end)
        .ok_or_else(|| optional_invalid_error(record_type, "XLNameUnicodeString is truncated"))?;
    let value = if width == 1 {
        bytes.iter().map(|byte| char::from(*byte)).collect()
    } else {
        let units = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units).map_err(|_| {
            optional_invalid_error(record_type, "XLNameUnicodeString contains invalid UTF-16")
        })?
    };
    if value.contains('\0') {
        return optional_invalid(record_type, "XLNameUnicodeString contains NUL");
    }
    Ok((value, count, end))
}

pub(crate) fn unicode_name_eq(left: &str, right: &str) -> bool {
    left.to_lowercase() == right.to_lowercase()
}

fn parse_no_cch_string(data: &[u8], offset: usize, count: usize) -> XlsResult<(String, usize)> {
    let flags = *data.get(offset).ok_or_else(|| {
        optional_invalid_error(NAME_CMT_RECORD_TYPE, "NameCmt string flags are missing")
    })?;
    if flags & !1 != 0 {
        return optional_invalid(NAME_CMT_RECORD_TYPE, "NameCmt string flags are invalid");
    }
    let width = if flags == 0 { 1usize } else { 2 };
    let start = offset + 1;
    let end = start
        .checked_add(count.checked_mul(width).ok_or_else(|| {
            optional_invalid_error(NAME_CMT_RECORD_TYPE, "NameCmt string size overflows")
        })?)
        .ok_or_else(|| {
            optional_invalid_error(NAME_CMT_RECORD_TYPE, "NameCmt string end overflows")
        })?;
    let bytes = data.get(start..end).ok_or_else(|| {
        optional_invalid_error(NAME_CMT_RECORD_TYPE, "NameCmt string is truncated")
    })?;
    let value = if width == 1 {
        bytes.iter().map(|byte| char::from(*byte)).collect()
    } else {
        let units = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units).map_err(|_| {
            optional_invalid_error(NAME_CMT_RECORD_TYPE, "NameCmt contains invalid UTF-16")
        })?
    };
    Ok((value, end))
}

fn invalid<T>(message: &str) -> XlsResult<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: &str) -> XlsError {
    XlsError::InvalidRecord {
        record_type: LBL_RECORD_TYPE,
        message: message.to_string(),
    }
}

fn optional_invalid<T>(record_type: u16, message: &str) -> XlsResult<T> {
    Err(optional_invalid_error(record_type, message))
}
fn optional_invalid_error(record_type: u16, message: &str) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn lbl(flags: u16, itab: u16, name: &str, unicode: bool, formula: &[u8]) -> Vec<u8> {
        let units: Vec<u16> = name.encode_utf16().collect();
        let cch = if unicode { units.len() } else { name.len() };
        let mut data = Vec::new();
        data.extend_from_slice(&flags.to_le_bytes());
        data.push(0);
        data.push(cch as u8);
        data.extend_from_slice(&(formula.len() as u16).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&itab.to_le_bytes());
        data.extend_from_slice(&[0; 4]);
        data.push(u8::from(unicode));
        if unicode {
            for unit in units {
                data.extend_from_slice(&unit.to_le_bytes());
            }
        } else {
            data.extend_from_slice(name.as_bytes());
        }
        data.extend_from_slice(formula);
        data
    }

    #[test]
    fn parses_compressed_unicode_hidden_and_scoped_names() {
        let ordinary =
            DefinedNameSlot::parse(&lbl(FLAG_HIDDEN, 2, "Rate", false, &[0x1e, 2, 0]), 1).unwrap();
        assert_eq!(ordinary.name, "Rate");
        assert!(ordinary.hidden);
        assert_eq!(ordinary.itab, 2);
        assert_eq!(ordinary.formula_tokens, [0x1e, 2, 0]);

        let unicode = DefinedNameSlot::parse(&lbl(0, 0, "税率", true, &[0x1e, 3, 0]), 2).unwrap();
        assert_eq!(unicode.name, "税率");
    }

    #[test]
    fn parses_every_built_in_name() {
        for code in 0u8..=0x0d {
            let slot = DefinedNameSlot::parse(
                &lbl(
                    FLAG_BUILT_IN,
                    1,
                    &char::from(code).to_string(),
                    false,
                    &[0x1e, 1, 0],
                ),
                u32::from(code) + 1,
            )
            .unwrap();
            assert!(matches!(slot.kind, XlsDefinedNameKind::BuiltIn(_)));
        }
    }

    #[test]
    fn macro_slots_keep_indices_but_have_no_symbol_or_public_value() {
        let slot = DefinedNameSlot::parse(&lbl(FLAG_PROCEDURE, 0, "Macro", false, &[]), 7).unwrap();
        assert!(slot.symbol().is_none());
        assert!(
            slot.into_public(1, &FormulaContext::default())
                .unwrap()
                .is_macro()
        );
    }

    #[test]
    fn rejects_malformed_names_and_scope() {
        assert!(DefinedNameSlot::parse(&[0; 14], 1).is_err());
        let mut truncated = lbl(0, 0, "Name", true, &[0x1e, 1, 0]);
        truncated.truncate(17);
        assert!(DefinedNameSlot::parse(&truncated, 1).is_err());
        assert!(DefinedNameSlot::parse(&lbl(FLAG_BUILT_IN, 0, "x", false, &[]), 1).is_err());
        let invalid_scope = DefinedNameSlot::parse(&lbl(0, 3, "Name", false, &[]), 1).unwrap();
        assert!(
            invalid_scope
                .into_public(2, &FormulaContext::default())
                .is_err()
        );
    }

    #[test]
    fn name_comment_requires_exact_header_and_matching_preceding_name() {
        let mut slot = DefinedNameSlot::parse(&lbl(0, 0, "Rate", false, &[]), 1).unwrap();
        let mut comment = Vec::new();
        comment.extend_from_slice(&NAME_CMT_RECORD_TYPE.to_le_bytes());
        comment.extend_from_slice(&[0; 10]);
        comment.extend_from_slice(&5u16.to_le_bytes());
        comment.extend_from_slice(&1u16.to_le_bytes());
        comment.extend_from_slice(&[0, b'O', b't', b'h', b'e', b'r', 0, b'X']);
        assert!(slot.attach_comment(&comment).is_err());
        comment[16] = 2;
        assert!(slot.attach_comment(&comment).is_err());
    }
}
