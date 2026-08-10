//! Binary Ptg stream parser.
//!
//! This layer decodes token bytes and ancillary `RgbExtra` payloads into the
//! typed formula model while keeping host relationship resolution external.
use super::super::model::read_u32_le_at;
use super::super::model::*;
use super::super::{Error, Result};
use super::validation::{
    add_wrapped_offset, builtin_function_by_index, is_formula_error_code, parse_table_row_type,
    validate_xnum,
};
use super::wire::{read_f64_le_at, read_u16_le_at};

/// Formula parser
///
/// Parses binary formula bytes into a sequence of tokens.
pub struct Parser<'a> {
    data: &'a [u8],
    offset: usize,
    extra: &'a [u8],
    extra_offset: usize,
    validate_extra: bool,
    memory_expression_ends: Vec<usize>,
    control_flow_targets: Vec<usize>,
    base_cell: Option<(u32, u32)>,
}

impl<'a> Parser<'a> {
    /// Create a new formula parser
    pub fn new(data: &'a [u8]) -> Self {
        Parser {
            data,
            offset: 0,
            extra: &[],
            extra_offset: 0,
            validate_extra: false,
            memory_expression_ends: Vec::new(),
            control_flow_targets: Vec::new(),
            base_cell: None,
        }
    }

    /// Parse a formula together with its corresponding `RgbExtra` stream.
    pub fn with_extra(data: &'a [u8], extra: &'a [u8]) -> Self {
        Parser {
            data,
            offset: 0,
            extra,
            extra_offset: 0,
            validate_extra: true,
            memory_expression_ends: Vec::new(),
            control_flow_targets: Vec::new(),
            base_cell: None,
        }
    }

    /// Parse a shared-formula definition relative to a concrete target cell.
    pub fn with_base_cell(data: &'a [u8], row: u32, col: u32) -> Self {
        Parser {
            data,
            offset: 0,
            extra: &[],
            extra_offset: 0,
            validate_extra: false,
            memory_expression_ends: Vec::new(),
            control_flow_targets: Vec::new(),
            base_cell: Some((row, col)),
        }
    }

    /// Parse a shared formula and ancillary stream at a concrete target cell.
    pub fn with_base_cell_and_extra(data: &'a [u8], extra: &'a [u8], row: u32, col: u32) -> Self {
        Parser {
            data,
            offset: 0,
            extra,
            extra_offset: 0,
            validate_extra: true,
            memory_expression_ends: Vec::new(),
            control_flow_targets: Vec::new(),
            base_cell: Some((row, col)),
        }
    }

    /// Parse the formula into tokens
    ///
    /// Returns a vector of formula tokens in RPN order.
    pub fn parse(&mut self) -> Result<Vec<Token>> {
        Ok(self
            .parse_spanned()?
            .into_iter()
            .map(|(token, _span)| token)
            .collect())
    }

    /// Parse tokens together with their exact `Rgce` byte ranges.
    ///
    /// This remains crate-private because ranges are a checked rewrite seam,
    /// not a stable public representation of every future Ptg family.
    #[deny(
        clippy::cast_lossless,
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        clippy::cast_precision_loss,
        clippy::cast_sign_loss,
        clippy::checked_conversions,
        clippy::expect_used,
        clippy::let_underscore_must_use,
        clippy::unnecessary_unwrap,
        reason = "formula dependency spans use parser-proven offsets without unchecked conversions"
    )]
    pub(crate) fn parse_spanned(&mut self) -> Result<Vec<(Token, std::ops::Range<usize>)>> {
        let mut tokens = Vec::new();
        let mut boundaries = Vec::new();

        while self.offset < self.data.len() {
            let start = self.offset;
            let token = self.parse_token()?;
            tokens.push((token, start..self.offset));
            boundaries.push(self.offset);
        }

        if let Some(end) = self
            .memory_expression_ends
            .iter()
            .find(|end| boundaries.binary_search(end).is_err())
        {
            return Err(Error::InvalidFormula(format!(
                "memory expression ends at byte {end}, which is not a token boundary"
            )));
        }

        if let Some(target) = self
            .control_flow_targets
            .iter()
            .find(|target| boundaries.binary_search(target).is_err())
        {
            return Err(Error::InvalidFormula(format!(
                "control-flow target byte {target} is not a token boundary"
            )));
        }

        if self.validate_extra && self.extra_offset != self.extra.len() {
            return Err(Error::InvalidFormula(format!(
                "formula has {} unconsumed ancillary bytes",
                self.extra.len() - self.extra_offset
            )));
        }

        Ok(tokens)
    }

    /// Parse a single token
    fn parse_token(&mut self) -> Result<Token> {
        if self.offset >= self.data.len() {
            return Err(Error::InvalidFormula(
                "unexpected end of formula token stream".to_string(),
            ));
        }

        let ptg_type = self.data[self.offset];
        self.offset += 1;

        use ptg_types::*;

        match ptg_type {
            PTG_ADD => Ok(Token::BinaryOp(BinaryOperator::Add)),
            PTG_SUB => Ok(Token::BinaryOp(BinaryOperator::Subtract)),
            PTG_MUL => Ok(Token::BinaryOp(BinaryOperator::Multiply)),
            PTG_DIV => Ok(Token::BinaryOp(BinaryOperator::Divide)),
            PTG_POWER => Ok(Token::BinaryOp(BinaryOperator::Power)),
            PTG_CONCAT => Ok(Token::BinaryOp(BinaryOperator::Concat)),
            PTG_LT => Ok(Token::BinaryOp(BinaryOperator::LessThan)),
            PTG_LE => Ok(Token::BinaryOp(BinaryOperator::LessEqual)),
            PTG_EQ => Ok(Token::BinaryOp(BinaryOperator::Equal)),
            PTG_GE => Ok(Token::BinaryOp(BinaryOperator::GreaterEqual)),
            PTG_GT => Ok(Token::BinaryOp(BinaryOperator::GreaterThan)),
            PTG_NE => Ok(Token::BinaryOp(BinaryOperator::NotEqual)),
            PTG_ISECT => Ok(Token::BinaryOp(BinaryOperator::Intersection)),
            PTG_UNION => Ok(Token::BinaryOp(BinaryOperator::Union)),
            PTG_RANGE => Ok(Token::BinaryOp(BinaryOperator::Range)),

            PTG_UPLUS => Ok(Token::UnaryOp(UnaryOperator::Plus)),
            PTG_UMINUS => Ok(Token::UnaryOp(UnaryOperator::Minus)),
            PTG_PERCENT => Ok(Token::UnaryOp(UnaryOperator::Percent)),
            PTG_PAREN => Ok(Token::Parenthesis),
            PTG_MISSING_ARG => Ok(Token::MissingArg),
            PTG_EXTENDED => self.parse_extended(),
            PTG_ATTR => self.parse_attr(),

            PTG_INT => self.parse_int(),
            PTG_NUM => self.parse_num(),
            PTG_STR => self.parse_str(),
            PTG_BOOL => self.parse_bool(),
            PTG_ERR => self.parse_err(),

            _ if ptg_type >= 0x20 => match ptg_type & 0x1F {
                0x04 => self.parse_ref(false),
                0x0C => self.parse_ref(true),
                0x05 => self.parse_area(false),
                0x0D => self.parse_area(true),
                0x0A => self.parse_reference_error(false, false),
                0x0B => self.parse_reference_error(true, false),
                0x1C => self.parse_reference_error(false, true),
                0x1D => self.parse_reference_error(true, true),
                0x1A => self.parse_ref_3d(),
                0x1B => self.parse_area_3d(),
                0x00 => self.parse_array(),
                0x06 => self.parse_memory(MemoryKind::Area),
                0x07 => self.parse_memory(MemoryKind::Error(0)),
                0x08 => self.parse_memory(MemoryKind::NoMemory),
                0x09 => self.parse_memory(MemoryKind::Function),
                0x01 => self.parse_func(),
                0x02 => self.parse_func_var(),
                0x03 => self.parse_name(),
                0x19 => self.parse_name_x(),
                _ => Ok(Token::Unknown(ptg_type)),
            },

            _ => {
                // Unknown token type
                Ok(Token::Unknown(ptg_type))
            },
        }
    }

    fn require(&self, len: usize, context: &str) -> Result<()> {
        if self.offset + len <= self.data.len() {
            Ok(())
        } else {
            Err(Error::InvalidFormula(format!(
                "truncated {context} token at byte {}: need {len} bytes, have {}",
                self.offset.saturating_sub(1),
                self.data.len().saturating_sub(self.offset)
            )))
        }
    }

    fn validate_classed_token(&self, context: &str) -> Result<()> {
        let token = self.data[self.offset - 1];
        if token & 0x80 != 0 {
            return Err(Error::InvalidFormula(format!(
                "{context} token 0x{token:02X} has its reserved bit set"
            )));
        }
        Ok(())
    }

    /// Parse integer constant
    fn parse_int(&mut self) -> Result<Token> {
        self.require(2, "PtgInt")?;

        let value = read_u16_le_at(self.data, self.offset)?;
        self.offset += 2;

        Ok(Token::Int(value))
    }

    /// Parse floating point constant
    fn parse_num(&mut self) -> Result<Token> {
        self.require(8, "PtgNum")?;

        let value = read_f64_le_at(self.data, self.offset)?;
        self.offset += 8;
        validate_xnum(value, "PtgNum")?;

        Ok(Token::Number(value))
    }

    /// Parse string constant
    fn parse_str(&mut self) -> Result<Token> {
        self.require(2, "PtgStr length")?;
        let len = read_u16_le_at(self.data, self.offset)? as usize;
        self.offset += 2;
        let byte_len = len
            .checked_mul(2)
            .ok_or_else(|| Error::InvalidFormula("PtgStr UTF-16 length overflow".to_string()))?;
        self.require(byte_len, "PtgStr text")?;
        let units = self.data[self.offset..self.offset + byte_len]
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
        let string = char::decode_utf16(units)
            .collect::<std::result::Result<String, _>>()
            .map_err(|_| Error::Encoding("invalid UTF-16 in PtgStr".to_string()))?;
        self.offset += byte_len;

        Ok(Token::String(string))
    }

    /// Parse boolean constant
    fn parse_bool(&mut self) -> Result<Token> {
        self.require(1, "PtgBool")?;

        let raw = self.data[self.offset];
        self.offset += 1;
        if raw > 1 {
            return Err(Error::InvalidFormula(format!(
                "invalid PtgBool value {raw}"
            )));
        }

        Ok(Token::Bool(raw != 0))
    }

    /// Parse error constant
    fn parse_err(&mut self) -> Result<Token> {
        self.require(1, "PtgErr")?;

        let error_code = self.data[self.offset];
        self.offset += 1;
        if !is_formula_error_code(error_code) {
            return Err(Error::InvalidFormula(format!(
                "invalid PtgErr code 0x{error_code:02X}"
            )));
        }

        Ok(Token::Error(error_code))
    }

    /// Parse the selector-specific payload of the `PtgAttr` token family.
    fn parse_attr(&mut self) -> Result<Token> {
        self.require(1, "PtgAttr selector")?;
        let selector = self.data[self.offset];
        self.offset += 1;

        if selector == 0x04 {
            // PtgAttrChoose: cOffset is one less than the number of u16
            // offsets that follow it.
            self.require(2, "PtgAttrChoose count")?;
            let count = usize::from(read_u16_le_at(self.data, self.offset)?) + 1;
            self.offset += 2;
            let byte_len = count.checked_mul(2).ok_or_else(|| {
                Error::InvalidFormula("PtgAttrChoose offset count overflow".to_string())
            })?;
            self.require(byte_len, "PtgAttrChoose offsets")?;
            let offsets = self.data[self.offset..self.offset + byte_len]
                .chunks_exact(2)
                .map(|bytes| usize::from(u16::from_le_bytes([bytes[0], bytes[1]])))
                .collect::<Vec<_>>();
            self.offset += byte_len;
            if offsets[0] != byte_len {
                return Err(Error::InvalidFormula(format!(
                    "PtgAttrChoose first offset is {}, expected {byte_len}",
                    offsets[0]
                )));
            }
            self.control_flow_targets.push(self.offset);
            for offset in &offsets[1..] {
                let target = self.offset.checked_add(*offset).ok_or_else(|| {
                    Error::InvalidFormula("PtgAttrChoose target overflow".to_string())
                })?;
                if target > self.data.len() {
                    return Err(Error::InvalidFormula(format!(
                        "PtgAttrChoose target byte {target} exceeds formula length {}",
                        self.data.len()
                    )));
                }
                self.control_flow_targets.push(target);
            }
            return Ok(Token::Attribute(selector));
        }

        match selector {
            // Semi, If, GoTo, Sum, Baxcel, Space, SpaceSemi, IfError all have
            // a two-byte selector-specific payload after the selector byte.
            0x01 | 0x02 | 0x08 | 0x10 | 0x20 | 0x21 | 0x40 | 0x41 | 0x80 => {
                self.require(2, "PtgAttr payload")?;
                let offset = usize::from(read_u16_le_at(self.data, self.offset)?);
                self.offset += 2;
                if matches!(selector, 0x02 | 0x08 | 0x80) {
                    let adjustment = usize::from(selector == 0x08);
                    let target = self
                        .offset
                        .checked_add(offset)
                        .and_then(|value| value.checked_add(adjustment))
                        .ok_or_else(|| {
                            Error::InvalidFormula("PtgAttr target overflow".to_string())
                        })?;
                    if target > self.data.len() {
                        return Err(Error::InvalidFormula(format!(
                            "PtgAttr target byte {target} exceeds formula length {}",
                            self.data.len()
                        )));
                    }
                    self.control_flow_targets.push(target);
                }
                if selector == 0x10 {
                    // PtgAttrSum is semantically SUM(expression).
                    Ok(Token::Function {
                        index: 4,
                        arg_count: 1,
                        is_command: false,
                    })
                } else {
                    Ok(Token::Attribute(selector))
                }
            },
            _ => Err(Error::InvalidFormula(format!(
                "unknown PtgAttr selector 0x{selector:02X}"
            ))),
        }
    }

    fn parse_extended(&mut self) -> Result<Token> {
        self.require(1, "extended Ptg selector")?;
        let selector = self.data[self.offset];
        self.offset += 1;
        match selector {
            ptg_types::EPTG_LIST => self.parse_list(),
            ptg_types::EPTG_SX_NAME => {
                self.require(4, "PtgSxName")?;
                let index = read_u32_le_at(self.data, self.offset)?;
                self.offset += 4;
                Ok(Token::PivotName(index))
            },
            _ => Err(Error::InvalidFormula(format!(
                "unknown extended Ptg selector 0x{selector:02X}"
            ))),
        }
    }

    fn parse_list(&mut self) -> Result<Token> {
        self.require(12, "PtgList")?;
        let sheet_index = read_u16_le_at(self.data, self.offset)?;
        let flags = read_u16_le_at(self.data, self.offset + 2)?;
        let raw_list_index = read_u32_le_at(self.data, self.offset + 4)?;
        let col_first = read_u16_le_at(self.data, self.offset + 8)?;
        let col_last = read_u16_le_at(self.data, self.offset + 10)?;
        self.offset += 12;

        if flags & 0xC000 != 0 {
            return Err(Error::InvalidFormula(format!(
                "PtgList reserved flag bits are nonzero: 0x{:04X}",
                flags & 0xC000
            )));
        }
        let data_type = match (flags >> 10) & 0x03 {
            0 => TableDataType::Reference,
            1 => TableDataType::Value,
            2 => TableDataType::Array,
            _ => {
                return Err(Error::InvalidFormula(
                    "PtgList has reserved data type 3".to_string(),
                ));
            },
        };
        let invalid = flags & 0x1000 != 0;
        let nonresident = !invalid && flags & 0x2000 != 0;
        let (row_type, columns, list_index) = if invalid || nonresident {
            (None, None, None)
        } else {
            if raw_list_index == 0 || raw_list_index == u32::MAX {
                return Err(Error::InvalidFormula(format!(
                    "PtgList has invalid table identifier {raw_list_index}"
                )));
            }
            let row_type = parse_table_row_type(((flags >> 2) & 0x1F) as u8)?;
            let columns = match flags & 0x03 {
                0 => TableColumns::All,
                1 => {
                    if col_first >= 16_384 {
                        return Err(Error::InvalidFormula(
                            "PtgList first column is outside worksheet bounds".to_string(),
                        ));
                    }
                    TableColumns::One(col_first)
                },
                2 => {
                    if col_first > col_last || col_last >= 16_384 {
                        return Err(Error::InvalidFormula(
                            "PtgList column range is invalid".to_string(),
                        ));
                    }
                    TableColumns::Range {
                        first: col_first,
                        last: col_last,
                    }
                },
                _ => {
                    return Err(Error::InvalidFormula(
                        "PtgList has reserved column selector 3".to_string(),
                    ));
                },
            };
            (Some(row_type), Some(columns), Some(raw_list_index))
        };
        let external = if nonresident {
            Some(self.parse_extra_list()?)
        } else {
            None
        };
        Ok(Token::TableReference(TableReference {
            sheet_index,
            row_type,
            columns,
            square_bracket_space: flags & 0x0080 != 0,
            comma_space: flags & 0x0100 != 0,
            data_type,
            invalid,
            list_index,
            external,
        }))
    }

    fn parse_extra_list(&mut self) -> Result<ExternalTableReference> {
        self.require_extra(5, "PtgExtraList header")?;
        let has_columns = match self.extra[self.extra_offset] {
            0 => false,
            1 => true,
            value => {
                return Err(Error::InvalidFormula(format!(
                    "invalid PtgExtraList hasColumns {value}"
                )));
            },
        };
        let row_flags = read_u16_le_at(self.extra, self.extra_offset + 1)?;
        if row_flags & !0x001F != 0 {
            return Err(Error::InvalidFormula(format!(
                "PtgExtraList reserved row bits are nonzero: 0x{:04X}",
                row_flags & !0x001F
            )));
        }
        let row_type = parse_table_row_type((row_flags & 0x1F) as u8)?;
        let table_len = usize::from(read_u16_le_at(self.extra, self.extra_offset + 3)?);
        self.extra_offset += 5;
        if table_len == 0 || table_len >= 256 {
            return Err(Error::InvalidFormula(format!(
                "PtgExtraList table length {table_len} is outside 1..=255"
            )));
        }
        let table = self.parse_extra_utf16(table_len, "PtgExtraList table")?;
        let columns = if has_columns {
            self.require_extra(3, "SxSu header")?;
            let reserved = read_u16_le_at(self.extra, self.extra_offset)?;
            let count = self.extra[self.extra_offset + 2];
            self.extra_offset += 3;
            if reserved != 0 {
                return Err(Error::InvalidFormula(
                    "SxSu reserved field is nonzero".to_string(),
                ));
            }
            if !matches!(count, 1 | 2) {
                return Err(Error::InvalidFormula(format!(
                    "invalid SxSu column count {count}"
                )));
            }
            let first = self.parse_sxos(count == 2, "SxSu first column")?;
            if count == 1 {
                TableNamedColumns::One(first)
            } else {
                let last = self.parse_sxos(false, "SxSu last column")?;
                TableNamedColumns::Range { first, last }
            }
        } else {
            TableNamedColumns::All
        };
        Ok(ExternalTableReference {
            table,
            row_type,
            columns,
        })
    }

    fn parse_sxos(&mut self, expected_not_last: bool, context: &str) -> Result<String> {
        self.require_extra(7, context)?;
        let not_last = match self.extra[self.extra_offset] {
            0 => false,
            1 => true,
            value => {
                return Err(Error::InvalidFormula(format!(
                    "invalid {context} notLast {value}"
                )));
            },
        };
        let reserved = read_u16_le_at(self.extra, self.extra_offset + 1)?;
        let length = usize::try_from(read_u32_le_at(self.extra, self.extra_offset + 3)?)
            .map_err(|_| Error::InvalidFormula(format!("{context} length overflow")))?;
        self.extra_offset += 7;
        if not_last != expected_not_last {
            return Err(Error::InvalidFormula(format!(
                "{context} has inconsistent notLast flag"
            )));
        }
        if reserved != 2 {
            return Err(Error::InvalidFormula(format!(
                "{context} reserved field is {reserved}, expected 2"
            )));
        }
        if length == 0 {
            return Err(Error::InvalidFormula(format!("{context} name is empty")));
        }
        self.parse_extra_utf16(length, context)
    }

    fn parse_extra_utf16(&mut self, units: usize, context: &str) -> Result<String> {
        let byte_len = units
            .checked_mul(2)
            .ok_or_else(|| Error::InvalidFormula(format!("{context} length overflow")))?;
        self.require_extra(byte_len, context)?;
        let value = char::decode_utf16(
            self.extra[self.extra_offset..self.extra_offset + byte_len]
                .chunks_exact(2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]])),
        )
        .collect::<std::result::Result<String, _>>()
        .map_err(|_| Error::Encoding(format!("invalid UTF-16 in {context}")))?;
        self.extra_offset += byte_len;
        Ok(value)
    }

    fn require_extra(&self, len: usize, context: &str) -> Result<()> {
        if self.extra_offset + len <= self.extra.len() {
            Ok(())
        } else {
            Err(Error::InvalidFormula(format!(
                "truncated {context} at ancillary byte {}: need {len} bytes, have {}",
                self.extra_offset,
                self.extra.len().saturating_sub(self.extra_offset)
            )))
        }
    }

    fn parse_array(&mut self) -> Result<Token> {
        let token = self.data[self.offset - 1];
        if token & 0x80 != 0 || !matches!(token & 0x60, 0x40 | 0x60) {
            return Err(Error::InvalidFormula(format!(
                "PtgArray has invalid data type in token 0x{token:02X}"
            )));
        }
        self.require(14, "PtgArray")?;
        self.offset += 14;
        self.require_extra(8, "PtgExtraArray dimensions")?;
        let rows = read_u32_le_at(self.extra, self.extra_offset)?;
        let cols = read_u32_le_at(self.extra, self.extra_offset + 4)?;
        self.extra_offset += 8;
        if rows == 0 || cols == 0 || rows > 1_048_576 || cols > 16_384 {
            return Err(Error::InvalidFormula(format!(
                "PtgExtraArray dimensions {rows}x{cols} are invalid"
            )));
        }
        let count_u64 = u64::from(rows)
            .checked_mul(u64::from(cols))
            .ok_or_else(|| Error::InvalidFormula("array size overflow".to_string()))?;
        let count = usize::try_from(count_u64)
            .map_err(|_| Error::InvalidFormula("array is too large".to_string()))?;
        // Every SerAr uses at least two bytes. Reject impossible dimensions
        // before allocating based on attacker-controlled counts.
        if count > self.extra.len().saturating_sub(self.extra_offset) / 2 {
            return Err(Error::InvalidFormula(format!(
                "PtgExtraArray declares {count} values beyond its ancillary payload"
            )));
        }
        let mut values = Vec::with_capacity(count);
        for _ in 0..count {
            self.require_extra(1, "SerAr tag")?;
            let tag = self.extra[self.extra_offset];
            self.extra_offset += 1;
            let value = match tag {
                0x00 => {
                    self.require_extra(8, "SerNum")?;
                    let value = read_f64_le_at(self.extra, self.extra_offset)?;
                    self.extra_offset += 8;
                    validate_xnum(value, "SerNum")?;
                    ArrayValue::Number(value)
                },
                0x01 => {
                    self.require_extra(2, "SerStr length")?;
                    let len = usize::from(read_u16_le_at(self.extra, self.extra_offset)?);
                    self.extra_offset += 2;
                    if len >= 256 {
                        return Err(Error::InvalidFormula(format!(
                            "SerStr length {len} exceeds 255 UTF-16 code units"
                        )));
                    }
                    let byte_len = len.checked_mul(2).ok_or_else(|| {
                        Error::InvalidFormula("SerStr length overflow".to_string())
                    })?;
                    self.require_extra(byte_len, "SerStr text")?;
                    let units = self.extra[self.extra_offset..self.extra_offset + byte_len]
                        .chunks_exact(2)
                        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
                    let value = char::decode_utf16(units)
                        .collect::<std::result::Result<String, _>>()
                        .map_err(|_| Error::Encoding("invalid UTF-16 in SerStr".to_string()))?;
                    self.extra_offset += byte_len;
                    ArrayValue::String(value)
                },
                0x02 => {
                    self.require_extra(1, "SerBool")?;
                    let value = self.extra[self.extra_offset];
                    self.extra_offset += 1;
                    if value > 1 {
                        return Err(Error::InvalidFormula(format!(
                            "invalid SerBool value {value}"
                        )));
                    }
                    ArrayValue::Bool(value != 0)
                },
                0x04 => {
                    self.require_extra(4, "SerErr")?;
                    let error = self.extra[self.extra_offset];
                    if !matches!(error, 0x00 | 0x07 | 0x0F | 0x17 | 0x1D | 0x24 | 0x2A | 0x2B) {
                        return Err(Error::InvalidFormula(format!(
                            "invalid SerErr code 0x{error:02X}"
                        )));
                    }
                    if self.extra[self.extra_offset + 1..self.extra_offset + 4]
                        .iter()
                        .any(|byte| *byte != 0)
                    {
                        return Err(Error::InvalidFormula(
                            "SerErr reserved bytes are nonzero".to_string(),
                        ));
                    }
                    self.extra_offset += 4;
                    ArrayValue::Error(error)
                },
                _ => {
                    return Err(Error::InvalidFormula(format!(
                        "unknown SerAr tag 0x{tag:02X}"
                    )));
                },
            };
            values.push(value);
        }
        Ok(Token::Array { rows, cols, values })
    }

    fn parse_memory(&mut self, mut kind: MemoryKind) -> Result<Token> {
        let token = self.data[self.offset - 1];
        if token & 0x80 != 0 {
            return Err(Error::InvalidFormula(format!(
                "memory token 0x{token:02X} has its reserved bit set"
            )));
        }
        let (payload_len, cce_offset) = match kind {
            MemoryKind::Function => (2, 0),
            MemoryKind::Area | MemoryKind::NoMemory => (6, 4),
            MemoryKind::Error(_) => (6, 4),
        };
        self.require(payload_len, "memory token")?;
        if matches!(kind, MemoryKind::Error(_)) {
            let error = self.data[self.offset];
            if !matches!(error, 0x00 | 0x07 | 0x0F | 0x17 | 0x1D | 0x24 | 0x2A | 0x2B) {
                return Err(Error::InvalidFormula(format!(
                    "invalid PtgMemErr code 0x{error:02X}"
                )));
            }
            kind = MemoryKind::Error(error);
        }
        let expression_bytes = read_u16_le_at(self.data, self.offset + cce_offset)?;
        self.offset += payload_len;
        if expression_bytes == 0 {
            return Err(Error::InvalidFormula(
                "memory token has an empty reference expression".to_string(),
            ));
        }
        if usize::from(expression_bytes) > self.data.len().saturating_sub(self.offset) {
            return Err(Error::InvalidFormula(format!(
                "memory token declares {expression_bytes} expression bytes, but only {} remain",
                self.data.len().saturating_sub(self.offset)
            )));
        }
        self.memory_expression_ends
            .push(self.offset + usize::from(expression_bytes));

        let mut cached_ranges = Vec::new();
        if kind == MemoryKind::Area {
            self.require_extra(4, "PtgExtraMem count")?;
            let count = usize::try_from(read_u32_le_at(self.extra, self.extra_offset)?)
                .map_err(|_| Error::InvalidFormula("PtgExtraMem is too large".to_string()))?;
            self.extra_offset += 4;
            if count > self.extra.len().saturating_sub(self.extra_offset) / 16 {
                return Err(Error::InvalidFormula(format!(
                    "PtgExtraMem declares {count} ranges beyond its ancillary payload"
                )));
            }
            cached_ranges.reserve(count);
            for _ in 0..count {
                self.require_extra(16, "PtgExtraMem range")?;
                let range = [
                    read_u32_le_at(self.extra, self.extra_offset)?,
                    read_u32_le_at(self.extra, self.extra_offset + 4)?,
                    read_u32_le_at(self.extra, self.extra_offset + 8)?,
                    read_u32_le_at(self.extra, self.extra_offset + 12)?,
                ];
                self.extra_offset += 16;
                let invalid = range == [1_048_575, 1_048_575, 16_383, 16_383];
                if !invalid {
                    Range::new(range[0], range[1], range[2], range[3])?;
                }
                cached_ranges.push(range);
            }
        }
        Ok(Token::Memory {
            kind,
            expression_bytes,
            cached_ranges,
        })
    }

    /// Parse cell reference
    fn parse_ref(&mut self, offset_reference: bool) -> Result<Token> {
        self.validate_classed_token("PtgRef")?;
        self.require(6, "PtgRef")?;

        let row_data = read_u32_le_at(self.data, self.offset)?;
        let col_data = read_u16_le_at(self.data, self.offset + 4)?;
        self.offset += 6;

        // Extract row and column (with relative flags)
        let col_relative = (col_data & 0x4000) != 0;
        let row_relative = (col_data & 0x8000) != 0;
        let (row, col) = self.resolve_reference(
            row_data,
            col_data & 0x3FFF,
            row_relative,
            col_relative,
            offset_reference,
        )?;

        Ok(Token::CellRef {
            row,
            col,
            row_relative,
            col_relative,
        })
    }

    /// Parse area reference
    fn parse_area(&mut self, offset_reference: bool) -> Result<Token> {
        self.validate_classed_token("PtgArea")?;
        self.require(12, "PtgArea")?;

        let row_first_data = read_u32_le_at(self.data, self.offset)?;
        let row_last_data = read_u32_le_at(self.data, self.offset + 4)?;
        let col_first_data = read_u16_le_at(self.data, self.offset + 8)?;
        let col_last_data = read_u16_le_at(self.data, self.offset + 10)?;
        self.offset += 12;

        let col_first_relative = (col_first_data & 0x4000) != 0;
        let row_first_relative = (col_first_data & 0x8000) != 0;
        let col_last_relative = (col_last_data & 0x4000) != 0;
        let row_last_relative = (col_last_data & 0x8000) != 0;
        let (row_first, col_first) = self.resolve_reference(
            row_first_data,
            col_first_data & 0x3FFF,
            row_first_relative,
            col_first_relative,
            offset_reference,
        )?;
        let (row_last, col_last) = self.resolve_reference(
            row_last_data,
            col_last_data & 0x3FFF,
            row_last_relative,
            col_last_relative,
            offset_reference,
        )?;
        Range::new(row_first, row_last, col_first, col_last)?;

        Ok(Token::AreaRef {
            row_first,
            row_last,
            col_first,
            col_last,
            row_first_relative,
            row_last_relative,
            col_first_relative,
            col_last_relative,
        })
    }

    fn parse_reference_error(&mut self, is_area: bool, is_3d: bool) -> Result<Token> {
        let token = self.data[self.offset - 1];
        if token & 0x80 != 0 {
            return Err(Error::InvalidFormula(format!(
                "reference-error token 0x{token:02X} has its reserved bit set"
            )));
        }
        let sheet_index = if is_3d {
            self.require(2, "3D reference-error sheet index")?;
            let index = read_u16_le_at(self.data, self.offset)?;
            self.offset += 2;
            Some(index)
        } else {
            None
        };
        let unused_len = if is_area { 12 } else { 6 };
        self.require(unused_len, "reference-error payload")?;
        self.offset += unused_len;
        Ok(Token::ReferenceError {
            is_area,
            sheet_index,
        })
    }

    fn parse_ref_3d(&mut self) -> Result<Token> {
        let token = self.data[self.offset - 1];
        if token & 0x80 != 0 {
            return Err(Error::InvalidFormula(format!(
                "PtgRef3d token 0x{token:02X} has its reserved bit set"
            )));
        }
        self.require(8, "PtgRef3d")?;
        let sheet_index = read_u16_le_at(self.data, self.offset)?;
        let row = read_u32_le_at(self.data, self.offset + 2)?;
        let col_data = read_u16_le_at(self.data, self.offset + 6)?;
        self.offset += 8;
        let col = u32::from(col_data & 0x3FFF);
        if row >= 1_048_576 || col >= 16_384 {
            return Err(Error::InvalidFormula(format!(
                "PtgRef3d coordinate ({row}, {col}) is outside the worksheet"
            )));
        }
        Ok(Token::CellRef3d {
            sheet_index,
            row,
            col,
            row_relative: col_data & 0x8000 != 0,
            col_relative: col_data & 0x4000 != 0,
        })
    }

    fn parse_area_3d(&mut self) -> Result<Token> {
        let token = self.data[self.offset - 1];
        if token & 0x80 != 0 {
            return Err(Error::InvalidFormula(format!(
                "PtgArea3d token 0x{token:02X} has its reserved bit set"
            )));
        }
        self.require(14, "PtgArea3d")?;
        let sheet_index = read_u16_le_at(self.data, self.offset)?;
        let row_first = read_u32_le_at(self.data, self.offset + 2)?;
        let row_last = read_u32_le_at(self.data, self.offset + 6)?;
        let col_first_data = read_u16_le_at(self.data, self.offset + 10)?;
        let col_last_data = read_u16_le_at(self.data, self.offset + 12)?;
        self.offset += 14;
        let col_first = u32::from(col_first_data & 0x3FFF);
        let col_last = u32::from(col_last_data & 0x3FFF);
        Range::new(row_first, row_last, col_first, col_last)?;
        Ok(Token::AreaRef3d {
            sheet_index,
            row_first,
            row_last,
            col_first,
            col_last,
            row_first_relative: col_first_data & 0x8000 != 0,
            row_last_relative: col_last_data & 0x8000 != 0,
            col_first_relative: col_first_data & 0x4000 != 0,
            col_last_relative: col_last_data & 0x4000 != 0,
        })
    }

    fn resolve_reference(
        &self,
        row_data: u32,
        col_data: u16,
        row_relative: bool,
        col_relative: bool,
        offset_reference: bool,
    ) -> Result<(u32, u32)> {
        if !offset_reference {
            if row_data >= 1_048_576 || col_data >= 16_384 {
                return Err(Error::InvalidFormula(format!(
                    "reference ({row_data}, {col_data}) is outside the worksheet"
                )));
            }
            return Ok((row_data, u32::from(col_data)));
        }
        let (base_row, base_col) = self.base_cell.ok_or_else(|| {
            Error::InvalidFormula(
                "PtgRefN/PtgAreaN requires a target cell for offset resolution".to_string(),
            )
        })?;

        let row = if row_relative {
            add_wrapped_offset(base_row, row_data as i32, 1_048_576)
        } else {
            row_data
        };
        let col = if col_relative {
            let signed = if col_data & 0x2000 != 0 {
                i32::from(col_data) - 0x4000
            } else {
                i32::from(col_data)
            };
            add_wrapped_offset(base_col, signed, 16_384)
        } else {
            u32::from(col_data)
        };
        if row >= 1_048_576 || col >= 16_384 {
            return Err(Error::InvalidFormula(format!(
                "resolved reference ({row}, {col}) is outside the worksheet"
            )));
        }
        Ok((row, col))
    }

    /// Parse function with fixed arguments
    fn parse_func(&mut self) -> Result<Token> {
        self.validate_classed_token("PtgFunc")?;
        self.require(2, "PtgFunc")?;

        let index = read_u16_le_at(self.data, self.offset)?;
        self.offset += 2;

        let arg_count = Self::get_function_arg_count(index)?;

        Ok(Token::Function {
            index,
            arg_count,
            is_command: false,
        })
    }

    /// Parse function with variable arguments
    fn parse_func_var(&mut self) -> Result<Token> {
        self.validate_classed_token("PtgFuncVar")?;
        self.require(3, "PtgFuncVar")?;

        let arg_count = self.data[self.offset];
        let tab = read_u16_le_at(self.data, self.offset + 1)?;
        let index = tab & 0x7FFF;
        let is_command = tab & 0x8000 != 0;
        self.offset += 3;

        if !is_command && let Some(function) = builtin_function_by_index(index) {
            if function.min_args == function.max_args {
                return Err(Error::InvalidFormula(format!(
                    "PtgFuncVar specifies fixed-arity function {}",
                    function.name
                )));
            }
            if !function.accepts_arg_count(arg_count) {
                return Err(Error::InvalidFormula(format!(
                    "{} does not accept {arg_count} arguments",
                    function.name
                )));
            }
        }

        Ok(Token::Function {
            index,
            arg_count,
            is_command,
        })
    }

    /// Parse defined name reference
    fn parse_name(&mut self) -> Result<Token> {
        self.validate_classed_token("PtgName")?;
        self.require(4, "PtgName")?;

        let name_index = read_u32_le_at(self.data, self.offset)?;
        self.offset += 4;
        if name_index == 0 {
            return Err(Error::InvalidFormula(
                "PtgName index is one-based and cannot be zero".to_string(),
            ));
        }

        Ok(Token::Name(name_index))
    }

    fn parse_name_x(&mut self) -> Result<Token> {
        self.validate_classed_token("PtgNameX")?;
        self.require(6, "PtgNameX")?;
        let sheet_index = read_u16_le_at(self.data, self.offset)?;
        let name_index = read_u32_le_at(self.data, self.offset + 2)?;
        self.offset += 6;
        if name_index == 0 {
            return Err(Error::InvalidFormula(
                "PtgNameX name index is one-based and cannot be zero".to_string(),
            ));
        }
        Ok(Token::ExternalName {
            sheet_index,
            name_index,
        })
    }

    /// Resolve the parameter count of a fixed-arity `Ftab` function.
    fn get_function_arg_count(index: u16) -> Result<u8> {
        let function = builtin_function_by_index(index).ok_or_else(|| {
            Error::UnsupportedFeature(format!(
                "fixed-arity XLSB function index {index} is not a non-macro Ftab entry"
            ))
        })?;
        if function.min_args != function.max_args {
            return Err(Error::InvalidFormula(format!(
                "PtgFunc specifies variable-arity function {}",
                function.name
            )));
        }
        Ok(function.min_args)
    }
}
