use crate::{Error, Result};

pub(crate) const DCON_RECORD_TYPE: u16 = 0x0050;
pub(crate) const DCON_REF_RECORD_TYPE: u16 = 0x0051;
pub(crate) const DCON_NAME_RECORD_TYPE: u16 = 0x0052;
pub(crate) const DCON_BIN_RECORD_TYPE: u16 = 0x01B5;
const MAX_SOURCES: usize = 16_384;
const MAX_PATH_UNITS: usize = 4_096;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConsolidationFunction {
    Average,
    CountNumbers,
    Count,
    Maximum,
    Minimum,
    Product,
    StandardDeviation,
    StandardDeviationPopulation,
    Sum,
    Variance,
    VariancePopulation,
}

impl ConsolidationFunction {
    fn from_code(code: u16) -> Result<Self> {
        Ok(match code {
            0 => Self::Average,
            1 => Self::CountNumbers,
            2 => Self::Count,
            3 => Self::Maximum,
            4 => Self::Minimum,
            5 => Self::Product,
            6 => Self::StandardDeviation,
            7 => Self::StandardDeviationPopulation,
            8 => Self::Sum,
            9 => Self::Variance,
            10 => Self::VariancePopulation,
            _ => return invalid(DCON_RECORD_TYPE, "DCon aggregation function must be 0..=10"),
        })
    }
    pub(crate) const fn code(self) -> u16 {
        match self {
            Self::Average => 0,
            Self::CountNumbers => 1,
            Self::Count => 2,
            Self::Maximum => 3,
            Self::Minimum => 4,
            Self::Product => 5,
            Self::StandardDeviation => 6,
            Self::StandardDeviationPopulation => 7,
            Self::Sum => 8,
            Self::Variance => 9,
            Self::VariancePopulation => 10,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ConsolidationRange {
    first_row: u16,
    last_row: u16,
    first_column: u8,
    last_column: u8,
}
impl ConsolidationRange {
    pub fn new(first_row: u16, last_row: u16, first_column: u8, last_column: u8) -> Result<Self> {
        if first_row > last_row || first_column > last_column {
            return invalid(
                DCON_REF_RECORD_TYPE,
                "DConRef contains an inverted source range",
            );
        }
        Ok(Self {
            first_row,
            last_row,
            first_column,
            last_column,
        })
    }
    pub fn first_row(&self) -> u16 {
        self.first_row
    }
    pub fn last_row(&self) -> u16 {
        self.last_row
    }
    pub fn first_column(&self) -> u8 {
        self.first_column
    }
    pub fn last_column(&self) -> u8 {
        self.last_column
    }
}

/// Encoded `DConFile` metadata. It is retained but never opened or resolved.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ConsolidationFile {
    encoded_path: String,
}
impl ConsolidationFile {
    pub fn new(encoded_path: impl Into<String>) -> Result<Self> {
        let value = Self {
            encoded_path: encoded_path.into(),
        };
        value.validate(DCON_REF_RECORD_TYPE)?;
        Ok(value)
    }
    pub fn self_reference(sheet_name: &str) -> Result<Self> {
        Self::new(format!("\u{2}{sheet_name}"))
    }
    pub fn encoded_path(&self) -> &str {
        &self.encoded_path
    }
    pub fn is_self_reference(&self) -> bool {
        self.encoded_path.starts_with('\u{2}')
    }
    pub fn is_external(&self) -> bool {
        self.encoded_path.starts_with('\u{1}')
    }
    fn validate(&self, record_type: u16) -> Result<()> {
        let units = self.encoded_path.encode_utf16().collect::<Vec<_>>();
        if !(2..=MAX_PATH_UNITS).contains(&units.len()) {
            return invalid(
                record_type,
                "DConFile path length must be 2..=4096 UTF-16 code units",
            );
        }
        if !matches!(units[0], 1 | 2) || units[1..].contains(&0) {
            return invalid(
                record_type,
                "DConFile must be an encoded external or self-reference path",
            );
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConsolidationBuiltInName {
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
impl ConsolidationBuiltInName {
    fn from_code(code: u8) -> Result<Self> {
        Ok(match code {
            0 => Self::ConsolidateArea,
            1 => Self::AutoOpen,
            2 => Self::AutoClose,
            3 => Self::Extract,
            4 => Self::Database,
            5 => Self::Criteria,
            6 => Self::PrintArea,
            7 => Self::PrintTitles,
            8 => Self::Recorder,
            9 => Self::DataForm,
            10 => Self::AutoActivate,
            11 => Self::AutoDeactivate,
            12 => Self::SheetTitle,
            13 => Self::FilterDatabase,
            _ => return invalid(DCON_BIN_RECORD_TYPE, "DConBin built-in name must be 0..=13"),
        })
    }
    pub(crate) const fn code(self) -> u8 {
        match self {
            Self::ConsolidateArea => 0,
            Self::AutoOpen => 1,
            Self::AutoClose => 2,
            Self::Extract => 3,
            Self::Database => 4,
            Self::Criteria => 5,
            Self::PrintArea => 6,
            Self::PrintTitles => 7,
            Self::Recorder => 8,
            Self::DataForm => 9,
            Self::AutoActivate => 10,
            Self::AutoDeactivate => 11,
            Self::SheetTitle => 12,
            Self::FilterDatabase => 13,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ConsolidationSource {
    CellRange {
        range: ConsolidationRange,
        file: ConsolidationFile,
    },
    DefinedName {
        name: String,
        file: Option<ConsolidationFile>,
    },
    BuiltInName {
        name: ConsolidationBuiltInName,
        file: Option<ConsolidationFile>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Consolidation {
    function: ConsolidationFunction,
    use_left_labels: bool,
    use_top_labels: bool,
    create_links: bool,
    sources: Vec<ConsolidationSource>,
}
impl Consolidation {
    pub fn new(function: ConsolidationFunction) -> Self {
        Self {
            function,
            use_left_labels: false,
            use_top_labels: false,
            create_links: false,
            sources: Vec::new(),
        }
    }
    pub fn function(&self) -> ConsolidationFunction {
        self.function
    }
    pub fn uses_left_labels(&self) -> bool {
        self.use_left_labels
    }
    pub fn uses_top_labels(&self) -> bool {
        self.use_top_labels
    }
    pub fn creates_links(&self) -> bool {
        self.create_links
    }
    pub fn sources(&self) -> &[ConsolidationSource] {
        &self.sources
    }
    pub fn set_use_left_labels(&mut self, value: bool) {
        self.use_left_labels = value;
    }
    pub fn set_use_top_labels(&mut self, value: bool) {
        self.use_top_labels = value;
    }
    pub fn set_create_links(&mut self, value: bool) {
        self.create_links = value;
    }
    pub fn add_source(&mut self, source: ConsolidationSource) -> Result<()> {
        if self.sources.len() >= MAX_SOURCES {
            return invalid(
                DCON_RECORD_TYPE,
                "data-consolidation source count exceeds 16384",
            );
        }
        validate_source(&source)?;
        self.sources.push(source);
        Ok(())
    }
    pub(crate) fn validate_for_write(&self) -> Result<()> {
        if self.sources.len() > MAX_SOURCES {
            return invalid(
                DCON_RECORD_TYPE,
                "data-consolidation source count exceeds 16384",
            );
        }
        for source in &self.sources {
            validate_source(source)?;
        }
        Ok(())
    }
}

fn validate_source(source: &ConsolidationSource) -> Result<()> {
    let (base, file, record_type) = match source {
        ConsolidationSource::CellRange { range, file } => {
            ConsolidationRange::new(
                range.first_row,
                range.last_row,
                range.first_column,
                range.last_column,
            )?;
            file.validate(DCON_REF_RECORD_TYPE)?;
            (8, Some(file), DCON_REF_RECORD_TYPE)
        },
        ConsolidationSource::DefinedName { name, file } => {
            validate_defined_name(name)?;
            if let Some(file) = file {
                file.validate(DCON_NAME_RECORD_TYPE)?;
            }
            (
                5 + name.encode_utf16().count() * 2,
                file.as_ref(),
                DCON_NAME_RECORD_TYPE,
            )
        },
        ConsolidationSource::BuiltInName { file, .. } => {
            if let Some(file) = file {
                file.validate(DCON_BIN_RECORD_TYPE)?;
            }
            (6, file.as_ref(), DCON_BIN_RECORD_TYPE)
        },
    };
    let file_size = file.map_or(0, |file| {
        let units = file.encoded_path.encode_utf16().collect::<Vec<_>>();
        let wide = !units.iter().all(|unit| *unit <= 0xff);
        3 + units.len() * if wide { 2 } else { 1 }
            + if file.is_self_reference() {
                if wide { 2 } else { 1 }
            } else {
                0
            }
    });
    if base + file_size > 8_224 {
        return invalid(
            record_type,
            "data-consolidation source exceeds BIFF8 record size",
        );
    }
    Ok(())
}

fn validate_defined_name(name: &str) -> Result<()> {
    if !(1..=255).contains(&name.encode_utf16().count()) {
        return invalid(
            DCON_NAME_RECORD_TYPE,
            "DConName length must be 1..=255 UTF-16 code units",
        );
    }
    if name.eq_ignore_ascii_case("TRUE")
        || name.eq_ignore_ascii_case("FALSE")
        || is_a1(name)
        || is_r1c1(name)
    {
        return invalid(
            DCON_NAME_RECORD_TYPE,
            "DConName is a reserved value or cell reference",
        );
    }
    let mut chars = name.chars();
    let start = |ch: char| ch == '_' || ch == '\\' || ch.is_alphabetic();
    if !start(chars.next().expect("validated non-empty name"))
        || chars.any(|ch| !(start(ch) || ch.is_numeric() || matches!(ch, '.' | '?' | '\u{061f}')))
    {
        return invalid(
            DCON_NAME_RECORD_TYPE,
            "DConName contains invalid name characters",
        );
    }
    Ok(())
}
fn is_a1(value: &str) -> bool {
    let Some(split) = value.find(|ch: char| ch.is_ascii_digit()) else {
        return false;
    };
    let (column, row) = value.split_at(split);
    if column.is_empty() || column.len() > 2 || !column.bytes().all(|b| b.is_ascii_alphabetic()) {
        return false;
    }
    let column = column.bytes().fold(0u16, |n, b| {
        n * 26 + u16::from(b.to_ascii_uppercase() - b'A' + 1)
    });
    column <= 256
        && row
            .parse::<u32>()
            .is_ok_and(|row| (1..=65_536).contains(&row))
}
fn is_r1c1(value: &str) -> bool {
    let upper = value.to_ascii_uppercase();
    let Some(rest) = upper.strip_prefix('R') else {
        return false;
    };
    let Some((row, column)) = rest.split_once('C') else {
        return false;
    };
    row.parse::<u32>().is_ok_and(|v| (1..=65_536).contains(&v))
        && column.parse::<u16>().is_ok_and(|v| (1..=256).contains(&v))
}

#[derive(Debug, Default)]
pub(crate) struct ConsolidationCollector {
    value: Option<Consolidation>,
    open: bool,
    closed: bool,
}
impl ConsolidationCollector {
    pub(crate) fn new() -> Self {
        Self::default()
    }
    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> Result<()> {
        let source = matches!(
            record_type,
            DCON_REF_RECORD_TYPE | DCON_NAME_RECORD_TYPE | DCON_BIN_RECORD_TYPE
        );
        if record_type == DCON_RECORD_TYPE {
            if self.value.is_some() {
                return invalid(
                    record_type,
                    "worksheet contains more than one DCon directory",
                );
            }
            self.value = Some(parse_dcon(data)?);
            self.open = true;
            return Ok(());
        }
        if source {
            if !self.open || self.closed {
                return invalid(
                    record_type,
                    "data-consolidation source is detached from DCon",
                );
            }
            let source = match record_type {
                DCON_REF_RECORD_TYPE => parse_ref(data)?,
                DCON_NAME_RECORD_TYPE => parse_name(data)?,
                DCON_BIN_RECORD_TYPE => parse_bin(data)?,
                _ => unreachable!(),
            };
            self.value
                .as_mut()
                .expect("open directory")
                .add_source(source)?;
        } else if self.open {
            self.open = false;
            self.closed = true;
        }
        Ok(())
    }
    pub(crate) fn finish(self) -> Option<Consolidation> {
        self.value
    }
}

fn parse_dcon(data: &[u8]) -> Result<Consolidation> {
    if data.len() != 8 {
        return invalid(DCON_RECORD_TYPE, "DCon payload must be exactly 8 bytes");
    }
    let mut value = Consolidation::new(ConsolidationFunction::from_code(u16::from_le_bytes([
        data[0], data[1],
    ]))?);
    value.use_left_labels = boolean(&data[2..4], DCON_RECORD_TYPE)?;
    value.use_top_labels = boolean(&data[4..6], DCON_RECORD_TYPE)?;
    value.create_links = boolean(&data[6..8], DCON_RECORD_TYPE)?;
    Ok(value)
}
fn parse_ref(data: &[u8]) -> Result<ConsolidationSource> {
    let mut c = Cursor::new(data, DCON_REF_RECORD_TYPE);
    let range = ConsolidationRange::new(c.u16()?, c.u16()?, c.u8()?, c.u8()?)?;
    let count = usize::from(c.u16()?);
    if count < 2 {
        return invalid(DCON_REF_RECORD_TYPE, "DConRef cchFile must be at least 2");
    }
    let file = parse_file(&mut c, count)?;
    c.finish()?;
    Ok(ConsolidationSource::CellRange { range, file })
}
fn parse_name(data: &[u8]) -> Result<ConsolidationSource> {
    let mut c = Cursor::new(data, DCON_NAME_RECORD_TYPE);
    let name = c.string()?;
    validate_defined_name(&name)?;
    let count = usize::from(c.u16()?);
    let file = if count == 0 {
        None
    } else {
        if count < 2 {
            return invalid(
                DCON_NAME_RECORD_TYPE,
                "DConName cchFile must be zero or at least 2",
            );
        }
        Some(parse_file(&mut c, count)?)
    };
    c.finish()?;
    Ok(ConsolidationSource::DefinedName { name, file })
}
fn parse_bin(data: &[u8]) -> Result<ConsolidationSource> {
    let mut c = Cursor::new(data, DCON_BIN_RECORD_TYPE);
    let name = ConsolidationBuiltInName::from_code(c.u8()?)?;
    if c.u16()? != 0 || c.u8()? != 0 {
        return invalid(DCON_BIN_RECORD_TYPE, "DConBin reserved fields must be zero");
    }
    let count = usize::from(c.u16()?);
    let file = if count == 0 {
        None
    } else {
        if count < 2 {
            return invalid(
                DCON_BIN_RECORD_TYPE,
                "DConBin cchFile must be zero or at least 2",
            );
        }
        Some(parse_file(&mut c, count)?)
    };
    c.finish()?;
    Ok(ConsolidationSource::BuiltInName { name, file })
}
fn parse_file(c: &mut Cursor<'_>, count: usize) -> Result<ConsolidationFile> {
    if count > MAX_PATH_UNITS {
        return invalid(c.record_type, "DConFile exceeds the path resource limit");
    }
    let flags = c.u8()?;
    if flags & !1 != 0 {
        return invalid(c.record_type, "DConFile contains reserved string flag bits");
    }
    let wide = flags & 1 != 0;
    let raw = c.take(
        count
            .checked_mul(if wide { 2 } else { 1 })
            .ok_or_else(|| Error::InvalidData("DConFile size overflow".into()))?,
    )?;
    let encoded_path = if wide {
        let units = raw
            .chunks_exact(2)
            .map(|p| u16::from_le_bytes([p[0], p[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map_err(|_| Error::InvalidData("DConFile contains invalid UTF-16".into()))?
    } else {
        raw.iter().map(|byte| char::from(*byte)).collect()
    };
    let file = ConsolidationFile { encoded_path };
    file.validate(c.record_type)?;
    if file.is_self_reference()
        && c.take(if wide { 2 } else { 1 })?
            .iter()
            .any(|byte| *byte != 0)
    {
        return invalid(
            c.record_type,
            "DConFile self-reference padding must be zero",
        );
    }
    Ok(file)
}
fn boolean(data: &[u8], record_type: u16) -> Result<bool> {
    match u16::from_le_bytes([data[0], data[1]]) {
        0 => Ok(false),
        1 => Ok(true),
        _ => invalid(record_type, "DCon Boolean must be 0 or 1"),
    }
}
fn invalid<T>(record_type: u16, message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidRecord {
        record_type,
        message: message.into(),
    })
}

struct Cursor<'a> {
    data: &'a [u8],
    offset: usize,
    record_type: u16,
}
impl<'a> Cursor<'a> {
    fn new(data: &'a [u8], record_type: u16) -> Self {
        Self {
            data,
            offset: 0,
            record_type,
        }
    }
    fn take(&mut self, count: usize) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| Error::InvalidData("data-consolidation size overflow".into()))?;
        let value = self
            .data
            .get(self.offset..end)
            .ok_or_else(|| Error::InvalidRecord {
                record_type: self.record_type,
                message: "truncated data-consolidation record".into(),
            })?;
        self.offset = end;
        Ok(value)
    }
    fn u8(&mut self) -> Result<u8> {
        Ok(self.take(1)?[0])
    }
    fn u16(&mut self) -> Result<u16> {
        let b = self.take(2)?;
        Ok(u16::from_le_bytes([b[0], b[1]]))
    }
    fn string(&mut self) -> Result<String> {
        let count = usize::from(self.u16()?);
        if !(1..=255).contains(&count) {
            return invalid(self.record_type, "DConName string count must be 1..=255");
        }
        let flags = self.u8()?;
        if flags & !1 != 0 {
            return invalid(
                self.record_type,
                "DConName contains reserved string flag bits",
            );
        }
        let wide = flags & 1 != 0;
        let raw = self.take(count * if wide { 2 } else { 1 })?;
        if wide {
            let units = raw
                .chunks_exact(2)
                .map(|p| u16::from_le_bytes([p[0], p[1]]))
                .collect::<Vec<_>>();
            String::from_utf16(&units)
                .map_err(|_| Error::InvalidData("DConName contains invalid UTF-16".into()))
        } else {
            Ok(raw.iter().map(|byte| char::from(*byte)).collect())
        }
    }
    fn finish(&self) -> Result<()> {
        if self.offset != self.data.len() {
            return invalid(
                self.record_type,
                "data-consolidation record contains trailing bytes",
            );
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn malformed_dcon_and_detached_sources_are_rejected() {
        let mut c = ConsolidationCollector::new();
        assert!(c.feed_record(DCON_REF_RECORD_TYPE, &[0; 12]).is_err());
        assert!(
            c.feed_record(DCON_RECORD_TYPE, &[8, 0, 0, 0, 0, 0, 0])
                .is_err()
        );
        assert!(
            ConsolidationCollector::new()
                .feed_record(DCON_RECORD_TYPE, &[11, 0, 0, 0, 0, 0, 0, 0])
                .is_err()
        );
        assert!(
            ConsolidationCollector::new()
                .feed_record(DCON_RECORD_TYPE, &[8, 0, 2, 0, 0, 0, 0, 0])
                .is_err()
        );
    }
    #[test]
    fn adjacency_reserved_fields_and_ranges_are_strict() {
        let mut c = ConsolidationCollector::new();
        c.feed_record(DCON_RECORD_TYPE, &[8, 0, 0, 0, 0, 0, 0, 0])
            .unwrap();
        c.feed_record(0x0200, &[]).unwrap();
        assert!(c.feed_record(DCON_REF_RECORD_TYPE, &[0; 12]).is_err());
        let mut c = ConsolidationCollector::new();
        c.feed_record(DCON_RECORD_TYPE, &[8, 0, 0, 0, 0, 0, 0, 0])
            .unwrap();
        assert!(
            c.feed_record(DCON_BIN_RECORD_TYPE, &[0, 1, 0, 0, 0, 0])
                .is_err()
        );
        let mut c = ConsolidationCollector::new();
        c.feed_record(DCON_RECORD_TYPE, &[8, 0, 0, 0, 0, 0, 0, 0])
            .unwrap();
        assert!(
            c.feed_record(
                DCON_REF_RECORD_TYPE,
                &[2, 0, 1, 0, 0, 0, 2, 0, 0, 2, b'S', 0]
            )
            .is_err()
        );
    }
}
