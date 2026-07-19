//! Passive AutoText and formatted AutoCorrect metadata for glossary-only DOC files.

use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;

const STTBF_GLSY_FIB_INDEX: usize = 9;
const PLCF_GLSY_FIB_INDEX: usize = 10;
const STTB_GLSY_STYLE_FIB_INDEX: usize = 83;
const MAX_ITEM_NAME_UNITS: usize = 32;
const MAX_STYLE_USE_COUNT: u8 = 0x32;
const MAX_TABLE_BYTES: usize = 16 * 1024 * 1024;

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn decode_utf16(data: &[u8], context: &str) -> Result<String> {
    let units = data
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    String::from_utf16(&units).map_err(|_| corrupted(format!("{context} contains invalid UTF-16")))
}

/// The inert classification recorded by `LEGOXTR_V11.flego`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GlossaryItemKind {
    NamedAutoText,
    FormattedAutoCorrect,
}

impl GlossaryItemKind {
    fn from_byte(value: u8) -> Result<Self> {
        match value {
            0x00 => Ok(Self::NamedAutoText),
            0x0A => Ok(Self::FormattedAutoCorrect),
            _ => Err(corrupted(format!(
                "SttbfGlsy contains invalid flego value 0x{value:02X}"
            ))),
        }
    }

    fn as_byte(self) -> u8 {
        match self {
            Self::NamedAutoText => 0x00,
            Self::FormattedAutoCorrect => 0x0A,
        }
    }
}

/// One AutoText or formatted AutoCorrect item and its main-story CP range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GlossaryItem {
    name: String,
    kind: GlossaryItemKind,
    style_index: Option<u16>,
    start_cp: u32,
    end_cp: u32,
}

impl GlossaryItem {
    pub fn try_new(
        name: impl Into<String>,
        kind: GlossaryItemKind,
        style_index: Option<u16>,
        start_cp: u32,
        end_cp: u32,
    ) -> Result<Self> {
        let item = Self {
            name: name.into(),
            kind,
            style_index,
            start_cp,
            end_cp,
        };
        validate_item_shape(&item, 0)?;
        Ok(item)
    }

    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn kind(&self) -> GlossaryItemKind {
        self.kind
    }
    pub fn style_index(&self) -> Option<u16> {
        self.style_index
    }
    pub fn start_cp(&self) -> u32 {
        self.start_cp
    }
    pub fn end_cp(&self) -> u32 {
        self.end_cp
    }
}

/// One style-name slot parallel to the style indices in `SttbfGlsy`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GlossaryStyle {
    name: String,
    use_count: u8,
}

impl GlossaryStyle {
    pub fn try_new(name: impl Into<String>, use_count: u8) -> Result<Self> {
        let style = Self {
            name: name.into(),
            use_count,
        };
        validate_style(&style, 0)?;
        Ok(style)
    }

    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn use_count(&self) -> u8 {
        self.use_count
    }
}

/// Serialized forms of the three glossary tables, ready for FIB slots 9, 10, and 83.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GlossaryTables {
    sttbf_glsy: Vec<u8>,
    plcf_glsy: Vec<u8>,
    sttb_glsy_style: Vec<u8>,
}

impl GlossaryTables {
    pub fn item_table(&self) -> &[u8] {
        &self.sttbf_glsy
    }
    pub fn position_table(&self) -> &[u8] {
        &self.plcf_glsy
    }
    pub fn style_table(&self) -> &[u8] {
        &self.sttb_glsy_style
    }
    pub fn into_parts(self) -> (Vec<u8>, Vec<u8>, Vec<u8>) {
        (self.sttbf_glsy, self.plcf_glsy, self.sttb_glsy_style)
    }
}

/// Cross-validated metadata from `SttbfGlsy`, `PlcfGlsy`, and `SttbGlsyStyle`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GlossaryMetadata {
    items: Vec<GlossaryItem>,
    styles: Vec<GlossaryStyle>,
    terminal_cp: u32,
    ignored_cp: u32,
    main_text_length: u32,
}

impl GlossaryMetadata {
    pub fn try_new(
        items: Vec<GlossaryItem>,
        styles: Vec<GlossaryStyle>,
        terminal_cp: u32,
        ignored_cp: u32,
        main_text_length: u32,
    ) -> Result<Self> {
        let value = Self {
            items,
            styles,
            terminal_cp,
            ignored_cp,
            main_text_length,
        };
        validate_metadata(&value)?;
        Ok(value)
    }

    /// Parse the three FIB-addressed tables only for an AutoText-only document.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        let pointers = [
            (STTBF_GLSY_FIB_INDEX, "SttbfGlsy"),
            (PLCF_GLSY_FIB_INDEX, "PlcfGlsy"),
            (STTB_GLSY_STYLE_FIB_INDEX, "SttbGlsyStyle"),
        ];
        if !fib.is_glossary_document() {
            if pointers.iter().any(|(index, _)| {
                fib.get_table_pointer(*index)
                    .is_some_and(|(_, length)| length != 0)
            }) {
                return Err(corrupted(
                    "glossary table data is present while FibBase.fGlsy is clear",
                ));
            }
            return Ok(None);
        }

        let item = table_range(fib, table_stream, pointers[0])?;
        let positions = table_range(fib, table_stream, pointers[1])?;
        let styles = table_range(fib, table_stream, pointers[2])?;
        Self::parse_table_bytes(item, positions, styles, fib.get_main_doc_range().1).map(Some)
    }

    /// Parse complete raw payloads for the three parallel glossary tables.
    pub fn parse_table_bytes(
        sttbf_glsy: &[u8],
        plcf_glsy: &[u8],
        sttb_glsy_style: &[u8],
        main_text_length: u32,
    ) -> Result<Self> {
        let styles = parse_styles(sttb_glsy_style)?;
        let raw_items = parse_items(sttbf_glsy)?;
        let expected_cp_count = raw_items
            .len()
            .checked_add(2)
            .ok_or_else(|| corrupted("PlcfGlsy CP count overflows"))?;
        let expected_size = expected_cp_count
            .checked_mul(4)
            .ok_or_else(|| corrupted("PlcfGlsy byte count overflows"))?;
        if plcf_glsy.len() != expected_size {
            return Err(corrupted(format!(
                "PlcfGlsy has {} bytes; expected {expected_size}",
                plcf_glsy.len()
            )));
        }
        if plcf_glsy.len() > MAX_TABLE_BYTES {
            return Err(corrupted("PlcfGlsy exceeds the table size cap"));
        }
        let mut cps = Vec::with_capacity(expected_cp_count);
        for index in 0..expected_cp_count {
            cps.push(read_u32(plcf_glsy, index * 4, "PlcfGlsy CP")?);
        }
        let items = raw_items
            .into_iter()
            .enumerate()
            .map(|(index, raw)| GlossaryItem {
                name: raw.name,
                kind: raw.kind,
                style_index: raw.style_index,
                start_cp: cps[index],
                end_cp: cps[index + 1],
            })
            .collect();
        Self::try_new(
            items,
            styles,
            cps[expected_cp_count - 2],
            cps[expected_cp_count - 1],
            main_text_length,
        )
    }

    pub fn items(&self) -> &[GlossaryItem] {
        &self.items
    }
    pub fn styles(&self) -> &[GlossaryStyle] {
        &self.styles
    }
    pub fn terminal_cp(&self) -> u32 {
        self.terminal_cp
    }
    pub fn ignored_cp(&self) -> u32 {
        self.ignored_cp
    }
    pub fn main_text_length(&self) -> u32 {
        self.main_text_length
    }

    pub fn style_for_item(&self, index: usize) -> Option<&GlossaryStyle> {
        self.items
            .get(index)?
            .style_index
            .and_then(|style| self.styles.get(usize::from(style)))
    }

    /// Serialize the three complete table payloads deterministically.
    pub fn to_table_bytes(&self) -> Result<GlossaryTables> {
        validate_metadata(self)?;
        let item_size = sttb_size(
            self.items.iter().map(|item| item.name.as_str()),
            4,
            "SttbfGlsy",
        )?;
        let mut sttbf_glsy = Vec::with_capacity(item_size);
        write_sttb_header(&mut sttbf_glsy, self.items.len(), 4, "SttbfGlsy")?;
        for item in &self.items {
            write_string(&mut sttbf_glsy, &item.name, "SttbfGlsy")?;
            sttbf_glsy.push(item.kind.as_byte());
            sttbf_glsy.push(0);
            sttbf_glsy.extend_from_slice(&item.style_index.unwrap_or(u16::MAX).to_le_bytes());
        }

        let mut plcf_glsy = Vec::with_capacity((self.items.len() + 2) * 4);
        for item in &self.items {
            plcf_glsy.extend_from_slice(&item.start_cp.to_le_bytes());
        }
        plcf_glsy.extend_from_slice(&self.terminal_cp.to_le_bytes());
        plcf_glsy.extend_from_slice(&self.ignored_cp.to_le_bytes());

        let style_size = sttb_size(
            self.styles.iter().map(|style| style.name.as_str()),
            1,
            "SttbGlsyStyle",
        )?;
        let mut sttb_glsy_style = Vec::with_capacity(style_size);
        write_sttb_header(&mut sttb_glsy_style, self.styles.len(), 1, "SttbGlsyStyle")?;
        for style in &self.styles {
            write_string(&mut sttb_glsy_style, &style.name, "SttbGlsyStyle")?;
            sttb_glsy_style.push(style.use_count);
        }
        Ok(GlossaryTables {
            sttbf_glsy,
            plcf_glsy,
            sttb_glsy_style,
        })
    }
}

struct RawItem {
    name: String,
    kind: GlossaryItemKind,
    style_index: Option<u16>,
}

fn table_range<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    (index, name): (usize, &str),
) -> Result<&'a [u8]> {
    let (offset, length) = fib
        .get_table_pointer(index)
        .ok_or_else(|| corrupted(format!("FIB does not contain the {name} pointer")))?;
    if length == 0 {
        return Err(corrupted(format!("glossary document has no {name}")));
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset is too large")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length is too large")))?;
    if length > MAX_TABLE_BYTES {
        return Err(corrupted(format!("{name} exceeds the table size cap")));
    }
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}

fn parse_items(data: &[u8]) -> Result<Vec<RawItem>> {
    let count = parse_sttb_header(data, 4, "SttbfGlsy")?;
    let mut items = Vec::with_capacity(count);
    let mut offset = 6usize;
    for index in 0..count {
        let (name, next) = parse_string(data, offset, MAX_ITEM_NAME_UNITS, "SttbfGlsy", index)?;
        offset = next;
        let extra = data
            .get(offset..offset + 4)
            .ok_or_else(|| corrupted(format!("SttbfGlsy item {index} extra data is truncated")))?;
        let kind = GlossaryItemKind::from_byte(extra[0])?;
        let raw_style = u16::from_le_bytes([extra[2], extra[3]]);
        let style_index = if raw_style == u16::MAX {
            None
        } else if raw_style <= i16::MAX as u16 {
            Some(raw_style)
        } else {
            return Err(corrupted(format!(
                "SttbfGlsy item {index} has a negative style index"
            )));
        };
        if kind == GlossaryItemKind::FormattedAutoCorrect && style_index.is_some() {
            return Err(corrupted(format!(
                "SttbfGlsy AutoCorrect item {index} uses a style"
            )));
        }
        items.push(RawItem {
            name,
            kind,
            style_index,
        });
        offset += 4;
    }
    if offset != data.len() {
        return Err(corrupted("SttbfGlsy has trailing bytes"));
    }
    Ok(items)
}

fn parse_styles(data: &[u8]) -> Result<Vec<GlossaryStyle>> {
    let count = parse_sttb_header(data, 1, "SttbGlsyStyle")?;
    let mut styles = Vec::with_capacity(count);
    let mut offset = 6usize;
    for index in 0..count {
        let (name, next) = parse_string(data, offset, u16::MAX as usize, "SttbGlsyStyle", index)?;
        offset = next;
        let use_count = *data
            .get(offset)
            .ok_or_else(|| corrupted(format!("SttbGlsyStyle entry {index} is truncated")))?;
        styles.push(GlossaryStyle { name, use_count });
        offset += 1;
    }
    if offset != data.len() {
        return Err(corrupted("SttbGlsyStyle has trailing bytes"));
    }
    Ok(styles)
}

fn parse_sttb_header(data: &[u8], extra: u16, name: &str) -> Result<usize> {
    if data.len() > MAX_TABLE_BYTES {
        return Err(corrupted(format!("{name} exceeds the table size cap")));
    }
    if data.len() < 6
        || read_u16(data, 0, &format!("{name} fExtend"))? != u16::MAX
        || read_u16(data, 4, &format!("{name} cbExtra"))? != extra
    {
        return Err(corrupted(format!("{name} has an invalid header")));
    }
    Ok(usize::from(read_u16(data, 2, &format!("{name} cData"))?))
}

fn parse_string(
    data: &[u8],
    offset: usize,
    max_units: usize,
    table: &str,
    index: usize,
) -> Result<(String, usize)> {
    let units = usize::from(read_u16(
        data,
        offset,
        &format!("{table} string {index} length"),
    )?);
    if units > max_units {
        return Err(corrupted(format!(
            "{table} string {index} exceeds {max_units} UTF-16 code units"
        )));
    }
    let start = offset
        .checked_add(2)
        .ok_or_else(|| corrupted(format!("{table} string offset overflows")))?;
    let end = start
        .checked_add(
            units
                .checked_mul(2)
                .ok_or_else(|| corrupted(format!("{table} string size overflows")))?,
        )
        .ok_or_else(|| corrupted(format!("{table} string range overflows")))?;
    let bytes = data
        .get(start..end)
        .ok_or_else(|| corrupted(format!("{table} string {index} is truncated")))?;
    Ok((
        decode_utf16(bytes, &format!("{table} string {index}"))?,
        end,
    ))
}

fn validate_item_shape(item: &GlossaryItem, index: usize) -> Result<()> {
    if item.name.encode_utf16().count() > MAX_ITEM_NAME_UNITS {
        return Err(corrupted(format!(
            "glossary item {index} name exceeds 32 UTF-16 code units"
        )));
    }
    if item.start_cp >= item.end_cp {
        return Err(corrupted(format!(
            "glossary item {index} has an empty or reversed CP range"
        )));
    }
    if item.kind == GlossaryItemKind::FormattedAutoCorrect && item.style_index.is_some() {
        return Err(corrupted(format!(
            "formatted AutoCorrect item {index} cannot use a style"
        )));
    }
    if item
        .style_index
        .is_some_and(|value| value > i16::MAX as u16)
    {
        return Err(corrupted(format!(
            "glossary item {index} style index exceeds the signed 16-bit range"
        )));
    }
    Ok(())
}

fn validate_style(style: &GlossaryStyle, index: usize) -> Result<()> {
    if style.name.encode_utf16().count() > u16::MAX as usize {
        return Err(corrupted(format!(
            "glossary style {index} name exceeds 65535 UTF-16 code units"
        )));
    }
    if style.use_count > MAX_STYLE_USE_COUNT {
        return Err(corrupted(format!(
            "glossary style {index} use count exceeds 0x32"
        )));
    }
    Ok(())
}

fn validate_metadata(value: &GlossaryMetadata) -> Result<()> {
    if value.items.len() > u16::MAX as usize || value.styles.len() > u16::MAX as usize {
        return Err(corrupted("glossary STTB count exceeds 65535 entries"));
    }
    if value.terminal_cp >= value.ignored_cp || value.ignored_cp >= value.main_text_length {
        return Err(corrupted(
            "PlcfGlsy terminal CPs are not strictly increasing within ccpText",
        ));
    }
    let mut actual_uses = vec![0u8; value.styles.len()];
    for (index, item) in value.items.iter().enumerate() {
        validate_item_shape(item, index)?;
        let expected_end = value
            .items
            .get(index + 1)
            .map_or(value.terminal_cp, |next| next.start_cp);
        if item.end_cp != expected_end {
            return Err(corrupted(format!(
                "glossary item {index} range is not contiguous with PlcfGlsy"
            )));
        }
        if item.end_cp >= value.main_text_length {
            return Err(corrupted(format!(
                "glossary item {index} range is outside ccpText"
            )));
        }
        if let Some(style_index) = item.style_index {
            let count = actual_uses
                .get_mut(usize::from(style_index))
                .ok_or_else(|| {
                    corrupted(format!("glossary item {index} style index is out of range"))
                })?;
            *count = count
                .checked_add(1)
                .ok_or_else(|| corrupted("glossary style use count overflows"))?;
        }
    }
    for (index, (style, actual)) in value.styles.iter().zip(actual_uses).enumerate() {
        validate_style(style, index)?;
        if style.use_count != actual {
            return Err(corrupted(format!(
                "glossary style {index} records {} uses but {actual} items refer to it",
                style.use_count
            )));
        }
    }
    Ok(())
}

fn sttb_size<'a>(
    strings: impl Iterator<Item = &'a str>,
    extra: usize,
    name: &str,
) -> Result<usize> {
    let mut size = 6usize;
    for string in strings {
        size = size
            .checked_add(2)
            .and_then(|value| value.checked_add(string.encode_utf16().count().checked_mul(2)?))
            .and_then(|value| value.checked_add(extra))
            .ok_or_else(|| corrupted(format!("{name} serialized size overflows")))?;
    }
    if size > MAX_TABLE_BYTES {
        return Err(corrupted(format!("{name} exceeds the table size cap")));
    }
    Ok(size)
}

fn write_sttb_header(data: &mut Vec<u8>, count: usize, extra: u16, name: &str) -> Result<()> {
    let count = u16::try_from(count)
        .map_err(|_| corrupted(format!("{name} contains more than 65535 strings")))?;
    data.extend_from_slice(&u16::MAX.to_le_bytes());
    data.extend_from_slice(&count.to_le_bytes());
    data.extend_from_slice(&extra.to_le_bytes());
    Ok(())
}

fn write_string(data: &mut Vec<u8>, value: &str, table: &str) -> Result<()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    let count = u16::try_from(units.len())
        .map_err(|_| corrupted(format!("{table} string length exceeds u16")))?;
    data.extend_from_slice(&count.to_le_bytes());
    data.extend(units.into_iter().flat_map(u16::to_le_bytes));
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> GlossaryMetadata {
        GlossaryMetadata::try_new(
            vec![
                GlossaryItem::try_new("Greeting", GlossaryItemKind::NamedAutoText, Some(0), 1, 5)
                    .unwrap(),
                GlossaryItem::try_new("teh", GlossaryItemKind::FormattedAutoCorrect, None, 5, 9)
                    .unwrap(),
            ],
            vec![GlossaryStyle::try_new("Normal", 1).unwrap()],
            9,
            10,
            11,
        )
        .unwrap()
    }

    #[test]
    fn round_trips_all_three_parallel_tables() {
        let metadata = sample();
        let tables = metadata.to_table_bytes().unwrap();
        let parsed = GlossaryMetadata::parse_table_bytes(
            tables.item_table(),
            tables.position_table(),
            tables.style_table(),
            11,
        )
        .unwrap();
        assert_eq!(parsed, metadata);
        assert_eq!(parsed.style_for_item(0).unwrap().name(), "Normal");
        assert!(parsed.style_for_item(1).is_none());
    }

    #[test]
    fn rejects_cross_table_and_lexical_inconsistencies() {
        let metadata = sample();
        let tables = metadata.to_table_bytes().unwrap();

        let mut items = tables.item_table().to_vec();
        let first_extra = 6 + 2 + "Greeting".encode_utf16().count() * 2;
        items[first_extra] = 0x09;
        assert!(
            GlossaryMetadata::parse_table_bytes(
                &items,
                tables.position_table(),
                tables.style_table(),
                11,
            )
            .is_err()
        );

        let mut positions = tables.position_table().to_vec();
        positions[4..8].copy_from_slice(&1u32.to_le_bytes());
        assert!(
            GlossaryMetadata::parse_table_bytes(
                tables.item_table(),
                &positions,
                tables.style_table(),
                11,
            )
            .is_err()
        );

        let wrong_count = vec![GlossaryStyle::try_new("Normal", 0).unwrap()];
        assert!(
            GlossaryMetadata::try_new(metadata.items.clone(), wrong_count, 9, 10, 11,).is_err()
        );
    }

    #[test]
    fn rejects_autocorrect_style_and_resource_limit_violations() {
        assert!(
            GlossaryItem::try_new("bad", GlossaryItemKind::FormattedAutoCorrect, Some(0), 0, 1,)
                .is_err()
        );
        assert!(
            GlossaryItem::try_new("x".repeat(33), GlossaryItemKind::NamedAutoText, None, 0, 1,)
                .is_err()
        );
        assert!(GlossaryStyle::try_new("Normal", 0x33).is_err());
    }
}
