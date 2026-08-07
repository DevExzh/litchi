//! Bounded, inert PowerPoint EOT parsing and publication.
//!
//! This module only validates and moves font bytes. It never installs, loads,
//! renders, decompresses, decrypts, or executes font programs.

use super::Prepared;
use crate::{FontError, License, Permission};

type Result<T> = std::result::Result<T, FontError>;

const VERSION_1: u32 = 0x0001_0000;
const MAGIC: u16 = 0x504C;
const FIXED_BYTES: usize = 82;
const SUBSET: u32 = 0x0000_0001;

fn invalid(message: impl Into<String>) -> FontError {
    FontError::EmbeddingFailed(message.into())
}

fn limit(resource: &'static str, actual: usize, maximum: usize) -> Result<()> {
    if actual > maximum {
        return Err(FontError::LimitExceeded {
            resource,
            limit: maximum,
            actual,
        });
    }
    Ok(())
}

/// Intended use of a document containing the embedded font.
///
/// This is deliberately distinct from the font's permission. A caller must
/// state whether recipients may edit the document; preview/print permission is
/// never silently promoted to editable use.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Intent {
    PreviewPrint,
    Editable,
}

impl Intent {
    const fn description(self) -> &'static str {
        match self {
            Self::PreviewPrint => "preview/print use",
            Self::Editable => "editable use",
        }
    }
}

/// Bounds used by the borrowed parser and canonical encoder.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    pub max_input_bytes: usize,
    pub max_output_bytes: usize,
    pub max_font_bytes: usize,
    pub max_sfnt_tables: usize,
    pub max_name_records: usize,
    pub max_name_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_input_bytes: 64 * 1024 * 1024,
            max_output_bytes: 65 * 1024 * 1024,
            max_font_bytes: 64 * 1024 * 1024,
            max_sfnt_tables: 4_096,
            max_name_records: 4_096,
            max_name_bytes: 4 * usize::from(u16::MAX),
        }
    }
}

/// One borrowed UTF-16LE EOT name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Name<'a> {
    bytes: &'a [u8],
}

impl<'a> Name<'a> {
    pub const fn bytes(self) -> &'a [u8] {
        self.bytes
    }

    pub fn units(self) -> impl Iterator<Item = u16> + 'a {
        self.bytes
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
    }

    pub fn decode(self) -> Result<String> {
        decode_utf16(self.units(), "EOT name")
    }
}

/// A validated, allocation-free view of canonical uncompressed EOT 1.0 data.
#[derive(Debug, Clone, Copy)]
pub struct View<'a> {
    bytes: &'a [u8],
    font: &'a [u8],
    names: [Name<'a>; 4],
    flags: u32,
    license: License,
    charset: u8,
    italic: bool,
    weight: u32,
    panose: &'a [u8; 10],
}

impl<'a> View<'a> {
    pub fn parse(bytes: &'a [u8]) -> Result<Self> {
        Self::parse_with(bytes, Limits::default())
    }

    pub fn parse_with(bytes: &'a [u8], limits: Limits) -> Result<Self> {
        limit("EOT input bytes", bytes.len(), limits.max_input_bytes)?;
        if bytes.len() < FIXED_BYTES {
            return Err(invalid("truncated EOT 1.0 fixed header"));
        }

        let declared_size = usize::try_from(le_u32(bytes, 0, "EOT size")?)
            .map_err(|_| invalid("EOT size does not fit this platform"))?;
        if declared_size != bytes.len() {
            return Err(invalid("EOT size does not match the input length"));
        }
        let font_size = usize::try_from(le_u32(bytes, 4, "EOT font size")?)
            .map_err(|_| invalid("EOT font size does not fit this platform"))?;
        limit("EOT font bytes", font_size, limits.max_font_bytes)?;
        if le_u32(bytes, 8, "EOT version")? != VERSION_1 {
            return Err(invalid(
                "unsupported EOT version; PowerPoint authoring requires 1.0",
            ));
        }
        let flags = le_u32(bytes, 12, "EOT flags")?;
        if flags & !SUBSET != 0 {
            return Err(invalid("unsupported EOT processing flags"));
        }
        if le_u16(bytes, 34, "EOT magic")? != MAGIC {
            return Err(invalid("invalid EOT magic number"));
        }
        for offset in [64, 68, 72, 76] {
            if le_u32(bytes, offset, "EOT reserved field")? != 0 {
                return Err(invalid("nonzero EOT reserved field"));
            }
        }
        if le_u16(bytes, 80, "EOT padding")? != 0 {
            return Err(invalid("nonzero EOT fixed-header padding"));
        }

        let mut cursor = FIXED_BYTES;
        let mut total_name_bytes = 0usize;
        let mut names = [Name { bytes: &[] }; 4];
        for (index, slot) in names.iter_mut().enumerate() {
            let name_size = usize::from(le_u16(bytes, cursor, "EOT name size")?);
            cursor = cursor
                .checked_add(2)
                .ok_or_else(|| invalid("EOT name offset overflow"))?;
            if name_size % 2 != 0 {
                return Err(invalid("EOT UTF-16 name has an odd byte length"));
            }
            total_name_bytes = total_name_bytes
                .checked_add(name_size)
                .ok_or_else(|| invalid("EOT name size overflow"))?;
            limit("EOT name bytes", total_name_bytes, limits.max_name_bytes)?;
            let end = cursor
                .checked_add(name_size)
                .ok_or_else(|| invalid("EOT name range overflow"))?;
            let name = bytes
                .get(cursor..end)
                .ok_or_else(|| invalid("truncated EOT name"))?;
            // Validate UTF-16 without allocating while retaining the borrowed view.
            if char::decode_utf16(
                name.chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]])),
            )
            .any(|value| value.is_err())
            {
                return Err(invalid("EOT name contains malformed UTF-16"));
            }
            *slot = Name { bytes: name };
            cursor = end;
            if index != 3 {
                if le_u16(bytes, cursor, "EOT name padding")? != 0 {
                    return Err(invalid("nonzero EOT name padding"));
                }
                cursor = cursor
                    .checked_add(2)
                    .ok_or_else(|| invalid("EOT padding offset overflow"))?;
            }
        }

        let end = cursor
            .checked_add(font_size)
            .ok_or_else(|| invalid("EOT font range overflow"))?;
        if end != bytes.len() {
            return Err(invalid("EOT FontDataSize does not match trailing data"));
        }
        let font = bytes
            .get(cursor..end)
            .ok_or_else(|| invalid("truncated EOT font data"))?;
        let sfnt = Sfnt::parse(font, limits)?;
        let os2 = Os2::parse(
            sfnt.table(*b"OS/2")?
                .ok_or_else(|| invalid("OpenType font has no OS/2 table"))?,
        )?;
        let actual_license = os2.license()?;
        let header_license = License::new(le_u16(bytes, 32, "EOT fsType")?)?;
        if header_license != actual_license {
            return Err(invalid("EOT fsType does not match embedded OS/2.fsType"));
        }
        validate_outline_license("embedded EOT font", actual_license)?;

        let panose = bytes[16..26]
            .try_into()
            .expect("validated fixed EOT header");
        Ok(Self {
            bytes,
            font,
            names,
            flags,
            license: actual_license,
            charset: bytes[26],
            italic: match bytes[27] {
                0 => false,
                1 => true,
                _ => return Err(invalid("EOT italic field is not zero or one")),
            },
            weight: le_u32(bytes, 28, "EOT weight")?,
            panose,
        })
    }

    pub const fn bytes(self) -> &'a [u8] {
        self.bytes
    }

    pub const fn font_data(self) -> &'a [u8] {
        self.font
    }

    pub const fn family_name(self) -> Name<'a> {
        self.names[0]
    }

    pub const fn style_name(self) -> Name<'a> {
        self.names[1]
    }

    pub const fn version_name(self) -> Name<'a> {
        self.names[2]
    }

    pub const fn full_name(self) -> Name<'a> {
        self.names[3]
    }

    pub const fn license(self) -> License {
        self.license
    }

    pub const fn subsetted(self) -> bool {
        self.flags & SUBSET != 0
    }

    pub const fn charset(self) -> u8 {
        self.charset
    }

    pub const fn italic(self) -> bool {
        self.italic
    }

    pub const fn weight(self) -> u32 {
        self.weight
    }

    pub const fn panose(self) -> &'a [u8; 10] {
        self.panose
    }
}

/// Compatibility wrapper with the least-privileged document intent.
///
/// New integrations should call [`encode`] and state their intent explicitly.
pub fn data(font: &mut Prepared) -> Result<Vec<u8>> {
    encode(font, Intent::PreviewPrint, Limits::default())
}

/// Restore a successfully encoded program after a later publication failure.
///
/// The EOT allocation is reused in place. This function refuses to overwrite a
/// nonempty `Prepared` program or restore bytes whose license/subset metadata no
/// longer matches the prepared face.
pub fn restore(font: &mut Prepared, eot: Vec<u8>) -> Result<()> {
    restore_with(font, eot, Limits::default())
}

/// Restore with the same explicit bounds used for encoding.
pub fn restore_with(font: &mut Prepared, mut eot: Vec<u8>, limits: Limits) -> Result<()> {
    if !font.data.is_empty() {
        return Err(invalid("cannot restore EOT over a nonempty prepared font"));
    }
    let view = View::parse_with(&eot, limits)?;
    if view.license() != font.properties.license() {
        return Err(FontError::LicenseMismatch {
            name: font.name.clone(),
            declared: font.properties.license().bits(),
            actual: view.license().bits(),
        });
    }
    if view.subsetted() != font.subsetted {
        return Err(invalid("EOT subset flag does not match the prepared font"));
    }
    let font_offset = view.font_data().as_ptr() as usize - eot.as_ptr() as usize;
    let font_size = view.font_data().len();
    eot.copy_within(font_offset..font_offset + font_size, 0);
    eot.truncate(font_size);
    font.data = eot;
    Ok(())
}

/// Canonically wrap one standalone, uncompressed OpenType face in EOT 1.0.
///
/// The actual `OS/2.fsType` is authoritative. Resolver metadata must agree
/// exactly, restricted or bitmap-only fonts are refused, and preview/print
/// permission cannot satisfy editable intent. The source program is moved only
/// after all validation and allocation succeeds.
pub fn encode(font: &mut Prepared, intent: Intent, limits: Limits) -> Result<Vec<u8>> {
    limit(
        "OpenType input bytes",
        font.data.len(),
        limits.max_input_bytes,
    )?;
    limit(
        "OpenType font bytes",
        font.data.len(),
        limits.max_font_bytes,
    )?;
    let fallback_name_bytes = utf16_bytes(&font.name)?;
    limit(
        "EOT fallback name bytes",
        fallback_name_bytes,
        limits.max_name_bytes,
    )?;
    let sfnt = Sfnt::parse(&font.data, limits)?;
    let os2 = Os2::parse(
        sfnt.table(*b"OS/2")?
            .ok_or_else(|| invalid("OpenType font has no OS/2 table"))?,
    )?;
    let license = os2.license()?;
    let declared = font.properties.license();
    if declared != license {
        return Err(FontError::LicenseMismatch {
            name: font.name.clone(),
            declared: declared.bits(),
            actual: license.bits(),
        });
    }
    validate_outline_license(&font.name, license)?;
    if font.subsetted && !license.may_subset() {
        return Err(invalid(
            "font is marked as subsetted but OS/2.fsType forbids subsetting",
        ));
    }
    if intent == Intent::Editable && license.permission() == Permission::PreviewPrint {
        return Err(FontError::EmbeddingUseForbidden {
            name: font.name.clone(),
            intent: intent.description(),
            permission: license.permission(),
        });
    }

    let head = sfnt
        .table(*b"head")?
        .ok_or_else(|| invalid("OpenType font has no head table"))?;
    let check_sum_adjustment = be_u32(head, 8, "head checkSumAdjustment")?;
    let name_table = sfnt
        .table(*b"name")?
        .ok_or_else(|| invalid("OpenType font has no name table"))?;
    let mut decoded_name_bytes = 0usize;
    let family = name_string(name_table, 1, limits, &mut decoded_name_bytes)?
        .unwrap_or_else(|| font.name.clone());
    let style = name_string(name_table, 2, limits, &mut decoded_name_bytes)?
        .unwrap_or_else(|| "Regular".into());
    let version = name_string(name_table, 5, limits, &mut decoded_name_bytes)?.unwrap_or_default();
    let full = name_string(name_table, 4, limits, &mut decoded_name_bytes)?
        .unwrap_or_else(|| font.name.clone());
    let names = [&family, &style, &version, &full];

    let mut total_name_bytes = 0usize;
    let name_bytes = names
        .iter()
        .enumerate()
        .try_fold(0usize, |total, (index, value)| {
            let bytes = utf16_bytes(value)?;
            total_name_bytes = total_name_bytes
                .checked_add(bytes)
                .ok_or_else(|| invalid("EOT name data is too large"))?;
            limit("EOT name bytes", total_name_bytes, limits.max_name_bytes)?;
            let overhead = if index + 1 == names.len() { 2 } else { 4 };
            total
                .checked_add(overhead)
                .and_then(|value| value.checked_add(bytes))
                .ok_or_else(|| invalid("EOT name data is too large"))
        })?;
    let output_size = FIXED_BYTES
        .checked_add(name_bytes)
        .and_then(|size| size.checked_add(font.data.len()))
        .ok_or_else(|| invalid("EOT payload size overflow"))?;
    limit("EOT output bytes", output_size, limits.max_output_bytes)?;
    let eot_size =
        u32::try_from(output_size).map_err(|_| invalid("EOT payload exceeds the 32-bit format"))?;
    let font_size = u32::try_from(font.data.len())
        .map_err(|_| invalid("OpenType payload exceeds the 32-bit EOT format"))?;
    let header_size = output_size
        .checked_sub(font.data.len())
        .ok_or_else(|| invalid("EOT header size underflow"))?;

    let mut header = Vec::new();
    header
        .try_reserve_exact(header_size)
        .map_err(|source| FontError::Allocation {
            resource: "PowerPoint EOT header",
            source,
        })?;
    push_u32(&mut header, eot_size);
    push_u32(&mut header, font_size);
    push_u32(&mut header, VERSION_1);
    push_u32(&mut header, if font.subsetted { SUBSET } else { 0 });
    header.extend_from_slice(os2.panose());
    header.push(font.properties.charset().map_or(1, crate::Charset::code));
    header.push(u8::from(os2.italic()));
    push_u32(&mut header, u32::from(os2.weight()));
    push_u16(&mut header, license.bits());
    push_u16(&mut header, MAGIC);
    for range in os2.unicode_ranges() {
        push_u32(&mut header, range);
    }
    let (code_page1, code_page2) = os2.code_pages();
    push_u32(&mut header, code_page1);
    push_u32(&mut header, code_page2);
    push_u32(&mut header, check_sum_adjustment);
    for _ in 0..4 {
        push_u32(&mut header, 0);
    }
    push_u16(&mut header, 0);
    for (index, value) in names.into_iter().enumerate() {
        push_utf16(&mut header, value)?;
        if index + 1 != 4 {
            push_u16(&mut header, 0);
        }
    }
    if header.len() != header_size {
        return Err(invalid("constructed EOT header has an inconsistent size"));
    }

    font.data
        .try_reserve_exact(header_size)
        .map_err(|source| FontError::Allocation {
            resource: "PowerPoint EOT payload",
            source,
        })?;
    let mut output = std::mem::take(&mut font.data);
    let source_size = output.len();
    output.resize(output_size, 0);
    output.copy_within(0..source_size, header_size);
    output[..header_size].copy_from_slice(&header);
    Ok(output)
}

fn validate_outline_license(name: &str, license: License) -> Result<()> {
    if license.permission() == Permission::Restricted {
        return Err(FontError::EmbeddingForbidden {
            name: name.to_owned(),
        });
    }
    if license
        .restrictions()
        .contains(crate::Restrictions::BITMAP_ONLY)
    {
        return Err(FontError::BitmapOnly {
            name: name.to_owned(),
        });
    }
    Ok(())
}

struct Sfnt<'a> {
    data: &'a [u8],
    table_count: usize,
}

impl<'a> Sfnt<'a> {
    fn parse(data: &'a [u8], limits: Limits) -> Result<Self> {
        limit("sfnt input bytes", data.len(), limits.max_font_bytes)?;
        let signature = data
            .get(..4)
            .ok_or_else(|| invalid("OpenType font is missing an sfnt signature"))?;
        if signature == b"ttcf" {
            return Err(FontError::RequiresStandaloneFace);
        }
        if !matches!(signature, b"\0\x01\0\0" | b"OTTO" | b"true" | b"typ1") {
            return Err(invalid("invalid standalone OpenType sfnt signature"));
        }
        let table_count = usize::from(be_u16(data, 4, "sfnt table count")?);
        limit("sfnt table count", table_count, limits.max_sfnt_tables)?;
        let directory_len = table_count
            .checked_mul(16)
            .and_then(|length| 12usize.checked_add(length))
            .ok_or_else(|| invalid("sfnt table directory overflows"))?;
        if data.len() < directory_len {
            return Err(invalid("truncated sfnt table directory"));
        }
        for index in 0..table_count {
            let record = 12 + index * 16;
            let offset = usize::try_from(be_u32(data, record + 8, "sfnt table offset")?)
                .map_err(|_| invalid("sfnt table offset does not fit this platform"))?;
            let length = usize::try_from(be_u32(data, record + 12, "sfnt table length")?)
                .map_err(|_| invalid("sfnt table length does not fit this platform"))?;
            let end = offset
                .checked_add(length)
                .ok_or_else(|| invalid("sfnt table range overflows"))?;
            if offset < directory_len || end > data.len() {
                return Err(invalid("sfnt table range is outside the font program"));
            }
        }
        Ok(Self { data, table_count })
    }

    fn table(&self, wanted: [u8; 4]) -> Result<Option<&'a [u8]>> {
        let mut found = None;
        for index in 0..self.table_count {
            let record = 12 + index * 16;
            if self.data[record..record + 4] != wanted {
                continue;
            }
            if found.is_some() {
                return Err(invalid(format!(
                    "duplicate sfnt table {}",
                    String::from_utf8_lossy(&wanted)
                )));
            }
            let offset = usize::try_from(be_u32(self.data, record + 8, "sfnt table offset")?)
                .map_err(|_| invalid("sfnt table offset does not fit this platform"))?;
            let length = usize::try_from(be_u32(self.data, record + 12, "sfnt table length")?)
                .map_err(|_| invalid("sfnt table length does not fit this platform"))?;
            let end = offset
                .checked_add(length)
                .ok_or_else(|| invalid("sfnt table range overflows"))?;
            found = Some(&self.data[offset..end]);
        }
        Ok(found)
    }
}

struct Os2<'a> {
    bytes: &'a [u8],
    version: u16,
}

impl<'a> Os2<'a> {
    fn parse(bytes: &'a [u8]) -> Result<Self> {
        let version = be_u16(bytes, 0, "OS/2 version")?;
        let minimum = match version {
            0 => 78,
            1 => 86,
            2..=4 => 96,
            5 => 100,
            _ => return Err(FontError::UnsupportedOs2Version(version)),
        };
        if bytes.len() < minimum {
            return Err(FontError::TruncatedOs2 {
                version,
                expected: minimum,
                actual: bytes.len(),
            });
        }
        Ok(Self { bytes, version })
    }

    fn license(&self) -> Result<License> {
        // PowerPoint authoring is intentionally strict even for legacy OS/2
        // tables: ambiguous permissions and reserved bits are not copied into a
        // newly authored EOT container.
        Ok(License::new(be_u16(self.bytes, 8, "OS/2.fsType")?)?)
    }

    fn panose(&self) -> &'a [u8] {
        &self.bytes[32..42]
    }

    fn italic(&self) -> bool {
        u16::from_be_bytes(
            self.bytes[62..64]
                .try_into()
                .expect("validated OS/2 length"),
        ) & 1
            != 0
    }

    fn weight(&self) -> u16 {
        u16::from_be_bytes(self.bytes[4..6].try_into().expect("validated OS/2 length"))
    }

    fn unicode_ranges(&self) -> [u32; 4] {
        [
            u32::from_be_bytes(
                self.bytes[42..46]
                    .try_into()
                    .expect("validated OS/2 length"),
            ),
            u32::from_be_bytes(
                self.bytes[46..50]
                    .try_into()
                    .expect("validated OS/2 length"),
            ),
            u32::from_be_bytes(
                self.bytes[50..54]
                    .try_into()
                    .expect("validated OS/2 length"),
            ),
            u32::from_be_bytes(
                self.bytes[54..58]
                    .try_into()
                    .expect("validated OS/2 length"),
            ),
        ]
    }

    fn code_pages(&self) -> (u32, u32) {
        if self.version == 0 {
            return (0, 0);
        }
        (
            u32::from_be_bytes(
                self.bytes[78..82]
                    .try_into()
                    .expect("validated OS/2 length"),
            ),
            u32::from_be_bytes(
                self.bytes[82..86]
                    .try_into()
                    .expect("validated OS/2 length"),
            ),
        )
    }
}

fn name_string(
    table: &[u8],
    wanted: u16,
    limits: Limits,
    decoded_bytes: &mut usize,
) -> Result<Option<String>> {
    let count = usize::from(be_u16(table, 2, "name record count")?);
    limit("sfnt name records", count, limits.max_name_records)?;
    let strings = usize::from(be_u16(table, 4, "name string offset")?);
    let records_end = count
        .checked_mul(12)
        .and_then(|size| 6usize.checked_add(size))
        .ok_or_else(|| invalid("name table record range overflows"))?;
    if records_end > table.len() || strings > table.len() {
        return Err(invalid("truncated name table"));
    }

    let mut selected: Option<(u8, String)> = None;
    for index in 0..count {
        let record = 6 + index * 12;
        let platform = be_u16(table, record, "name platform")?;
        let encoding = be_u16(table, record + 2, "name encoding")?;
        let language = be_u16(table, record + 4, "name language")?;
        if be_u16(table, record + 6, "name id")? != wanted {
            continue;
        }
        let rank = match (platform, encoding, language) {
            (3, 0..=10, 0x0409) => 0,
            (0, _, _) => 1,
            (3, 0..=10, _) => 2,
            _ => 3,
        };
        // An equal or lower-priority alias cannot change the selected value.
        // Avoid repeatedly decoding attacker-controlled overlapping records.
        if selected.as_ref().is_some_and(|(best, _)| rank >= *best) {
            continue;
        }
        let length = usize::from(be_u16(table, record + 8, "name string length")?);
        limit("one sfnt name", length, limits.max_name_bytes)?;
        *decoded_bytes = decoded_bytes
            .checked_add(length)
            .ok_or(FontError::LimitExceeded {
                resource: "decoded sfnt name bytes",
                limit: limits.max_name_bytes,
                actual: usize::MAX,
            })?;
        limit(
            "decoded sfnt name bytes",
            *decoded_bytes,
            limits.max_name_bytes,
        )?;
        let offset = usize::from(be_u16(table, record + 10, "name string offset")?);
        let start = strings
            .checked_add(offset)
            .ok_or_else(|| invalid("name string offset overflows"))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| invalid("name string range overflows"))?;
        let value = table
            .get(start..end)
            .ok_or_else(|| invalid("truncated name string"))?;
        let decoded = if matches!(platform, 0 | 3) {
            if value.len() % 2 != 0 {
                return Err(invalid("sfnt UTF-16 name has an odd byte length"));
            }
            decode_utf16(
                value
                    .chunks_exact(2)
                    .map(|pair| u16::from_be_bytes([pair[0], pair[1]])),
                "sfnt name",
            )?
        } else {
            std::str::from_utf8(value)
                .map_err(|_| invalid("sfnt name is neither valid UTF-16 nor UTF-8"))?
                .to_owned()
        };
        if selected.as_ref().is_none_or(|(best, _)| rank < *best) {
            selected = Some((rank, decoded));
        }
    }
    Ok(selected.map(|(_, value)| value))
}

fn decode_utf16(units: impl Iterator<Item = u16>, resource: &'static str) -> Result<String> {
    let (minimum, maximum) = units.size_hint();
    let capacity = maximum
        .unwrap_or(minimum)
        .checked_mul(3)
        .ok_or_else(|| invalid(format!("{resource} allocation size overflows")))?;
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| FontError::Allocation { resource, source })?;
    for value in char::decode_utf16(units) {
        output.push(value.map_err(|_| invalid(format!("{resource} contains malformed UTF-16")))?);
    }
    Ok(output)
}

fn be_u16(bytes: &[u8], offset: usize, field: &str) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| invalid(format!("{field} offset overflows")))?;
    let value = bytes
        .get(offset..end)
        .ok_or_else(|| invalid(format!("truncated {field}")))?;
    Ok(u16::from_be_bytes(
        value.try_into().expect("two-byte slice"),
    ))
}

fn be_u32(bytes: &[u8], offset: usize, field: &str) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| invalid(format!("{field} offset overflows")))?;
    let value = bytes
        .get(offset..end)
        .ok_or_else(|| invalid(format!("truncated {field}")))?;
    Ok(u32::from_be_bytes(
        value.try_into().expect("four-byte slice"),
    ))
}

fn le_u16(bytes: &[u8], offset: usize, field: &str) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| invalid(format!("{field} offset overflows")))?;
    let value = bytes
        .get(offset..end)
        .ok_or_else(|| invalid(format!("truncated {field}")))?;
    Ok(u16::from_le_bytes(
        value.try_into().expect("two-byte slice"),
    ))
}

fn le_u32(bytes: &[u8], offset: usize, field: &str) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| invalid(format!("{field} offset overflows")))?;
    let value = bytes
        .get(offset..end)
        .ok_or_else(|| invalid(format!("truncated {field}")))?;
    Ok(u32::from_le_bytes(
        value.try_into().expect("four-byte slice"),
    ))
}

fn utf16_bytes(value: &str) -> Result<usize> {
    let units = value.encode_utf16().count();
    let bytes = units
        .checked_mul(2)
        .ok_or_else(|| invalid("EOT name length overflow"))?;
    u16::try_from(bytes)
        .map(|_| bytes)
        .map_err(|_| invalid("EOT name exceeds 65535 bytes"))
}

fn push_utf16(output: &mut Vec<u8>, value: &str) -> Result<()> {
    let bytes = utf16_bytes(value)?;
    push_u16(
        output,
        u16::try_from(bytes).map_err(|_| invalid("EOT name exceeds 65535 bytes"))?,
    );
    for unit in value.encode_utf16() {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}

fn push_u16(output: &mut Vec<u8>, value: u16) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn push_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn canonical_eot_is_borrowed_and_round_trips_structural_fields() {
        let sfnt = test_sfnt(0x0008);
        let mut font = prepared(sfnt.clone(), 0x0008, true);
        let eot = encode(&mut font, Intent::Editable, Limits::default()).unwrap();
        assert!(font.data.is_empty());

        let view = View::parse(&eot).unwrap();
        assert_eq!(view.font_data(), sfnt);
        assert_eq!(view.family_name().decode().unwrap(), "Litchi Test");
        assert_eq!(view.style_name().decode().unwrap(), "Regular");
        assert_eq!(view.license().bits(), 0x0008);
        assert!(view.subsetted());
        assert_eq!(view.weight(), 400);
        assert!(std::ptr::eq(
            view.font_data().as_ptr(),
            eot[view.font_data().as_ptr() as usize - eot.as_ptr() as usize..].as_ptr()
        ));
    }

    #[test]
    fn actual_fstype_is_authoritative_and_intent_is_explicit() {
        let mut forged = prepared(test_sfnt(0x0004), 0x0008, false);
        assert!(matches!(
            encode(&mut forged, Intent::PreviewPrint, Limits::default()),
            Err(FontError::LicenseMismatch {
                declared: 8,
                actual: 4,
                ..
            })
        ));
        assert!(!forged.data.is_empty());

        let mut preview = prepared(test_sfnt(0x0004), 0x0004, false);
        assert!(matches!(
            encode(&mut preview, Intent::Editable, Limits::default()),
            Err(FontError::EmbeddingUseForbidden { .. })
        ));
        assert!(!preview.data.is_empty());
        assert!(encode(&mut preview, Intent::PreviewPrint, Limits::default()).is_ok());
    }

    #[test]
    fn unsafe_and_ambiguous_licenses_are_rejected_from_font_bytes() {
        for bits in [0x0002, 0x0208, 0x8008, 0x000C] {
            let declared = License::new(bits).unwrap_or_else(|_| License::new(0).unwrap());
            let mut font = prepared_with_license(test_sfnt(bits), declared);
            assert!(encode(&mut font, Intent::PreviewPrint, Limits::default()).is_err());
            assert!(!font.data.is_empty());
        }
    }

    #[test]
    fn limits_fail_before_source_program_is_moved() {
        let sfnt = test_sfnt(0);
        let mut font = prepared(sfnt, 0, false);
        let mut limits = Limits::default();
        limits.max_output_bytes = 16;
        assert!(matches!(
            encode(&mut font, Intent::PreviewPrint, limits),
            Err(FontError::LimitExceeded {
                resource: "EOT output bytes",
                ..
            })
        ));
        assert!(!font.data.is_empty());
    }

    #[test]
    fn structural_limits_accept_exact_boundaries_and_reject_one_less() {
        let sfnt = test_sfnt(0);
        let font_size = sfnt.len();

        let mut exact_font = prepared(sfnt.clone(), 0, false);
        let mut limits = Limits::default();
        limits.max_font_bytes = font_size;
        assert!(encode(&mut exact_font, Intent::Editable, limits).is_ok());
        let mut over_font = prepared(sfnt.clone(), 0, false);
        limits.max_font_bytes = font_size - 1;
        assert!(matches!(
            encode(&mut over_font, Intent::Editable, limits),
            Err(FontError::LimitExceeded {
                resource: "OpenType font bytes",
                ..
            })
        ));

        let mut exact_tables = prepared(sfnt.clone(), 0, false);
        limits = Limits::default();
        limits.max_sfnt_tables = 3;
        assert!(encode(&mut exact_tables, Intent::Editable, limits).is_ok());
        let mut over_tables = prepared(sfnt.clone(), 0, false);
        limits.max_sfnt_tables = 2;
        assert!(matches!(
            encode(&mut over_tables, Intent::Editable, limits),
            Err(FontError::LimitExceeded {
                resource: "sfnt table count",
                limit: 2,
                actual: 3,
            })
        ));

        // The four selected synthetic names contain exactly 96 UTF-16 bytes.
        let mut exact_names = prepared(sfnt.clone(), 0, false);
        limits = Limits::default();
        limits.max_name_bytes = 96;
        assert!(encode(&mut exact_names, Intent::Editable, limits).is_ok());
        let mut over_names = prepared(sfnt.clone(), 0, false);
        limits.max_name_bytes = 95;
        assert!(matches!(
            encode(&mut over_names, Intent::Editable, limits),
            Err(FontError::LimitExceeded {
                resource: "decoded sfnt name bytes",
                limit: 95,
                actual: 96,
            })
        ));

        let mut sized = prepared(sfnt.clone(), 0, false);
        let output_size = encode(&mut sized, Intent::Editable, Limits::default())
            .unwrap()
            .len();
        let mut exact_output = prepared(sfnt.clone(), 0, false);
        limits = Limits::default();
        limits.max_output_bytes = output_size;
        assert!(encode(&mut exact_output, Intent::Editable, limits).is_ok());
        let mut over_output = prepared(sfnt, 0, false);
        limits.max_output_bytes = output_size - 1;
        assert!(matches!(
            encode(&mut over_output, Intent::Editable, limits),
            Err(FontError::LimitExceeded {
                resource: "EOT output bytes",
                ..
            })
        ));
    }

    #[test]
    fn fallback_name_is_bounded_before_license_error_cloning() {
        let mut font = prepared(test_sfnt(0), 0x0008, false);
        font.name = "AB".into();
        let mut limits = Limits::default();
        limits.max_name_bytes = 2;
        assert!(matches!(
            encode(&mut font, Intent::PreviewPrint, limits),
            Err(FontError::LimitExceeded {
                resource: "EOT fallback name bytes",
                limit: 2,
                actual: 4,
            })
        ));
        assert!(!font.data.is_empty());
    }

    #[test]
    fn publication_failure_can_restore_the_same_allocation() {
        let sfnt = test_sfnt(0x0008);
        let mut font = prepared(sfnt.clone(), 0x0008, true);
        let eot = encode(&mut font, Intent::Editable, Limits::default()).unwrap();
        let allocation = eot.as_ptr();
        assert!(font.data.is_empty());
        restore(&mut font, eot).unwrap();
        assert_eq!(font.data, sfnt);
        assert_eq!(font.data.as_ptr(), allocation);

        let replacement = encode(&mut font, Intent::Editable, Limits::default()).unwrap();
        let mut occupied = prepared(test_sfnt(0x0008), 0x0008, true);
        assert!(restore(&mut occupied, replacement).is_err());
        assert!(!occupied.data.is_empty());
    }

    #[test]
    fn parser_refuses_header_license_forgery_reserved_data_and_truncation() {
        let mut font = prepared(test_sfnt(0), 0, false);
        let eot = encode(&mut font, Intent::Editable, Limits::default()).unwrap();

        let mut forged = eot.clone();
        forged[32..34].copy_from_slice(&0x0008u16.to_le_bytes());
        assert!(View::parse(&forged).is_err());
        let mut reserved = eot.clone();
        reserved[64] = 1;
        assert!(View::parse(&reserved).is_err());
        assert!(View::parse(&eot[..eot.len() - 1]).is_err());

        let mut limits = Limits::default();
        limits.max_input_bytes = eot.len() - 1;
        assert!(matches!(
            View::parse_with(&eot, limits),
            Err(FontError::LimitExceeded {
                resource: "EOT input bytes",
                ..
            })
        ));
    }

    #[test]
    fn name_record_limit_precedes_hostile_overlapping_record_traversal() {
        let mut sfnt = test_sfnt(0);
        let name_offset = usize::try_from(u32::from_be_bytes(
            sfnt[52..56].try_into().expect("name table offset"),
        ))
        .expect("test offset fits");
        // The advertised records alias the same small backing table. The
        // record-count cap must fire before range traversal or UTF-16 decoding.
        set_u16(&mut sfnt, name_offset + 2, 1_000);
        let mut font = prepared(sfnt, 0, false);
        let mut limits = Limits::default();
        limits.max_name_records = 8;
        assert!(matches!(
            encode(&mut font, Intent::PreviewPrint, limits),
            Err(FontError::LimitExceeded {
                resource: "sfnt name records",
                limit: 8,
                actual: 1_000,
            })
        ));
        assert!(!font.data.is_empty());
    }

    #[test]
    fn overlapping_selected_names_share_one_decoding_budget() {
        let mut sfnt = test_sfnt(0);
        let name_offset = usize::try_from(u32::from_be_bytes(
            sfnt[52..56].try_into().expect("name table offset"),
        ))
        .expect("test offset fits");
        // Make the four selected IDs alias the same two-byte UTF-16 string.
        // Every selected semantic name is charged even though storage overlaps.
        for index in 0..4 {
            let record = name_offset + 6 + index * 12;
            set_u16(&mut sfnt, record + 8, 2);
            set_u16(&mut sfnt, record + 10, 0);
        }
        let mut font = prepared(sfnt, 0, false);
        font.name = "L".into();
        let mut limits = Limits::default();
        limits.max_name_bytes = 6;
        assert!(matches!(
            encode(&mut font, Intent::PreviewPrint, limits),
            Err(FontError::LimitExceeded {
                resource: "decoded sfnt name bytes",
                limit: 6,
                actual: 8,
            })
        ));
        assert!(!font.data.is_empty());
    }

    fn prepared(data: Vec<u8>, license: u16, subsetted: bool) -> Prepared {
        prepared_with_license_and_subset(data, License::new(license).unwrap(), subsetted)
    }

    fn prepared_with_license(data: Vec<u8>, license: License) -> Prepared {
        prepared_with_license_and_subset(data, license, false)
    }

    fn prepared_with_license_and_subset(
        data: Vec<u8>,
        license: License,
        subsetted: bool,
    ) -> Prepared {
        Prepared {
            name: "Litchi Test".into(),
            style: crate::Style::Regular,
            data,
            properties: crate::FontProperties::new(
                license,
                crate::Panose::new([2, 11, 6, 4, 2, 2, 2, 2, 2, 4]),
                Some(crate::Charset::ANSI),
                crate::Family::Roman,
                crate::Pitch::Variable,
                crate::Signature::new([1, 2, 3, 4], [5, 6]),
            ),
            subsetted,
        }
    }

    fn test_sfnt(fs_type: u16) -> Vec<u8> {
        let mut os2 = vec![0; 96];
        set_u16(&mut os2, 0, 2);
        set_u16(&mut os2, 4, 400);
        set_u16(&mut os2, 6, 5);
        set_u16(&mut os2, 8, fs_type);
        os2[32..42].copy_from_slice(&[2, 11, 6, 4, 2, 2, 2, 2, 2, 4]);
        set_u32(&mut os2, 42, 1);
        set_u32(&mut os2, 78, 1);

        let mut head = vec![0; 54];
        set_u32(&mut head, 0, 0x0001_0000);
        set_u32(&mut head, 8, 0x1234_5678);
        set_u32(&mut head, 12, 0x5F0F_3CF5);
        set_u16(&mut head, 18, 1000);

        let name = name_table(&[
            (1, "Litchi Test"),
            (2, "Regular"),
            (4, "Litchi Test Regular"),
            (5, "Version 1.0"),
        ]);
        sfnt(&[(b"OS/2", os2), (b"head", head), (b"name", name)])
    }

    fn name_table(values: &[(u16, &str)]) -> Vec<u8> {
        let string_offset = 6 + values.len() * 12;
        let mut strings = Vec::new();
        let mut records = Vec::new();
        for (id, value) in values {
            let offset = strings.len();
            for unit in value.encode_utf16() {
                strings.extend_from_slice(&unit.to_be_bytes());
            }
            records.push((*id, offset, strings.len() - offset));
        }
        let mut output = vec![0; string_offset];
        set_u16(&mut output, 2, values.len() as u16);
        set_u16(&mut output, 4, string_offset as u16);
        for (index, (id, offset, length)) in records.into_iter().enumerate() {
            let start = 6 + index * 12;
            set_u16(&mut output, start, 3);
            set_u16(&mut output, start + 2, 1);
            set_u16(&mut output, start + 4, 0x0409);
            set_u16(&mut output, start + 6, id);
            set_u16(&mut output, start + 8, length as u16);
            set_u16(&mut output, start + 10, offset as u16);
        }
        output.extend_from_slice(&strings);
        output
    }

    fn sfnt(tables: &[(&[u8; 4], Vec<u8>)]) -> Vec<u8> {
        let directory = 12 + tables.len() * 16;
        let mut offsets = Vec::new();
        let mut length = directory;
        for (_, table) in tables {
            length = (length + 3) & !3;
            offsets.push(length);
            length += table.len();
        }
        let mut output = vec![0; length];
        set_u32(&mut output, 0, 0x0001_0000);
        set_u16(&mut output, 4, tables.len() as u16);
        for (index, ((tag, table), offset)) in tables.iter().zip(offsets).enumerate() {
            let record = 12 + index * 16;
            output[record..record + 4].copy_from_slice(*tag);
            set_u32(&mut output, record + 8, offset as u32);
            set_u32(&mut output, record + 12, table.len() as u32);
            output[offset..offset + table.len()].copy_from_slice(table);
        }
        output
    }

    fn set_u16(bytes: &mut [u8], offset: usize, value: u16) {
        bytes[offset..offset + 2].copy_from_slice(&value.to_be_bytes());
    }

    fn set_u32(bytes: &mut [u8], offset: usize, value: u32) {
        bytes[offset..offset + 4].copy_from_slice(&value.to_be_bytes());
    }
}
