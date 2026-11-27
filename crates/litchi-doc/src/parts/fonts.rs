//! The document's font table: `SttbfFfn` ([MS-DOC] §2.9.286) and its `FFN`
//! records (§2.9.82).
//!
//! Character properties name fonts by index (`ftc`: `sprmCRgFtc0`,
//! `sprmCRgFtc1`, `sprmCRgFtc2`, `sprmCFtcBi`); this table maps an index to
//! the font's name and the attributes Word records for substitution.

use crate::package::{Error as PackageError, Result};

/// `FibRgFcLcb97` index of `fcSttbfFfn`/`lcbSttbfFfn`.
pub(crate) const FIB_INDEX: usize = 15;
/// Largest `cData` a `SttbfFfn` may declare.
const MAX_FONTS: usize = 0x7FF0;
/// Bytes of an `FFN` before its names: `ffid`, `wWeight`, `chs`,
/// `ixchSzAlt`, `panose` and `fs`.
const FFN_FIXED_BYTES: usize = 1 + 2 + 1 + 1 + 10 + 24;
/// Largest `SttbfFfn`: the header and `MAX_FONTS` records of at most 255
/// bytes, each after its one-byte length.
pub(crate) const MAX_TABLE_BYTES: usize = 4 + MAX_FONTS * 256;

/// The character pitch a font requests (`FFID.prq`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum FontPitch {
    /// Default pitch (0).
    Default,
    /// Fixed pitch (1).
    Fixed,
    /// Variable pitch (2).
    Variable,
    /// A value [MS-DOC] does not define (3), kept as stored.
    Other(u8),
}

/// A font's family (`FFID.ff`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum FontFamily {
    /// The family is unspecified (0).
    Unspecified,
    /// Roman, serif (1).
    Roman,
    /// Swiss, sans-serif (2).
    Swiss,
    /// Modern, monospace (3).
    Modern,
    /// Script, cursive (4).
    Script,
    /// Decorative, fantasy (5).
    Decorative,
    /// A value [MS-DOC] does not define (6 or 7), kept as stored.
    Other(u8),
}

/// One font of the font table (an `FFN`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Font {
    name: String,
    alternate_name: Option<String>,
    ffid: u8,
    weight: i16,
    charset: u8,
    panose: [u8; 10],
    signature: [u8; 24],
}

impl Font {
    /// The font's name (`xszFfn`).
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The font to substitute when this one is not available (`xszAlt`).
    #[must_use]
    pub fn alternate_name(&self) -> Option<&str> {
        self.alternate_name.as_deref()
    }

    /// The character pitch the font requests.
    #[must_use]
    pub const fn pitch(&self) -> FontPitch {
        match self.ffid & 0x03 {
            0 => FontPitch::Default,
            1 => FontPitch::Fixed,
            2 => FontPitch::Variable,
            other => FontPitch::Other(other),
        }
    }

    /// Whether the font is a TrueType font.
    #[must_use]
    pub const fn is_true_type(&self) -> bool {
        self.ffid & 0x04 != 0
    }

    /// The font's family.
    #[must_use]
    pub const fn family(&self) -> FontFamily {
        match (self.ffid >> 4) & 0x07 {
            0 => FontFamily::Unspecified,
            1 => FontFamily::Roman,
            2 => FontFamily::Swiss,
            3 => FontFamily::Modern,
            4 => FontFamily::Script,
            5 => FontFamily::Decorative,
            other => FontFamily::Other(other),
        }
    }

    /// The stored `FFID` byte the pitch, TrueType flag and family come from.
    #[must_use]
    pub const fn ffid(&self) -> u8 {
        self.ffid
    }

    /// The font's visual weight (`wWeight`): 400 is normal, 700 bold.
    #[must_use]
    pub const fn weight(&self) -> i16 {
        self.weight
    }

    /// The font's character set (`chs`), a Windows `*_CHARSET` value.
    #[must_use]
    pub const fn charset(&self) -> u8 {
        self.charset
    }

    /// The font's PANOSE classification bytes.
    #[must_use]
    pub const fn panose(&self) -> [u8; 10] {
        self.panose
    }

    /// The font's `FONTSIGNATURE` bytes (Unicode subset and code page
    /// bitfields).
    #[must_use]
    pub const fn signature(&self) -> [u8; 24] {
        self.signature
    }
}

/// The document's fonts, in font-index (`ftc`) order.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FontTable {
    fonts: Vec<Font>,
}

impl FontTable {
    /// Parse a complete `SttbfFfn`.
    ///
    /// The table is a non-extended STTB: a two-byte count of at most
    /// 0x7FF0, a two-byte `cbExtra` that must be 0, then one record per font,
    /// each a one-byte length and an `FFN` of that many bytes. A name must be
    /// null-terminated UTF-16 inside its record, an alternate name exists
    /// exactly when `ixchSzAlt` is nonzero, and the records must fill the
    /// table.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::Corrupted`] for a table that breaks any of
    /// those rules.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let corrupted = |message: &str| PackageError::Corrupted(format!("SttbfFfn {message}"));
        let header = data
            .get(..4)
            .ok_or_else(|| corrupted("is shorter than its header"))?;
        let count = usize::from(u16::from_le_bytes([header[0], header[1]]));
        if count > MAX_FONTS {
            return Err(corrupted("declares more than 0x7FF0 fonts"));
        }
        if u16::from_le_bytes([header[2], header[3]]) != 0 {
            return Err(corrupted("has extra data"));
        }
        let mut fonts = Vec::new();
        fonts
            .try_reserve_exact(count)
            .map_err(|_| corrupted("font list allocation failed"))?;
        let mut offset = 4usize;
        for _ in 0..count {
            let length = usize::from(
                *data
                    .get(offset)
                    .ok_or_else(|| corrupted("ends before its last font"))?,
            );
            let start = offset + 1;
            let record = data
                .get(start..start + length)
                .ok_or_else(|| corrupted("has a truncated font record"))?;
            fonts.push(Self::parse_ffn(record).map_err(corrupted)?);
            offset = start + length;
        }
        if offset != data.len() {
            return Err(corrupted("has bytes after its last font"));
        }
        Ok(Self { fonts })
    }

    fn parse_ffn(record: &[u8]) -> std::result::Result<Font, &'static str> {
        if record.len() < FFN_FIXED_BYTES {
            return Err("has a font record shorter than an FFN");
        }
        let names = &record[FFN_FIXED_BYTES..];
        if names.len() % 2 != 0 {
            return Err("has a font name of an odd byte length");
        }
        let units = names
            .as_chunks::<2>()
            .0
            .iter()
            .map(|unit| u16::from_le_bytes(*unit))
            .collect::<Vec<_>>();
        let name = null_terminated(&units, 0).ok_or("has an unterminated font name")?;
        let alternate_start = usize::from(record[4]);
        let alternate_name = if alternate_start == 0 {
            None
        } else {
            Some(
                null_terminated(&units, alternate_start)
                    .ok_or("has an unterminated alternate font name")?,
            )
        };
        let mut panose = [0u8; 10];
        panose.copy_from_slice(&record[5..15]);
        let mut signature = [0u8; 24];
        signature.copy_from_slice(&record[15..39]);
        Ok(Font {
            name,
            alternate_name,
            ffid: record[0],
            weight: i16::from_le_bytes([record[1], record[2]]),
            charset: record[3],
            panose,
            signature,
        })
    }

    /// The number of fonts.
    #[must_use]
    pub fn len(&self) -> usize {
        self.fonts.len()
    }

    /// Whether the table lists no font.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.fonts.is_empty()
    }

    /// The font at index `ftc`.
    #[must_use]
    pub fn get(&self, ftc: u16) -> Option<&Font> {
        self.fonts.get(usize::from(ftc))
    }

    /// The name of the font at index `ftc`.
    #[must_use]
    pub fn name(&self, ftc: u16) -> Option<&str> {
        self.get(ftc).map(Font::name)
    }

    /// Every font, in index order.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Font> {
        self.fonts.iter()
    }
}

/// The UTF-16 string that starts at `start` in `units` and ends before the
/// next null, or `None` when no null follows or the string is not valid
/// UTF-16.
fn null_terminated(units: &[u16], start: usize) -> Option<String> {
    let tail = units.get(start..)?;
    let end = tail.iter().position(|&unit| unit == 0)?;
    String::from_utf16(&tail[..end]).ok()
}

#[cfg(test)]
mod tests {
    use super::*;

    /// An `FFN` record with its one-byte length, as Word writes it.
    fn record(ffid: u8, weight: i16, charset: u8, name: &str, alternate: Option<&str>) -> Vec<u8> {
        let name_units = name.encode_utf16().collect::<Vec<_>>();
        let alternate_at = alternate.map_or(0, |_| name_units.len() + 1);
        let mut body = vec![ffid];
        body.extend_from_slice(&weight.to_le_bytes());
        body.push(charset);
        body.push(u8::try_from(alternate_at).unwrap());
        body.extend((1..=10).map(|value: u8| value));
        body.extend((0..24).map(|value: u8| value.wrapping_mul(3)));
        for unit in name_units.iter().copied().chain([0]) {
            body.extend_from_slice(&unit.to_le_bytes());
        }
        if let Some(alternate) = alternate {
            for unit in alternate.encode_utf16().chain([0]) {
                body.extend_from_slice(&unit.to_le_bytes());
            }
        }
        let mut record = vec![u8::try_from(body.len()).unwrap()];
        record.extend(body);
        record
    }

    fn table(records: &[Vec<u8>]) -> Vec<u8> {
        let mut data = u16::try_from(records.len()).unwrap().to_le_bytes().to_vec();
        data.extend_from_slice(&0u16.to_le_bytes());
        for record in records {
            data.extend_from_slice(record);
        }
        data
    }

    #[test]
    fn parses_names_alternates_and_attributes_by_index() {
        let data = table(&[
            record(0x16, 400, 0, "Times New Roman", None),
            record(0x26, 700, 2, "Symbol", None),
            record(0x25, 400, 0, "Arial", Some("Helvetica")),
            record(0x00, 400, 128, "ＭＳ 明朝", None),
        ]);
        let fonts = FontTable::parse(&data).unwrap();
        assert_eq!(fonts.len(), 4);
        assert_eq!(fonts.name(0), Some("Times New Roman"));
        assert_eq!(fonts.name(3), Some("ＭＳ 明朝"));
        assert_eq!(fonts.name(4), None);
        let times = fonts.get(0).unwrap();
        assert_eq!(times.pitch(), FontPitch::Variable);
        assert!(times.is_true_type());
        assert_eq!(times.family(), FontFamily::Roman);
        assert_eq!(times.weight(), 400);
        assert_eq!(times.alternate_name(), None);
        assert_eq!(times.panose(), [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
        let symbol = fonts.get(1).unwrap();
        assert_eq!(symbol.family(), FontFamily::Swiss);
        assert_eq!((symbol.weight(), symbol.charset()), (700, 2));
        let arial = fonts.get(2).unwrap();
        assert_eq!(arial.alternate_name(), Some("Helvetica"));
        assert_eq!(arial.pitch(), FontPitch::Fixed);
        assert_eq!(fonts.iter().count(), 4);
        assert!(FontTable::parse(&table(&[])).unwrap().is_empty());
    }

    #[test]
    fn matches_the_writer_font_table() {
        let mut builder = crate::writer::font_table::FontTableBuilder::new();
        assert_eq!(builder.get_or_add("Arial"), 1);
        let fonts = FontTable::parse(&builder.generate()).unwrap();
        assert_eq!(fonts.name(0), Some("Times New Roman"));
        assert_eq!(fonts.name(1), Some("Arial"));
        assert_eq!(fonts.get(1).unwrap().weight(), 400);
    }

    #[test]
    fn refuses_malformed_tables() {
        let valid = record(0x16, 400, 0, "Times New Roman", None);
        let cases: Vec<(&str, Vec<u8>)> = vec![
            ("short header", vec![1, 0, 0]),
            ("extra data", {
                let mut data = table(std::slice::from_ref(&valid));
                data[2] = 1;
                data
            }),
            ("too many fonts", {
                let mut data = table(&[]);
                data[..2].copy_from_slice(&0x7FF1u16.to_le_bytes());
                data
            }),
            ("missing record", {
                let mut data = table(std::slice::from_ref(&valid));
                data[..2].copy_from_slice(&2u16.to_le_bytes());
                data
            }),
            ("truncated record", {
                let mut data = table(std::slice::from_ref(&valid));
                data.pop();
                data
            }),
            ("trailing bytes", {
                let mut data = table(std::slice::from_ref(&valid));
                data.push(0);
                data
            }),
            ("short record", table(&[vec![2, 0, 0]])),
            ("unterminated name", {
                let mut record = valid.clone();
                let length = record.len();
                record[length - 2..].copy_from_slice(&u16::from(b'x').to_le_bytes());
                table(&[record])
            }),
            ("alternate past the record", {
                let mut record = valid.clone();
                record[5] = 200;
                table(&[record])
            }),
            ("invalid UTF-16", {
                let mut record = valid;
                record[40..42].copy_from_slice(&0xD800u16.to_le_bytes());
                table(&[record])
            }),
        ];
        for (case, data) in cases {
            assert!(FontTable::parse(&data).is_err(), "{case}");
        }
    }
}
