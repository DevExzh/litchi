//! Legacy LOGFONT character-set decoding for metafile text records.
//!
//! WMF and EMF text records carry single-byte code-page strings plus the
//! character set of the font selected when the record was recorded. Windows
//! maps the character set to an installed code page and converts the string
//! with it, so the same byte sequence means different text in different
//! character sets. [`decode_logfont_text`] performs that conversion, and
//! reports when the character set cannot be resolved exactly.

use litchi_codepage::Page;

/// LOGFONT character-set values and the code page each one selects.
mod charset {
    /// `ANSI_CHARSET`: the system ANSI code page.
    pub(super) const ANSI: u8 = 0;
    /// `DEFAULT_CHARSET`: the system default code page.
    pub(super) const DEFAULT: u8 = 1;
    /// `SYMBOL_CHARSET`: glyph order of a symbol font, not a code page.
    pub(super) const SYMBOL: u8 = 2;
    /// `MAC_CHARSET`.
    pub(super) const MAC: u8 = 77;
    /// `SHIFTJIS_CHARSET`.
    pub(super) const SHIFT_JIS: u8 = 128;
    /// `HANGUL_CHARSET`.
    pub(super) const HANGUL: u8 = 129;
    /// `JOHAB_CHARSET`.
    pub(super) const JOHAB: u8 = 130;
    /// `GB2312_CHARSET`.
    pub(super) const GB2312: u8 = 134;
    /// `CHINESEBIG5_CHARSET`.
    pub(super) const BIG5: u8 = 136;
    /// `GREEK_CHARSET`.
    pub(super) const GREEK: u8 = 161;
    /// `TURKISH_CHARSET`.
    pub(super) const TURKISH: u8 = 162;
    /// `VIETNAMESE_CHARSET`.
    pub(super) const VIETNAMESE: u8 = 163;
    /// `HEBREW_CHARSET`.
    pub(super) const HEBREW: u8 = 177;
    /// `ARABIC_CHARSET`.
    pub(super) const ARABIC: u8 = 178;
    /// `BALTIC_CHARSET`.
    pub(super) const BALTIC: u8 = 186;
    /// `RUSSIAN_CHARSET`.
    pub(super) const RUSSIAN: u8 = 204;
    /// `THAI_CHARSET`.
    pub(super) const THAI: u8 = 222;
    /// `EASTEUROPE_CHARSET`.
    pub(super) const EAST_EUROPE: u8 = 238;
    /// `OEM_CHARSET`: the system OEM code page.
    pub(super) const OEM: u8 = 255;
}

/// Decoded text plus what the caller needs to know about the conversion.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct DecodedText {
    /// The text to draw.
    pub(crate) text: String,
    /// Whether the conversion reproduces what Windows draws for this font.
    pub(crate) exact: bool,
    /// Whether the recorded face draws glyphs from a symbol character set and
    /// the text therefore needs a Unicode-capable face instead.
    pub(crate) unicode_face: bool,
}

/// Code page selected by a character set, when the mapping is unambiguous.
///
/// `DEFAULT_CHARSET` and `ANSI_CHARSET` name the system code page; a metafile
/// recorded on a Western Windows system — the overwhelming majority — uses
/// windows-1252, which is also the fallback for a character set this function
/// does not resolve. `OEM_CHARSET` names the system OEM code page, which is
/// not recorded anywhere in the metafile.
fn code_page(charset: u8) -> Option<u32> {
    Some(match charset {
        charset::ANSI | charset::DEFAULT => 1252,
        charset::MAC => 10000,
        charset::SHIFT_JIS => 932,
        charset::HANGUL => 949,
        charset::GB2312 => 936,
        charset::BIG5 => 950,
        charset::GREEK => 1253,
        charset::TURKISH => 1254,
        charset::VIETNAMESE => 1258,
        charset::HEBREW => 1255,
        charset::ARABIC => 1256,
        charset::BALTIC => 1257,
        charset::RUSSIAN => 1251,
        charset::THAI => 874,
        charset::EAST_EUROPE => 1250,
        // JOHAB has no exact codec here, and the OEM code page is a property
        // of the recording system that the metafile does not record.
        charset::JOHAB | charset::OEM => return None,
        _ => return None,
    })
}

/// Glyph ordering of a face that is not a text font.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SymbolFace {
    /// Adobe Symbol encoding, which has direct Unicode equivalents.
    Adobe,
    /// A symbol cmap whose glyphs Windows places in the U+F000 private-use
    /// range: dingbat and equation faces (Wingdings, MT Extra, ...).
    PrivateUse,
}

/// Recognizes faces whose glyph order is a symbol order rather than a text
/// encoding.
///
/// The names are the symbol and dingbat faces Windows ships. Producers do not
/// always set `SYMBOL_CHARSET` for them: MathType records `MT Extra` with the
/// undocumented character set 160, so only the face name identifies the glyph
/// order. A face that is not listed here is decoded with the code page of its
/// character set.
fn symbol_face(face_name: &str) -> Option<SymbolFace> {
    for (name, kind) in [
        ("Symbol", SymbolFace::Adobe),
        ("SymbolMT", SymbolFace::Adobe),
        ("MT Extra", SymbolFace::PrivateUse),
        ("Wingdings", SymbolFace::PrivateUse),
        ("Wingdings 2", SymbolFace::PrivateUse),
        ("Wingdings 3", SymbolFace::PrivateUse),
        ("Webdings", SymbolFace::PrivateUse),
        ("Marlett", SymbolFace::PrivateUse),
        ("Monotype Sorts", SymbolFace::PrivateUse),
        ("Zapf Dingbats", SymbolFace::PrivateUse),
    ] {
        if face_name.eq_ignore_ascii_case(name) {
            return Some(kind);
        }
    }
    None
}

/// Decodes one text record's bytes with the character set of the font that is
/// selected for it.
pub(crate) fn decode_logfont_text(bytes: &[u8], charset: u8, face_name: &str) -> DecodedText {
    if let Some(face) = symbol_face(face_name) {
        return decode_symbol_text(bytes, face);
    }
    if charset == charset::SYMBOL {
        return decode_symbol_text(bytes, SymbolFace::PrivateUse);
    }
    let page = match code_page(charset) {
        Some(id) => Page::new(id),
        None => None,
    };
    let Some(page) = page else {
        // The character set names a code page this crate cannot provide
        // (JOHAB, OEM). Windows-1252 keeps the picture readable for the ASCII
        // range, but the non-ASCII range is a guess.
        return DecodedText {
            text: without_nul(Page::WINDOWS_1252.recover(bytes).0.as_ref()),
            exact: bytes.iter().all(u8::is_ascii),
            unicode_face: false,
        };
    };
    if MULTIBYTE_CODE_PAGES.contains(&page.id()) {
        let (text, recovered) = page.recover(bytes);
        return DecodedText {
            text: without_nul(&text),
            exact: !recovered,
            unicode_face: false,
        };
    }
    DecodedText {
        text: without_nul(&decode_single_byte(page, bytes)),
        exact: true,
        unicode_face: false,
    }
}

/// Single-byte code pages whose positions are not all assigned.
const MULTIBYTE_CODE_PAGES: [u32; 4] = [932, 936, 949, 950];

/// Decodes a single-byte code page one byte at a time.
///
/// A code page can leave positions unassigned (0x81 and 0x8D in windows-1252,
/// 0xA1 and 0xA5 in windows-1257). Windows converts such a position to the
/// same code point rather than to a replacement character, so the byte's own
/// value keeps the text exact and the picture unchanged.
fn decode_single_byte(page: Page, bytes: &[u8]) -> String {
    let mut text = String::with_capacity(bytes.len());
    for &byte in bytes {
        let buffer = [byte];
        let (decoded, unassigned) = page.recover(&buffer);
        if unassigned {
            text.push(char::from(byte));
        } else {
            text.push_str(&decoded);
        }
    }
    text
}

/// Decodes symbol-font text.
///
/// A symbol face draws glyph *n* for byte *n*, so the text is not a code-page
/// string: the Adobe Symbol order maps onto real Unicode characters, while
/// every other symbol face is reported in the Unicode private-use area,
/// exactly where its Windows cmap places the glyphs.
fn decode_symbol_text(bytes: &[u8], face: SymbolFace) -> DecodedText {
    if face == SymbolFace::Adobe {
        DecodedText {
            text: bytes.iter().map(|&byte| adobe_symbol(byte)).collect(),
            exact: true,
            unicode_face: true,
        }
    } else {
        DecodedText {
            text: bytes
                .iter()
                .map(|&byte| char::from_u32(0xf000 + u32::from(byte)).unwrap_or('\u{fffd}'))
                .collect(),
            exact: true,
            unicode_face: false,
        }
    }
}

fn without_nul(text: &str) -> String {
    text.chars()
        .filter(|&character| character != '\0')
        .collect()
}

/// Adobe Symbol encoding, the glyph order of the Windows "Symbol" face.
///
/// Greek letters occupy the letter positions, arithmetic and set operators the
/// punctuation and high positions; the table follows the encoding vector of
/// the Adobe Symbol font.
fn adobe_symbol(byte: u8) -> char {
    /// Uppercase Greek in Adobe Symbol letter order: alpha, beta, chi, ...
    const UPPER: [char; 26] = [
        '\u{0391}', '\u{0392}', '\u{03a7}', '\u{0394}', '\u{0395}', '\u{03a6}', '\u{0393}',
        '\u{0397}', '\u{0399}', '\u{03d1}', '\u{039a}', '\u{039b}', '\u{039c}', '\u{039d}',
        '\u{039f}', '\u{03a0}', '\u{0398}', '\u{03a1}', '\u{03a3}', '\u{03a4}', '\u{03a5}',
        '\u{03c2}', '\u{03a9}', '\u{039e}', '\u{03a8}', '\u{0396}',
    ];
    /// Lowercase Greek in Adobe Symbol letter order.
    const LOWER: [char; 26] = [
        '\u{03b1}', '\u{03b2}', '\u{03c7}', '\u{03b4}', '\u{03b5}', '\u{03c6}', '\u{03b3}',
        '\u{03b7}', '\u{03b9}', '\u{03d5}', '\u{03ba}', '\u{03bb}', '\u{03bc}', '\u{03bd}',
        '\u{03bf}', '\u{03c0}', '\u{03b8}', '\u{03c1}', '\u{03c3}', '\u{03c4}', '\u{03c5}',
        '\u{03d6}', '\u{03c9}', '\u{03be}', '\u{03c8}', '\u{03b6}',
    ];
    /// Positions 0xA0 through 0xFF, sixteen per row.
    const HIGH: [char; 96] = [
        '\u{20ac}', '\u{03d2}', '\u{2032}', '\u{2264}', '\u{2044}', '\u{221e}', '\u{0192}',
        '\u{2663}', '\u{2666}', '\u{2665}', '\u{2660}', '\u{2194}', '\u{2190}', '\u{2191}',
        '\u{2192}', '\u{2193}', '\u{00b0}', '\u{00b1}', '\u{2033}', '\u{2265}', '\u{00d7}',
        '\u{221d}', '\u{2202}', '\u{2022}', '\u{00f7}', '\u{2260}', '\u{2261}', '\u{2248}',
        '\u{2026}', '\u{23d0}', '\u{23af}', '\u{21b5}', '\u{2135}', '\u{2111}', '\u{211c}',
        '\u{2118}', '\u{2297}', '\u{2295}', '\u{2205}', '\u{2229}', '\u{222a}', '\u{2283}',
        '\u{2287}', '\u{2284}', '\u{2282}', '\u{2286}', '\u{2208}', '\u{2209}', '\u{2220}',
        '\u{2207}', '\u{00ae}', '\u{00a9}', '\u{2122}', '\u{220f}', '\u{221a}', '\u{22c5}',
        '\u{00ac}', '\u{2227}', '\u{2228}', '\u{21d4}', '\u{21d0}', '\u{21d1}', '\u{21d2}',
        '\u{21d3}', '\u{25ca}', '\u{27e8}', '\u{00ae}', '\u{00a9}', '\u{2122}', '\u{2211}',
        '\u{239b}', '\u{239c}', '\u{239d}', '\u{23a1}', '\u{23a2}', '\u{23a3}', '\u{23a7}',
        '\u{23a8}', '\u{23a9}', '\u{23aa}', '\u{fffd}', '\u{27e9}', '\u{222b}', '\u{2320}',
        '\u{23ae}', '\u{2321}', '\u{239e}', '\u{239f}', '\u{23a0}', '\u{23a4}', '\u{23a5}',
        '\u{23a6}', '\u{23ab}', '\u{23ac}', '\u{23ad}', '\u{fffd}',
    ];
    match byte {
        0x20..=0x2f => match byte {
            0x22 => '\u{2200}',
            0x24 => '\u{2203}',
            0x27 => '\u{220b}',
            0x2a => '\u{2217}',
            0x2d => '\u{2212}',
            _ => char::from(byte),
        },
        0x30..=0x3f => char::from(byte),
        0x40 => '\u{2245}',
        b'A'..=b'Z' => UPPER[usize::from(byte - b'A')],
        0x5b..=0x60 => match byte {
            0x5c => '\u{2234}',
            0x5e => '\u{22a5}',
            0x60 => '\u{203e}',
            _ => char::from(byte),
        },
        b'a'..=b'z' => LOWER[usize::from(byte - b'a')],
        0x7b..=0x7e => match byte {
            0x7e => '\u{223c}',
            _ => char::from(byte),
        },
        0xa0..=0xff => HIGH[usize::from(byte - 0xa0)],
        // Control positions draw no glyph in the Symbol face.
        _ => '\u{fffd}',
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn decode(bytes: &[u8], charset: u8, face: &str) -> DecodedText {
        decode_logfont_text(bytes, charset, face)
    }

    #[test]
    fn windows_code_pages_are_selected_by_character_set() {
        assert_eq!(decode(b"caf\xe9", charset::ANSI, "").text, "café");
        assert_eq!(decode(b"caf\xe9", charset::DEFAULT, "").text, "café");
        assert_eq!(decode(b"\xc0\xe1", charset::GREEK, "").text, "ΐα");
        assert_eq!(decode(b"\xc0\xe1", charset::TURKISH, "").text, "Àá");
        assert_eq!(decode(b"\xc0\xe1", charset::RUSSIAN, "").text, "Аб");
        assert_eq!(decode(b"\xa5\xa2", charset::EAST_EUROPE, "").text, "Ą˘");
        assert_eq!(decode(b"\xc0\xe1", charset::BALTIC, "").text, "Ąį");
        assert_eq!(decode(b"\xc0\xe1", charset::HEBREW, "").text, "\u{5b0}ב");
        assert_eq!(decode(b"\xc0\xe1", charset::ARABIC, "").text, "ہل");
        assert_eq!(decode(b"\xa1", charset::THAI, "").text, "ก");
        assert_eq!(decode(b"\xa1", charset::VIETNAMESE, "").text, "¡");
        assert_eq!(decode(b"\x80", charset::MAC, "").text, "Ä");
        for text in [
            decode(b"abc", charset::ANSI, ""),
            decode(b"abc", charset::DEFAULT, ""),
            decode(b"abc", charset::RUSSIAN, ""),
        ] {
            assert!(text.exact);
            assert!(!text.unicode_face);
        }
    }

    #[test]
    fn multibyte_character_sets_use_their_code_page() {
        assert_eq!(decode(b"\x82\xa0", charset::SHIFT_JIS, "").text, "あ");
        assert_eq!(decode(b"\xb0\xa1", charset::HANGUL, "").text, "가");
        assert_eq!(decode(b"\xd6\xd0", charset::GB2312, "").text, "中");
        assert_eq!(decode(b"\xa4\xa4", charset::BIG5, "").text, "中");

        let malformed = decode(b"\x82", charset::SHIFT_JIS, "");
        assert!(!malformed.exact, "a dangling lead byte is not exact");
    }

    #[test]
    fn unassigned_single_byte_positions_keep_their_code_point() {
        // windows-1257 leaves 0xA1 and 0xA5 unassigned.
        let text = decode(b"a\xa1\xc0", charset::BALTIC, "");
        assert_eq!(text.text, "a\u{a1}Ą");
        assert!(text.exact);
    }

    #[test]
    fn unresolved_character_sets_fall_back_and_report_recovery() {
        let ascii = decode(b"abc", charset::OEM, "");
        assert_eq!(ascii.text, "abc");
        assert!(ascii.exact, "ASCII is code-page independent");

        let high = decode(b"\x82\xa0", charset::OEM, "");
        assert!(!high.exact);
        assert_eq!(high.text, "‚\u{a0}");

        let johab = decode(b"\x88\x61", charset::JOHAB, "");
        assert!(!johab.exact);
    }

    #[test]
    fn adobe_symbol_letters_decode_to_greek() {
        let text = decode(b"abgd", charset::SYMBOL, "Symbol");
        assert_eq!(text.text, "αβγδ");
        assert!(text.exact);
        assert!(text.unicode_face);

        assert_eq!(
            decode("ABGWD".as_bytes(), charset::SYMBOL, "Symbol").text,
            "ΑΒΓΩΔ"
        );
        // The face name selects the encoding even without the charset.
        assert_eq!(decode(b"a", charset::ANSI, "Symbol").text, "α");
        assert_eq!(decode(b"a", charset::ANSI, "symbolMT").text, "α");
    }

    #[test]
    fn adobe_symbol_operators_and_fences_decode_to_unicode() {
        let text = decode(
            b"\xb1\xb4\xb8\xa3\xb3\xb9\xbb\xe5\xf2\xf5\xcc\xcd\xff",
            charset::SYMBOL,
            "Symbol",
        );
        assert_eq!(text.text, "±×÷≤≥≠≈∑∫⌡⊂⊆\u{fffd}");
    }

    #[test]
    fn other_symbol_faces_report_private_use_glyph_positions() {
        let text = decode(b"\x21\x61", charset::SYMBOL, "Wingdings");
        assert_eq!(text.text, "\u{f021}\u{f061}");
        assert!(text.exact);
        assert!(!text.unicode_face);

        let math = decode(b"\x61", 160, "MT Extra");
        assert_eq!(math.text, "\u{f061}");
        assert!(!math.unicode_face);
    }

    #[test]
    fn nul_bytes_draw_nothing() {
        assert_eq!(decode(b"a\0b", charset::ANSI, "").text, "ab");
    }
}
