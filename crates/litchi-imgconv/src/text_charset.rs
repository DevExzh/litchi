//! Legacy LOGFONT character-set decoding and font naming for metafile text.
//!
//! WMF and EMF text records carry single-byte code-page strings plus the
//! character set of the font selected when the record was recorded. Windows
//! maps the character set to an installed code page and converts the string
//! with it, so the same byte sequence means different text in different
//! character sets. [`decode_logfont_text`] performs that conversion, and
//! reports when the character set cannot be resolved exactly.
//! [`svg_font_family`] names the recorded face for SVG together with a
//! generic family that stands in for it where it is not installed.

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
    /// For each character of `text`, the index of the first code unit of the
    /// recorded string it was decoded from: a byte of an ANSI string, or a
    /// UTF-16 unit of a Unicode string.
    ///
    /// Text records give one advance per code unit, and characters that draw
    /// nothing are left out of `text`, so a character's origin is the sum of
    /// the advances before its first unit.
    pub(crate) unit_starts: Vec<usize>,
    /// Whether the conversion reproduces what Windows draws for this font.
    pub(crate) exact: bool,
    /// Whether the recorded face draws glyphs from a symbol character set and
    /// the text therefore needs a Unicode-capable face instead.
    pub(crate) unicode_face: bool,
}

impl DecodedText {
    /// Offsets of the character origins from the reference point, given the
    /// recorded advance of every code unit.
    ///
    /// A record may carry fewer advances than characters; the offsets stop at
    /// the first character whose origin they do not determine, and the
    /// characters after it follow the font's own advances.
    pub(crate) fn origin_offsets(&self, advances: &[(f64, f64)]) -> Vec<(f64, f64)> {
        let mut offsets = Vec::with_capacity(self.unit_starts.len().min(advances.len() + 1));
        let mut sum = (0.0, 0.0);
        let mut summed = 0;
        for &start in &self.unit_starts {
            let Some(before) = advances.get(summed..start) else {
                break;
            };
            for (x, y) in before {
                sum.0 += x;
                sum.1 += y;
            }
            summed = start;
            offsets.push(sum);
        }
        offsets
    }

    /// Removes the characters whose origin `visible` rejects, and with them
    /// the characters after the last origin when its character is removed,
    /// since those follow it. Returns the origins of the characters kept.
    ///
    /// Corrupt advances can place characters at coordinates far beyond the
    /// picture, where they add no ink but exceed what renderers can represent.
    pub(crate) fn retain_visible(
        &mut self,
        origins: &[(f64, f64)],
        visible: impl Fn((f64, f64)) -> bool,
    ) -> Vec<(f64, f64)> {
        let keep: Vec<bool> = origins.iter().map(|&origin| visible(origin)).collect();
        if keep.iter().all(|&kept| kept) {
            return origins.to_vec();
        }
        let follow_last = keep.last().copied().unwrap_or(true);
        let mut text = String::with_capacity(self.text.len());
        let mut unit_starts = Vec::with_capacity(self.unit_starts.len());
        let mut kept_origins = Vec::with_capacity(origins.len());
        for (index, (character, &start)) in self.text.chars().zip(&self.unit_starts).enumerate() {
            let kept = keep.get(index).copied().unwrap_or(follow_last);
            if kept {
                text.push(character);
                unit_starts.push(start);
                if let Some(&origin) = origins.get(index) {
                    kept_origins.push(origin);
                }
            }
        }
        self.text = text;
        self.unit_starts = unit_starts;
        kept_origins
    }
}

/// Whether a character whose origin is `origin`, in a string whose reference
/// point is `reference`, can put ink on the canvas `(left, top, right,
/// bottom)` when its glyph extends at most `reach` from its origin.
///
/// The test holds whatever rotation about the reference point the text is
/// drawn with, since that rotation keeps the distance from the reference
/// point.
pub(crate) fn origin_reaches_canvas(
    origin: (f64, f64),
    reference: (f64, f64),
    canvas: (f64, f64, f64, f64),
    reach: f64,
) -> bool {
    let center = ((canvas.0 + canvas.2) / 2.0, (canvas.1 + canvas.3) / 2.0);
    let radius = (canvas.2 - canvas.0).hypot(canvas.3 - canvas.1) / 2.0 + reach;
    let from_reference = (origin.0 - reference.0).hypot(origin.1 - reference.1);
    let reference_to_center = (reference.0 - center.0).hypot(reference.1 - center.1);
    // NaN distances fail the comparison, so they are not reachable either.
    from_reference <= reference_to_center + radius
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
#[derive(Debug, Clone, Copy)]
enum SymbolFace {
    /// Adobe Symbol encoding, which has direct Unicode equivalents.
    Adobe,
    /// A symbol cmap whose glyphs Windows places in the U+F000 private-use
    /// range: dingbat and equation faces (Wingdings, MT Extra, ...). The
    /// function names the standard Unicode character a glyph position draws,
    /// where one exists.
    PrivateUse(fn(u8) -> Option<char>),
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
        ("MT Extra", SymbolFace::PrivateUse(mt_extra)),
        ("Wingdings", SymbolFace::PrivateUse(wingdings)),
        ("Wingdings 2", SymbolFace::PrivateUse(no_equivalent)),
        ("Wingdings 3", SymbolFace::PrivateUse(wingdings_3)),
        ("Webdings", SymbolFace::PrivateUse(arrowhead_digits)),
        ("Marlett", SymbolFace::PrivateUse(arrowhead_digits)),
        ("Monotype Sorts", SymbolFace::PrivateUse(zapf_dingbats)),
        ("Zapf Dingbats", SymbolFace::PrivateUse(zapf_dingbats)),
    ] {
        if face_name.eq_ignore_ascii_case(name) {
            return Some(kind);
        }
    }
    None
}

/// Character drawn by glyph `position` of a private-use symbol face: the
/// standard equivalent where the face has one, else the private-use code
/// point through which Windows reaches the glyph.
fn private_use_glyph(equivalent: fn(u8) -> Option<char>, position: u8) -> char {
    if position == 0x20 {
        // Every symbol face keeps the space at the ASCII position.
        return ' ';
    }
    equivalent(position)
        .or_else(|| char::from_u32(0xf000 + u32::from(position)))
        .unwrap_or('\u{fffd}')
}

/// A symbol face without known standard equivalents.
fn no_equivalent(_: u8) -> Option<char> {
    None
}

/// Standard characters drawn by MathType's "MT Extra" face.
///
/// MathType stores the Unicode value of every character of an equation next
/// to the MT Extra position it draws it with (MTEF 5 `CHAR` records); these
/// pairs are the ones those records give. Positions that MathType itself
/// assigns to the private-use area (such as 0x67) stay there. MathType draws
/// the accents of its embellishments (dot, tilde, hat, vector arrow, vector
/// harpoon) with positions 0x24 through 0x26, 0x72 and 0x76; they are the
/// combining characters of those accents, placed where MathType drew the
/// glyph.
fn mt_extra(position: u8) -> Option<char> {
    Some(match position {
        0x24 => '\u{0302}',
        0x25 => '\u{0303}',
        0x26 => '\u{0307}',
        0x3a => '\u{223c}',
        0x3b => '\u{2243}',
        0x3d => '\u{226a}',
        0x3f => '\u{226b}',
        0x40 => '\u{225c}',
        0x41 => '\u{2259}',
        0x44 => '\u{019b}',
        0x49 => '\u{2229}',
        0x4b => '\u{2026}',
        0x4c => '\u{22ef}',
        0x4d => '\u{22ee}',
        0x4f => '\u{22f1}',
        0x50 => '\u{2225}',
        0x51 => '\u{2235}',
        0x53 => '\u{2222}',
        0x55 => '\u{222a}',
        0x56 => '\u{25b3}',
        0x57 => '\u{25a1}',
        0x5a => '\u{2197}',
        0x5d => '\u{2198}',
        0x64 => '\u{25cb}',
        0x65 => '\u{2299}',
        0x66 => '\u{227b}',
        0x68 => '\u{210f}',
        0x6c => '\u{2113}',
        0x6d => '\u{2213}',
        0x6f => '\u{2218}',
        0x70 => '\u{227a}',
        0x72 => '\u{20d7}',
        0x76 => '\u{20d1}',
        0x80 => '\u{21c4}',
        0xa1 => '\u{211d}',
        0xa2 => '\u{2124}',
        0xa5 => '\u{2115}',
        _ => return None,
    })
}

/// Standard characters drawn by the Wingdings face.
///
/// The glyphs follow Microsoft's Wingdings character set; the positions
/// listed draw a glyph that has long been encoded in Unicode's Dingbats,
/// Geometric Shapes, Miscellaneous Symbols, Enclosed Alphanumerics or Arrows
/// blocks. Office uses most of them as list bullets and check marks. The
/// pictographs without such an equivalent stay in the private-use area.
fn wingdings(position: u8) -> Option<char> {
    Some(match position {
        0x22 => '\u{2702}',
        0x23 => '\u{2701}',
        0x28 => '\u{260e}',
        0x29 => '\u{2706}',
        0x36 => '\u{231b}',
        0x37 => '\u{2328}',
        0x3e => '\u{2707}',
        0x3f => '\u{270d}',
        0x41 => '\u{270c}',
        0x45 => '\u{261c}',
        0x46 => '\u{261e}',
        0x47 => '\u{261d}',
        0x48 => '\u{261f}',
        0x4a => '\u{263a}',
        0x4c => '\u{2639}',
        0x4e => '\u{2620}',
        0x51 => '\u{2708}',
        0x52 => '\u{263c}',
        0x54 => '\u{2744}',
        0x58 => '\u{2720}',
        0x59 => '\u{2721}',
        0x5a => '\u{262a}',
        0x5b => '\u{262f}',
        0x5c => '\u{0950}',
        0x5d => '\u{2638}',
        // The twelve zodiac signs, Aries through Pisces.
        0x5e..=0x69 => char::from_u32(0x2648 + u32::from(position - 0x5e))?,
        0x6c => '\u{25cf}',
        0x6d => '\u{274d}',
        0x6e => '\u{25a0}',
        0x6f => '\u{25a1}',
        0x71 => '\u{2751}',
        0x72 => '\u{2752}',
        0x75 => '\u{25c6}',
        0x76 => '\u{2756}',
        0x77 => '\u{2b25}',
        0x78 => '\u{2327}',
        0x7a => '\u{2318}',
        0x7d => '\u{275d}',
        0x7e => '\u{275e}',
        0x80 => '\u{24ea}',
        // Circled digits one through ten.
        0x81..=0x8a => char::from_u32(0x2460 + u32::from(position - 0x81))?,
        0x8b => '\u{24ff}',
        // Negative circled digits one through ten.
        0x8c..=0x95 => char::from_u32(0x2776 + u32::from(position - 0x8c))?,
        0x9f => '\u{2022}',
        0xa4 => '\u{25c9}',
        0xa5 => '\u{25ce}',
        0xa7 => '\u{25aa}',
        0xa8 => '\u{25fb}',
        0xaa => '\u{2726}',
        0xab => '\u{2605}',
        0xac => '\u{2736}',
        0xae => '\u{2739}',
        0xaf => '\u{2735}',
        0xb5 => '\u{272a}',
        0xb6 => '\u{2730}',
        0xd5 => '\u{232b}',
        0xd6 => '\u{2326}',
        0xd8 => '\u{27a2}',
        0xe8 => '\u{2794}',
        0xef => '\u{21e6}',
        0xf0 => '\u{21e8}',
        0xf1 => '\u{21e7}',
        0xf2 => '\u{21e9}',
        0xf3 => '\u{2b04}',
        0xf4 => '\u{21f3}',
        0xfb => '\u{2717}',
        0xfc => '\u{2713}',
        0xfd => '\u{2612}',
        0xfe => '\u{2611}',
        _ => return None,
    })
}

/// Standard characters drawn by the Wingdings 3 face: the solid and outlined
/// triangles of Microsoft's Wingdings 3 character set, which Office uses as
/// list bullets. Its arrows have no single Unicode equivalent each and stay in
/// the private-use area.
fn wingdings_3(position: u8) -> Option<char> {
    Some(match position {
        0x70 => '\u{25b2}',
        0x71 => '\u{25bc}',
        0x72 => '\u{25b3}',
        0x73 => '\u{25bd}',
        0x74 => '\u{25c0}',
        0x75 => '\u{25b6}',
        0x76 => '\u{25c1}',
        0x77 => '\u{25b7}',
        0x78 => '\u{25e3}',
        0x79 => '\u{25e2}',
        0x7a => '\u{25e4}',
        0x7b => '\u{25e5}',
        _ => return None,
    })
}

/// Standard characters drawn by Webdings and Marlett: both faces draw the
/// solid arrowheads of the Windows user interface (left, right, up and down)
/// at the digits 3 through 6.
fn arrowhead_digits(position: u8) -> Option<char> {
    Some(match position {
        b'3' => '\u{25c0}',
        b'4' => '\u{25b6}',
        b'5' => '\u{25b2}',
        b'6' => '\u{25bc}',
        _ => return None,
    })
}

/// Standard characters drawn by the ITC Zapf Dingbats glyph order, which
/// Monotype Sorts shares.
///
/// Unicode's Dingbats block encodes Zapf Dingbats in code order: position
/// 0x21 + n is U+2701 + n, except for the glyphs Unicode had already encoded
/// elsewhere (the telephone, pointing hands, star, circle, square, triangles,
/// diamond and half circle). The upper half holds ornaments, card suits,
/// circled digits and arrows in the same order as U+2761 onwards.
fn zapf_dingbats(position: u8) -> Option<char> {
    let offset = |base: u32, first: u8| char::from_u32(base + u32::from(position - first));
    match position {
        0x25 => Some('\u{260e}'),
        0x2a => Some('\u{261b}'),
        0x2b => Some('\u{261e}'),
        0x48 => Some('\u{2605}'),
        0x6c => Some('\u{25cf}'),
        0x6e => Some('\u{25a0}'),
        0x73 => Some('\u{25b2}'),
        0x74 => Some('\u{25bc}'),
        0x75 => Some('\u{25c6}'),
        0x77 => Some('\u{25d7}'),
        0x21..=0x7e => offset(0x2701, 0x21),
        0xa1..=0xa7 => offset(0x2761, 0xa1),
        0xa8 => Some('\u{2663}'),
        0xa9 => Some('\u{2666}'),
        0xaa => Some('\u{2665}'),
        0xab => Some('\u{2660}'),
        0xac..=0xb5 => offset(0x2460, 0xac),
        0xb6..=0xd4 => offset(0x2776, 0xb6),
        0xd5 => Some('\u{2192}'),
        0xd6 => Some('\u{2194}'),
        0xd7 => Some('\u{2195}'),
        0xd8..=0xef => offset(0x2798, 0xd8),
        0xf1..=0xfe => offset(0x27b1, 0xf1),
        _ => None,
    }
}

/// Generic CSS family for the family class of a LOGFONT `PitchAndFamily`
/// byte, whose high four bits hold a FamilyFont value ([MS-WMF] 2.1.1.8,
/// 2.2.2.14).
fn class_generic_family(pitch_and_family: u8) -> Option<&'static str> {
    Some(match pitch_and_family >> 4 {
        1 => "serif",
        2 => "sans-serif",
        3 => "monospace",
        4 => "cursive",
        5 => "fantasy",
        _ => return None,
    })
}

/// Generic CSS family of common Windows and Macintosh faces, for fonts whose
/// LOGFONT leaves the family class at FF_DONTCARE.
fn face_generic_family(face_name: &str) -> Option<&'static str> {
    const SERIF: &[&str] = &[
        "Times New Roman",
        "Times",
        "Georgia",
        "Garamond",
        "Cambria",
        "Book Antiqua",
        "Bookman Old Style",
        "Century Schoolbook",
        "Palatino",
        "Palatino Linotype",
        "MS Serif",
    ];
    const SANS_SERIF: &[&str] = &[
        "Arial",
        "Arial Black",
        "Arial Narrow",
        "Helvetica",
        "Helv",
        "Calibri",
        "Tahoma",
        "Verdana",
        "Trebuchet MS",
        "Segoe UI",
        "Century Gothic",
        "Franklin Gothic Medium",
        "Geneva",
        "Microsoft Sans Serif",
        "MS Sans Serif",
    ];
    const MONOSPACE: &[&str] = &[
        "Courier New",
        "Courier",
        "Consolas",
        "Lucida Console",
        "Monaco",
    ];
    let known = |names: &[&str]| {
        names
            .iter()
            .any(|name| name.eq_ignore_ascii_case(face_name))
    };
    if known(SERIF) {
        Some("serif")
    } else if known(SANS_SERIF) {
        Some("sans-serif")
    } else if known(MONOSPACE) {
        Some("monospace")
    } else {
        None
    }
}

/// SVG `font-family` list (not yet XML-escaped) for a recorded face.
///
/// The recorded face comes first, so a consumer that has it draws with it.
/// The generic family of the LOGFONT family class, or of a well-known face
/// when the class is not recorded, follows it so that any consumer falls back
/// to a face of the same kind. A face name is quoted when CSS would not read
/// it as one family name.
pub(crate) fn svg_font_family(face_name: &str, pitch_and_family: u8) -> String {
    let mut family = String::with_capacity(face_name.len() + 16);
    let plain = face_name
        .split(' ')
        .all(|word| word.chars().next().is_some_and(char::is_alphabetic))
        && face_name
            .chars()
            .all(|character| character.is_alphanumeric() || character == ' ' || character == '-');
    let keyword = [
        "serif",
        "sans-serif",
        "monospace",
        "cursive",
        "fantasy",
        "inherit",
        "initial",
        "unset",
        "default",
    ]
    .iter()
    .any(|keyword| keyword.eq_ignore_ascii_case(face_name));
    if plain && !keyword {
        family.push_str(face_name);
    } else {
        // CSS strings take either quote; a quote of the delimiting kind is
        // dropped from the name.
        let quote = if face_name.contains('\'') { '"' } else { '\'' };
        family.push(quote);
        family.extend(face_name.chars().filter(|&character| character != quote));
        family.push(quote);
    }
    let generic = class_generic_family(pitch_and_family).or_else(|| face_generic_family(face_name));
    if let Some(generic) = generic {
        family.push_str(", ");
        family.push_str(generic);
    }
    family
}

/// Collects decoded characters together with the code unit each one starts
/// at.
struct TextBuilder {
    text: String,
    unit_starts: Vec<usize>,
}

impl TextBuilder {
    fn with_capacity(units: usize) -> Self {
        Self {
            text: String::with_capacity(units),
            unit_starts: Vec::with_capacity(units),
        }
    }

    /// Appends the characters decoded from the code units starting at
    /// `unit`. A NUL draws nothing and is left out; the advance recorded for
    /// it still moves the characters after it.
    fn push_str(&mut self, unit: usize, decoded: &str) {
        for character in decoded.chars() {
            self.push(unit, character);
        }
    }

    fn push(&mut self, unit: usize, character: char) {
        if character != '\0' {
            self.text.push(character);
            self.unit_starts.push(unit);
        }
    }

    fn finish(self, exact: bool, unicode_face: bool) -> DecodedText {
        DecodedText {
            text: self.text,
            unit_starts: self.unit_starts,
            exact,
            unicode_face,
        }
    }
}

/// Decodes one text record's bytes with the character set of the font that is
/// selected for it.
pub(crate) fn decode_logfont_text(bytes: &[u8], charset: u8, face_name: &str) -> DecodedText {
    if let Some(face) = symbol_face(face_name) {
        return decode_symbol_text(bytes, face);
    }
    if charset == charset::SYMBOL {
        return decode_symbol_text(bytes, SymbolFace::PrivateUse(no_equivalent));
    }
    let mut text = TextBuilder::with_capacity(bytes.len());
    let Some(page) = code_page(charset).and_then(Page::new) else {
        // The character set names a code page this crate cannot provide
        // (JOHAB, OEM). Windows-1252 keeps the picture readable for the ASCII
        // range, but the non-ASCII range is a guess.
        for (index, &byte) in bytes.iter().enumerate() {
            text.push_str(index, &Page::WINDOWS_1252.recover(&[byte]).0);
        }
        return text.finish(bytes.iter().all(u8::is_ascii), false);
    };
    if MULTIBYTE_CODE_PAGES.contains(&page.id()) {
        let mut exact = true;
        let mut index = 0;
        while index < bytes.len() {
            let length = if is_lead_byte(page.id(), bytes[index]) {
                2.min(bytes.len() - index)
            } else {
                1
            };
            let (decoded, recovered) = page.recover(&bytes[index..index + length]);
            exact &= !recovered;
            text.push_str(index, &decoded);
            index += length;
        }
        return text.finish(exact, false);
    }
    for (index, &byte) in bytes.iter().enumerate() {
        let buffer = [byte];
        let (decoded, unassigned) = page.recover(&buffer);
        // A code page can leave positions unassigned (0x81 and 0x8D in
        // windows-1252, 0xA1 and 0xA5 in windows-1257). Windows converts such
        // a position to the same code point rather than to a replacement
        // character, so the byte's own value keeps the text exact and the
        // picture unchanged.
        if unassigned {
            text.push(index, char::from(byte));
        } else {
            text.push_str(index, &decoded);
        }
    }
    text.finish(true, false)
}

/// Double-byte code pages, whose characters take one or two bytes.
const MULTIBYTE_CODE_PAGES: [u32; 4] = [932, 936, 949, 950];

/// Whether `byte` starts a two-byte character in double-byte code page
/// `page`, following the lead-byte ranges of the Windows code page tables
/// ([MS-UCODEREF] 2.2.2.1.3).
fn is_lead_byte(page: u32, byte: u8) -> bool {
    match page {
        932 => matches!(byte, 0x81..=0x9f | 0xe0..=0xfc),
        936 | 949 | 950 => matches!(byte, 0x81..=0xfe),
        _ => false,
    }
}

/// Decodes the UTF-16 string of a Unicode text record drawn with `face_name`.
///
/// Windows reaches the glyphs of a symbol face through the U+F000 private-use
/// range, and it also accepts the plain 0x20 through 0xFF positions there, so
/// a record that draws with a symbol face names its characters as U+F0xx (or
/// U+00xx). They are mapped to the characters those glyphs draw; any other
/// character, and any other face, keeps its text.
pub(crate) fn decode_utf16_logfont_text(units: &[u16], face_name: &str) -> DecodedText {
    let face = symbol_face(face_name);
    let mut text = TextBuilder::with_capacity(units.len());
    let mut index = 0;
    for decoded in char::decode_utf16(units.iter().copied()) {
        let character = decoded.unwrap_or(char::REPLACEMENT_CHARACTER);
        let position = match u32::from(character) {
            value @ (0x20..=0xff | 0xf020..=0xf0ff) => u8::try_from(value & 0xff).ok(),
            _ => None,
        };
        let drawn = match (face, position) {
            (Some(SymbolFace::Adobe), Some(position)) => adobe_symbol(position),
            (Some(SymbolFace::PrivateUse(equivalent)), Some(position)) => {
                private_use_glyph(equivalent, position)
            },
            _ => character,
        };
        text.push(index, drawn);
        index += character.len_utf16();
    }
    text.finish(true, matches!(face, Some(SymbolFace::Adobe)))
}

/// Decodes symbol-font text.
///
/// A symbol face draws glyph *n* for byte *n*, so the text is not a code-page
/// string: the Adobe Symbol order maps onto real Unicode characters. Every
/// other symbol face draws the standard character of a glyph where it has
/// one, and otherwise reports the glyph in the Unicode private-use area,
/// exactly where its Windows cmap places it; the recorded face stays first in
/// the font family so that a system that has it draws the glyphs left there.
fn decode_symbol_text(bytes: &[u8], face: SymbolFace) -> DecodedText {
    let mut text = TextBuilder::with_capacity(bytes.len());
    for (index, &byte) in bytes.iter().enumerate() {
        let drawn = match face {
            SymbolFace::Adobe => adobe_symbol(byte),
            SymbolFace::PrivateUse(equivalent) => private_use_glyph(equivalent, byte),
        };
        text.push(index, drawn);
    }
    text.finish(true, matches!(face, SymbolFace::Adobe))
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
    fn private_use_faces_draw_the_standard_equivalent_of_their_glyphs() {
        let bullets = decode(b"\x6c\x9f\xa7\xd8\xfc\x21", charset::SYMBOL, "Wingdings");
        assert_eq!(bullets.text, "●•▪➢✓\u{f021}");
        assert!(bullets.exact);
        assert!(!bullets.unicode_face, "the recorded face stays selectable");

        // MathType records MT Extra with character set 160.
        let math = decode(b"\x4c\x68\x6c\x72\x67 ", 160, "MT Extra");
        assert_eq!(math.text, "⋯ℏℓ\u{20d7}\u{f067} ");
        assert!(!math.unicode_face);

        let sorts = decode(b"\x21\x34\x6c\x75\x76\xac\xd4\xfe", 2, "Monotype Sorts");
        assert_eq!(sorts.text, "✁✔●◆❖①➔➾");
        assert_eq!(decode(b"\x70\x75", 2, "Wingdings 3").text, "▲▶");
        assert_eq!(decode(b"\x34\x61", 2, "Webdings").text, "▶\u{f061}");
        assert_eq!(decode(b"\x54", 2, "Wingdings 2").text, "\u{f054}");
        // A symbol character set without a known face keeps every glyph.
        assert_eq!(decode(b"\x6c", charset::SYMBOL, "Unknown").text, "\u{f06c}");
    }

    #[test]
    fn unicode_text_in_a_symbol_face_maps_its_glyph_positions() {
        let units = |text: &str| text.encode_utf16().collect::<Vec<_>>();
        let mapped = decode_utf16_logfont_text(&units("\u{f070}\u{f065}a\u{3b1}"), "Symbol");
        assert_eq!(mapped.text, "πεαα");
        assert!(mapped.unicode_face);

        let dingbat = decode_utf16_logfont_text(&units("\u{f06c}\u{f0fc}l"), "Wingdings");
        assert_eq!(dingbat.text, "●✓●");
        assert!(!dingbat.unicode_face);

        let text = decode_utf16_logfont_text(&units("\u{f06c}"), "Arial");
        assert_eq!(text.text, "\u{f06c}");
    }

    #[test]
    fn nul_bytes_draw_nothing_but_keep_their_advance() {
        let text = decode(b"a\0b", charset::ANSI, "");
        assert_eq!(text.text, "ab");
        assert_eq!(text.unit_starts, [0, 2]);
        let units = [0x61, 0, 0x62, 0];
        let wide = decode_utf16_logfont_text(&units, "Arial");
        assert_eq!(wide.text, "ab");
        assert_eq!(wide.unit_starts, [0, 2]);
    }

    #[test]
    fn characters_record_the_code_unit_they_start_at() {
        let text = decode(b"a\x82\xa0b\x82", charset::SHIFT_JIS, "");
        assert_eq!(text.text, "aあb\u{fffd}");
        assert_eq!(text.unit_starts, [0, 1, 3, 4]);
        assert!(!text.exact);

        let units = "a\u{1d400}b".encode_utf16().collect::<Vec<_>>();
        let wide = decode_utf16_logfont_text(&units, "Arial");
        assert_eq!(wide.text, "a\u{1d400}b");
        assert_eq!(wide.unit_starts, [0, 1, 3]);
    }

    #[test]
    fn origin_offsets_sum_the_advances_before_each_character() {
        let text = decode(b"a\x82\xa0b", charset::SHIFT_JIS, "");
        let advances = [(10.0, 0.0), (6.0, 1.0), (6.0, 1.0), (8.0, 0.0)];
        assert_eq!(
            text.origin_offsets(&advances),
            [(0.0, 0.0), (10.0, 0.0), (22.0, 2.0)]
        );
        // Missing advances leave the remaining characters unplaced.
        assert_eq!(
            text.origin_offsets(&advances[..2]),
            [(0.0, 0.0), (10.0, 0.0)]
        );
    }

    #[test]
    fn font_family_lists_the_face_then_its_generic_family() {
        assert_eq!(svg_font_family("Arial", 0x22), "Arial, sans-serif");
        assert_eq!(svg_font_family("Arial", 0), "Arial, sans-serif");
        assert_eq!(svg_font_family("Some Face", 0x12), "Some Face, serif");
        assert_eq!(svg_font_family("Some Face", 0x31), "Some Face, monospace");
        assert_eq!(svg_font_family("Some Face", 0x42), "Some Face, cursive");
        assert_eq!(svg_font_family("Some Face", 0x52), "Some Face, fantasy");
        assert_eq!(svg_font_family("Some Face", 0x02), "Some Face");
        assert_eq!(svg_font_family("3 of 9", 0), "'3 of 9'");
        assert_eq!(svg_font_family("serif", 0x10), "'serif', serif");
        assert_eq!(svg_font_family("O'Neil, Sans", 0), "\"O'Neil, Sans\"");
    }

    #[test]
    fn characters_that_cannot_reach_the_canvas_are_removed() {
        let mut text = decode(b"abcd", charset::ANSI, "");
        let origins = [(0.0, 0.0), (10.0, 0.0), (1.0e9, 0.0), (20.0, 0.0)];
        let canvas = (0.0, 0.0, 100.0, 100.0);
        let kept = text.retain_visible(&origins, |origin| {
            origin_reaches_canvas(origin, (0.0, 0.0), canvas, 40.0)
        });
        assert_eq!(text.text, "abd");
        assert_eq!(text.unit_starts, [0, 1, 3]);
        assert_eq!(kept, [(0.0, 0.0), (10.0, 0.0), (20.0, 0.0)]);

        // Characters without an origin follow the last placed one.
        let mut text = decode(b"abcd", charset::ANSI, "");
        let kept = text.retain_visible(&[(0.0, 0.0), (f64::NAN, 0.0)], |origin| {
            origin_reaches_canvas(origin, (0.0, 0.0), canvas, 40.0)
        });
        assert_eq!(text.text, "a");
        assert_eq!(kept, [(0.0, 0.0)]);
    }

    #[test]
    fn reachability_holds_under_rotation_about_the_reference_point() {
        let canvas = (0.0, 0.0, 100.0, 100.0);
        // Right of the canvas, but a rotation of 45 degrees about the corner
        // brings it over the canvas.
        assert!(origin_reaches_canvas(
            (130.0, 0.0),
            (0.0, 0.0),
            canvas,
            10.0
        ));
        // No rotation about the center brings a point 100 units away onto a
        // canvas whose corners are 71 units away.
        assert!(!origin_reaches_canvas(
            (150.0, 50.0),
            (50.0, 50.0),
            canvas,
            10.0
        ));
    }
}
