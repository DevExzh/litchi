//! Table integrity: sorted lower-case keys, the named-color sources, and the
//! vocabulary counts.

use super::super::color::names::{
    DRAWING_SCHEMES, NAMED, Origin, SPREADSHEET_THEMES, SYSTEMS, TRANSFORMS, WORD_SCHEMES,
};
use super::super::text::{is_lower_ascii, str_lt};
use super::super::{Rgba, System};

/// OfficeCLI's `ParseHelpers.NamedColors` at `dced0d74`: 148 CSS names plus
/// `transparent`, transcribed as the parity oracle.
const OFFICECLI_NAMED: &[(&str, u32)] = &[
    ("aliceblue", 0xF0F8FF),
    ("antiquewhite", 0xFAEBD7),
    ("aqua", 0x00FFFF),
    ("aquamarine", 0x7FFFD4),
    ("azure", 0xF0FFFF),
    ("beige", 0xF5F5DC),
    ("bisque", 0xFFE4C4),
    ("black", 0x000000),
    ("blanchedalmond", 0xFFEBCD),
    ("blue", 0x0000FF),
    ("blueviolet", 0x8A2BE2),
    ("brown", 0xA52A2A),
    ("burlywood", 0xDEB887),
    ("cadetblue", 0x5F9EA0),
    ("chartreuse", 0x7FFF00),
    ("chocolate", 0xD2691E),
    ("coral", 0xFF7F50),
    ("cornflowerblue", 0x6495ED),
    ("cornsilk", 0xFFF8DC),
    ("crimson", 0xDC143C),
    ("cyan", 0x00FFFF),
    ("darkblue", 0x00008B),
    ("darkcyan", 0x008B8B),
    ("darkgoldenrod", 0xB8860B),
    ("darkgray", 0xA9A9A9),
    ("darkgreen", 0x006400),
    ("darkgrey", 0xA9A9A9),
    ("darkkhaki", 0xBDB76B),
    ("darkmagenta", 0x8B008B),
    ("darkolivegreen", 0x556B2F),
    ("darkorange", 0xFF8C00),
    ("darkorchid", 0x9932CC),
    ("darkred", 0x8B0000),
    ("darksalmon", 0xE9967A),
    ("darkseagreen", 0x8FBC8F),
    ("darkslateblue", 0x483D8B),
    ("darkslategray", 0x2F4F4F),
    ("darkslategrey", 0x2F4F4F),
    ("darkturquoise", 0x00CED1),
    ("darkviolet", 0x9400D3),
    ("deeppink", 0xFF1493),
    ("deepskyblue", 0x00BFFF),
    ("dimgray", 0x696969),
    ("dimgrey", 0x696969),
    ("dodgerblue", 0x1E90FF),
    ("firebrick", 0xB22222),
    ("floralwhite", 0xFFFAF0),
    ("forestgreen", 0x228B22),
    ("fuchsia", 0xFF00FF),
    ("gainsboro", 0xDCDCDC),
    ("ghostwhite", 0xF8F8FF),
    ("gold", 0xFFD700),
    ("goldenrod", 0xDAA520),
    ("gray", 0x808080),
    ("green", 0x008000),
    ("greenyellow", 0xADFF2F),
    ("grey", 0x808080),
    ("honeydew", 0xF0FFF0),
    ("hotpink", 0xFF69B4),
    ("indianred", 0xCD5C5C),
    ("indigo", 0x4B0082),
    ("ivory", 0xFFFFF0),
    ("khaki", 0xF0E68C),
    ("lavender", 0xE6E6FA),
    ("lavenderblush", 0xFFF0F5),
    ("lawngreen", 0x7CFC00),
    ("lemonchiffon", 0xFFFACD),
    ("lightblue", 0xADD8E6),
    ("lightcoral", 0xF08080),
    ("lightcyan", 0xE0FFFF),
    ("lightgoldenrodyellow", 0xFAFAD2),
    ("lightgray", 0xD3D3D3),
    ("lightgreen", 0x90EE90),
    ("lightgrey", 0xD3D3D3),
    ("lightpink", 0xFFB6C1),
    ("lightsalmon", 0xFFA07A),
    ("lightseagreen", 0x20B2AA),
    ("lightskyblue", 0x87CEFA),
    ("lightslategray", 0x778899),
    ("lightslategrey", 0x778899),
    ("lightsteelblue", 0xB0C4DE),
    ("lightyellow", 0xFFFFE0),
    ("lime", 0x00FF00),
    ("limegreen", 0x32CD32),
    ("linen", 0xFAF0E6),
    ("magenta", 0xFF00FF),
    ("maroon", 0x800000),
    ("mediumaquamarine", 0x66CDAA),
    ("mediumblue", 0x0000CD),
    ("mediumorchid", 0xBA55D3),
    ("mediumpurple", 0x9370DB),
    ("mediumseagreen", 0x3CB371),
    ("mediumslateblue", 0x7B68EE),
    ("mediumspringgreen", 0x00FA9A),
    ("mediumturquoise", 0x48D1CC),
    ("mediumvioletred", 0xC71585),
    ("midnightblue", 0x191970),
    ("mintcream", 0xF5FFFA),
    ("mistyrose", 0xFFE4E1),
    ("moccasin", 0xFFE4B5),
    ("navajowhite", 0xFFDEAD),
    ("navy", 0x000080),
    ("oldlace", 0xFDF5E6),
    ("olive", 0x808000),
    ("olivedrab", 0x6B8E23),
    ("orange", 0xFFA500),
    ("orangered", 0xFF4500),
    ("orchid", 0xDA70D6),
    ("palegoldenrod", 0xEEE8AA),
    ("palegreen", 0x98FB98),
    ("paleturquoise", 0xAFEEEE),
    ("palevioletred", 0xDB7093),
    ("papayawhip", 0xFFEFD5),
    ("peachpuff", 0xFFDAB9),
    ("peru", 0xCD853F),
    ("pink", 0xFFC0CB),
    ("plum", 0xDDA0DD),
    ("powderblue", 0xB0E0E6),
    ("purple", 0x800080),
    ("rebeccapurple", 0x663399),
    ("red", 0xFF0000),
    ("rosybrown", 0xBC8F8F),
    ("royalblue", 0x4169E1),
    ("saddlebrown", 0x8B4513),
    ("salmon", 0xFA8072),
    ("sandybrown", 0xF4A460),
    ("seagreen", 0x2E8B57),
    ("seashell", 0xFFF5EE),
    ("sienna", 0xA0522D),
    ("silver", 0xC0C0C0),
    ("skyblue", 0x87CEEB),
    ("slateblue", 0x6A5ACD),
    ("slategray", 0x708090),
    ("slategrey", 0x708090),
    ("snow", 0xFFFAFA),
    ("springgreen", 0x00FF7F),
    ("steelblue", 0x4682B4),
    ("tan", 0xD2B48C),
    ("teal", 0x008080),
    ("thistle", 0xD8BFD8),
    ("tomato", 0xFF6347),
    ("transparent", 0x000000),
    ("turquoise", 0x40E0D0),
    ("violet", 0xEE82EE),
    ("wheat", 0xF5DEB3),
    ("white", 0xFFFFFF),
    ("whitesmoke", 0xF5F5F5),
    ("yellow", 0xFFFF00),
    ("yellowgreen", 0x9ACD32),
];

/// The ISO/IEC 29500-4:2016 `ST_PresetColorVal` enumeration, in schema order.
pub(super) const PRESET_TOKENS: &[&str] = &[
    "aliceBlue",
    "antiqueWhite",
    "aqua",
    "aquamarine",
    "azure",
    "beige",
    "bisque",
    "black",
    "blanchedAlmond",
    "blue",
    "blueViolet",
    "brown",
    "burlyWood",
    "cadetBlue",
    "chartreuse",
    "chocolate",
    "coral",
    "cornflowerBlue",
    "cornsilk",
    "crimson",
    "cyan",
    "darkBlue",
    "darkCyan",
    "darkGoldenrod",
    "darkGray",
    "darkGrey",
    "darkGreen",
    "darkKhaki",
    "darkMagenta",
    "darkOliveGreen",
    "darkOrange",
    "darkOrchid",
    "darkRed",
    "darkSalmon",
    "darkSeaGreen",
    "darkSlateBlue",
    "darkSlateGray",
    "darkSlateGrey",
    "darkTurquoise",
    "darkViolet",
    "dkBlue",
    "dkCyan",
    "dkGoldenrod",
    "dkGray",
    "dkGrey",
    "dkGreen",
    "dkKhaki",
    "dkMagenta",
    "dkOliveGreen",
    "dkOrange",
    "dkOrchid",
    "dkRed",
    "dkSalmon",
    "dkSeaGreen",
    "dkSlateBlue",
    "dkSlateGray",
    "dkSlateGrey",
    "dkTurquoise",
    "dkViolet",
    "deepPink",
    "deepSkyBlue",
    "dimGray",
    "dimGrey",
    "dodgerBlue",
    "firebrick",
    "floralWhite",
    "forestGreen",
    "fuchsia",
    "gainsboro",
    "ghostWhite",
    "gold",
    "goldenrod",
    "gray",
    "grey",
    "green",
    "greenYellow",
    "honeydew",
    "hotPink",
    "indianRed",
    "indigo",
    "ivory",
    "khaki",
    "lavender",
    "lavenderBlush",
    "lawnGreen",
    "lemonChiffon",
    "lightBlue",
    "lightCoral",
    "lightCyan",
    "lightGoldenrodYellow",
    "lightGray",
    "lightGrey",
    "lightGreen",
    "lightPink",
    "lightSalmon",
    "lightSeaGreen",
    "lightSkyBlue",
    "lightSlateGray",
    "lightSlateGrey",
    "lightSteelBlue",
    "lightYellow",
    "ltBlue",
    "ltCoral",
    "ltCyan",
    "ltGoldenrodYellow",
    "ltGray",
    "ltGrey",
    "ltGreen",
    "ltPink",
    "ltSalmon",
    "ltSeaGreen",
    "ltSkyBlue",
    "ltSlateGray",
    "ltSlateGrey",
    "ltSteelBlue",
    "ltYellow",
    "lime",
    "limeGreen",
    "linen",
    "magenta",
    "maroon",
    "medAquamarine",
    "medBlue",
    "medOrchid",
    "medPurple",
    "medSeaGreen",
    "medSlateBlue",
    "medSpringGreen",
    "medTurquoise",
    "medVioletRed",
    "mediumAquamarine",
    "mediumBlue",
    "mediumOrchid",
    "mediumPurple",
    "mediumSeaGreen",
    "mediumSlateBlue",
    "mediumSpringGreen",
    "mediumTurquoise",
    "mediumVioletRed",
    "midnightBlue",
    "mintCream",
    "mistyRose",
    "moccasin",
    "navajoWhite",
    "navy",
    "oldLace",
    "olive",
    "oliveDrab",
    "orange",
    "orangeRed",
    "orchid",
    "paleGoldenrod",
    "paleGreen",
    "paleTurquoise",
    "paleVioletRed",
    "papayaWhip",
    "peachPuff",
    "peru",
    "pink",
    "plum",
    "powderBlue",
    "purple",
    "red",
    "rosyBrown",
    "royalBlue",
    "saddleBrown",
    "salmon",
    "sandyBrown",
    "seaGreen",
    "seaShell",
    "sienna",
    "silver",
    "skyBlue",
    "slateBlue",
    "slateGray",
    "slateGrey",
    "snow",
    "springGreen",
    "steelBlue",
    "tan",
    "teal",
    "thistle",
    "tomato",
    "turquoise",
    "violet",
    "wheat",
    "white",
    "whiteSmoke",
    "yellow",
    "yellowGreen",
];

fn sorted<T>(table: &[T], key: impl Fn(&T) -> &str) -> bool {
    table.iter().all(|entry| is_lower_ascii(key(entry)))
        && table
            .windows(2)
            .all(|pair| str_lt(key(&pair[0]), key(&pair[1])))
}

#[test]
fn every_table_is_sorted_and_lower_case() {
    assert!(sorted(NAMED, |named| named.name));
    assert!(sorted(DRAWING_SCHEMES, |(name, _)| name));
    assert!(sorted(WORD_SCHEMES, |(name, _)| name));
    assert!(sorted(SPREADSHEET_THEMES, |(name, _)| name));
    assert!(sorted(SYSTEMS, |(name, _)| name));
    assert!(sorted(TRANSFORMS, |(name, _)| name));
}

#[test]
fn named_colors_match_officecli_exactly() {
    assert_eq!(OFFICECLI_NAMED.len(), 149);
    for &(name, rgb) in OFFICECLI_NAMED {
        let entry = NAMED
            .iter()
            .find(|named| named.name == name)
            .unwrap_or_else(|| panic!("{name} missing"));
        let [r, g, b] = entry.rgb;
        assert_eq!(u32::from_be_bytes([0, r, g, b]), rgb, "{name}");
        let expected_origin = match name {
            "transparent" => Origin::Transparent,
            "rebeccapurple" => Origin::Css,
            _ => Origin::CssAndPreset,
        };
        assert_eq!(entry.origin, expected_origin, "{name}");
    }
    let css = NAMED
        .iter()
        .filter(|named| matches!(named.origin, Origin::Css | Origin::CssAndPreset))
        .count();
    assert_eq!(css, 148, "148 CSS named colors besides transparent");
    assert_eq!(
        NAMED.len(),
        192,
        "148 CSS + transparent + 43 legacy presets"
    );
}

/// The long DrawingML spelling of a legacy preset (`dkBlue` -> `darkBlue`).
fn long_form(token: &str) -> String {
    for (short, long) in [("dk", "dark"), ("lt", "light"), ("med", "medium")] {
        if let Some(rest) = token.strip_prefix(short) {
            if rest.starts_with(|c: char| c.is_ascii_uppercase()) {
                return format!("{long}{rest}");
            }
        }
    }
    token.to_owned()
}

#[test]
fn every_preset_token_resolves() {
    assert_eq!(PRESET_TOKENS.len(), 190);
    let mut legacy = 0;
    for token in PRESET_TOKENS {
        let color = Rgba::preset(token).unwrap_or_else(|| panic!("{token}"));
        let long = long_form(token);
        if long != *token {
            legacy += 1;
        }
        // Every preset, legacy spellings included, has the CSS color of its
        // long form.
        assert_eq!(Rgba::named(&long), Some(color), "{token}");
        let entry = NAMED
            .iter()
            .find(|named| named.name.eq_ignore_ascii_case(token))
            .expect("listed");
        assert_eq!(
            entry.preset,
            Some(*token),
            "the exact schema spelling is kept"
        );
    }
    assert_eq!(legacy, 43);
    assert_eq!(
        NAMED.iter().filter(|named| named.preset.is_some()).count(),
        190
    );
    assert_eq!(Rgba::preset("rebeccapurple"), None);
    assert_eq!(Rgba::preset("transparent"), None);
    assert_eq!(Rgba::preset("DKBLUE"), Rgba::named("darkblue"));
}

#[test]
fn css_reference_values() {
    // Spot values from CSS Color Module Level 4, independent of both tables.
    let cases: &[(&str, [u8; 3])] = &[
        ("rebeccapurple", [0x66, 0x33, 0x99]),
        ("darkgray", [0xA9, 0xA9, 0xA9]),
        ("gray", [0x80, 0x80, 0x80]),
        ("lightgray", [0xD3, 0xD3, 0xD3]),
        ("green", [0x00, 0x80, 0x00]),
        ("lime", [0x00, 0xFF, 0x00]),
        ("navy", [0x00, 0x00, 0x80]),
        ("white", [0xFF, 0xFF, 0xFF]),
        ("black", [0x00, 0x00, 0x00]),
        ("cornflowerblue", [0x64, 0x95, 0xED]),
    ];
    for &(name, [r, g, b]) in cases {
        assert_eq!(Rgba::named(name), Some(Rgba::new(r, g, b)), "{name}");
    }
    assert_eq!(
        Rgba::named("transparent"),
        Some(Rgba::new(0, 0, 0).with_alpha(0))
    );
    assert_eq!(Rgba::named("TRANSPARENT"), Rgba::named("transparent"));
    assert_eq!(Rgba::named("grey"), Rgba::named("gray"));
    assert_eq!(Rgba::named("not-a-color"), None);
    assert_eq!(Rgba::named(""), None);
}

#[test]
fn vocabulary_counts() {
    assert_eq!(SYSTEMS.len(), 30, "ST_SystemColorVal");
    for &(name, system) in SYSTEMS {
        assert!(system.token().eq_ignore_ascii_case(name), "{name}");
        assert_eq!(System::from_token(system.token()), Some(system));
    }
    assert_eq!(TRANSFORMS.len(), 9);
    assert_eq!(DRAWING_SCHEMES.len(), 27);
    assert_eq!(WORD_SCHEMES.len(), 28);
    assert_eq!(SPREADSHEET_THEMES.len(), 26);
}
