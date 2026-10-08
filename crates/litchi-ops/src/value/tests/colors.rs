//! Colors: OfficeCLI `ParseHelpers` (`TryResolveColorInput`,
//! `NormalizeArgbColor`, `SanitizeColorForOoxml`, `SplitExcelColorTint`,
//! `NormalizeSchemeColorName`), `DrawingColorBuilder` and the Word run color.

use super::super::color::names::NAMED;
use super::super::{
    Base, Color, ColorContext, ErrorKind, Rgba, Scheme, System, ThemeSlot, Transform,
    TransformKind, What,
};
use super::kind;
use super::tables::PRESET_TOKENS;

const CONTEXTS: [ColorContext; 3] = [
    ColorContext::Drawing,
    ColorContext::Spreadsheet,
    ColorContext::Word,
];

fn parse(text: &str, context: ColorContext) -> Color {
    Color::parse(text, context).unwrap_or_else(|error| panic!("{text:?} in {context:?}: {error}"))
}

fn drawing(text: &str) -> String {
    parse(text, ColorContext::Drawing).to_string()
}

fn sheet(text: &str) -> String {
    parse(text, ColorContext::Spreadsheet).to_string()
}

fn word(text: &str) -> String {
    parse(text, ColorContext::Word).to_string()
}

fn rgb(text: &str, context: ColorContext) -> Rgba {
    match parse(text, context).base() {
        Base::Rgb(rgba) => rgba,
        other => panic!("{text:?}: expected RGB, got {other:?}"),
    }
}

fn fails(text: &str, context: ColorContext) -> ErrorKind {
    kind(Color::parse(text, context))
}

#[test]
fn every_named_color_in_every_context() {
    for named in NAMED {
        let [r, g, b] = named.rgb;
        let alpha = if named.name == "transparent" { 0 } else { 255 };
        let expected = Rgba::new(r, g, b).with_alpha(alpha);
        for context in CONTEXTS {
            assert_eq!(
                rgb(named.name, context),
                expected,
                "{} in {context:?}",
                named.name
            );
            assert_eq!(
                rgb(&named.name.to_ascii_uppercase(), context),
                expected,
                "{}",
                named.name
            );
        }
    }
}

#[test]
fn every_preset_token_in_every_context() {
    for token in PRESET_TOKENS {
        let expected = Rgba::preset(token).expect("preset");
        for context in CONTEXTS {
            assert_eq!(rgb(token, context), expected, "{token} in {context:?}");
        }
    }
}

#[test]
fn hex_forms() {
    let cases: &[(&str, Rgba)] = &[
        ("F00", Rgba::new(0xFF, 0, 0)),
        ("#f00", Rgba::new(0xFF, 0, 0)),
        ("##F00", Rgba::new(0xFF, 0, 0)),
        ("#F00A", Rgba::new(0xFF, 0, 0).with_alpha(0xAA)),
        ("FF0000", Rgba::new(0xFF, 0, 0)),
        ("#ff0000", Rgba::new(0xFF, 0, 0)),
        ("1F4E79", Rgba::new(0x1F, 0x4E, 0x79)),
        // Eight digits: AARRGGBB without `#` (OOXML), RRGGBBAA with it (CSS).
        ("80FF0000", Rgba::new(0xFF, 0, 0).with_alpha(0x80)),
        ("#80FF0000", Rgba::new(0x80, 0xFF, 0).with_alpha(0)),
        ("#FF000080", Rgba::new(0xFF, 0, 0).with_alpha(0x80)),
        ("FFFF0000", Rgba::new(0xFF, 0, 0)),
        ("  #00ff00  ", Rgba::new(0, 0xFF, 0)),
        ("\u{3000}abc\u{3000}", Rgba::new(0xAA, 0xBB, 0xCC)),
    ];
    for &(text, expected) in cases {
        for context in CONTEXTS {
            assert_eq!(rgb(text, context), expected, "{text:?} in {context:?}");
        }
    }
    // `#RGBA` needs the `#` and a non-zero alpha digit; other lengths fail.
    for text in [
        "#F000",
        "F00A",
        "FF000",
        "GG0000",
        "#",
        "##",
        "FF00000",
        "FF0000000",
        "#ff 0000",
    ] {
        for context in CONTEXTS {
            assert_eq!(
                fails(text, context),
                ErrorKind::Syntax,
                "{text:?} in {context:?}"
            );
        }
    }
    for context in CONTEXTS {
        assert_eq!(fails("", context), ErrorKind::Empty);
        assert_eq!(fails("  ", context), ErrorKind::Empty);
    }
}

#[test]
fn rgb_functions() {
    let cases: &[(&str, Rgba)] = &[
        ("rgb(255,0,0)", Rgba::new(255, 0, 0)),
        ("rgb( 255 , 0 , 0 )", Rgba::new(255, 0, 0)),
        ("rgb(255 0 0)", Rgba::new(255, 0, 0)),
        ("RGB(255,0,0)", Rgba::new(255, 0, 0)),
        ("rgb(255, 0 0)", Rgba::new(255, 0, 0)),
        ("rgb(100%,0%,0%)", Rgba::new(255, 0, 0)),
        // 50% of 255 is 127.5, which rounds to even: 128.
        ("rgb(50%,0,0)", Rgba::new(128, 0, 0)),
        ("rgba(255,0,0,0.5)", Rgba::new(255, 0, 0).with_alpha(128)),
        ("rgba(255 0 0 / 50%)", Rgba::new(255, 0, 0).with_alpha(128)),
        ("rgb(255,0,0 / 0.25)", Rgba::new(255, 0, 0).with_alpha(64)),
        ("rgb(255 0 0, 1)", Rgba::new(255, 0, 0)),
        ("rgba(10,20,30)", Rgba::new(10, 20, 30)),
        ("rgba(0,0,0,0)", Rgba::new(0, 0, 0).with_alpha(0)),
        ("rgb(+5,0,0)", Rgba::new(5, 0, 0)),
        ("rgb(1e2%,0,0)", Rgba::new(255, 0, 0)),
    ];
    for &(text, expected) in cases {
        // A `+` splits a DrawingML transform first, as in OfficeCLI.
        let contexts: &[ColorContext] = if text.contains('+') {
            &[ColorContext::Spreadsheet, ColorContext::Word]
        } else {
            &CONTEXTS
        };
        for &context in contexts {
            assert_eq!(rgb(text, context), expected, "{text:?} in {context:?}");
        }
    }
    assert_eq!(
        fails("rgb(+5,0,0)", ColorContext::Drawing),
        ErrorKind::Syntax
    );
    for (text, expected) in [
        ("rgb(256,0,0)", ErrorKind::Component),
        ("rgb(-1,0,0)", ErrorKind::Component),
        ("rgb(1.5,0,0)", ErrorKind::Component),
        ("rgb(101%,0,0)", ErrorKind::Component),
        ("rgb(1,2,3))", ErrorKind::Component),
        ("rgba(0,0,0,1.5)", ErrorKind::Alpha),
        ("rgba(0,0,0,101%)", ErrorKind::Alpha),
        ("rgba(0,0,0,-0.1)", ErrorKind::Alpha),
        // Not shaped like OfficeCLI's pattern: no function at all.
        ("rgb(1,2)", ErrorKind::Syntax),
        ("rgb(1 2 3 4)", ErrorKind::Syntax),
        ("rgb(1,,2,3)", ErrorKind::Syntax),
        ("rgb (1,2,3)", ErrorKind::Syntax),
        ("rgb(1,2,3", ErrorKind::Syntax),
        ("rgb()", ErrorKind::Syntax),
    ] {
        for context in CONTEXTS {
            assert_eq!(fails(text, context), expected, "{text:?} in {context:?}");
        }
    }
    let error = Color::parse("rgb(256,0,0)", ColorContext::Drawing).unwrap_err();
    assert_eq!(
        error.to_string(),
        "invalid color: RGB component the component must be 0-255 or 0%-100%"
    );
}

#[test]
fn hsl_functions() {
    let cases: &[(&str, Rgba)] = &[
        ("hsl(0,100%,50%)", Rgba::new(0xFF, 0, 0)),
        ("hsl(120, 100%, 25%)", Rgba::new(0, 0x80, 0)),
        ("hsl(240deg 100% 50%)", Rgba::new(0, 0, 0xFF)),
        ("HSL(240DEG,100%,50%)", Rgba::new(0, 0, 0xFF)),
        ("hsl(-120,100%,50%)", Rgba::new(0, 0, 0xFF)),
        ("hsl(480,100%,50%)", Rgba::new(0, 0xFF, 0)),
        ("hsl(360,100%,50%)", Rgba::new(0xFF, 0, 0)),
        ("hsl(30,100%,50%)", Rgba::new(0xFF, 0x80, 0)),
        ("hsl(0,0%,50%)", Rgba::new(0x80, 0x80, 0x80)),
        ("hsl(0,0%,100%)", Rgba::new(0xFF, 0xFF, 0xFF)),
        ("hsl(0,0%,0%)", Rgba::new(0, 0, 0)),
        ("hsl(210,50%,40%)", Rgba::new(0x33, 0x66, 0x99)),
        // 229.5 and 76.5 round to even.
        ("hsl(300,75%,60%)", Rgba::new(0xE6, 0x4C, 0xE6)),
        (
            "hsla(0,100%,50%,0.5)",
            Rgba::new(0xFF, 0, 0).with_alpha(128),
        ),
        (
            "hsl(0 100% 50% / 25%)",
            Rgba::new(0xFF, 0, 0).with_alpha(64),
        ),
    ];
    for &(text, expected) in cases {
        for context in CONTEXTS {
            assert_eq!(rgb(text, context), expected, "{text:?} in {context:?}");
        }
    }
    for (text, expected) in [
        ("hsl(0,100,50)", ErrorKind::HslPercent),
        ("hsl(0,101%,50%)", ErrorKind::HslPercent),
        ("hsl(0,50%,-1%)", ErrorKind::HslPercent),
        ("hsl(abc,100%,50%)", ErrorKind::Hue),
        ("hsl(1e999,100%,50%)", ErrorKind::Hue),
        ("hsla(0,100%,50%,2)", ErrorKind::Alpha),
        ("hsl(0,100%)", ErrorKind::Syntax),
    ] {
        for context in CONTEXTS {
            assert_eq!(fails(text, context), expected, "{text:?} in {context:?}");
        }
    }
}

#[test]
fn drawing_scheme_names_print_long() {
    let cases: &[(&str, &str)] = &[
        ("accent1", "accent1"),
        ("Accent6", "accent6"),
        ("dk1", "dark1"),
        ("dark1", "dark1"),
        ("lt1", "light1"),
        ("dk2", "dark2"),
        ("lt2", "light2"),
        ("tx1", "text1"),
        ("text2", "text2"),
        ("bg1", "background1"),
        ("BG2", "background2"),
        ("hlink", "hyperlink"),
        ("hyperlink", "hyperlink"),
        ("folHlink", "followedHyperlink"),
        ("followedhyperlink", "followedHyperlink"),
        ("phClr", "phClr"),
        ("#accent1", "accent1"),
        (" accent1 ", "accent1"),
    ];
    for &(text, expected) in cases {
        assert_eq!(drawing(text), expected, "{text:?}");
    }
    assert_eq!(
        parse("tx1", ColorContext::Drawing).base(),
        Base::Scheme(Scheme::Text1)
    );
    assert_eq!(
        parse("dk1", ColorContext::Drawing).base(),
        Base::Scheme(Scheme::Dark1)
    );
    assert_eq!(Scheme::Text1.drawing_token(), "tx1");
    assert_eq!(Scheme::FollowedHyperlink.drawing_token(), "folHlink");
}

#[test]
fn drawing_transforms() {
    let cases: &[(&str, &str)] = &[
        ("accent1+lumMod75", "accent1+lumMod75"),
        ("accent1+lumMod75+lumOff25", "accent1+lumMod75+lumOff25"),
        ("ACCENT1+LUMMOD75", "accent1+lumMod75"),
        ("accent1 + lumMod75 + lumOff25", "accent1+lumMod75+lumOff25"),
        ("accent1:lumMod75", "accent1+lumMod75"),
        ("accent1+lumMod=75000", "accent1+lumMod75"),
        ("accent1+lumMod=75500", "accent1+lumMod=75500"),
        ("accent1+lumMod=750", "accent1+lumMod=750"),
        ("accent1+satOff-10", "accent1+satOff-10"),
        ("accent1+satMod200", "accent1+satMod200"),
        ("accent1+satMod1000", "accent1+satMod1000"),
        ("accent1+lumMod=-1000000", "accent1+lumMod-1000"),
        ("accent1+shade=100000", "accent1+shade100"),
        ("accent1+tint0", "accent1+tint0"),
        ("accent2+shade60+alpha80", "accent2+shade60+alpha80"),
        ("accent1+hueMod50+hueOff10", "accent1+hueMod50+hueOff10"),
        ("accent1+lumMod 75", "accent1+lumMod75"),
        ("accent1++lumMod75+", "accent1+lumMod75"),
        ("dk1+lumMod50", "dark1+lumMod50"),
        ("FF0000+lumMod50", "#FF0000+lumMod50"),
        ("#FF000080+alpha50", "#FF000080+alpha50"),
        ("window+lumMod50", "window+lumMod50"),
        ("red+alpha50", "#FF0000+alpha50"),
    ];
    for &(text, expected) in cases {
        assert_eq!(drawing(text), expected, "{text:?}");
    }
    let color = parse("accent1+lumMod75+lumOff25", ColorContext::Drawing);
    assert_eq!(
        color.transforms(),
        &[
            Transform::new(TransformKind::LumMod, 75_000).expect("valid"),
            Transform::new(TransformKind::LumOff, 25_000).expect("valid"),
        ]
    );
    // `:` is only a separator after a theme name and only for one transform:
    // OfficeCLI splits at the first `+` when there is one.
    assert_eq!(
        fails("accent1:lumMod75+lumOff25", ColorContext::Drawing),
        ErrorKind::Syntax
    );
    assert_eq!(
        fails("FF0000:lumMod75", ColorContext::Drawing),
        ErrorKind::Syntax
    );
}

#[test]
fn drawing_transform_errors() {
    for (text, expected) in [
        ("accent1+satMod1001", ErrorKind::TransformRange),
        ("accent1+lumMod=1000001", ErrorKind::TransformRange),
        ("accent1+tint-5", ErrorKind::TransformRange),
        ("accent1+alpha101", ErrorKind::TransformRange),
        ("accent1+shade=-1", ErrorKind::TransformRange),
        ("accent1+lumMod99999999999", ErrorKind::TransformNumber),
        ("accent1+lumMod7.5", ErrorKind::TransformNumber),
        ("accent1+lumMod=x", ErrorKind::TransformNumber),
        ("accent1+foo50", ErrorKind::UnknownTransform),
        ("accent1+lumMd75", ErrorKind::UnknownTransform),
        ("accent1+75", ErrorKind::TransformName),
        ("accent1+=75", ErrorKind::TransformName),
        ("accent1+lumMod", ErrorKind::TransformValue),
        (
            "accent1+lumMod75+lumOff25+satMod1+satOff1+hueMod1+hueOff1+shade1+tint1+alpha1",
            ErrorKind::TooManyTransforms,
        ),
    ] {
        assert_eq!(fails(text, ColorContext::Drawing), expected, "{text:?}");
    }
    let eight = "accent1+lumMod75+lumOff25+satMod1+satOff1+hueMod1+hueOff1+shade1+tint1";
    assert_eq!(
        parse(eight, ColorContext::Drawing).transforms().len(),
        Color::MAX_TRANSFORMS
    );
    let error = Color::parse("accent1+lumMd75", ColorContext::Drawing).unwrap_err();
    assert_eq!(
        error.suggestion().as_deref(),
        Some("did you mean 'lumMod'?")
    );
    assert_eq!(error.to_string(), "invalid color: unknown color transform");
    let error = Color::parse("accent1+tint-5", ColorContext::Drawing).unwrap_err();
    assert_eq!(
        error.to_string(),
        "invalid color: the component is out of range (0 to 100 as a percentage, 0 to 100000 after '=')"
    );
    let error = Color::parse("accent1+75", ColorContext::Drawing).unwrap_err();
    assert!(
        error
            .suggestion()
            .is_some_and(|text| text.contains("lumMod"))
    );
}

#[test]
fn drawing_system_colors() {
    assert_eq!(drawing("window"), "window");
    assert_eq!(drawing("WINDOWTEXT"), "windowText");
    assert_eq!(drawing("3dDkShadow"), "3dDkShadow");
    assert_eq!(drawing("menuHighlight"), "menuHighlight");
    assert_eq!(
        parse("background", ColorContext::Drawing).base(),
        Base::System(System::Background)
    );
    assert_eq!(
        System::Highlight.default_last_color(),
        Some(Rgba::new(0x00, 0x78, 0xD7))
    );
    assert_eq!(
        System::GrayText.default_last_color(),
        Some(Rgba::new(0x6D, 0x6D, 0x6D))
    );
    assert_eq!(System::ScrollBar.default_last_color(), None);
    // `auto` has no DrawingML form.
    let error = Color::parse("auto", ColorContext::Drawing).unwrap_err();
    assert_eq!(error.kind(), ErrorKind::Auto);
    assert!(error.suggestion().is_some());
}

#[test]
fn spreadsheet_theme_slots_and_tints() {
    let cases: &[(&str, &str)] = &[
        ("accent1", "accent1"),
        ("lt1", "lt1"),
        ("light1", "lt1"),
        ("bg1", "lt1"),
        ("background1", "lt1"),
        ("dark1", "dk1"),
        ("dk1", "dk1"),
        ("tx1", "dk1"),
        ("text1", "dk1"),
        ("lt2", "lt2"),
        ("TX2", "dk2"),
        ("hyperlink", "hlink"),
        ("followedHyperlink", "folHlink"),
        ("accent1+tint40", "accent1+tint40"),
        ("dk2+tint-25", "dk2+tint-25"),
        ("accent1+TINT=40", "accent1+tint40"),
        ("accent1 + tint40", "accent1+tint40"),
        ("accent1+tint33.333", "accent1+tint33.33"),
        ("accent1+tint0.4", "accent1"),
        ("accent1+tint100", "accent1+tint100"),
        ("accent1+tint-100", "accent1+tint-100"),
        ("FF0000+tint40", "#FF0000+tint40"),
        ("red", "#FF0000"),
        ("#80FF0000", "#80FF0000"),
        ("transparent", "#00000000"),
    ];
    for &(text, expected) in cases {
        assert_eq!(sheet(text), expected, "{text:?}");
    }
    assert_eq!(
        parse("accent1+tint40", ColorContext::Spreadsheet).spreadsheet_tint(),
        Some(0.4)
    );
    assert_eq!(
        parse("dk2+tint-25", ColorContext::Spreadsheet).spreadsheet_tint(),
        Some(-0.25)
    );
    assert_eq!(
        parse("bg1", ColorContext::Spreadsheet).base(),
        Base::Theme(ThemeSlot::Light1)
    );
    for (text, expected) in [
        ("accent1+tint101", ErrorKind::TintRange),
        ("accent1+tint-100.5", ErrorKind::TintRange),
        ("accent1+tint", ErrorKind::Tint),
        ("accent1+tintabc", ErrorKind::Tint),
        ("accent1+lumMod75", ErrorKind::NoTransforms),
        ("phClr", ErrorKind::Syntax),
        ("windowText", ErrorKind::Syntax),
        ("auto", ErrorKind::Syntax),
        ("accnet1", ErrorKind::Syntax),
    ] {
        assert_eq!(fails(text, ColorContext::Spreadsheet), expected, "{text:?}");
    }
    let error = Color::parse("accent1+lumMod75", ColorContext::Spreadsheet).unwrap_err();
    assert!(
        error
            .suggestion()
            .is_some_and(|text| text.contains("+tint"))
    );
    for index in 0..12 {
        let slot = ThemeSlot::from_index(index).expect("in range");
        assert_eq!(slot.index(), index);
        assert_eq!(sheet(slot.name()), slot.name());
    }
    assert_eq!(ThemeSlot::from_index(12), None);
}

#[test]
fn word_colors() {
    let cases: &[(&str, &str)] = &[
        ("auto", "auto"),
        ("AUTO", "auto"),
        ("accent1", "accent1"),
        ("dk1", "dark1"),
        ("windowText", "dark1"),
        ("windowBackground", "light1"),
        ("tx1", "text1"),
        ("bg1", "background1"),
        ("hlink", "hyperlink"),
        ("FF0000", "#FF0000"),
        ("red", "#FF0000"),
        ("80FF0000", "#FF000080"),
    ];
    for &(text, expected) in cases {
        assert_eq!(word(text), expected, "{text:?}");
    }
    assert_eq!(parse("auto", ColorContext::Word), Color::AUTO);
    for (text, expected) in [
        ("#accent1", ErrorKind::Syntax),
        ("phClr", ErrorKind::Syntax),
        ("none", ErrorKind::Syntax),
        ("window", ErrorKind::Syntax),
        ("accent1+lumMod75", ErrorKind::NoTransforms),
        ("FF0000+tint40", ErrorKind::NoTransforms),
        ("FF0000;themeColor=accent1", ErrorKind::ThemeTail),
        ("rgb(300,0,0)", ErrorKind::Component),
    ] {
        assert_eq!(fails(text, ColorContext::Word), expected, "{text:?}");
    }
}

#[test]
fn suggestions_name_the_nearest_color() {
    let suggest = |text: &str, context| {
        Color::parse(text, context)
            .unwrap_err()
            .suggestion()
            .map(|text| text.into_owned())
    };
    assert_eq!(
        suggest("reed", ColorContext::Drawing).as_deref(),
        Some("did you mean 'red'?")
    );
    assert_eq!(
        suggest("bleu", ColorContext::Drawing).as_deref(),
        Some("did you mean 'blue'?")
    );
    assert_eq!(
        suggest("accnet1", ColorContext::Drawing).as_deref(),
        Some("did you mean 'accent1'?")
    );
    assert_eq!(
        suggest("accnet1", ColorContext::Spreadsheet).as_deref(),
        Some("did you mean 'accent1'?")
    );
    assert_eq!(
        suggest("dakr1", ColorContext::Word).as_deref(),
        Some("did you mean 'dark1'?")
    );
    assert_eq!(
        suggest("windwoText", ColorContext::Drawing).as_deref(),
        Some("did you mean 'windowText'?")
    );
    assert_eq!(
        suggest("folowedhyperlink", ColorContext::Drawing).as_deref(),
        Some("did you mean 'followedHyperlink'?")
    );
    assert_eq!(suggest("xyzzyq", ColorContext::Drawing), None);
    // A theme or system name another context accepts gets a pointed hint.
    for (text, context) in [
        ("phClr", ColorContext::Spreadsheet),
        ("windowText", ColorContext::Spreadsheet),
        ("window", ColorContext::Word),
        ("#accent1", ColorContext::Word),
        ("windowBackground", ColorContext::Drawing),
        ("none", ColorContext::Word),
        ("none", ColorContext::Drawing),
    ] {
        let hint = suggest(text, context).expect("a hint");
        assert!(
            !hint.starts_with("did you mean"),
            "{text:?} in {context:?}: {hint}"
        );
    }
    assert_eq!(suggest("#12345", ColorContext::Drawing), None);
    let error = Color::parse("reed", ColorContext::Drawing).unwrap_err();
    assert_eq!(error.what(), What::Color(ColorContext::Drawing));
    let unified = crate::Error::from(error);
    assert_eq!(unified.code(), crate::Code::InvalidValue);
    assert_eq!(unified.suggestion(), Some("did you mean 'red'?"));
    assert!(unified.message().starts_with("invalid color: expected hex"));
    let named = Color::parse("reed", ColorContext::Drawing)
        .unwrap_err()
        .into_error_for("fill");
    assert!(
        named
            .message()
            .starts_with("invalid 'fill' value: invalid color")
    );
}

#[test]
fn constructors_accessors_and_equality() {
    assert_eq!(Color::rgb(1, 2, 3).to_string(), "#010203");
    assert_eq!(Color::rgba(1, 2, 3, 4).to_string(), "#01020304");
    assert_eq!(Color::scheme(Scheme::Accent3).to_string(), "accent3");
    assert_eq!(Color::theme(ThemeSlot::Hyperlink).to_string(), "hlink");
    assert_eq!(Color::system(System::Window).to_string(), "window");
    assert_eq!(Color::AUTO.to_string(), "auto");
    let shaded = Color::scheme(Scheme::Accent1)
        .with_transform(Transform::new(TransformKind::Shade, 50_000).expect("valid"))
        .expect("room");
    assert_eq!(shaded.to_string(), "accent1+shade50");
    let mut full = Color::rgb(0, 0, 0);
    for _ in 0..Color::MAX_TRANSFORMS {
        full = full
            .with_transform(Transform::new(TransformKind::Alpha, 1000).expect("valid"))
            .expect("room");
    }
    assert!(
        full.with_transform(Transform::new(TransformKind::Alpha, 1000).expect("valid"))
            .is_none()
    );
    assert!(
        Color::theme(ThemeSlot::Accent1)
            .with_spreadsheet_tint(1.5)
            .is_none()
    );
    assert!(
        Color::theme(ThemeSlot::Accent1)
            .with_spreadsheet_tint(f64::NAN)
            .is_none()
    );
    let tinted = Color::theme(ThemeSlot::Accent1)
        .with_spreadsheet_tint(0.399_975_585_192_419_2)
        .expect("in range");
    assert_eq!(tinted.to_string(), "accent1+tint40");
    assert!(tinted.matches(&parse("accent1+tint40", ColorContext::Spreadsheet)));
    assert!(tinted.matches(&parse("accent1+tint39.6", ColorContext::Spreadsheet)));
    assert!(!tinted.matches(&parse("accent1+tint39.4", ColorContext::Spreadsheet)));
    assert!(!tinted.matches(&parse("accent1", ColorContext::Spreadsheet)));
    assert_eq!(
        Color::theme(ThemeSlot::Dark1)
            .with_spreadsheet_tint(0.0)
            .expect("zero")
            .spreadsheet_tint(),
        None
    );
    assert_eq!(
        "accent1+lumMod75"
            .parse::<Color>()
            .expect("drawing")
            .to_string(),
        "accent1+lumMod75"
    );
    assert_eq!(
        "bleu".parse::<Color>().unwrap_err().code(),
        crate::Code::InvalidValue
    );
    assert!(Transform::new(TransformKind::LumMod, 1_000_001).is_none());
    assert_eq!(TransformKind::Alpha.range(), (0, 100_000));
}
