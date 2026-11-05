//! Round trips: the canonical text of a value parses back to the value
//! (exact renderings) or to a value with the same text (rounded renderings),
//! and arbitrary input never panics.

use super::super::color::names::NAMED;
use super::super::{
    Base, Color, ColorContext, Length, LengthContext, LengthStyle, Rgba, Scheme, System, ThemeSlot,
    Transform, TransformKind,
};

struct Rng(u64);

impl Rng {
    fn next(&mut self) -> u64 {
        self.0 = self.0.wrapping_add(0x9E37_79B9_7F4A_7C15);
        let mut z = self.0;
        z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
        z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
        z ^ (z >> 31)
    }

    fn below(&mut self, bound: u64) -> u64 {
        self.next() % bound
    }

    /// A length up to 2^50 EMU (about 31 km) in magnitude, biased toward the
    /// unit grids so every rendering is exercised.
    fn length(&mut self) -> i64 {
        let grid =
            [1_i64, 127, 635, 3_600, 6_350, 9_144, 9_525, 12_700, 360_000][self.below(9) as usize];
        let magnitude = (self.below(1 << 50) as i64) / grid * grid;
        if self.below(2) == 0 {
            magnitude
        } else {
            -magnitude
        }
    }
}

const SCHEMES: [Scheme; 17] = [
    Scheme::Background1,
    Scheme::Text1,
    Scheme::Background2,
    Scheme::Text2,
    Scheme::Accent1,
    Scheme::Accent2,
    Scheme::Accent3,
    Scheme::Accent4,
    Scheme::Accent5,
    Scheme::Accent6,
    Scheme::Hyperlink,
    Scheme::FollowedHyperlink,
    Scheme::Placeholder,
    Scheme::Dark1,
    Scheme::Light1,
    Scheme::Dark2,
    Scheme::Light2,
];

const KINDS: [TransformKind; 9] = [
    TransformKind::LumMod,
    TransformKind::LumOff,
    TransformKind::SatMod,
    TransformKind::SatOff,
    TransformKind::HueMod,
    TransformKind::HueOff,
    TransformKind::Shade,
    TransformKind::Tint,
    TransformKind::Alpha,
];

#[test]
fn exact_length_renderings_parse_back_to_the_same_emu() {
    let mut rng = Rng(1);
    let mut values: Vec<i64> = (-5_000..=5_000).collect();
    values.extend((0..50_000).map(|_| rng.length()));
    for emu in values {
        for style in [LengthStyle::Auto, LengthStyle::PointsFirst] {
            let text = Length::from_emu(emu).display(style).to_string();
            let back = Length::parse_in(&text, LengthContext::Emu)
                .unwrap_or_else(|error| panic!("{emu} as {text}: {error}"));
            assert_eq!(back.emu(), emu, "{emu} printed {text}");
        }
    }
    for emu in [i64::MAX, i64::MIN, i64::MAX - 1] {
        let text = Length::from_emu(emu).to_string();
        assert_eq!(Length::parse(&text).expect("parses").emu(), emu, "{text}");
    }
}

#[test]
fn rounded_length_renderings_are_fixed_points() {
    let mut rng = Rng(2);
    for _ in 0..50_000 {
        let emu = rng.length();
        for (style, context) in [
            (LengthStyle::Lossy, LengthContext::Emu),
            (LengthStyle::LineWidth, LengthContext::Emu),
            (LengthStyle::Points, LengthContext::Emu),
        ] {
            let text = Length::from_emu(emu).display(style).to_string();
            let back = Length::parse_in(&text, context)
                .unwrap_or_else(|error| panic!("{emu} as {text}: {error}"));
            assert_eq!(back.display(style).to_string(), text, "{emu}");
            // Re-reading never moves by more than half a hundredth of the unit.
            assert!((back.emu() - emu).abs() <= 1_800, "{emu} -> {}", back.emu());
        }
    }
}

#[test]
fn line_width_and_font_size_texts_parse_back() {
    for points in 0..=4_000 {
        for hundredths in [0, 1, 25, 50, 75, 99] {
            let emu = points * 12_700 + hundredths * 127;
            let length = Length::from_emu(emu);
            let line = length.display(LengthStyle::LineWidth).to_string();
            let back = Length::parse_in(&line, LengthContext::LineWidth).expect("line width");
            assert_eq!(back, length, "{line}");
            if (6_350..=50_800_000).contains(&emu) {
                let size = length.display(LengthStyle::Points).to_string();
                let back = Length::parse_in(&size, LengthContext::FontSize).expect("font size");
                assert_eq!(back, length, "{size}");
            }
        }
    }
}

fn random_rgba(rng: &mut Rng) -> Rgba {
    let bits = rng.next().to_le_bytes();
    let alpha = if rng.below(3) == 0 { bits[3] } else { 255 };
    Rgba::new(bits[0], bits[1], bits[2]).with_alpha(alpha)
}

fn random_transform(rng: &mut Rng) -> Transform {
    let kind = KINDS[rng.below(9) as usize];
    let (low, high) = kind.range();
    let span = u64::try_from(i64::from(high) - i64::from(low)).expect("positive") + 1;
    let mut value = low + i32::try_from(rng.below(span)).expect("in range");
    if rng.below(2) == 0 {
        // Whole percentages print in the short form.
        value -= value % 1000;
    }
    Transform::new(kind, value).expect("in range")
}

fn assert_round_trip(color: Color, context: ColorContext) {
    let text = color.to_string();
    let back = Color::parse(&text, context)
        .unwrap_or_else(|error| panic!("{color:?} printed {text}: {error}"));
    assert_eq!(back, color, "{text} in {context:?}");
    assert_eq!(back.to_string(), text);
}

#[test]
fn drawing_colors_round_trip() {
    let mut rng = Rng(3);
    for _ in 0..20_000 {
        let base = match rng.below(3) {
            0 => Base::Rgb(random_rgba(&mut rng)),
            1 => Base::Scheme(SCHEMES[rng.below(17) as usize]),
            _ => Base::System(
                System::from_token(super::super::color::names::SYSTEMS[rng.below(30) as usize].0)
                    .expect("token"),
            ),
        };
        let mut color = Color::new(base);
        for _ in 0..rng.below(Color::MAX_TRANSFORMS as u64 + 1) {
            color = color
                .with_transform(random_transform(&mut rng))
                .expect("room");
        }
        assert_round_trip(color, ColorContext::Drawing);
    }
    for named in NAMED {
        let color = Color::parse(named.name, ColorContext::Drawing).expect("named");
        assert_round_trip(color, ColorContext::Drawing);
    }
}

#[test]
fn spreadsheet_colors_round_trip() {
    let mut rng = Rng(4);
    for _ in 0..20_000 {
        let base = if rng.below(2) == 0 {
            Base::Rgb(random_rgba(&mut rng))
        } else {
            Base::Theme(
                ThemeSlot::from_index(u32::try_from(rng.below(12)).expect("small")).expect("slot"),
            )
        };
        let mut color = Color::new(base);
        if rng.below(2) == 0 {
            // Tints printable with two decimals of a percentage, at least
            // 0.5% from zero (smaller tints are not printed).
            let hundredths = i32::try_from(rng.below(9_951)).expect("small") + 50;
            let signed = if rng.below(2) == 0 {
                hundredths
            } else {
                -hundredths
            };
            let percent = f64::from(signed) / 100.0;
            color = color
                .with_spreadsheet_tint(percent / 100.0)
                .expect("in range");
        }
        let text = color.to_string();
        let back = Color::parse(&text, ColorContext::Spreadsheet)
            .unwrap_or_else(|error| panic!("{text}: {error}"));
        assert!(back.matches(&color), "{text}");
        assert_eq!(back.to_string(), text);
    }
}

#[test]
fn word_colors_round_trip() {
    let mut rng = Rng(5);
    for _ in 0..5_000 {
        assert_round_trip(
            Color::new(Base::Rgb(random_rgba(&mut rng))),
            ColorContext::Word,
        );
    }
    for scheme in SCHEMES {
        if scheme != Scheme::Placeholder {
            assert_round_trip(Color::scheme(scheme), ColorContext::Word);
        }
    }
    assert_round_trip(Color::AUTO, ColorContext::Word);
}

/// Mutated inputs either parse or fail with a typed error; nothing panics.
#[test]
fn hostile_inputs_never_panic() {
    let seeds = [
        "accent1+lumMod75+lumOff25",
        "rgba(255, 0, 0, 0.5)",
        "hsl(240deg 100% 50%)",
        "#FF000080",
        "2.54cm",
        "1e-7in",
        "-12pt",
        "lt1+tint-25",
        "windowText",
        "transparent",
        "0x1F",
        "true",
        "1,234.5",
        "(5)",
        "\u{a4}5",
        "１２ｐｔ",
        "∞",
        "NaN",
        "Ölfarbe",
        "accent1:lumMod75",
    ];
    let alphabet: Vec<char> = "0123456789.,+-=:;#()%/ eEptcmixnqQaAdlr\u{a0}\u{3000}é\u{1F600}"
        .chars()
        .collect();
    let mut rng = Rng(6);
    for seed in seeds {
        for round in 0..2_000 {
            // Suggestions run edit distance over the vocabulary; sample them.
            let suggest = round % 16 == 0;
            let mut chars: Vec<char> = seed.chars().collect();
            for _ in 0..=rng.below(4) {
                let position = rng.below(chars.len() as u64 + 1) as usize;
                match rng.below(3) {
                    0 => chars.insert(
                        position,
                        alphabet[rng.below(alphabet.len() as u64) as usize],
                    ),
                    1 if position < chars.len() => {
                        chars.remove(position);
                    },
                    _ if position < chars.len() => {
                        chars[position] = alphabet[rng.below(alphabet.len() as u64) as usize];
                    },
                    _ => {},
                }
            }
            let text: String = chars.into_iter().collect();
            for context in [
                ColorContext::Drawing,
                ColorContext::Spreadsheet,
                ColorContext::Word,
            ] {
                if let Err(error) = Color::parse(&text, context) {
                    let _ = error.to_string();
                    if suggest {
                        let _ = error.suggestion();
                    }
                }
            }
            for context in [
                LengthContext::Emu,
                LengthContext::Extent,
                LengthContext::LineWidth,
                LengthContext::FontSize,
            ] {
                if let Err(error) = Length::parse_in(&text, context) {
                    let _ = (error.to_string(), error.suggestion());
                }
            }
            let _ = super::super::parse_bool(&text);
            let _ = super::super::parse_number(&text);
            let _ = super::super::parse_numeric_text(&text);
            let _ = super::super::parse_opacity(&text);
        }
    }
}
