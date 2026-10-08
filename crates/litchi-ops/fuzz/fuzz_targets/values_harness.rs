//! Bounded coercion and canonical round-trip exercise.
use litchi_ops::value::{
    Color, ColorContext, Length, LengthContext, LengthStyle, parse_bool, parse_integer,
    parse_number, parse_numeric_text, parse_opacity,
};

pub fn exercise(data: &[u8]) -> usize {
    if data.len() > 64 * 1024 {
        return 0;
    }
    let Ok(text) = std::str::from_utf8(data) else {
        return 0;
    };
    let mut accepted = 0;
    for context in [
        ColorContext::Drawing,
        ColorContext::Word,
        ColorContext::Spreadsheet,
    ] {
        if let Ok(color) = Color::parse(text, context) {
            let canonical = color.to_string();
            assert_eq!(
                Color::parse(&canonical, context).expect("canonical color"),
                color
            );
            accepted += 1;
        }
    }
    for context in [
        LengthContext::Emu,
        LengthContext::Extent,
        LengthContext::LineWidth,
        LengthContext::FontSize,
        LengthContext::WordSpacing,
    ] {
        if let Ok(length) = Length::parse_in(text, context) {
            // Exact EMU is the common lossless interchange representation.
            let canonical = format!("{}emu", length.emu());
            assert_eq!(
                Length::parse_in(&canonical, LengthContext::Emu).expect("canonical length"),
                length
            );
            let _ = length.display(LengthStyle::Points).to_string();
            accepted += 1;
        }
    }
    accepted += usize::from(parse_bool(text).is_ok());
    accepted += usize::from(parse_number(text).is_ok());
    accepted += usize::from(parse_integer(text).is_ok());
    accepted += usize::from(parse_opacity(text).is_ok());
    accepted += usize::from(parse_numeric_text(text).is_some());
    accepted
}
