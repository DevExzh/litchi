//! ODF helpers of the ODP adapter (record 0278): the conversion of unified
//! lengths to ODF length attributes. The `/metadata` set path moved to the
//! shared `format::odf_metadata` (record 0288).

use crate::value::{Length, LengthStyle};

/// An ODF length attribute value for `length`: the unified canonical
/// spelling (`2.54cm`, `72pt`, `1in`), or centimetres with six decimals when
/// no two-decimal unit is exact.
pub(crate) fn odf_length(length: Length) -> String {
    let display = length.display(LengthStyle::Auto).to_string();
    if !display.ends_with("emu") {
        return display;
    }
    let micro = i128::from(length.emu()) * 1_000_000 / 360_000;
    let negative = micro < 0;
    let micro = micro.unsigned_abs();
    let mut text = format!(
        "{}{}.{:06}",
        if negative { "-" } else { "" },
        micro / 1_000_000,
        micro % 1_000_000
    );
    while text.ends_with('0') {
        text.pop();
    }
    if text.ends_with('.') {
        text.pop();
    }
    text.push_str("cm");
    text
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn odf_lengths_round_trip() {
        for emu in [
            0, 914_400, 12_700, 360_000, 1_143_360, 4_525_920, -360_001, 1,
        ] {
            let length = Length::from_emu(emu);
            let text = odf_length(length);
            let back = Length::parse(&text).unwrap();
            assert!(
                (back.emu() - emu).abs() <= 1,
                "{emu} -> {text} -> {}",
                back.emu()
            );
        }
        assert_eq!(odf_length(Length::from_emu(914_400)), "2.54cm");
    }
}
