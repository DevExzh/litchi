//! Booleans and opacities.

use super::error::{ErrorKind, ValueError, What};
use super::number::{Styles, parse_f64};

/// Zero-width and no-break characters OfficeCLI strips around a boolean after
/// trimming white space (`ParseHelpers.TrimInvisible`): the byte-order mark,
/// zero-width space, zero-width non-joiner, zero-width joiner, word joiner
/// and no-break space.
const INVISIBLE: [char; 6] = [
    '\u{feff}', '\u{200b}', '\u{200c}', '\u{200d}', '\u{2060}', '\u{a0}',
];

/// Parses a boolean the way OfficeCLI's `ParseHelpers.IsTruthy` does:
/// `true`, `1`, `yes` and `on` are true; `false`, `0`, `no`, `off` and the
/// empty string are false; matching ignores ASCII case and surrounding white
/// space and invisible characters.
///
/// # Errors
///
/// Returns an `invalid_value` [`ValueError`] for any other text.
///
/// ```
/// use litchi_ops::value::parse_bool;
///
/// assert!(parse_bool(" YES ")?);
/// assert!(!parse_bool("off")?);
/// assert!(!parse_bool("")?);
/// assert!(parse_bool("maybe").is_err());
/// # Ok::<(), litchi_ops::Error>(())
/// ```
pub fn parse_bool(text: &str) -> Result<bool, ValueError<'_>> {
    // `Trim()` first, then the invisible set, exactly in OfficeCLI's order:
    // a zero-width character outside a space keeps that space inside.
    let word = text.trim().trim_matches(INVISIBLE.as_slice());
    const TRUE: [&str; 4] = ["true", "1", "yes", "on"];
    const FALSE: [&str; 5] = ["false", "0", "no", "off", ""];
    if TRUE
        .iter()
        .any(|candidate| word.eq_ignore_ascii_case(candidate))
    {
        Ok(true)
    } else if FALSE
        .iter()
        .any(|candidate| word.eq_ignore_ascii_case(candidate))
    {
        Ok(false)
    } else {
        Err(ValueError::new(What::Boolean, ErrorKind::Syntax, text))
    }
}

/// Parses an opacity the way OfficeCLI's shape, picture and cell `opacity`
/// properties do: a finite decimal from 0.0 to 1.0, or a percentage from 2 to
/// 100 (`30` means 0.30). Values strictly between 1 and 2 are refused as
/// ambiguous.
///
/// # Errors
///
/// Returns an `invalid_value` [`ValueError`] for anything else.
///
/// ```
/// use litchi_ops::value::parse_opacity;
///
/// assert_eq!(parse_opacity("0.25")?, 0.25);
/// assert_eq!(parse_opacity("30")?, 0.3);
/// assert!(parse_opacity("1.5").is_err());
/// # Ok::<(), litchi_ops::Error>(())
/// ```
pub fn parse_opacity(text: &str) -> Result<f64, ValueError<'_>> {
    let error = |kind| ValueError::new(What::Opacity, kind, text);
    let value = match parse_f64(text, Styles::FLOAT) {
        Some(value) if value.is_finite() => value,
        Some(_) => return Err(error(ErrorKind::NotFinite)),
        None if text.trim().is_empty() => return Err(error(ErrorKind::Empty)),
        None => return Err(error(ErrorKind::Syntax)),
    };
    if value > 1.0 && value < 2.0 {
        return Err(error(ErrorKind::Ambiguous));
    }
    let fraction = if value > 1.0 { value / 100.0 } else { value };
    if (0.0..=1.0).contains(&fraction) {
        Ok(fraction)
    } else {
        Err(error(ErrorKind::OutOfRange))
    }
}
