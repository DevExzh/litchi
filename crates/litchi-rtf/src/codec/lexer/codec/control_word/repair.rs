//! Word-compatible interpretation of malformed control-word parameters.
//!
//! RTF 1.9.1 defines toggles as taking 0 or 1, flags as taking no parameter,
//! and value controls as requiring one, but prescribes no reader behavior when
//! a producer deviates. LibreOffice's tokenizer (`RTFTokenizer::dispatchKeyword`)
//! documents Word's interpretation: "flags ignore any parameter by
//! definition", a value control without a parameter takes the keyword's
//! default, and toggles treat every nonzero parameter as on. These helpers
//! apply that interpretation for the controls the parser validates strictly
//! and report the repair, which the lexer records as a [`ParameterDefect`] at
//! the control word's offset.

use crate::leniency::ParameterDefect;

/// A flag, symbol, or destination: any parameter is ignored.
#[inline]
pub(super) fn flag(param: Option<i32>, repair: &mut Option<ParameterDefect>) -> Option<i32> {
    if param.is_some() {
        *repair = Some(ParameterDefect::UnexpectedParameter);
    }
    None
}

/// A toggle whose parameter is optional: every nonzero value is on.
#[inline]
pub(super) fn toggle(param: Option<i32>, repair: &mut Option<ParameterDefect>) -> Option<i32> {
    match param {
        None | Some(0 | 1) => param,
        Some(_) => {
            *repair = Some(ParameterDefect::NonBinaryToggle);
            Some(1)
        },
    }
}

/// A 0/1 control that requires its parameter: a missing parameter takes
/// `default`, and every nonzero value is on.
#[inline]
pub(super) fn binary(
    param: Option<i32>,
    default: i32,
    repair: &mut Option<ParameterDefect>,
) -> Option<i32> {
    match param {
        None => {
            *repair = Some(ParameterDefect::MissingParameter);
            Some(default)
        },
        Some(0 | 1) => param,
        Some(_) => {
            *repair = Some(ParameterDefect::NonBinaryToggle);
            Some(1)
        },
    }
}

/// A value control that requires its parameter: a missing parameter takes
/// `default`.
#[inline]
pub(super) fn value(
    param: Option<i32>,
    default: i32,
    repair: &mut Option<ParameterDefect>,
) -> Option<i32> {
    if param.is_none() {
        *repair = Some(ParameterDefect::MissingParameter);
        return Some(default);
    }
    param
}

/// A selector the model treats as a flag but Word and LibreOffice treat as a
/// toggle (the emphasis-mark and associated-underline selectors): a zero
/// parameter turns the attribute off, any other parameter is ignored.
/// Returns whether the selector is on.
#[inline]
pub(super) fn toggled_selector(param: Option<i32>, repair: &mut Option<ParameterDefect>) -> bool {
    if param.is_some() {
        *repair = Some(ParameterDefect::UnexpectedParameter);
    }
    param != Some(0)
}
