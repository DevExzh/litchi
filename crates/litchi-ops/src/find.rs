//! Formatting-aware find/replace: the `find=`/`replace=` properties of `set`
//! (record 0273).
//!
//! OfficeCLI models find/replace as keywords of its `set` command rather
//! than as a verb (`WordHandler.Set.cs`, `PowerPointHandler.Set.cs`,
//! `ExcelHandler.Set.cs`), so Litchi does too: a `set` whose properties name
//! `find` is a text search over the subtree of its target, and every other
//! property of the same call is either a find option or a run format applied
//! to the replaced (or, without `replace`, the matched) text.
//!
//! This module is the format-agnostic engine. A format adapter splits each
//! text container (a paragraph, a shape paragraph, a cell) into ordered
//! [`Piece`]s: the text of one run (one leaf text node), grouped so that a
//! match never crosses a group boundary (a tab, a line break, a field or a
//! hyperlink edge). [`plan`] finds the matches in each group and returns, per
//! changed piece, its new text and the byte ranges that a format property
//! must cover. The adapter applies those rewrites through its leaf; untouched
//! pieces are never written, so their runs stay byte-for-byte.
//!
//! # Semantics
//!
//! | Key | Value | Meaning |
//! |---|---|---|
//! | `find` | non-empty text | the literal to search for |
//! | `replace` | text, may be empty | the replacement; empty deletes each match |
//! | `matchCase` | bool, default `true` | `false` compares characters by their simple lowercase mapping |
//! | `wholeWord` | bool, default `false` | `true` keeps only matches not adjacent to a letter, digit or `_` |
//! | `regex` | bool | `true` is refused: Litchi has no approved regular-expression engine (record 0010 Q3) |
//! | anything else | run property | applied to the replacement text, or to the matched text without `replace` |
//!
//! Matching is leftmost and non-overlapping, scanning left to right, as
//! OfficeCLI's `FindHelpers.FindMatchRanges` does. A match that spans
//! several pieces collapses into the piece where it starts: the replacement
//! takes that piece's formatting, and the matched characters of the
//! following pieces are removed from them.
//!
//! A `set` with `find` and neither `replace` nor a format property is
//! refused (`missing_property` naming `replace`), as OfficeCLI refuses it.
//! Zero matches succeed as an exact no-op; the match count is reported as
//! [`crate::Applied::matched`].

#![cfg_attr(
    not(any(
        feature = "docx",
        feature = "xlsx",
        feature = "pptx",
        feature = "xlsx",
        feature = "rtf",
        feature = "doc",
        feature = "ppt",
        feature = "xls",
        feature = "xlsb",
        feature = "odt",
        feature = "ods",
        feature = "odp"
    )),
    allow(dead_code, unused_imports, unused_variables)
)]

mod matcher;
mod plan;
#[cfg(any(
    feature = "xlsx",
    feature = "rtf",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "odt",
    feature = "ods",
    feature = "odp"
))]
pub(crate) mod surface;
#[cfg(test)]
mod tests;

use std::borrow::Cow;

pub(crate) use matcher::Matcher;
#[cfg(feature = "docx")]
pub(crate) use plan::matches_with;
#[cfg(test)]
pub(crate) use plan::plan;
#[cfg(any(
    feature = "pptx",
    feature = "xlsx",
    feature = "rtf",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "odt",
    feature = "ods",
    feature = "odp"
))]
pub(crate) use plan::plan_with;
pub(crate) use plan::{Piece, Scratch};
#[cfg(any(
    feature = "pptx",
    feature = "xlsx",
    feature = "rtf",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "odt",
    feature = "ods",
    feature = "odp"
))]
pub(crate) use plan::{Values, rewrite_ranges};

use crate::error::{Code, Error, Result};
use crate::props::{PropValue, Props};

/// The property that turns a `set` into a find.
pub(crate) const FIND: &str = "find";
/// The replacement text.
pub(crate) const REPLACE: &str = "replace";
/// The regular-expression switch (refused when true).
pub(crate) const REGEX: &str = "regex";
/// Case-sensitive comparison (default true).
pub(crate) const MATCH_CASE: &str = "matchCase";
/// Whole-word matches only (default false).
pub(crate) const WHOLE_WORD: &str = "wholeWord";

/// One keyword a find admits, as the schema export lists it.
pub(crate) struct Keyword {
    /// The key.
    pub(crate) name: &'static str,
    /// The OfficeCLI help-schema value type.
    pub(crate) ty: &'static str,
    /// The value an absent boolean option takes.
    pub(crate) default: Option<bool>,
    /// Whether a value must be non-empty text.
    pub(crate) not_empty: bool,
    /// The refusal of a `true` value, when the option is a gap.
    pub(crate) gap: Option<(&'static str, &'static str)>,
}

/// The find keywords, in the order the schema lists them. Every other
/// property of a find call is a run format of the target format's run
/// registry.
pub(crate) const KEYWORDS: [Keyword; 5] = [
    Keyword {
        name: FIND,
        ty: "string",
        default: None,
        not_empty: true,
        gap: None,
    },
    Keyword {
        name: REPLACE,
        ty: "string",
        default: None,
        not_empty: false,
        gap: None,
    },
    Keyword {
        name: MATCH_CASE,
        ty: "bool",
        default: Some(true),
        not_empty: false,
        gap: None,
    },
    Keyword {
        name: WHOLE_WORD,
        ty: "bool",
        default: Some(false),
        not_empty: false,
        gap: None,
    },
    Keyword {
        name: REGEX,
        ty: "bool",
        default: Some(false),
        not_empty: false,
        gap: Some((
            REGEX_FOLLOW_UP,
            "Litchi has no approved regular-expression engine; `regex=true` and OfficeCLI's `r\"...\"` patterns are refused",
        )),
    },
];

/// What makes `regex=true` available.
const REGEX_FOLLOW_UP: &str = "record 0010 Q3: an owner-approved regular-expression dependency";

/// The longest accepted `find` or `replace` text, in bytes.
pub(crate) const MAX_TEXT_BYTES: usize = 64 * 1024;

/// Whether `props` asks for a find: any key spelled `find`, ignoring ASCII
/// case, as OfficeCLI's case-insensitive property dictionary does.
pub(crate) fn requested(props: &Props<'_>) -> bool {
    props.iter().any(|(key, _)| key.eq_ignore_ascii_case(FIND))
}

/// An admitted find: the pattern, the replacement, the options and the
/// remaining (format) properties, in their original order.
pub(crate) struct Spec<'a> {
    pattern: Cow<'a, str>,
    replacement: Option<Cow<'a, str>>,
    match_case: bool,
    whole_word: bool,
    format: Props<'a>,
}

impl std::fmt::Debug for Spec<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        // Pattern and replacement are authored content: keep them out of logs.
        formatter
            .debug_struct("Spec")
            .field("replace", &self.replacement.is_some())
            .field("match_case", &self.match_case)
            .field("whole_word", &self.whole_word)
            .field("format", &self.format.len())
            .finish()
    }
}

impl<'a> Spec<'a> {
    /// Admit the find keys of `props`. Every other key is kept, in order, as
    /// a format property for the adapter to admit through its registry.
    ///
    /// # Errors
    ///
    /// `invalid_argument` for a repeated find key, `invalid_value` for an
    /// empty or control-bearing pattern or a non-text value,
    /// `unsupported_capability` for a regular expression, `limit_exceeded`
    /// for an over-long text, and `missing_property` (`replace`) when the
    /// call neither replaces nor formats.
    pub(crate) fn admit(props: &Props<'a>) -> Result<Self> {
        let mut pattern = None;
        let mut replacement = None;
        let mut match_case = None;
        let mut whole_word = None;
        let mut regex = None;
        let mut format = Props::new();
        for (key, value) in props.iter() {
            let name = key.as_str();
            let slot = if name.eq_ignore_ascii_case(FIND) {
                admit_text(&mut pattern, FIND, value)?;
                continue;
            } else if name.eq_ignore_ascii_case(REPLACE) {
                admit_text(&mut replacement, REPLACE, value)?;
                continue;
            } else if name.eq_ignore_ascii_case(MATCH_CASE) {
                (&mut match_case, MATCH_CASE)
            } else if name.eq_ignore_ascii_case(WHOLE_WORD) {
                (&mut whole_word, WHOLE_WORD)
            } else if name.eq_ignore_ascii_case(REGEX) {
                (&mut regex, REGEX)
            } else {
                format.push(key.clone(), value.clone());
                continue;
            };
            let (slot, name) = slot;
            if slot.is_some() {
                return Err(repeated(name));
            }
            *slot = Some(admit_bool(name, value)?);
        }
        let Some(pattern) = pattern else {
            return Err(
                Error::new(Code::MissingProperty, "a find needs `find`").with_property(FIND)
            );
        };
        if regex == Some(true) || is_raw_regex(&pattern) {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "regular-expression find is not supported",
            )
            .with_property(REGEX)
            .with_suggestion(
                "search for literal text; Litchi has no approved regular-expression engine (record 0010 Q3)",
            ));
        }
        if pattern.is_empty() {
            return Err(
                Error::new(Code::InvalidValue, "the find text is empty").with_property(FIND)
            );
        }
        if replacement.is_none() && format.is_empty() {
            return Err(Error::new(
                Code::MissingProperty,
                "a find needs `replace` or a format property to apply",
            )
            .with_property(REPLACE)
            .with_suggestion("add `replace=TEXT`, or a run property such as `bold=true`"));
        }
        Ok(Self {
            pattern,
            replacement,
            match_case: match_case.unwrap_or(true),
            whole_word: whole_word.unwrap_or(false),
            format,
        })
    }

    /// The matcher of this find.
    pub(crate) fn matcher(&self) -> Matcher<'_> {
        Matcher::new(&self.pattern, self.match_case, self.whole_word)
    }

    /// The replacement text, or `None` for a format-only find.
    pub(crate) fn replacement(&self) -> Option<&str> {
        self.replacement.as_deref()
    }

    /// The format properties, in their original order.
    pub(crate) const fn format(&self) -> &Props<'a> {
        &self.format
    }

    /// Whether the call applies run formatting.
    pub(crate) fn formats(&self) -> bool {
        !self.format.is_empty()
    }
}

/// A `find` value written as OfficeCLI's raw regular expression, `r"..."`
/// or `r'...'`. Treating it as a literal would silently search for the
/// quotes, so it is refused with `regex=true`.
fn is_raw_regex(pattern: &str) -> bool {
    let bytes = pattern.as_bytes();
    bytes.len() >= 3
        && bytes[0] == b'r'
        && matches!(bytes[1], b'"' | b'\'')
        && bytes[bytes.len() - 1] == bytes[1]
}

/// `set` with `find=` on a format whose adapter has no find yet.
pub(crate) fn unsupported(file: &crate::File) -> Error {
    let error = Error::new(
        Code::UnsupportedCapability,
        "find/replace is not integrated for this format",
    )
    .with_property(FIND)
    .with_suggestion("see the format's capability report (record 0273)");
    match file.format() {
        Some(format) => error.with_format(format),
        None => error,
    }
}

/// [`crate::File::set`] with `find=`: one native commit, the count dropped.
pub(crate) fn set(
    file: &crate::File,
    target: crate::TargetArg<'_>,
    props: &Props<'_>,
) -> Result<crate::File> {
    let spec = Spec::admit(props)?;
    let limits = crate::transaction::SemanticLimits::default();
    let _ = (&target, &limits);
    match file {
        #[cfg(feature = "docx")]
        crate::File::Docx(handle) => handle
            .find_replace(&target, &spec, &limits)
            .map(|(commit, _)| crate::File::Docx(commit.into_file())),
        #[cfg(feature = "xlsx")]
        crate::File::Xlsx(handle) => handle
            .find_replace(&target, &spec, &limits)
            .map(|(commit, _)| crate::File::Xlsx(commit.into_file())),
        #[cfg(feature = "pptx")]
        crate::File::Pptx(handle) => handle
            .find_replace(&target, &spec, &limits)
            .map(|(commit, _)| crate::File::Pptx(commit.into_file())),
        #[cfg(feature = "rtf")]
        crate::File::Rtf(handle) => handle
            .find_replace(&target, &spec, &limits)
            .map(|(commit, _)| crate::File::Rtf(commit.after)),
        #[cfg(feature = "doc")]
        crate::File::Doc(handle) => handle
            .find_replace(&target, &spec, &limits)
            .map(|(commit, _)| crate::File::Doc(commit.after)),
        #[cfg(feature = "xlsb")]
        crate::File::Xlsb(handle) => {
            crate::format::xlsb::write::find::find_replace(handle, &target, &spec, &limits)
                .map(|(commit, _)| crate::File::Xlsb(commit.into_file()))
        },
        #[cfg(feature = "xls")]
        crate::File::Xls(handle) => crate::format::xls::write::find::find_replace(
            handle, &target, &spec, &limits,
        )
        .map(|(commits, _)| {
            crate::File::Xls(
                commits
                    .last()
                    .map_or_else(|| handle.clone(), |commit| commit.after.clone()),
            )
        }),
        #[cfg(feature = "ods")]
        crate::File::Ods(handle) => {
            crate::format::ods::write::find::find_replace(handle, &target, &spec, &limits)
                .map(|(commit, _)| crate::File::Ods(commit.after))
        },
        #[cfg(feature = "odt")]
        crate::File::Odt(handle) | crate::File::Oth(handle) | crate::File::Odm(handle) => {
            crate::format::odt::write::find::find_replace(handle, &target, &spec, &limits)
                .map(|(commit, _)| commit.after.into_file())
        },
        #[cfg(feature = "odp")]
        crate::File::Odp(handle) => {
            crate::format::odp::write::find::find_replace(handle, &target, &spec, &limits)
                .map(|(commit, _)| crate::File::Odp(commit.after))
        },
        #[cfg(feature = "ppt")]
        crate::File::Ppt(handle) => {
            crate::format::ppt::write::find::find_replace(handle, &target, &spec, &limits)
                .map(|(commit, _)| crate::File::Ppt(commit.after))
        },
        _ => Err(unsupported(file)),
    }
}

fn repeated(name: &'static str) -> Error {
    Error::new(Code::InvalidArgument, "a find key is given more than once").with_property(name)
}

fn admit_text<'a>(
    slot: &mut Option<Cow<'a, str>>,
    name: &'static str,
    value: &PropValue<'a>,
) -> Result<()> {
    if slot.is_some() {
        return Err(repeated(name));
    }
    let text: Cow<'a, str> = match value {
        PropValue::Text(text) => text.clone(),
        PropValue::Int(value) => Cow::Owned(value.to_string()),
        PropValue::Unsigned(value) => Cow::Owned(value.to_string()),
        _ => {
            return Err(
                Error::new(Code::InvalidValue, "a find value must be text").with_property(name)
            );
        },
    };
    if text.len() > MAX_TEXT_BYTES {
        return Err(Error::new(
            Code::LimitExceeded,
            format!("a find value is limited to {MAX_TEXT_BYTES} bytes"),
        )
        .with_property(name));
    }
    if text.chars().any(char::is_control) {
        return Err(Error::new(
            Code::InvalidValue,
            "find and replace text cannot hold tabs, line breaks or other control characters",
        )
        .with_property(name)
        .with_suggestion(
            "tabs and breaks separate the text find searches; match the text on either side",
        ));
    }
    *slot = Some(text);
    Ok(())
}

fn admit_bool(name: &'static str, value: &PropValue<'_>) -> Result<bool> {
    match value {
        PropValue::Bool(value) => Ok(*value),
        PropValue::Text(text) => Ok(crate::value::parse_bool(text).map_err(|_error| {
            Error::new(Code::InvalidValue, "expected a boolean").with_property(name)
        })?),
        _ => Err(Error::new(Code::InvalidValue, "expected a boolean").with_property(name)),
    }
}
