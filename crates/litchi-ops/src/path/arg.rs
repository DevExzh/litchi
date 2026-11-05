//! [`IntoPath`]: the one way `get` and `find` take a path, as text or as a
//! parsed [`Path`] (record 0010 §8.1, decision D9).

use std::borrow::Cow;
use std::fmt;

use litchi_core::selector::{Condition, Expr, Key, Path, Position, Predicate, Step, limits};

use crate::{Code, Error, Result};

/// A value `get` and `find` accept as a path.
///
/// Text is parsed once, borrowing from the input, so short paths allocate
/// nothing (record 0005). A parsed [`Path`], by value or by reference, is used
/// as is: parse once with [`Path::parse`] and resolve it many times without
/// parsing again.
///
/// Text that starts with `/` is an absolute path. Text in Excel notation
/// (`Sheet1!A1`, `'My Data'!A1:C5`, `Sheet1!3:3`, `Sheet1!B:B`) becomes the
/// absolute workbook path `/Sheet1/A1`, `/My Data/A1:C5`, `/Sheet1/row[3]` or
/// `/Sheet1/col[B]`. Any other text is a relative path, resolved from a node.
///
/// ```
/// use litchi_ops::path::{IntoPath, Path};
///
/// let text = "/body/p[3]/r[1]".into_path()?;
/// assert_eq!(text.path().steps().len(), 3);
///
/// let excel = "'My Data'!B:B".into_path()?;
/// assert_eq!(excel.to_string(), "/My Data/col[B]");
///
/// let parsed = Path::parse("/slide[2]/shape[1]")?;
/// assert_eq!((&parsed).into_path()?.path(), &parsed);
/// # Ok::<(), litchi_ops::Error>(())
/// ```
pub trait IntoPath<'p> {
    /// Parses text, or wraps an already parsed path.
    ///
    /// # Errors
    ///
    /// [`Code::InvalidPath`] for text that is not path syntax, with the byte
    /// offset in the message and the grammar's typed error as the source;
    /// [`Code::MaxDepthExceeded`] for paths beyond the step limit.
    fn into_path(self) -> Result<PathArg<'p>>;
}

impl<'p> IntoPath<'p> for &'p str {
    fn into_path(self) -> Result<PathArg<'p>> {
        PathArg::parse(self)
    }
}

impl<'p> IntoPath<'p> for &'p String {
    fn into_path(self) -> Result<PathArg<'p>> {
        PathArg::parse(self)
    }
}

impl<'p> IntoPath<'p> for Path<'p> {
    fn into_path(self) -> Result<PathArg<'p>> {
        Ok(PathArg::from(self))
    }
}

impl<'p, 'q: 'p> IntoPath<'p> for &'p Path<'q> {
    fn into_path(self) -> Result<PathArg<'p>> {
        Ok(PathArg::from(self))
    }
}

impl<'p> IntoPath<'p> for PathArg<'p> {
    fn into_path(self) -> Result<PathArg<'p>> {
        Ok(self)
    }
}

/// A path ready to resolve: the parsed [`Path`] and, when it came from
/// text, that text, so errors can point at the failing step's byte offset.
///
/// ```
/// use litchi_ops::path::PathArg;
///
/// let arg = PathArg::parse("Sheet1!3:3")?;
/// assert!(arg.is_absolute());
/// assert_eq!(arg.text(), Some("Sheet1!3:3"));
/// assert_eq!(arg.to_string(), "/Sheet1/row[3]");
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Clone)]
pub struct PathArg<'p> {
    held: Held<'p>,
    origin: Origin<'p>,
    pub(crate) raw: bool,
}

#[derive(Clone)]
#[allow(
    clippy::large_enum_variant,
    reason = "a PathArg is built once per call and moved into the resolver; boxing the \
              parsed path would add the heap allocation short paths must not make"
)]
enum Held<'p> {
    Owned(Path<'p>),
    Borrowed(&'p Path<'p>),
}

/// Where a [`PathArg`] came from, which decides how a step's byte offset is
/// found on the error path.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum Origin<'p> {
    /// A parsed `Path` value: offsets count in its canonical spelling.
    Typed,
    /// Path text, parsed as written.
    Text(&'p str),
    /// Excel notation: step 0 is the sheet before `!`; the steps after it
    /// start at byte `rest`.
    Sheet {
        /// The text as written.
        text: &'p str,
        /// The byte offset just past `!`.
        rest: usize,
    },
}

impl<'p> PathArg<'p> {
    /// Parses path text, borrowing from it.
    ///
    /// # Errors
    ///
    /// As [`IntoPath::into_path`].
    ///
    /// ```
    /// use litchi_ops::path::PathArg;
    ///
    /// let relative = PathArg::parse("tr[2]/tc[1]")?;
    /// assert!(!relative.is_absolute());
    ///
    /// let error = PathArg::parse("/body/p[0]").unwrap_err();
    /// assert_eq!(error.code().as_str(), "invalid_path");
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    #[inline]
    pub fn parse(text: &'p str) -> Result<Self> {
        if text.len() > limits::MAX_INPUT_BYTES {
            Path::parse(text)?;
            return Err(Error::new(
                Code::InvalidPath,
                "path exceeds the input byte limit",
            ));
        }
        if text.starts_with('/') {
            return Ok(Self::parsed(Path::parse(text)?, Origin::Text(text)));
        }
        if let Some((sheet, rest)) = split_sheet(text) {
            return sheet_path(text, sheet, rest);
        }
        Ok(Self::parsed(
            Path::parse_relative(text)?,
            Origin::Text(text),
        ))
    }

    #[inline]
    fn parsed(path: Path<'p>, origin: Origin<'p>) -> Self {
        Self {
            held: Held::Owned(path),
            origin,
            raw: false,
        }
    }

    /// The parsed path.
    ///
    /// ```
    /// use litchi_ops::path::PathArg;
    ///
    /// let arg = PathArg::parse("/body/p[1]")?;
    /// assert_eq!(arg.path().steps()[1].name(), "p");
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    #[must_use]
    pub fn path(&self) -> &Path<'p> {
        match &self.held {
            Held::Owned(path) => path,
            Held::Borrowed(path) => path,
        }
    }

    /// The text the path was parsed from, or `None` for a parsed [`Path`].
    ///
    /// ```
    /// use litchi_ops::path::{IntoPath, Path, PathArg};
    ///
    /// assert_eq!(PathArg::parse("/body")?.text(), Some("/body"));
    /// assert_eq!(Path::parse("/body")?.into_path()?.text(), None);
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    #[must_use]
    pub fn text(&self) -> Option<&'p str> {
        match self.origin {
            Origin::Typed => None,
            Origin::Text(text) | Origin::Sheet { text, .. } => Some(text),
        }
    }

    /// Whether the path starts at the document root.
    ///
    /// ```
    /// use litchi_ops::path::PathArg;
    ///
    /// assert!(PathArg::parse("/")?.is_absolute());
    /// assert!(!PathArg::parse("p[@paraId=1A2B3C4D]")?.is_absolute());
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    #[must_use]
    pub fn is_absolute(&self) -> bool {
        self.path().is_absolute()
    }

    pub(crate) fn origin(&self) -> Origin<'p> {
        self.origin
    }

    /// The same path, borrowed: no step is copied (record 0291).
    pub(crate) fn reborrow(&self) -> PathArg<'_> {
        PathArg {
            held: Held::Borrowed(self.path()),
            origin: self.origin,
            raw: self.raw,
        }
    }
}

impl<'p> From<Path<'p>> for PathArg<'p> {
    fn from(path: Path<'p>) -> Self {
        Self {
            held: Held::Owned(path),
            origin: Origin::Typed,
            raw: false,
        }
    }
}

impl<'p, 'q: 'p> From<&'p Path<'q>> for PathArg<'p> {
    fn from(path: &'p Path<'q>) -> Self {
        Self {
            held: Held::Borrowed(path),
            origin: Origin::Typed,
            raw: false,
        }
    }
}

/// The canonical spelling of the path: a fixed point of [`PathArg::parse`]
/// (record 0338 F2).
///
/// It is the core grammar's spelling, except for a relative path whose first
/// step name the core grammar writes bare but [`PathArg::parse`] would read
/// as Excel sheet notation (`bod! my`, read back as `/bod/ my`): that name is
/// written quoted (`"bod! my"`), which both grammars read as the same step.
///
/// ```
/// use litchi_ops::path::PathArg;
///
/// let arg = PathArg::parse("'bod! my'")?;
/// assert_eq!(arg.to_string(), "\"bod! my\"");
/// let again = arg.to_string();
/// assert_eq!(PathArg::parse(&again)?.path(), arg.path());
/// # Ok::<(), litchi_ops::Error>(())
/// ```
impl fmt::Display for PathArg<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let path = self.path();
        match path.steps().first() {
            Some(first) if !path.is_absolute() && reads_as_sheet(first) => {
                write_quoted(formatter, first.name())?;
                for predicate in first.predicates() {
                    fmt::Display::fmt(predicate, formatter)?;
                }
                for step in &path.steps()[1..] {
                    formatter.write_str("/")?;
                    fmt::Display::fmt(step, formatter)?;
                }
                Ok(())
            },
            _ => path.fmt(formatter),
        }
    }
}

/// Whether the core grammar's spelling of a relative path starting with
/// `first` is read by [`split_sheet`] as sheet notation. The core grammar
/// writes a name bare only when it has no leading quote and no `/`, `[` or
/// `]`, so the first `!` of the spelling, if `split_sheet` finds one before
/// any `/`, `[`, `]` or `"`, is in that name; the byte after it is the
/// name's next byte, or the `[`, `/` or end that follows the name.
fn reads_as_sheet(first: &Step<'_>) -> bool {
    let name = first.name();
    let Some(bang) = name.find('!') else {
        return false;
    };
    let bytes = name.as_bytes();
    let bare = bytes
        .first()
        .is_some_and(|&byte| byte != b'"' && byte != b'\'')
        && !bytes
            .iter()
            .any(|&byte| matches!(byte, b'/' | b'[' | b']') || byte.is_ascii_control());
    bare && bang > 0 && !bytes[..bang].contains(&b'"') && bytes.get(bang + 1) != Some(&b'=')
}

/// Writes `text` double-quoted with `"` and `\` escaped, as the core grammar
/// writes a name that needs quotes.
fn write_quoted(formatter: &mut fmt::Formatter<'_>, text: &str) -> fmt::Result {
    use fmt::Write as _;
    formatter.write_char('"')?;
    for c in text.chars() {
        if c == '"' || c == '\\' {
            formatter.write_char('\\')?;
        }
        formatter.write_char(c)?;
    }
    formatter.write_char('"')
}

/// Splits Excel notation into the sheet name and the byte offset just past
/// `!`, as OfficeCLI's `NormalizeExcelPath` does for text without a leading
/// `/`: a quoted name `'My Data'` (with `''` for an apostrophe) must be
/// followed by `!`; an unquoted name runs to the first `!`. Unlike OfficeCLI,
/// an unquoted name containing `/`, `[`, `]` or `"`, or a `!` that begins
/// `!=`, is not sheet notation, so relative paths such as
/// `shape[@name=Hello!]` keep their meaning.
fn split_sheet(text: &str) -> Option<(Cow<'_, str>, usize)> {
    let bytes = text.as_bytes();
    if bytes.first() == Some(&b'\'') {
        let mut index = 1;
        let mut escaped = false;
        loop {
            match bytes.get(index)? {
                b'\'' if bytes.get(index + 1) == Some(&b'\'') => {
                    escaped = true;
                    index += 2;
                },
                b'\'' => break,
                _ => index += 1,
            }
        }
        if bytes.get(index + 1) != Some(&b'!') {
            return None;
        }
        let inner = &text[1..index];
        let name = if escaped {
            Cow::Owned(inner.replace("''", "'"))
        } else {
            Cow::Borrowed(inner)
        };
        return Some((name, index + 2));
    }
    let bang = text.find('!')?;
    let name = &text[..bang];
    let plain = bang > 0
        && !name
            .bytes()
            .any(|byte| matches!(byte, b'/' | b'[' | b']' | b'"'))
        && bytes.get(bang + 1) != Some(&b'=');
    plain.then_some((Cow::Borrowed(name), bang + 1))
}

/// Builds `/sheet/...` from Excel notation: whole rows (`3:3`) and columns
/// (`B:B`, either case) become `row[3]` and `col[B]`; anything else after `!`
/// is a relative path (`A1`, `A1:C5`, `row[3]`, `chart[1]/series[2]`).
fn sheet_path<'p>(text: &'p str, sheet: Cow<'p, str>, rest: usize) -> Result<PathArg<'p>> {
    let tail = &text[rest..];
    if sheet.is_empty() {
        return Err(notation_error(
            "the sheet name before '!' is empty",
            text,
            0,
            "name the sheet, e.g. Sheet1!A1 or 'My Data'!A1",
        ));
    }
    if tail.is_empty() {
        return Err(notation_error(
            "nothing follows '!'",
            text,
            rest,
            "write a cell, range, row or column after '!', e.g. Sheet1!A1, Sheet1!3:3 or Sheet1!B:B",
        ));
    }
    if sheet.chars().any(char::is_control) {
        return Err(notation_error(
            "sheet name contains a control character",
            text,
            0,
            "remove control characters from the sheet name",
        ));
    }
    let sheet = Step::new(sheet);
    let path = if let Some(row) = whole(tail, |byte| byte.is_ascii_digit()) {
        Path::from_steps([sheet, row_step(text, rest, row)?])
    } else if let Some(letters) = whole(tail, |byte| byte.is_ascii_alphabetic()) {
        let column = Predicate::Filter(Expr::Condition(Condition::Exists(Key::new(letters))));
        Path::from_steps([sheet, Step::new("col").with(column)])
    } else {
        let relative = Path::parse_relative(tail).map_err(|error| error.offset_by(rest))?;
        if relative.steps().len() >= limits::MAX_STEPS {
            return Err(Error::new(
                Code::MaxDepthExceeded,
                format!(
                    "path has more than {} steps once the sheet is added",
                    limits::MAX_STEPS
                ),
            ));
        }
        Path::from_steps(std::iter::once(sheet).chain(relative.steps().iter().cloned()))
    };
    let arg = PathArg::parsed(path, Origin::Sheet { text, rest });
    canonical_bound(arg.path())?;
    Ok(arg)
}

/// `N:N` or `L:L`: the same run of bytes on both sides of one `:` (letters
/// compare ASCII case-insensitively, as OfficeCLI's case-insensitive regex
/// does).
fn whole(tail: &str, class: impl Fn(u8) -> bool) -> Option<&str> {
    let (start, end) = tail.split_once(':')?;
    let ok = !start.is_empty() && start.bytes().all(&class) && start.eq_ignore_ascii_case(end);
    ok.then_some(start)
}

fn row_step<'p>(text: &str, rest: usize, digits: &str) -> Result<Step<'p>> {
    let position = digits
        .parse::<usize>()
        .ok()
        .filter(|row| *row <= super::report::MAX_ROW)
        .and_then(Position::from_one_based)
        .ok_or_else(|| {
            notation_error(
                "the row number must be between 1 and 1048576",
                text,
                rest,
                "rows are numbered from 1, e.g. Sheet1!3:3",
            )
        })?;
    Ok(Step::new("row").with(Predicate::Index(position)))
}

fn notation_error(problem: &str, _text: &str, offset: usize, suggestion: &'static str) -> Error {
    Error::new(Code::InvalidPath, format!("{problem} at byte {offset}")).with_suggestion(suggestion)
}

impl fmt::Debug for PathArg<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("PathArg")
            .field("steps", &self.path().steps().len())
            .field("raw", &self.raw)
            .finish()
    }
}
impl fmt::Debug for Origin<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("Origin { .. }")
    }
}

/// Admit hand-built paths before adapter I/O. Content filters are rejected
/// without recursively visiting arbitrary caller-owned expression trees.
pub(crate) fn validate(arg: &PathArg<'_>) -> Result<()> {
    let mut bytes = 0usize;
    for step in arg.path().steps() {
        if step.predicates().len() > limits::MAX_PREDICATES {
            return Err(Error::new(
                Code::InvalidPath,
                "path exceeds the predicate limit",
            ));
        }
        bytes = bytes.saturating_add(step.name().len());
        for predicate in step.predicates() {
            if let Predicate::Filter(expr) = predicate {
                match expr {
                    Expr::Condition(condition) => {
                        bytes = bytes.saturating_add(condition.key().name().len());
                        if let Condition::Compare { value, .. } = condition {
                            bytes = bytes.saturating_add(value.as_str().len());
                        }
                    },
                    _ => {
                        return Err(Error::new(
                            Code::InvalidPath,
                            "content filters belong to query",
                        ));
                    },
                }
            }
        }
        if bytes > limits::MAX_INPUT_BYTES {
            return Err(Error::new(
                Code::InvalidPath,
                "path exceeds the input byte limit",
            ));
        }
    }
    // Writing the canonical spelling is most of a short path's admission
    // cost (record 0280); a path whose spelling cannot reach the limit skips
    // it.
    if spelling_bound(arg.path()) <= limits::MAX_INPUT_BYTES {
        return Ok(());
    }
    canonical_bound(arg.path())
}

/// An upper bound of the canonical spelling's length of an admitted path
/// (positions, `last()` and single conditions only). Quoting at most doubles
/// a name, key or value and adds its quotes; the other parts have a fixed
/// width.
fn spelling_bound(path: &Path<'_>) -> usize {
    let quoted = |text: &str| text.len().saturating_mul(2).saturating_add(3);
    path.steps().iter().fold(1usize, |total, step| {
        let predicates = step.predicates().iter().fold(0usize, |sum, predicate| {
            let width = match predicate {
                // `[` + at most 39 digits (u128) + `]`.
                Predicate::Index(_) => 41,
                Predicate::Last => 8,
                Predicate::Filter(Expr::Condition(Condition::Exists(key))) => {
                    quoted(key.name()).saturating_add(3)
                },
                Predicate::Filter(Expr::Condition(Condition::Compare { key, value, .. })) => {
                    quoted(key.name())
                        .saturating_add(quoted(value.as_str()))
                        .saturating_add(8)
                },
                // Rejected before this point, or a future form: keep the
                // exact check.
                _ => usize::MAX,
            };
            sum.saturating_add(width)
        });
        total
            .saturating_add(1)
            .saturating_add(quoted(step.name()))
            .saturating_add(predicates)
    })
}

fn canonical_bound(path: &Path<'_>) -> Result<()> {
    struct Bound(usize);
    impl fmt::Write for Bound {
        fn write_str(&mut self, text: &str) -> fmt::Result {
            self.0 = self.0.saturating_add(text.len());
            if self.0 > limits::MAX_INPUT_BYTES {
                Err(fmt::Error)
            } else {
                Ok(())
            }
        }
    }
    use std::fmt::Write as _;
    write!(Bound(0), "{path}").map_err(|_| {
        Error::new(
            Code::InvalidPath,
            "canonical path exceeds the input byte limit",
        )
    })
}

#[cfg(test)]
mod spelling_tests {
    use litchi_core::selector::Path;

    use super::spelling_bound;

    /// The admission shortcut's bound never undercounts the canonical
    /// spelling (record 0280).
    #[test]
    fn the_spelling_bound_covers_the_canonical_spelling() {
        let names = ["p", "Sheet 1", "a/b", "\"q\"", "x\\y]", "Ünïcødé", ""];
        let mut checked = 0;
        for name in names {
            for predicate in [
                "",
                "[3]",
                "[last()]",
                "[@id=\"a\\\"b\"]",
                "[style=Heading1]",
                "[@paraId]",
                "[\"Full Name\"=\"x/y\"]",
                "[184467440737095516]",
            ] {
                let quoted = name.replace('\\', "\\\\").replace('"', "\\\"");
                let text = format!("/body/\"{quoted}\"{predicate}/r[2]");
                let Ok(path) = Path::parse(&text) else {
                    continue;
                };
                let spelling = path.to_string();
                assert!(
                    spelling_bound(&path) >= spelling.len(),
                    "{text}: {} < {}",
                    spelling_bound(&path),
                    spelling.len()
                );
                checked += 1;
            }
        }
        assert!(checked > 30);
    }
}

/// The canonical spelling is a fixed point of [`PathArg::parse`] (record
/// 0338 F2): reparsing it gives the same path, and spelling that again gives
/// the same text.
#[cfg(test)]
mod display_tests {
    use litchi_core::selector::{Path, Step};
    use proptest::prelude::*;

    use super::PathArg;

    fn assert_fixed_point(arg: &PathArg<'_>, origin: &str) {
        let canonical = arg.to_string();
        let again = PathArg::parse(&canonical)
            .unwrap_or_else(|error| panic!("{origin:?} -> {canonical:?}: {error}"));
        assert_eq!(again.path(), arg.path(), "{origin:?} -> {canonical:?}");
        assert_eq!(again.to_string(), canonical, "{origin:?}");
    }

    #[test]
    fn a_relative_name_with_a_bang_is_quoted() {
        for (text, canonical) in [
            ("'bod! my'", "\"bod! my\""),
            ("\"a!b\"[2]/c", "\"a!b\"[2]/c"),
            ("'x!'/y", "\"x!\"/y"),
            // Not sheet notation as spelled: unchanged.
            ("\"a!=b\"", "a!=b"),
            ("\"!a\"", "!a"),
            ("a/b!c", "a/b!c"),
            ("shape[@name=Hello!]", "shape[@name=Hello!]"),
        ] {
            let arg = PathArg::parse(text).unwrap_or_else(|error| panic!("{text}: {error}"));
            assert_eq!(arg.to_string(), canonical, "{text}");
            assert_fixed_point(&arg, text);
        }
        // Absolute paths and sheet notation keep the core spelling.
        let arg = PathArg::parse("'Q1!Results'!A1").expect("sheet notation");
        assert_eq!(arg.to_string(), "/Q1!Results/A1");
        assert_fixed_point(&arg, "'Q1!Results'!A1");
        // A typed relative path whose bare spelling ends at `!`.
        let typed = Path::relative([Step::new("abc!")]);
        assert!(!typed.is_absolute());
        let arg = PathArg::from(&typed);
        assert_eq!(arg.to_string(), "\"abc!\"");
        assert_fixed_point(&arg, "abc!");
    }

    proptest! {
        #![proptest_config(ProptestConfig::with_cases(4096))]

        /// Text over the bytes that matter to sheet notation and quoting.
        #[test]
        fn parsed_text_spells_a_fixed_point(text in "[a1 !'\"/=\\[\\]:@\\\\]{1,12}") {
            if let Ok(arg) = PathArg::parse(&text) {
                assert_fixed_point(&arg, &text);
            }
        }

        /// Typed relative and absolute paths from arbitrary step names.
        #[test]
        fn typed_paths_spell_a_fixed_point(
            names in proptest::collection::vec("[a1 !'\"=:\\[\\]/]{1,6}", 1..4),
            absolute in any::<bool>(),
        ) {
            let steps = names.iter().map(|name| Step::new(name.as_str()));
            let path = if absolute {
                Path::from_steps(steps)
            } else {
                Path::relative(steps)
            };
            let arg = PathArg::from(&path);
            let canonical = arg.to_string();
            if let Ok(again) = PathArg::parse(&canonical) {
                prop_assert_eq!(again.path(), &path, "{:?}", canonical);
                prop_assert_eq!(again.to_string(), canonical);
            }
        }
    }
}
