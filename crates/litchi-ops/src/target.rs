//! [`IntoTarget`]: what `set` acts on (record 0015 §2.1): one
//! element by path, or every element a selector matches.

use litchi_core::selector::{Path, Query, Scope, limits};

use crate::error::{Code, Error, Result};
use crate::path::{IntoPath, PathArg};
use crate::query::{IntoQuery, Plan, QueryArg};

/// A value [`crate::File::set`] accepts as its target.
///
/// A path names one element; a selector names every element it matches,
/// and the verb applies to each match. The matches of one call run in
/// one transaction.
///
/// Text is classified as OfficeCLI classifies it, parsed once and
/// borrowing from the input:
///
/// | Text | Target |
/// |---|---|
/// | a `/`-rooted path or Excel notation whose predicates address by position or identity (`/body/p[3]`, `/body/p[@paraId=1A2B3C4D]`, `Sheet1!B2`) | that element ([`TargetArg::path`]) |
/// | the same with a content filter (`/body/p[style=Heading1]`, `Sheet1!row[Amount>5000]`), or selector syntax (`/slide[1] > shape:contains(Total)`) | every match ([`TargetArg::query`]) |
/// | either of the above without the leading `/` but with a top-level one (`Sheet1/row[3]`, `Sheet1/cell[value>5]`) | as above, with the slash restored, as OfficeCLI's query restores it |
/// | a bare selector (`run`, `cell[value>5]`) | refused: [`Code::BareSelectorRejected`] |
///
/// A bare selector would match across the whole document, so text must
/// name where it applies (OfficeCLI's `MutationSelectorGuard`). Parsed
/// [`Query`] and compiled [`Plan`] values must carry the same global absolute
/// scope; a parsed [`Path`] is a path.
///
/// ```
/// use litchi_ops::Code;
/// use litchi_ops::{IntoTarget, TargetArg};
///
/// let one = "/body/p[3]".into_target()?;
/// assert!(one.path().is_some() && !one.is_selector());
///
/// let excel = "Sheet1!B2".into_target()?;
/// assert_eq!(excel.path().map(ToString::to_string), Some("/Sheet1/B2".to_owned()));
///
/// let every = "/body/p[style=Heading1]".into_target()?;
/// assert!(every.is_selector());
///
/// let error = "run[bold=true]".into_target().unwrap_err();
/// assert_eq!(error.code(), Code::BareSelectorRejected);
/// # Ok::<(), litchi_ops::Error>(())
/// ```
pub trait IntoTarget<'p> {
    /// Classifies and parses text, or wraps a parsed path, query or plan.
    ///
    /// # Errors
    ///
    /// [`Code::BareSelectorRejected`] for unscoped selector text or a parsed
    /// query/plan without a global scope;
    /// [`Code::InvalidPath`] (or [`Code::InvalidSelector`] for selector
    /// syntax) for text that does not parse, with the byte offset;
    /// [`Code::MaxDepthExceeded`] beyond the grammar's limits.
    fn into_target(self) -> Result<TargetArg<'p>>;
}

/// A target ready to resolve: a path, or a selector.
#[derive(Debug, Clone)]
pub struct TargetArg<'p>(Held<'p>);

#[derive(Debug, Clone)]
#[allow(
    clippy::large_enum_variant,
    reason = "a target is built once per call and moved into the verb; boxing the parsed path \
              or query would add the allocation short paths must not make"
)]
enum Held<'p> {
    Path(PathArg<'p>),
    Query(QueryArg<'p>),
}

impl<'p> TargetArg<'p> {
    /// Classifies and parses target text; see [`IntoTarget`].
    ///
    /// # Errors
    ///
    /// As [`IntoTarget::into_target`].
    ///
    /// ```
    /// use litchi_ops::TargetArg;
    ///
    /// assert!(TargetArg::parse("/Sheet1/A1:C3")?.path().is_some());
    /// assert!(TargetArg::parse("Sheet1/row[Amount>5000]")?.is_selector());
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn parse(text: &'p str) -> Result<Self> {
        // The rooted-relative spelling below prepends `/` before handing the
        // text to the bounded grammar. Reject first so an oversized caller
        // input cannot force an allocation proportional to its length.
        if text.len() > limits::MAX_INPUT_BYTES {
            return Err(Error::new(
                Code::InvalidPath,
                "target exceeds the selector input byte limit",
            ));
        }
        if text.starts_with('/') || excel_notation(text) {
            if has_query_syntax(text) {
                return if has_mixed_absolute_union(text) {
                    Err(mixed_union_scope())
                } else {
                    QueryArg::parse(text).and_then(Self::from_query_arg)
                };
            }
            return match PathArg::parse(text) {
                Ok(path) if path.path().is_structural() => Ok(Self(Held::Path(path))),
                // A content filter makes the text a selector.
                Ok(_) => QueryArg::parse(text).and_then(Self::from_query_arg),
                Err(error) => Err(error),
            };
        }
        if text.is_empty() {
            return PathArg::parse(text).map(|path| Self(Held::Path(path)));
        }
        if top_level_slash(text) {
            return rooted(text);
        }
        Err(bare_selector())
    }

    /// The path, when the target is one element.
    #[must_use]
    pub fn path(&self) -> Option<&PathArg<'p>> {
        match &self.0 {
            Held::Path(path) => Some(path),
            Held::Query(_) => None,
        }
    }

    /// The selector, when the target is every element it matches.
    #[must_use]
    pub fn query(&self) -> Option<&QueryArg<'p>> {
        match &self.0 {
            Held::Path(_) => None,
            Held::Query(query) => Some(query),
        }
    }

    /// Whether the target is a selector, which may match any number of
    /// elements.
    #[must_use]
    pub fn is_selector(&self) -> bool {
        matches!(self.0, Held::Query(_))
    }

    fn from_query_arg(query: QueryArg<'p>) -> Result<Self> {
        validate_query_arg(&query)?;
        Ok(Self(Held::Query(query)))
    }

    fn from_path_arg(path: PathArg<'p>) -> Result<Self> {
        crate::path::arg::validate(&path)?;
        if !path.path().is_structural() {
            return Err(Error::new(
                Code::InvalidPath,
                "content filters belong to query",
            )
            .with_suggestion(
                "use a scoped query for content filters; paths address one element by position or identity",
            ));
        }
        Ok(Self(Held::Path(path)))
    }
}

/// A scoped selector that lost its leading slash (`Sheet1/row[3]`): parsed
/// with the slash restored, as OfficeCLI's query restores it. The only
/// target text that allocates, once, for the restored copy.
fn rooted<'p>(text: &str) -> Result<TargetArg<'p>> {
    let rooted = format!("/{text}");
    if !has_query_syntax(&rooted) {
        match Path::parse(&rooted) {
            Ok(path) if path.is_structural() => {
                return TargetArg::from_path_arg(PathArg::from(path.into_owned()));
            },
            Ok(_) => {},
            Err(error) => return Err(Error::from(error)),
        }
    }
    // Query syntax, or a path with a content filter, uses the query grammar.
    if has_mixed_absolute_union(&rooted) {
        return Err(mixed_union_scope());
    }
    let query = rooted.parse::<Query<'static>>()?;
    query.into_query().and_then(TargetArg::from_query_arg)
}

/// A query target must carry a document-wide scope before a mutation can use
/// it. This also applies to parsed and compiled selectors, so callers cannot
/// bypass the mutation selector guard by wrapping an unscoped value.
fn validate_query_arg(query: &QueryArg<'_>) -> Result<()> {
    if let Some(query) = query.query() {
        return validate_query(query);
    }
    if let Some(plan) = query.plan() {
        return if plan.is_scoped() {
            Ok(())
        } else {
            Err(unscoped_query())
        };
    }
    Err(unscoped_query())
}

fn validate_query(query: &Query<'_>) -> Result<()> {
    let scoped = match query.scope() {
        Some(Scope::Path(path)) => path.is_absolute(),
        Some(Scope::Sheet(_)) => true,
        _ => false,
    };
    if scoped {
        Ok(())
    } else {
        Err(unscoped_query())
    }
}

fn unscoped_query() -> Error {
    Error::new(
        Code::BareSelectorRejected,
        "a mutation query must have an absolute scope",
    )
    .with_suggestion(
        "scope the query to a document path or worksheet, e.g. '/body/p[1]' or 'Sheet1!row[1]'",
    )
}

fn mixed_union_scope() -> Error {
    Error::new(
        Code::InvalidSelector,
        "a selector union must use one global scope",
    )
}

/// Detects query-only punctuation outside path predicates and quoted names.
/// The path grammar deliberately accepts broad step names, so a plain path
/// parse would otherwise misclassify unions, combinators and pseudo-classes
/// as one structural step.
///
/// `(`, `)` and whitespace outside brackets are query syntax only after the
/// last top-level `/` (outside brackets and parentheses). One pass keeps
/// whether such a byte follows the latest top-level slash seen so far
/// (record 0291; it used to find the last slash in a separate pass first).
/// The bytes that can make [`has_query_syntax`] answer `true`.
static QUERY_BYTES: [bool; 256] = {
    let mut table = [false; 256];
    let mut byte = 0;
    while byte < 256 {
        #[allow(clippy::cast_possible_truncation, reason = "byte < 256")]
        let b = byte as u8;
        table[byte] = matches!(b, b'"' | b'\'' | b',' | b'>' | b'*' | b':' | b'(' | b')')
            || b.is_ascii_whitespace();
        byte += 1;
    }
    table
};

fn has_query_syntax(text: &str) -> bool {
    // Only these bytes can make the answer `true` (brackets, slashes and
    // the rest only change state); a plain path has none (record 0302). One
    // table lookup per byte, folded without a branch (record 0326).
    if !text
        .bytes()
        .fold(false, |seen, byte| seen | QUERY_BYTES[usize::from(byte)])
    {
        return false;
    }
    let mut brackets = 0_usize;
    let mut parens = 0_usize;
    let mut quote: Option<u8> = None;
    let mut tail = false;
    let mut bytes = text.bytes().enumerate();
    while let Some((index, byte)) = bytes.next() {
        if let Some(open) = quote {
            if byte == b'\\' {
                bytes.next();
            } else if byte == open {
                quote = None;
            }
            continue;
        }
        match byte {
            b'"' | b'\'' => quote = Some(byte),
            b'[' => brackets += 1,
            b']' => brackets = brackets.saturating_sub(1),
            b',' | b'>' | b'*' if brackets == 0 => return true,
            b':' if brackets == 0 && is_query_pseudo(text, index) => return true,
            b'(' | b')' => {
                if byte == b'(' {
                    parens += 1;
                } else {
                    parens = parens.saturating_sub(1);
                }
                if brackets == 0 {
                    tail = true;
                }
            },
            b'/' if brackets == 0 && parens == 0 => tail = false,
            byte if brackets == 0 && byte.is_ascii_whitespace() => tail = true,
            _ => {},
        }
    }
    tail
}

fn is_query_pseudo(text: &str, colon: usize) -> bool {
    let rest = &text[colon.saturating_add(1)..];
    let end = rest
        .bytes()
        .position(|byte| !byte.is_ascii_alphanumeric() && byte != b'-')
        .unwrap_or(rest.len());
    matches!(
        &rest[..end],
        name if name.eq_ignore_ascii_case("contains")
            || name.eq_ignore_ascii_case("empty")
            || name.eq_ignore_ascii_case("no-alt")
            || name.eq_ignore_ascii_case("has")
            || name.eq_ignore_ascii_case("not")
    )
}

/// Rejects a second absolute path after a union comma. Query syntax has one
/// scope for the complete union; allowing `/body/p, /body/tbl` would make the
/// query parser absorb the first alternative into its scope path.
fn has_mixed_absolute_union(text: &str) -> bool {
    let mut brackets = 0_usize;
    let mut parens = 0_usize;
    let mut quote: Option<u8> = None;
    let mut saw_comma = false;
    let mut bytes = text.bytes();
    while let Some(byte) = bytes.next() {
        if let Some(open) = quote {
            if byte == b'\\' {
                bytes.next();
            } else if byte == open {
                quote = None;
            }
            continue;
        }
        match byte {
            b'"' | b'\'' => quote = Some(byte),
            b'[' => brackets += 1,
            b']' => brackets = brackets.saturating_sub(1),
            b'(' if brackets == 0 => parens += 1,
            b')' if brackets == 0 => parens = parens.saturating_sub(1),
            b',' if brackets == 0 && parens == 0 => saw_comma = true,
            b'/' if brackets == 0 && parens == 0 && saw_comma => return true,
            _ => {},
        }
    }
    false
}

/// OfficeCLI's Excel notation test: a sheet name without `/`, `[` or `]`,
/// then `!` (`Sheet1!A1`, `'My Data'!B:B`).
fn excel_notation(text: &str) -> bool {
    for (at, byte) in text.bytes().enumerate() {
        match byte {
            b'!' => return at > 0,
            b'/' | b'[' | b']' => return false,
            _ => {},
        }
    }
    false
}

/// Whether `/` occurs outside every `[]` and `()` group and outside quotes
/// (a backslash escapes the next character inside quotes): a scoped
/// selector that lost its leading slash (`Sheet1/row[...]`), as OfficeCLI's
/// `SelectorCommaSplit.ContainsTopLevelChar` decides.
fn top_level_slash(text: &str) -> bool {
    last_top_level_slash_index(text).is_some()
}

fn last_top_level_slash_index(text: &str) -> Option<usize> {
    let mut brackets = 0_usize;
    let mut parens = 0_usize;
    let mut quote: Option<u8> = None;
    let mut last = None;
    let mut bytes = text.bytes().enumerate();
    while let Some((index, byte)) = bytes.next() {
        if let Some(open) = quote {
            if byte == b'\\' {
                bytes.next();
            } else if byte == open {
                quote = None;
            }
            continue;
        }
        match byte {
            b'"' | b'\'' => quote = Some(byte),
            b'[' => brackets += 1,
            b']' => brackets = brackets.saturating_sub(1),
            b'(' => parens += 1,
            b')' => parens = parens.saturating_sub(1),
            b'/' if brackets == 0 && parens == 0 => last = Some(index),
            _ => {},
        }
    }
    last
}

/// The refusal of an unscoped selector on a mutation.
fn bare_selector() -> Error {
    Error::new(
        Code::BareSelectorRejected,
        "a bare mutation selector would match across the whole document",
    )
    .with_suggestion(
        "scope the selector to an absolute document path or worksheet; bare selectors remain available to File::query",
    )
}

impl<'p> IntoTarget<'p> for &'p str {
    fn into_target(self) -> Result<TargetArg<'p>> {
        TargetArg::parse(self)
    }
}

impl<'p> IntoTarget<'p> for &'p String {
    fn into_target(self) -> Result<TargetArg<'p>> {
        TargetArg::parse(self)
    }
}

impl<'p> IntoTarget<'p> for Path<'p> {
    fn into_target(self) -> Result<TargetArg<'p>> {
        self.into_path().and_then(TargetArg::from_path_arg)
    }
}

impl<'p, 'q: 'p> IntoTarget<'p> for &'p Path<'q> {
    fn into_target(self) -> Result<TargetArg<'p>> {
        self.into_path().and_then(TargetArg::from_path_arg)
    }
}

impl<'p> IntoTarget<'p> for PathArg<'p> {
    fn into_target(self) -> Result<TargetArg<'p>> {
        TargetArg::from_path_arg(self)
    }
}

impl<'p> IntoTarget<'p> for Query<'p> {
    fn into_target(self) -> Result<TargetArg<'p>> {
        self.into_query().and_then(TargetArg::from_query_arg)
    }
}

impl<'p> IntoTarget<'p> for &'p Query<'p> {
    fn into_target(self) -> Result<TargetArg<'p>> {
        self.into_query().and_then(TargetArg::from_query_arg)
    }
}

impl<'p> IntoTarget<'p> for Plan<'p> {
    fn into_target(self) -> Result<TargetArg<'p>> {
        self.into_query().and_then(TargetArg::from_query_arg)
    }
}

impl<'p> IntoTarget<'p> for &'p Plan<'p> {
    fn into_target(self) -> Result<TargetArg<'p>> {
        self.into_query().and_then(TargetArg::from_query_arg)
    }
}

impl<'p> IntoTarget<'p> for QueryArg<'p> {
    fn into_target(self) -> Result<TargetArg<'p>> {
        TargetArg::from_query_arg(self)
    }
}

impl<'p> IntoTarget<'p> for TargetArg<'p> {
    fn into_target(self) -> Result<TargetArg<'p>> {
        Ok(self)
    }
}

/// A target parsed once runs again without parsing, classifying or
/// copying it: the path or selector is borrowed (record 0291).
///
/// ```
/// use litchi_ops::{IntoTarget, TargetArg};
///
/// let target = TargetArg::parse("/body/p[1]")?;
/// for _ in 0..2 {
///     let again = (&target).into_target()?;
///     assert_eq!(again.path().map(ToString::to_string), Some("/body/p[1]".to_owned()));
/// }
/// # Ok::<(), litchi_ops::Error>(())
/// ```
impl<'a, 'p: 'a> IntoTarget<'a> for &'a TargetArg<'p> {
    fn into_target(self) -> Result<TargetArg<'a>> {
        Ok(TargetArg(match &self.0 {
            Held::Path(path) => Held::Path(path.reborrow()),
            Held::Query(query) => Held::Query(query.into_query()?),
        }))
    }
}

#[cfg(test)]
mod tests {
    use litchi_core::selector::{Path, Query, limits};

    use super::{IntoTarget, TargetArg};
    use crate::Code;
    use crate::query::QueryArg;

    #[test]
    fn text_paths_and_scoped_queries_are_classified() {
        let path = TargetArg::parse("/body/p[3]").expect("structural path");
        assert!(path.path().is_some());
        assert!(!path.is_selector());

        let filtered = TargetArg::parse("/body/p[style=Heading1]").expect("scoped query");
        assert!(filtered.query().is_some());
        assert!(filtered.is_selector());

        let pseudo = TargetArg::parse("/body/p:contains(Heading1)").expect("pseudo query");
        assert!(pseudo.is_selector());

        let wildcard = TargetArg::parse("/body/*").expect("wildcard query");
        assert!(wildcard.is_selector());

        let excel = TargetArg::parse("Sheet1!row[Amount>5000]").expect("sheet query");
        assert!(excel.is_selector());

        let rooted = TargetArg::parse("Sheet1/row[3]").expect("rooted path");
        assert_eq!(
            rooted.path().map(ToString::to_string),
            Some("/Sheet1/row[3]".into())
        );

        let union = TargetArg::parse("/body/p, tbl").expect("scoped union");
        assert!(union.is_selector());

        assert!(
            TargetArg::parse("/body/p, /body/tbl").is_err(),
            "each union member must share the query's one global scope"
        );
    }

    #[test]
    fn bare_text_selectors_and_unscoped_typed_queries_are_rejected() {
        for text in ["run", "cell[value>5]", "p, r", "shape[name=title/a]"] {
            let error = TargetArg::parse(text).expect_err("bare selector");
            assert_eq!(error.code(), Code::BareSelectorRejected, "{text}");
        }

        let query = Query::parse("run").expect("query");
        assert_eq!(
            (&query)
                .into_target()
                .expect_err("unscoped typed query")
                .code(),
            Code::BareSelectorRejected
        );
        assert_eq!(
            query
                .into_target()
                .expect_err("unscoped owned query")
                .code(),
            Code::BareSelectorRejected
        );

        let scoped = Query::parse("/body/run").expect("scoped query");
        assert!((&scoped).into_target().expect("scoped query").is_selector());

        let arg = QueryArg::parse("run").expect("query argument");
        assert_eq!(
            arg.into_target()
                .expect_err("unscoped query argument")
                .code(),
            Code::BareSelectorRejected
        );

        let path = Path::parse_relative("run").expect("typed path");
        assert!((&path).into_target().expect("typed path").path().is_some());

        let filtered_path = Path::parse_relative("run[value>5]").expect("filtered path syntax");
        assert_eq!(
            filtered_path
                .into_target()
                .expect_err("content-filter path")
                .code(),
            Code::InvalidPath
        );
    }

    #[test]
    fn malformed_scoped_text_keeps_parser_classification() {
        for text in ["/body/p[", "Sheet1/row[", "Sheet1!"] {
            let error = TargetArg::parse(text).expect_err("malformed target");
            assert!(
                matches!(error.code(), Code::InvalidPath | Code::InvalidSelector),
                "{text}: {error:?}"
            );
        }
    }

    #[test]
    fn target_parser_enforces_selector_limits_without_format_features() {
        let deep = format!("/{}", ["p"].repeat(limits::MAX_STEPS + 1).join("/"));
        assert_eq!(
            TargetArg::parse(&deep).expect_err("depth limit").code(),
            Code::MaxDepthExceeded
        );

        let too_long = format!("/{}", "p".repeat(limits::MAX_INPUT_BYTES));
        assert!(
            matches!(
                TargetArg::parse(&too_long).expect_err("input limit").code(),
                Code::InvalidPath | Code::MaxDepthExceeded
            ),
            "target input should be bounded"
        );
    }

    #[test]
    fn bare_selector_errors_do_not_expose_authored_text() {
        let secret = "PRIVATE_SELECTOR_PAYLOAD";
        let error = TargetArg::parse(secret).unwrap_err();
        assert_eq!(error.code(), Code::BareSelectorRejected);
        assert!(!error.message().contains(secret));
        assert!(!error.suggestion().unwrap_or_default().contains(secret));
        assert!(!error.to_string().contains(secret));
        assert!(!format!("{error:?}").contains(secret));
    }

    /// The two-pass classifier `has_query_syntax` replaced (record 0291).
    fn two_pass_query_syntax(text: &str) -> bool {
        let mut brackets = 0_usize;
        let mut quote: Option<u8> = None;
        let last_slash = super::last_top_level_slash_index(text);
        let mut bytes = text.bytes().enumerate();
        while let Some((index, byte)) = bytes.next() {
            if let Some(open) = quote {
                if byte == b'\\' {
                    bytes.next();
                } else if byte == open {
                    quote = None;
                }
                continue;
            }
            match byte {
                b'"' | b'\'' => quote = Some(byte),
                b'[' => brackets += 1,
                b']' => brackets = brackets.saturating_sub(1),
                b',' | b'>' | b'*' if brackets == 0 => return true,
                b':' if brackets == 0 && super::is_query_pseudo(text, index) => return true,
                b'(' | b')' if brackets == 0 => {
                    if last_slash.is_none_or(|slash| index > slash) {
                        return true;
                    }
                },
                byte if brackets == 0
                    && byte.is_ascii_whitespace()
                    && last_slash.is_none_or(|slash| index > slash) =>
                {
                    return true;
                },
                _ => {},
            }
        }
        false
    }

    #[test]
    fn one_pass_query_syntax_matches_the_two_pass_classifier() {
        // Every string of up to five bytes over the bytes the classifier
        // treats specially, plus a letter.
        let alphabet = b"/[]()'\\\" ,>*:pa";
        let mut text = Vec::new();
        let mut checked = 0;
        for len in 0..=5_u32 {
            for mut code in 0..alphabet.len().pow(len) {
                text.clear();
                for _ in 0..len {
                    text.push(alphabet[code % alphabet.len()]);
                    code /= alphabet.len();
                }
                let text = std::str::from_utf8(&text).expect("ASCII");
                assert_eq!(
                    super::has_query_syntax(text),
                    two_pass_query_syntax(text),
                    "{text:?}"
                );
                checked += 1;
            }
        }
        assert!(checked > 800_000);
        for text in [
            "/body/p[1]",
            "/body/p:contains(x)",
            "/Sheet1/row[1] > cell",
            "/a(b)/c",
            "/a/(b)",
            "/a(b/c)",
            "/a/b c",
            "/a b/c",
            "Sheet1!A1",
            "/body/p[@paraId='1 2']/r",
        ] {
            assert_eq!(
                super::has_query_syntax(text),
                two_pass_query_syntax(text),
                "{text}"
            );
        }
    }

    #[test]
    fn target_debug_does_not_expose_selector_text() {
        let target = TargetArg::parse("/body/p:contains(PRIVATE)").expect("selector");
        assert!(!format!("{target:?}").contains("PRIVATE"));
    }
}
