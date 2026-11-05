//! Text surfaces: the adapter contract find/replace and template merge
//! share for formats whose leaf rewrites whole runs, paragraphs or cells of
//! plain text (record 0295).
//!
//! An adapter lists the searchable text of its scope as [`Run`]s, in
//! document order: the text of one leaf unit its native editor can replace
//! (a DOC character run, a PPT paragraph, an ODT span, an ODP run, a
//! spreadsheet cell). Runs of one `group` are adjacent text a match or a
//! placeholder may span; a new group starts wherever one must not cross (a
//! paragraph, a cell, a field). A run the leaf cannot rewrite as text (an
//! ODT span holding a field, a computed ODP field) is listed with
//! `editable: false`: find refuses a match that touches it, so a broad find
//! never silently leaves text behind, and merge treats it as fixed content,
//! as DOCX merge treats a drawing.
//!
//! [`find`] and [`merge`] return the new text of every changed run; the
//! adapter stages one native replacement per change in one native edit.
//! Untouched runs are never written.

#![cfg_attr(
    not(any(
        feature = "rtf",
        feature = "doc",
        feature = "ppt",
        feature = "xls",
        feature = "xlsb",
        feature = "odt",
        feature = "ods",
        feature = "odp"
    )),
    allow(dead_code, reason = "XLSX merges rich-text runs only")
)]

use std::borrow::Cow;
use std::ops::Range;

use litchi_core::FileFormat;

use super::{Piece, Scratch, Spec, Values, plan_with, rewrite_ranges};
use crate::error::{Code, Error, Result};
use crate::merge::{Planner, normalize_breaks};
use crate::transaction::{Error as TransactionError, SemanticLimits};

/// One replaceable unit of text.
#[derive(Debug, Clone)]
pub(crate) struct Run<'t> {
    /// The unit's current text.
    pub(crate) text: Cow<'t, str>,
    /// The group: adjacent runs with the same group are contiguous text.
    pub(crate) group: usize,
    /// Whether the run is inside the find's scope (always for merge).
    pub(crate) scoped: bool,
    /// Whether the leaf can replace the run's text.
    pub(crate) editable: bool,
}

impl<'t> Run<'t> {
    /// An editable, scoped run of `group`.
    pub(crate) fn new(text: impl Into<Cow<'t, str>>, group: usize) -> Self {
        Self {
            text: text.into(),
            group,
            scoped: true,
            editable: true,
        }
    }

    /// The same run with its scope flag set.
    #[must_use]
    #[cfg_attr(
        not(any(
            feature = "doc",
            feature = "ppt",
            feature = "rtf",
            feature = "odt",
            feature = "odp"
        )),
        allow(dead_code, reason = "spreadsheet scopes select whole cells")
    )]
    pub(crate) const fn scoped(mut self, scoped: bool) -> Self {
        self.scoped = scoped;
        self
    }

    /// The same run with its editability set.
    #[must_use]
    #[cfg_attr(
        not(any(
            feature = "odt",
            feature = "odp",
            feature = "ods",
            feature = "xlsb",
            feature = "doc",
            feature = "ppt",
        )),
        allow(dead_code, reason = "only adapters with fixed runs mark them")
    )]
    pub(crate) const fn editable(mut self, editable: bool) -> Self {
        self.editable = editable;
        self
    }
}

/// The new text of one run, by its index in the adapter's list.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Change {
    /// Index of the run.
    pub(crate) run: usize,
    /// Its new text.
    pub(crate) text: String,
}

/// The replacement text of a find whose leaf applies no formatting: the
/// format properties are refused (`unsupported_property`, naming the first)
/// and a format-only find needs `replace`.
///
/// # Errors
///
/// As described.
pub(crate) fn text_only<'s>(spec: &'s Spec<'_>, format: FileFormat, why: &str) -> Result<&'s str> {
    if spec.formats() {
        let key = spec
            .format()
            .iter()
            .next()
            .map_or(super::FIND, |(key, _)| key.as_str());
        return Err(Error::new(
            Code::UnsupportedProperty,
            format!("this find replaces text and applies no formatting: {why}"),
        )
        .with_property(key.to_owned())
        .with_format(format)
        .with_suggestion(
            "drop the format properties, or set them on the matched runs afterwards",
        ));
    }
    spec.replacement().ok_or_else(|| {
        Error::new(
            Code::MissingProperty,
            "this find needs `replace` (format-only finds need a native run split)",
        )
        .with_property(super::REPLACE)
        .with_format(format)
    })
}

/// Plan a find over `runs`: the changed runs and the match count.
///
/// # Errors
///
/// `unsupported_capability` when a match touches a run the leaf cannot
/// rewrite, `limit_exceeded` beyond `max_staged_bytes`, and allocation
/// failures.
pub(crate) fn find(
    runs: &[Run<'_>],
    spec: &Spec<'_>,
    replacement: &str,
    format: FileFormat,
    limits: &SemanticLimits,
) -> Result<(Vec<Change>, usize)> {
    let pieces = pieces(runs)?;
    let matcher = spec.matcher();
    let planned = plan_with(
        &mut Scratch::default(),
        &pieces,
        &matcher,
        Some(replacement),
        false,
        limits.max_staged_bytes(),
    )?;
    let mut changes = Vec::new();
    changes
        .try_reserve_exact(planned.rewrites.len())
        .map_err(allocation)?;
    for rewrite in planned.rewrites {
        let Some(text) = rewrite.text else {
            continue;
        };
        if !runs[rewrite.piece].editable {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "a match touches text the format cannot rewrite in place (a field, or markup inside the run)",
            )
            .with_property(super::FIND)
            .with_format(format)
            .with_suggestion("narrow the find's target to text outside fields and nested markup (record 0295)"));
        }
        changes.push(Change {
            run: rewrite.piece,
            text,
        });
    }
    Ok((changes, planned.matches))
}

/// Plan a template merge over `runs`. Placeholders that touch a run the
/// leaf cannot rewrite are not placeholders. With `breaks` false, a value
/// holding a line break is refused (`invalid_value`); otherwise carriage
/// returns become line feeds and the adapter writes them.
///
/// # Errors
///
/// The planner's value and limit errors, and the refusal above.
pub(crate) fn merge(
    runs: &[Run<'_>],
    planner: &mut Planner<'_>,
    breaks: bool,
    format: FileFormat,
    limits: &SemanticLimits,
) -> Result<Vec<Change>> {
    let mut changes = Vec::new();
    let mut first = 0;
    let mut text = String::new();
    let mut fixed: Vec<Range<usize>> = Vec::new();
    while first < runs.len() {
        let group = runs[first].group;
        let end = runs[first..]
            .iter()
            .position(|run| run.group != group)
            .map_or(runs.len(), |offset| first + offset);
        let members = &runs[first..end];
        if !members.iter().any(|run| run.text.contains('{')) {
            first = end;
            continue;
        }
        text.clear();
        fixed.clear();
        for run in members {
            let start = text.len();
            text.push_str(&run.text);
            if !run.editable {
                fixed.push(start..text.len());
            }
        }
        let mut ranges: Vec<Range<usize>> = Vec::new();
        let mut values: Vec<Cow<'_, str>> = Vec::new();
        planner.text(
            &text,
            |range| {
                !fixed
                    .iter()
                    .any(|fixed| fixed.start < range.end && range.start < fixed.end)
            },
            &[],
            |range, value| {
                ranges.push(range);
                values.push(normalize_breaks(value));
            },
        )?;
        if !ranges.is_empty() {
            if !breaks && values.iter().any(|value| value.contains('\n')) {
                return Err(Error::new(
                    Code::InvalidValue,
                    "a merge value for this format cannot hold a line break",
                )
                .with_format(format)
                .with_suggestion("remove line breaks from the data"));
            }
            let pieces = pieces(members)?;
            let values: Vec<&str> = values.iter().map(AsRef::as_ref).collect();
            for rewrite in rewrite_ranges(
                &pieces,
                &ranges,
                Values::Each(&values),
                limits.max_staged_bytes(),
            )? {
                if let Some(text) = rewrite.text {
                    changes.try_reserve(1).map_err(allocation)?;
                    changes.push(Change {
                        run: first + rewrite.piece,
                        text,
                    });
                }
            }
        }
        first = end;
    }
    Ok(changes)
}

/// Whether the content of an ODF run (between its tags) holds only text,
/// spaces (`text:s`), tabs, line breaks and soft page breaks: what a leaf
/// text replacement re-encodes without losing markup.
#[cfg(any(feature = "odt", feature = "odp"))]
pub(crate) fn odf_text_only(content: &str) -> bool {
    let mut rest = content;
    while let Some(start) = rest.find('<') {
        let tag = &rest[start + 1..];
        let end = tag.find('>').unwrap_or(tag.len());
        let name = tag[..end]
            .trim_start_matches('/')
            .split(|character: char| character.is_whitespace() || character == '/')
            .next()
            .unwrap_or("");
        let local = name.rsplit(':').next().unwrap_or(name);
        if !matches!(local, "s" | "tab" | "line-break" | "soft-page-break") {
            return false;
        }
        rest = &tag[end..];
    }
    true
}

/// Call `push` with every node `target` selects: the node of a path, or
/// each match of a query.
///
/// # Errors
///
/// Resolution errors, `invalid_selector` for a query with unsupported
/// predicates, `invalid_argument` for an empty target, and `push`'s.
pub(crate) fn each_target<'f, N>(
    nav: N,
    target: &crate::target::TargetArg<'_>,
    mut push: impl FnMut(N::Node) -> Result<()>,
) -> Result<()>
where
    N: crate::query::inspect::Inspect<'f>,
    N::Node: Clone,
{
    if let Some(path) = target.path() {
        push(crate::path::resolve::get(&nav, path)?)
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(nav, query.clone())?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(Error::new(
                Code::InvalidSelector,
                "the mutation selector contains an unsupported predicate",
            ));
        }
        for node in walk {
            push(node?)?;
        }
        Ok(())
    } else {
        Err(Error::new(
            Code::InvalidArgument,
            "the mutation target is empty",
        ))
    }
}

fn pieces<'r>(runs: &'r [Run<'_>]) -> Result<Vec<Piece<'r>>> {
    let mut pieces = Vec::new();
    pieces.try_reserve_exact(runs.len()).map_err(allocation)?;
    pieces.extend(runs.iter().map(|run| Piece {
        text: &run.text,
        group: run.group,
        scoped: run.scoped,
    }));
    Ok(pieces)
}

fn allocation(source: std::collections::TryReserveError) -> Error {
    Error::from(TransactionError::Allocation { source })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::merge::{Options, Record};
    use crate::props::Props;

    fn spec<'a>(find: &'a str, replace: &'a str) -> Props<'a> {
        Props::new().with("find", find).with("replace", replace)
    }

    #[test]
    fn find_collapses_across_runs_and_refuses_fixed_runs() {
        let runs = [
            Run::new("Hel", 0),
            Run::new("lo world", 0),
            Run::new("Hello", 1),
        ];
        let props = spec("Hello", "Bye");
        let spec = Spec::admit(&props).expect("spec");
        let (changes, matches) = find(
            &runs,
            &spec,
            "Bye",
            FileFormat::Doc,
            &SemanticLimits::default(),
        )
        .expect("find");
        assert_eq!(matches, 2);
        assert_eq!(
            changes,
            [
                Change {
                    run: 0,
                    text: "Bye".into()
                },
                Change {
                    run: 1,
                    text: " world".into()
                },
                Change {
                    run: 2,
                    text: "Bye".into()
                },
            ]
        );
        let fixed = [Run::new("Hello", 0).editable(false)];
        let error = find(
            &fixed,
            &spec,
            "Bye",
            FileFormat::Odt,
            &SemanticLimits::default(),
        )
        .expect_err("fixed");
        assert_eq!(error.code(), Code::UnsupportedCapability);
    }

    #[test]
    fn merge_fills_groups_and_skips_fixed_runs() {
        let runs = [
            Run::new("Dear {{na", 0),
            Run::new("me}},", 0),
            Run::new("{{name}}", 1).editable(false),
            Run::new("{{city}}", 2),
        ];
        let record: Record = [("name", "Ada"), ("city", "A\r\nB")].into_iter().collect();
        let limits = SemanticLimits::default();
        let mut planner = Planner::new(&record, &Options::new(), &limits);
        let changes = merge(&runs, &mut planner, true, FileFormat::Odt, &limits).expect("merge");
        assert_eq!(
            changes,
            [
                Change {
                    run: 0,
                    text: "Dear Ada".into()
                },
                Change {
                    run: 1,
                    text: ",".into()
                },
                Change {
                    run: 3,
                    text: "A\nB".into()
                },
            ]
        );
        let mut planner = Planner::new(&record, &Options::new(), &limits);
        let error = merge(&runs, &mut planner, false, FileFormat::Ppt, &limits).expect_err("break");
        assert_eq!(error.code(), Code::InvalidValue);
    }
}
