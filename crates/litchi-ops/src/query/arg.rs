//! [`IntoQuery`]: the one way `query` takes a selector, as text, as a parsed
//! [`Query`], or as a compiled [`Plan`] (record 0010 §8.1, decision D9).

use super::inspect::Inspect;
use super::plan::Plan;
use super::walk::PlanRef;
use crate::Result;
use litchi_core::selector::Query;

/// A value `query` accepts as a selector.
///
/// Text is parsed once, borrowing from the input; a parsed [`Query`] is
/// compiled for the document's format when the query runs; a compiled
/// [`Plan`] is used as is, so a selector run on many documents of one format
/// is parsed and compiled once. A query the same thread runs repeatedly for
/// one format is also compiled once: its plan is kept in a small per-thread
/// cache from its second run on (record 0291).
///
/// ```
/// use litchi_ops::query::{IntoQuery, Query};
///
/// let text = "paragraph[style=Heading1] > run[bold=true]".into_query()?;
/// assert_eq!(text.text(), Some("paragraph[style=Heading1] > run[bold=true]"));
/// assert_eq!(text.query().map(|query| query.alternatives().len()), Some(1));
///
/// let parsed = Query::parse("shape:contains(Revenue), picture:no-alt")?;
/// let arg = (&parsed).into_query()?;
/// assert_eq!(arg.query(), Some(&parsed));
/// assert_eq!(arg.text(), None);
///
/// let error = "paragraph[".into_query().unwrap_err();
/// assert_eq!(error.code().as_str(), "invalid_selector");
/// # Ok::<(), litchi_ops::Error>(())
/// ```
pub trait IntoQuery<'q> {
    /// Parses text, or wraps a parsed query or a compiled plan.
    ///
    /// # Errors
    ///
    /// `invalid_selector` for text that is not selector syntax, with the
    /// byte offset in the message and the grammar's typed error as the
    /// source; `max_depth_exceeded` for selectors beyond the nesting limits.
    fn into_query(self) -> Result<QueryArg<'q>>;
}

impl<'q> IntoQuery<'q> for &'q str {
    fn into_query(self) -> Result<QueryArg<'q>> {
        QueryArg::parse(self)
    }
}

impl<'q> IntoQuery<'q> for &'q String {
    fn into_query(self) -> Result<QueryArg<'q>> {
        QueryArg::parse(self)
    }
}

impl<'q> IntoQuery<'q> for Query<'q> {
    fn into_query(self) -> Result<QueryArg<'q>> {
        Ok(QueryArg {
            held: Held::Query(self),
            text: None,
            raw: false,
        })
    }
}

impl<'q> IntoQuery<'q> for &'q Query<'q> {
    fn into_query(self) -> Result<QueryArg<'q>> {
        Ok(QueryArg {
            held: Held::QueryRef(self),
            text: None,
            raw: false,
        })
    }
}

impl<'q> IntoQuery<'q> for Plan<'q> {
    fn into_query(self) -> Result<QueryArg<'q>> {
        Ok(QueryArg {
            held: Held::Plan(self),
            text: None,
            raw: false,
        })
    }
}

impl<'q> IntoQuery<'q> for &'q Plan<'q> {
    fn into_query(self) -> Result<QueryArg<'q>> {
        Ok(QueryArg {
            held: Held::PlanRef(self),
            text: None,
            raw: false,
        })
    }
}

impl<'q> IntoQuery<'q> for QueryArg<'q> {
    fn into_query(self) -> Result<QueryArg<'q>> {
        Ok(self)
    }
}

/// A borrowed selector runs without copying it: a parsed query is compiled
/// from the reference and a compiled plan is used in place (record 0287).
impl<'a, 'q: 'a> IntoQuery<'a> for &'a QueryArg<'q> {
    fn into_query(self) -> Result<QueryArg<'a>> {
        Ok(QueryArg {
            held: match &self.held {
                Held::Query(query) => Held::QueryRef(query),
                Held::QueryRef(query) => Held::QueryRef(query),
                Held::Plan(plan) => Held::PlanRef(plan),
                Held::PlanRef(plan) => Held::PlanRef(plan),
            },
            text: self.text,
            raw: self.raw,
        })
    }
}

/// A selector ready to run: parsed text, a parsed [`Query`], or a compiled
/// [`Plan`].
///
/// ```
/// use litchi_ops::query::QueryArg;
///
/// let arg = QueryArg::parse("'My Data'!row[Amount>5000]")?;
/// assert_eq!(arg.text(), Some("'My Data'!row[Amount>5000]"));
/// assert!(arg.plan().is_none());
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Clone)]
pub struct QueryArg<'q> {
    held: Held<'q>,
    text: Option<&'q str>,
    pub(super) raw: bool,
}

#[derive(Clone)]
#[allow(
    clippy::large_enum_variant,
    reason = "a selector is wrapped once per query and moved into the walk; boxing the parsed \
              query or the plan would add an allocation to every query"
)]
enum Held<'q> {
    Query(Query<'q>),
    QueryRef(&'q Query<'q>),
    Plan(Plan<'q>),
    PlanRef(&'q Plan<'q>),
}

impl<'q> QueryArg<'q> {
    /// Parses selector text, borrowing from it.
    ///
    /// # Errors
    ///
    /// As [`IntoQuery::into_query`].
    pub fn parse(text: &'q str) -> Result<Self> {
        Ok(Self {
            held: Held::Query(Query::parse(text)?),
            text: Some(text),
            raw: false,
        })
    }

    /// The text the selector was parsed from, or `None` for a parsed query
    /// or a plan.
    #[must_use]
    pub fn text(&self) -> Option<&'q str> {
        self.text
    }

    /// The parsed query, or `None` for a compiled plan.
    #[must_use]
    pub fn query(&self) -> Option<&Query<'q>> {
        match &self.held {
            Held::Query(query) => Some(query),
            Held::QueryRef(query) => Some(query),
            Held::Plan(_) | Held::PlanRef(_) => None,
        }
    }

    /// The compiled plan, or `None` for a selector compiled when it runs.
    #[must_use]
    pub fn plan(&self) -> Option<&Plan<'q>> {
        match &self.held {
            Held::Plan(plan) => Some(plan),
            Held::PlanRef(plan) => Some(plan),
            Held::Query(_) | Held::QueryRef(_) => None,
        }
    }

    /// The plan to run on documents the adapter `N` reads: compiled now
    /// unless the caller passed one, or taken from this thread's plan
    /// cache when it compiled the same query for `N` recently (record 0291).
    pub(crate) fn into_plan<'f, N: Inspect<'f>>(self) -> Result<PlanRef<'q>> {
        match self.held {
            Held::Query(query) => super::cache::compile::<N>(&query, self.raw).map(PlanRef::from),
            Held::QueryRef(query) => super::cache::compile::<N>(query, self.raw).map(PlanRef::from),
            Held::Plan(mut plan) => {
                plan.raw |= self.raw;
                Ok(PlanRef::Owned(Box::new(plan)))
            },
            Held::PlanRef(plan) if self.raw && !plan.raw => {
                let mut raw_plan = plan.clone();
                raw_plan.raw = true;
                Ok(PlanRef::Owned(Box::new(raw_plan)))
            },
            Held::PlanRef(plan) => Ok(PlanRef::Borrowed(plan)),
        }
    }
}

impl std::fmt::Debug for QueryArg<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str("QueryArg { .. }")
    }
}
