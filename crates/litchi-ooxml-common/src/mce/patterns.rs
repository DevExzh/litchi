//! Targets of the `ProcessContent`, `PreserveElements` and
//! `PreserveAttributes` compatibility directives.
//!
//! An element may carry thousands of directive tokens and every ancestor with
//! a directive adds a layer, so testing a name against a layer's targets by
//! scanning them would cost a number of comparisons the input chooses, once per
//! element and per ignorable attribute. [`Patterns`] keeps exact names by
//! namespace and whole-namespace wildcards in separate tables, so that test is
//! two lookups. Each table is created with its first target, so an element
//! without directives pays for none.
//!
//! The namespace key is generic: the in-memory processor uses namespace
//! identities (`scope::UriId`), as does the stream. A directive lookup does
//! not hash namespace URI text for each name.

use core::borrow::Borrow;
use core::hash::Hash;
use std::collections::{HashMap, HashSet};

use super::model::Error;

/// One directive target: an exact expanded name, or every name in a
/// namespace (`prefix:*`).
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub(super) enum NamePattern<N> {
    Exact(N, String),
    Namespace(N),
}

impl<N> NamePattern<N> {
    /// The namespace the target belongs to.
    pub(super) const fn namespace(&self) -> &N {
        match self {
            Self::Exact(namespace, _) | Self::Namespace(namespace) => namespace,
        }
    }
}

/// The targets of one directive on one element.
#[derive(Debug)]
pub(super) struct Patterns<N> {
    exact: Option<HashMap<N, HashSet<String>>>,
    namespaces: Option<HashSet<N>>,
}

impl<N> Default for Patterns<N> {
    fn default() -> Self {
        Self {
            exact: None,
            namespaces: None,
        }
    }
}

impl<N: Eq + Hash> Patterns<N> {
    /// Whether the directive names no target.
    pub(super) fn is_empty(&self) -> bool {
        self.exact.as_ref().is_none_or(HashMap::is_empty)
            && self.namespaces.as_ref().is_none_or(HashSet::is_empty)
    }

    /// Add one target; `false` when the directive already names it.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`], naming `resource`, when a table cannot grow.
    pub(super) fn insert(
        &mut self,
        pattern: NamePattern<N>,
        resource: &'static str,
    ) -> Result<bool, Error> {
        let allocation = |source| Error::Allocation { resource, source };
        match pattern {
            NamePattern::Exact(namespace, local) => {
                let exact = self.exact.get_or_insert_with(HashMap::new);
                exact.try_reserve(1).map_err(allocation)?;
                let locals = exact.entry(namespace).or_default();
                locals.try_reserve(1).map_err(allocation)?;
                Ok(locals.insert(local))
            },
            NamePattern::Namespace(namespace) => {
                let namespaces = self.namespaces.get_or_insert_with(HashSet::new);
                namespaces.try_reserve(1).map_err(allocation)?;
                Ok(namespaces.insert(namespace))
            },
        }
    }

    /// Whether a target matches the name `local` in `namespace`: the exact
    /// name, or its namespace.
    pub(super) fn matches<Q>(&self, namespace: &Q, local: &str) -> bool
    where
        N: Borrow<Q>,
        Q: Eq + Hash + ?Sized,
    {
        self.exact
            .as_ref()
            .and_then(|exact| exact.get(namespace))
            .is_some_and(|locals| locals.contains(local))
            || self
                .namespaces
                .as_ref()
                .is_some_and(|namespaces| namespaces.contains(namespace))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn exact_and_namespace_targets_match_as_a_scan_of_both_would() {
        let mut patterns = Patterns::<String>::default();
        assert!(patterns.is_empty());
        let exact = || NamePattern::Exact("urn:a".to_owned(), "x".to_owned());
        let wildcard = || NamePattern::Namespace("urn:b".to_owned());
        assert!(patterns.insert(exact(), "test").unwrap());
        assert!(!patterns.insert(exact(), "test").unwrap());
        assert!(patterns.insert(wildcard(), "test").unwrap());
        assert!(!patterns.insert(wildcard(), "test").unwrap());
        assert!(!patterns.is_empty());

        assert!(patterns.matches("urn:a", "x"));
        assert!(!patterns.matches("urn:a", "y"));
        assert!(patterns.matches("urn:b", "anything"));
        assert!(!patterns.matches("urn:c", "x"));
        assert_eq!(wildcard().namespace(), "urn:b");
        assert_eq!(exact().namespace(), "urn:a");

        let mut by_identity = Patterns::<usize>::default();
        assert!(
            by_identity
                .insert(NamePattern::Exact(7, "x".to_owned()), "test")
                .unwrap()
        );
        assert!(by_identity.matches(&7, "x"));
        assert!(!by_identity.matches(&8, "x"));
    }
}
