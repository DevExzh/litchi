//! Targets of the `ProcessContent`, `PreserveElements` and
//! `PreserveAttributes` compatibility directives.
//!
//! An element may carry thousands of directive tokens and every ancestor with
//! a directive adds a layer, so testing a name against a layer's targets by
//! scanning them would cost a number of comparisons the input chooses, once per
//! element and per ignorable attribute. [`Patterns`] keeps exact names and
//! whole-namespace wildcards in separate sets, so that test is two lookups.

use std::collections::HashSet;

use super::model::{Error, Name};

/// One directive target: an exact expanded name, or every name in a
/// namespace (`prefix:*`).
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub(super) enum NamePattern {
    Exact(Name),
    Namespace(String),
}

impl NamePattern {
    /// The namespace the target belongs to.
    pub(super) fn namespace(&self) -> &str {
        match self {
            Self::Exact(name) => &name.namespace,
            Self::Namespace(namespace) => namespace,
        }
    }
}

/// The targets of one directive on one element.
#[derive(Debug, Default)]
pub(super) struct Patterns {
    exact: HashSet<Name>,
    namespaces: HashSet<String>,
}

impl Patterns {
    /// Whether the directive names no target.
    pub(super) fn is_empty(&self) -> bool {
        self.exact.is_empty() && self.namespaces.is_empty()
    }

    /// Add one target; `false` when the directive already names it.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`], naming `resource`, when the set cannot grow.
    pub(super) fn insert(
        &mut self,
        pattern: NamePattern,
        resource: &'static str,
    ) -> Result<bool, Error> {
        let allocation = |source| Error::Allocation { resource, source };
        match pattern {
            NamePattern::Exact(name) => {
                self.exact.try_reserve(1).map_err(allocation)?;
                Ok(self.exact.insert(name))
            },
            NamePattern::Namespace(namespace) => {
                self.namespaces.try_reserve(1).map_err(allocation)?;
                Ok(self.namespaces.insert(namespace))
            },
        }
    }

    /// Whether a target matches `name`: the exact name, or its namespace.
    pub(super) fn matches(&self, name: &Name) -> bool {
        self.exact.contains(name) || self.namespaces.contains(name.namespace.as_str())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn name(namespace: &str, local: &str) -> Name {
        Name {
            namespace: namespace.to_owned(),
            local_name: local.to_owned(),
        }
    }

    #[test]
    fn exact_and_namespace_targets_match_as_a_scan_of_both_would() {
        let mut patterns = Patterns::default();
        assert!(patterns.is_empty());
        assert!(
            patterns
                .insert(NamePattern::Exact(name("urn:a", "x")), "test")
                .unwrap()
        );
        assert!(
            !patterns
                .insert(NamePattern::Exact(name("urn:a", "x")), "test")
                .unwrap()
        );
        assert!(
            patterns
                .insert(NamePattern::Namespace("urn:b".to_owned()), "test")
                .unwrap()
        );
        assert!(
            !patterns
                .insert(NamePattern::Namespace("urn:b".to_owned()), "test")
                .unwrap()
        );
        assert!(!patterns.is_empty());

        assert!(patterns.matches(&name("urn:a", "x")));
        assert!(!patterns.matches(&name("urn:a", "y")));
        assert!(patterns.matches(&name("urn:b", "anything")));
        assert!(!patterns.matches(&name("urn:c", "x")));
        assert_eq!(
            NamePattern::Namespace("urn:b".to_owned()).namespace(),
            "urn:b"
        );
        assert_eq!(NamePattern::Exact(name("urn:a", "x")).namespace(), "urn:a");
    }
}
