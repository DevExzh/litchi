//! Namespace bindings in scope, indexed by prefix, for the MCE processors.
//!
//! Both processors also keep a persistent chain of per-element declaration
//! layers. Resolving a prefix by walking that chain costs one comparison per
//! declaration in scope, and the input chooses that number: the declarations
//! of every open ancestor count, shadowed re-declarations included. [`Scope`]
//! answers the same questions with a logarithmic number of comparisons.
//!
//! It is kept in step with the element stack: an element's declarations are
//! added with its depth once its start tag has been checked, and
//! [`Scope::truncate`] removes them when it closes. Depths are positive and
//! strictly increase from an element to its children, so the bindings of one
//! prefix are always ordered by depth, outermost first.

use std::collections::BTreeMap;

use super::model::{Error, XML_NS};

/// The bindings in scope at the element being processed.
#[derive(Debug, Default)]
pub(super) struct Scope {
    /// Each prefix's bindings, outermost first, with the depth of the element
    /// that declared each.
    prefixes: BTreeMap<String, Vec<(usize, String)>>,
    /// Every binding in `prefixes`, in the order it was added, as the depth
    /// that declared it and its prefix.
    added: Vec<(usize, String)>,
}

impl Scope {
    /// The namespace `prefix` is bound to in scope.
    pub(super) fn get(&self, prefix: &str) -> Option<&str> {
        if prefix == "xml" {
            return Some(XML_NS);
        }
        self.prefixes
            .get(prefix)
            .and_then(|bindings| bindings.last())
            .map(|(_, namespace)| namespace.as_str())
    }

    /// The namespace `prefix` is bound to by an element shallower than
    /// `depth`, that is in scope at the parent of the element at `depth`.
    pub(super) fn get_outside(&self, prefix: &str, depth: usize) -> Option<&str> {
        if prefix == "xml" {
            return Some(XML_NS);
        }
        // One element binds a prefix at most once, so the binding sought is
        // the last or the one before it.
        self.prefixes.get(prefix).and_then(|bindings| {
            bindings
                .iter()
                .rev()
                .take(2)
                .find(|(declared, _)| *declared < depth)
                .map(|(_, namespace)| namespace.as_str())
        })
    }

    /// Whether the element at `depth` itself declares `prefix`.
    pub(super) fn declared_at(&self, prefix: &str, depth: usize) -> bool {
        self.prefixes
            .get(prefix)
            .and_then(|bindings| bindings.last())
            .is_some_and(|(declared, _)| *declared == depth)
    }

    /// Add the declarations of the element at `depth`, which binds each
    /// prefix at most once and is deeper than every binding in scope.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`] when the log of added bindings cannot grow.
    pub(super) fn push(&mut self, depth: usize, local: &[(String, String)]) -> Result<(), Error> {
        self.added
            .try_reserve(local.len())
            .map_err(|source| Error::Allocation {
                resource: "MCE namespace scope",
                source,
            })?;
        for (prefix, namespace) in local {
            debug_assert!(
                self.added.last().is_none_or(|(last, _)| *last <= depth),
                "namespace scope depths must not decrease"
            );
            match self.prefixes.get_mut(prefix.as_str()) {
                Some(bindings) => bindings.push((depth, namespace.clone())),
                None => {
                    self.prefixes
                        .insert(prefix.clone(), vec![(depth, namespace.clone())]);
                },
            }
            self.added.push((depth, prefix.clone()));
        }
        Ok(())
    }

    /// Remove every binding declared at `depth` or deeper: those of an element
    /// that has closed and of anything it left open.
    pub(super) fn truncate(&mut self, depth: usize) {
        while self
            .added
            .last()
            .is_some_and(|(declared, _)| *declared >= depth)
        {
            let Some((declared, prefix)) = self.added.pop() else {
                break;
            };
            if let Some(bindings) = self.prefixes.get_mut(prefix.as_str()) {
                debug_assert!(
                    bindings.last().is_some_and(|(last, _)| *last == declared),
                    "a prefix's latest binding must be the latest one added"
                );
                bindings.pop();
                if bindings.is_empty() {
                    self.prefixes.remove(prefix.as_str());
                }
            }
        }
    }
}

/// Whether two of one element's declarations bind the same prefix.
///
/// Sorting borrowed prefixes keeps the check `O(n log n)` in the number of
/// declarations; a pairwise comparison would be quadratic in a count the input
/// chooses.
///
/// # Errors
///
/// [`Error::Allocation`] when the sort buffer cannot be reserved.
pub(super) fn has_duplicate_prefix(local: &[(String, String)]) -> Result<bool, Error> {
    if local.len() < 2 {
        return Ok(false);
    }
    let mut prefixes = sorted_prefixes(local)?;
    prefixes.dedup();
    Ok(prefixes.len() != local.len())
}

/// One element's declared prefixes, sorted for binary search.
///
/// # Errors
///
/// [`Error::Allocation`] when the buffer cannot be reserved.
pub(super) fn sorted_prefixes(local: &[(String, String)]) -> Result<Vec<&str>, Error> {
    let mut prefixes = Vec::new();
    prefixes
        .try_reserve_exact(local.len())
        .map_err(|source| Error::Allocation {
            resource: "MCE namespace declaration check",
            source,
        })?;
    prefixes.extend(local.iter().map(|(prefix, _)| prefix.as_str()));
    prefixes.sort_unstable();
    Ok(prefixes)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn declarations(pairs: &[(&str, &str)]) -> Vec<(String, String)> {
        pairs
            .iter()
            .map(|(prefix, namespace)| ((*prefix).to_owned(), (*namespace).to_owned()))
            .collect()
    }

    #[test]
    fn nested_bindings_shadow_and_truncation_restores_them() {
        let mut scope = Scope::default();
        scope
            .push(1, &declarations(&[("a", "urn:a1"), ("", "urn:d1")]))
            .unwrap();
        scope.push(2, &declarations(&[("b", "urn:b2")])).unwrap();
        scope
            .push(3, &declarations(&[("a", "urn:a3"), ("", "")]))
            .unwrap();
        assert_eq!(scope.get("a"), Some("urn:a3"));
        assert_eq!(scope.get(""), Some(""));
        assert_eq!(scope.get("b"), Some("urn:b2"));
        assert_eq!(scope.get("c"), None);
        assert_eq!(scope.get("xml"), Some(XML_NS));
        assert_eq!(scope.get_outside("a", 3), Some("urn:a1"));
        assert_eq!(scope.get_outside("b", 3), Some("urn:b2"));
        assert_eq!(scope.get_outside("b", 2), None);
        assert!(scope.declared_at("a", 3));
        assert!(!scope.declared_at("b", 3));

        scope.truncate(3);
        assert_eq!(scope.get("a"), Some("urn:a1"));
        assert_eq!(scope.get(""), Some("urn:d1"));
        scope.truncate(2);
        assert_eq!(scope.get("b"), None);
        scope.truncate(1);
        assert_eq!(scope.get("a"), None);
        assert!(scope.prefixes.is_empty());
        assert!(scope.added.is_empty());
    }

    #[test]
    fn truncating_an_element_without_declarations_removes_nothing() {
        let mut scope = Scope::default();
        scope.push(1, &declarations(&[("a", "urn:a")])).unwrap();
        scope.truncate(2);
        assert_eq!(scope.get("a"), Some("urn:a"));
    }

    #[test]
    fn duplicate_prefixes_are_found_in_any_position() {
        assert!(!has_duplicate_prefix(&declarations(&[])).unwrap());
        assert!(!has_duplicate_prefix(&declarations(&[("a", "1")])).unwrap());
        assert!(!has_duplicate_prefix(&declarations(&[("a", "1"), ("b", "1")])).unwrap());
        assert!(has_duplicate_prefix(&declarations(&[("a", "1"), ("a", "2")])).unwrap());
        assert!(
            has_duplicate_prefix(&declarations(&[
                ("", "1"),
                ("b", "1"),
                ("c", "1"),
                ("", "2")
            ]))
            .unwrap()
        );
    }
}
