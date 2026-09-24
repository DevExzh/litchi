//! A set of one start tag's attribute names whose cost stays bounded on
//! hostile tags.
//!
//! This mirrors `SeenNames` in `litchi_ooxml_common::xml::attributes`, which
//! this crate cannot depend on. Membership costs a linear scan while the set
//! holds at most 32 names and `O(log n)` comparisons after that, so checking
//! every attribute of a tag for a duplicate costs `O(n log n)` rather than the
//! `O(n²)` of scanning the attributes read so far.

use std::collections::BTreeSet;

/// Names the set holds before it leaves its linear scan for an ordered
/// index. It matches quick-xml's own threshold, so ordinary tags cost what
/// quick-xml's duplicate check costs.
const LINEAR_NAMES: usize = 32;

/// The names one tag has shown so far.
#[derive(Debug)]
pub(crate) struct SeenNames<K> {
    linear: Vec<K>,
    ordered: BTreeSet<K>,
}

impl<K> Default for SeenNames<K> {
    fn default() -> Self {
        Self {
            linear: Vec::new(),
            ordered: BTreeSet::new(),
        }
    }
}

impl<K: Ord> SeenNames<K> {
    /// Insert `key`; `false` when it was already present.
    pub(crate) fn insert(&mut self, key: K) -> bool {
        if self.ordered.is_empty() {
            if self.linear.contains(&key) {
                return false;
            }
            if self.linear.len() < LINEAR_NAMES {
                self.linear.push(key);
                return true;
            }
            self.ordered.extend(self.linear.drain(..));
        }
        self.ordered.insert(key)
    }
}

#[cfg(test)]
mod tests {
    use super::{LINEAR_NAMES, SeenNames};

    #[test]
    fn names_are_kept_across_the_switch_to_an_ordered_index() {
        let mut seen = SeenNames::default();
        for index in 0..100u32 {
            assert!(seen.insert(index), "{index}");
            assert!(!seen.insert(index), "{index}");
        }
        for index in 0..100u32 {
            assert!(!seen.insert(index), "{index}");
        }
        assert_eq!(seen.linear.len() + seen.ordered.len(), 100);
        assert!(seen.linear.is_empty());
        assert!(seen.insert(100));
    }

    #[test]
    fn a_small_set_stays_linear() {
        let mut seen = SeenNames::default();
        for index in 0..LINEAR_NAMES {
            assert!(seen.insert(index));
        }
        assert!(!seen.insert(0));
        assert_eq!(seen.linear.len(), LINEAR_NAMES);
        assert!(seen.ordered.is_empty());
    }
}
