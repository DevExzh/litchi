//! Allocation-conscious overlap helpers for disjoint transaction joins.

use std::collections::BTreeMap;

/// Return keys whose actions overlap in two sorted maps.
///
/// The merge walk keeps the join path linear in the number of staged actions,
/// while one helper serves cells, rows, and columns without duplicating three
/// subtly different map scans.
pub(super) fn overlapping_keys<K, V, F>(
    left: &BTreeMap<K, V>,
    right: &BTreeMap<K, V>,
    mut overlaps: F,
) -> Vec<K>
where
    K: Ord + Copy,
    F: FnMut(&V, &V) -> bool,
{
    let mut left_iter = left.iter();
    let mut right_iter = right.iter();
    let mut left_entry = left_iter.next();
    let mut right_entry = right_iter.next();
    let mut keys = Vec::new();

    while let (Some((left_key, left_action)), Some((right_key, right_action))) =
        (left_entry, right_entry)
    {
        match left_key.cmp(right_key) {
            std::cmp::Ordering::Less => left_entry = left_iter.next(),
            std::cmp::Ordering::Greater => right_entry = right_iter.next(),
            std::cmp::Ordering::Equal => {
                if overlaps(left_action, right_action) {
                    keys.push(*left_key);
                }
                left_entry = left_iter.next();
                right_entry = right_iter.next();
            },
        }
    }
    keys
}
