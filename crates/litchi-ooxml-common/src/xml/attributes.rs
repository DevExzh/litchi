//! Attribute iteration whose cost stays bounded on hostile start tags.
//!
//! quick-xml 0.41 checks a start tag's attribute names for duplicates by
//! default. From the 33rd name on it pre-filters with an unkeyed hash and, on
//! every hit, scans the names seen so far. A duplicate therefore costs a scan
//! of the tag's earlier names, and the checked iterator keeps going after it
//! reports one. A caller that skips errors and reads on, as a lenient reader
//! does, spends time quadratic in the tag's attributes on a tag of `D` distinct
//! names followed by `R` repeats of the last one. No crafting is needed.
//!
//! Its recovery after a duplicate is also unreliable: it resumes scanning at
//! the duplicate's `=` and skips only to the next whitespace, so a duplicate
//! whose quoted value contains whitespace makes it report names from inside
//! that value. `<e a="1" a="x b='y'"/>` yields `a="1"`, an error, and a `b`
//! attribute the tag does not have.
//!
//! [`first_wins`] yields what a lenient reader relies on — every qualified
//! name once, at its first occurrence — at a cost of `O(n log n)` comparisons
//! for `n` attributes, and parses every value as written.

use std::collections::{BTreeMap, BTreeSet};

use quick_xml::events::BytesStart;
use quick_xml::events::attributes::{AttrError, Attribute, Attributes};

/// Names a tag may carry before [`first_wins`] and [`SeenNames`] leave their
/// linear scan for an ordered index. It matches quick-xml's own threshold, so
/// ordinary tags cost what quick-xml's check costs.
const LINEAR_NAMES: usize = 32;

/// Iterate `tag`'s attributes with each qualified name's first occurrence
/// yielded as `Ok`.
///
/// Every later occurrence of a name is yielded as
/// `Err(AttrError::Duplicated(position, first_position))`, positions relative
/// to the start of the tag as quick-xml reports them, and its value is parsed
/// and skipped as written. Other lexical errors pass through unchanged.
///
/// On a tag without duplicate names this yields exactly what
/// `tag.attributes()` yields. On a tag with duplicates it yields the same
/// `Ok` items in the same order unless quick-xml's recovery would resume
/// inside a duplicate's value (see the module documentation), or unless a
/// name's first occurrence is itself malformed, in which case quick-xml
/// reports its later occurrences as duplicates while this yields the first
/// well-formed one.
///
/// It costs `O(log n)` name comparisons per attribute once a tag has more
/// than 32, and a linear scan of the names seen so far below that.
#[must_use]
pub fn first_wins<'a>(tag: &'a BytesStart<'_>) -> FirstWins<'a> {
    let mut attributes = tag.attributes();
    attributes.with_checks(false);
    FirstWins {
        attributes,
        base: tag,
        seen: FirstSeen::default(),
    }
}

/// The number of `tag`'s attributes, counting at most `maximum + 1`.
///
/// Items are counted as quick-xml's iterator yields them, errors included,
/// without its duplicate check, so the count costs `O(min(n, maximum))`
/// whatever the tag holds. A caller that refuses a tag over `maximum`
/// compares the result with it.
#[must_use]
pub fn count_up_to(tag: &BytesStart<'_>, maximum: usize) -> usize {
    tag.attributes()
        .with_checks(false)
        .take(maximum.saturating_add(1))
        .count()
}

/// The iterator [`first_wins`] returns.
#[derive(Debug)]
pub struct FirstWins<'a> {
    attributes: Attributes<'a>,
    base: &'a [u8],
    seen: FirstSeen<'a>,
}

impl<'a> Iterator for FirstWins<'a> {
    type Item = Result<Attribute<'a>, AttrError>;

    fn next(&mut self) -> Option<Self::Item> {
        let attribute = match self.attributes.next()? {
            Ok(attribute) => attribute,
            Err(error) => return Some(Err(error)),
        };
        let name = attribute.key.into_inner();
        let position = offset_in(self.base, name);
        match self.seen.first(name, position) {
            None => Some(Ok(attribute)),
            Some(first) => Some(Err(AttrError::Duplicated(position, first))),
        }
    }
}

impl core::iter::FusedIterator for FirstWins<'_> {}

/// The byte offset of `part` within `base`, of which it is a subslice.
fn offset_in(base: &[u8], part: &[u8]) -> usize {
    (part.as_ptr() as usize).wrapping_sub(base.as_ptr() as usize)
}

/// The names one tag has shown so far, with the position of each first
/// occurrence.
#[derive(Debug)]
enum FirstSeen<'a> {
    Linear(Vec<(&'a [u8], usize)>),
    Ordered(BTreeMap<Name<'a>, usize>),
}

impl Default for FirstSeen<'_> {
    fn default() -> Self {
        Self::Linear(Vec::new())
    }
}

impl<'a> FirstSeen<'a> {
    /// The position of `name`'s first occurrence if it was seen before;
    /// otherwise record it at `position`.
    fn first(&mut self, name: &'a [u8], position: usize) -> Option<usize> {
        match self {
            Self::Linear(names) => {
                if let Some((_, first)) = names.iter().find(|(seen, _)| same(seen, name)) {
                    return Some(*first);
                }
                if names.len() < LINEAR_NAMES {
                    names.push((name, position));
                    return None;
                }
                let mut ordered: BTreeMap<Name<'a>, usize> = names
                    .drain(..)
                    .map(|(seen, first)| (Name(seen), first))
                    .collect();
                ordered.insert(Name(name), position);
                *self = Self::Ordered(ordered);
                None
            },
            Self::Ordered(names) => match names.entry(Name(name)) {
                std::collections::btree_map::Entry::Occupied(entry) => Some(*entry.get()),
                std::collections::btree_map::Entry::Vacant(entry) => {
                    entry.insert(position);
                    None
                },
            },
        }
    }
}

/// A set of one tag's names, for a caller that checks names of its own
/// making (for example expanded names) for duplicates as it reads them.
///
/// Membership costs a linear scan while the set holds at most 32 names and
/// `O(log n)` comparisons after that, so checking every attribute of a tag
/// costs `O(n log n)` rather than the `O(n²)` of scanning a list.
#[derive(Clone, Debug)]
pub struct SeenNames<K> {
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
    /// An empty set.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Whether `key` has been inserted.
    #[must_use]
    pub fn contains(&self, key: &K) -> bool {
        if self.ordered.is_empty() {
            self.linear.contains(key)
        } else {
            self.ordered.contains(key)
        }
    }

    /// Insert `key`; `false` when it was already present.
    pub fn insert(&mut self, key: K) -> bool {
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

    /// The number of names inserted.
    #[must_use]
    pub fn len(&self) -> usize {
        self.linear.len() + self.ordered.len()
    }

    /// Whether no name has been inserted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }
}

/// A name ordered by its bytes, with every comparison counted in tests.
#[derive(Clone, Copy, Debug)]
struct Name<'a>(&'a [u8]);

impl PartialEq for Name<'_> {
    fn eq(&self, other: &Self) -> bool {
        same(self.0, other.0)
    }
}

impl Eq for Name<'_> {}

impl PartialOrd for Name<'_> {
    fn partial_cmp(&self, other: &Self) -> Option<core::cmp::Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for Name<'_> {
    fn cmp(&self, other: &Self) -> core::cmp::Ordering {
        #[cfg(test)]
        tests::count_comparison();
        self.0.cmp(other.0)
    }
}

fn same(left: &[u8], right: &[u8]) -> bool {
    #[cfg(test)]
    tests::count_comparison();
    left == right
}

#[cfg(test)]
mod tests {
    use core::cell::Cell;

    use quick_xml::events::BytesStart;
    use quick_xml::events::attributes::AttrError;

    use super::{SeenNames, count_up_to, first_wins};

    thread_local! {
        static COMPARISONS: Cell<usize> = const { Cell::new(0) };
    }

    pub(super) fn count_comparison() {
        COMPARISONS.with(|count| count.set(count.get() + 1));
    }

    fn counted<T>(run: impl FnOnce() -> T) -> (T, usize) {
        COMPARISONS.with(|count| count.set(0));
        let result = run();
        (result, COMPARISONS.with(Cell::get))
    }

    type Item = Result<(Vec<u8>, Vec<u8>), AttrError>;

    fn checked(tag: &BytesStart<'_>) -> Vec<Item> {
        tag.attributes()
            .map(|item| {
                item.map(|attribute| (attribute.key.as_ref().to_vec(), attribute.value.to_vec()))
            })
            .collect()
    }

    fn first(tag: &BytesStart<'_>) -> Vec<Item> {
        first_wins(tag)
            .map(|item| {
                item.map(|attribute| (attribute.key.as_ref().to_vec(), attribute.value.to_vec()))
            })
            .collect()
    }

    fn tag(content: &str) -> BytesStart<'static> {
        let name_len = content.find(' ').unwrap_or(content.len());
        BytesStart::from_content(content.to_owned(), name_len)
    }

    #[test]
    fn well_formed_tags_yield_what_the_checked_iterator_yields() {
        let mut cases = vec![
            "e".to_owned(),
            "e a=\"1\"".to_owned(),
            "w:rPr w:val=\"x\" w:eastAsia='y' xmlns:w=\"urn:w\"".to_owned(),
            "e a = \"1\"\tb='2'\n c=\"\" ".to_owned(),
            "e a=\"&amp;\" b=\"x y z\"".to_owned(),
        ];
        for count in [31, 32, 33, 64, 500] {
            let names = (0..count)
                .map(|index| format!("n{index}=\"{index}\""))
                .collect::<Vec<_>>();
            cases.push(format!("e {}", names.join(" ")));
        }
        for content in cases {
            let tag = tag(&content);
            assert_eq!(first(&tag), checked(&tag), "{content}");
        }
    }

    #[test]
    fn duplicates_report_the_positions_the_checked_iterator_reports() {
        for count in [2, 31, 32, 33, 100] {
            let names = (0..count)
                .map(|index| format!("n{index}=\"\""))
                .collect::<Vec<_>>();
            // Repeat the first, a middle and the last name, twice each.
            let repeats = [0, count / 2, count - 1, count - 1, count / 2, 0]
                .iter()
                .map(|index| format!("n{index}=\"r\""))
                .collect::<Vec<_>>();
            let content = format!("e {} {} tail=\"t\"", names.join(" "), repeats.join(" "));
            let tag = tag(&content);
            assert_eq!(first(&tag), checked(&tag), "{content}");
            let duplicates = first(&tag)
                .iter()
                .filter(|item| matches!(item, Err(AttrError::Duplicated(..))))
                .count();
            assert_eq!(duplicates, 6);
        }
    }

    #[test]
    fn a_duplicate_with_whitespace_in_its_value_is_skipped_whole() {
        let tag = tag(r#"e a="1" a="x b='evil'" c="3""#);
        // quick-xml resumes inside the duplicate's value and reports `b`.
        let lenient: Vec<_> = checked(&tag).into_iter().filter_map(Result::ok).collect();
        assert!(lenient.iter().any(|(key, _)| key == b"b"), "{lenient:?}");
        // The first occurrence wins and nothing inside the value is a name.
        assert_eq!(
            first(&tag),
            vec![
                Ok((b"a".to_vec(), b"1".to_vec())),
                Err(AttrError::Duplicated(8, 2)),
                Ok((b"c".to_vec(), b"3".to_vec())),
            ]
        );
    }

    #[test]
    fn syntax_errors_pass_through() {
        let tag = tag(r#"e a=x b="2" c"#);
        let items = first(&tag);
        assert!(
            matches!(items[0], Err(AttrError::UnquotedValue(_))),
            "{items:?}"
        );
        assert!(
            items
                .iter()
                .any(|item| matches!(item, Ok((key, _)) if key == b"b"))
        );
        assert!(
            matches!(items.last(), Some(Err(AttrError::ExpectedEq(_)))),
            "{items:?}"
        );
    }

    #[test]
    fn a_tag_of_repeated_names_costs_n_log_n_comparisons() {
        // 20,000 distinct names followed by 20,000 repeats of the last one:
        // quick-xml's checked iterator scans the distinct names for every
        // repeat, 4 * 10^8 comparisons; this costs a logarithmic number each.
        let distinct = 20_000;
        let mut content = String::from("e");
        for index in 0..distinct {
            content.push_str(&format!(" n{index:05}=\"\""));
        }
        let last = format!(" n{:05}=\"\"", distinct - 1);
        for _ in 0..distinct {
            content.push_str(&last);
        }
        let tag = tag(&content);
        let (items, comparisons) = counted(|| first(&tag));
        assert_eq!(items.iter().filter(|item| item.is_ok()).count(), distinct);
        assert_eq!(items.len(), 2 * distinct);
        // A B-tree compares a key with up to eleven keys on each of a few
        // levels; four comparisons per level of a binary tree bounds that
        // generously and stays two orders of magnitude below D * R.
        let levels = (usize::BITS - items.len().leading_zeros()) as usize;
        let bound = 4 * items.len() * levels;
        assert!(bound < distinct * distinct / 100);
        assert!(
            comparisons <= bound,
            "{comparisons} comparisons, bound {bound}"
        );
    }

    #[test]
    fn count_up_to_stops_at_the_bound_and_ignores_duplicates() {
        let tag = tag(r#"e a="" a="" a="" b="" c="""#);
        assert_eq!(count_up_to(&tag, 10), 5);
        assert_eq!(count_up_to(&tag, 2), 3);
        assert_eq!(count_up_to(&tag, 0), 1);
        assert_eq!(count_up_to(&tag, usize::MAX), 5);
    }

    #[test]
    fn seen_names_switch_to_an_ordered_index_without_losing_names() {
        let mut seen = SeenNames::new();
        for index in 0..100u32 {
            assert!(seen.insert(index), "{index}");
            assert!(!seen.insert(index), "{index}");
            assert!(seen.contains(&index));
        }
        assert_eq!(seen.len(), 100);
        for index in 0..100u32 {
            assert!(seen.contains(&index));
        }
        assert!(!seen.contains(&100));
        assert!(SeenNames::<u32>::new().is_empty());
    }
}
