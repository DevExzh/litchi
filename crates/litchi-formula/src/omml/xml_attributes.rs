//! Attribute iteration whose cost stays bounded on hostile start tags.
//!
//! This mirrors `first_wins` in `litchi_ooxml_common::xml::attributes`, with
//! the same semantics; this crate depends on no workspace crate, so it keeps
//! its own copy. Keep the two in step.
//!
//! quick-xml 0.41 checks a start tag's attribute names for duplicates by
//! default, and every duplicate costs a scan of the tag's earlier names. Its
//! iterator keeps going after it reports one, so a reader that skips errors,
//! as the OMML parser does, spends time quadratic in the tag's attributes on a
//! tag of `D` distinct names followed by `R` repeats of the last one. Its
//! recovery after a duplicate also resumes inside the duplicate's value, so
//! `<e a="1" a="x b='y'"/>` yields a `b` attribute the tag does not have.
//!
//! [`first_wins`] yields every qualified name once, at its first
//! occurrence, at a cost of `O(n log n)` comparisons for `n` attributes, and
//! parses every value as written.

use std::collections::BTreeMap;
use std::collections::btree_map::Entry;

use quick_xml::events::BytesStart;
use quick_xml::events::attributes::{AttrError, Attribute, Attributes};

/// Names a tag may carry before [`first_wins`] leaves its linear scan for an
/// ordered index; quick-xml's own threshold.
const LINEAR_NAMES: usize = 32;

/// Iterate `tag`'s attributes with each qualified name's first occurrence
/// yielded as `Ok`.
///
/// Every later occurrence of a name is yielded as
/// `Err(AttrError::Duplicated(position, first_position))`, positions relative
/// to the start of the tag as quick-xml reports them, and its value is parsed
/// and skipped as written. Other lexical errors pass through unchanged. On a
/// tag without duplicate names this yields exactly what `tag.attributes()`
/// yields.
#[must_use]
pub(crate) fn first_wins<'a>(tag: &'a BytesStart<'_>) -> FirstWins<'a> {
    let mut attributes = tag.attributes();
    attributes.with_checks(false);
    FirstWins {
        attributes,
        base: tag,
        seen: FirstSeen::Linear(Vec::new()),
    }
}

/// The iterator [`first_wins`] returns.
#[derive(Debug)]
pub(crate) struct FirstWins<'a> {
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
        // The name is a subslice of the tag, so this is its offset in it.
        let position = (name.as_ptr() as usize).wrapping_sub(self.base.as_ptr() as usize);
        match self.seen.first(name, position) {
            None => Some(Ok(attribute)),
            Some(first) => Some(Err(AttrError::Duplicated(position, first))),
        }
    }
}

/// The names one tag has shown so far, with the position of each first
/// occurrence.
#[derive(Debug)]
enum FirstSeen<'a> {
    Linear(Vec<(&'a [u8], usize)>),
    Ordered(BTreeMap<&'a [u8], usize>),
}

impl<'a> FirstSeen<'a> {
    /// The position of `name`'s first occurrence if it was seen before;
    /// otherwise record it at `position`.
    fn first(&mut self, name: &'a [u8], position: usize) -> Option<usize> {
        match self {
            Self::Linear(names) => {
                if let Some((_, first)) = names.iter().find(|(seen, _)| *seen == name) {
                    return Some(*first);
                }
                if names.len() < LINEAR_NAMES {
                    names.push((name, position));
                    return None;
                }
                let mut ordered: BTreeMap<&'a [u8], usize> = names.drain(..).collect();
                ordered.insert(name, position);
                *self = Self::Ordered(ordered);
                None
            },
            Self::Ordered(names) => match names.entry(name) {
                Entry::Occupied(entry) => Some(*entry.get()),
                Entry::Vacant(entry) => {
                    entry.insert(position);
                    None
                },
            },
        }
    }
}

#[cfg(test)]
mod tests {
    use quick_xml::events::BytesStart;
    use quick_xml::events::attributes::AttrError;

    use super::first_wins;

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
    fn tags_without_duplicates_yield_what_the_checked_iterator_yields() {
        let mut cases = vec![
            "m:e".to_owned(),
            "m:chr m:val=\"x\"".to_owned(),
            "m:d a = \"1\"\tb='2'\n c=\"\" ".to_owned(),
            "m:t a=\"&amp;\" b=\"x y z\"".to_owned(),
        ];
        for count in [31, 32, 33, 100] {
            let names = (0..count)
                .map(|index| format!("n{index}=\"{index}\""))
                .collect::<Vec<_>>();
            cases.push(format!("m:e {}", names.join(" ")));
        }
        for content in cases {
            let tag = tag(&content);
            assert_eq!(first(&tag), checked(&tag), "{content}");
        }
    }

    #[test]
    fn duplicates_are_reported_where_the_checked_iterator_reports_them() {
        for count in [2, 31, 32, 33, 100] {
            let names = (0..count)
                .map(|index| format!("n{index}=\"\""))
                .collect::<Vec<_>>();
            let repeats = [0, count / 2, count - 1, count - 1]
                .iter()
                .map(|index| format!("n{index}=\"r\""))
                .collect::<Vec<_>>();
            let content = format!("m:e {} {} tail=\"t\"", names.join(" "), repeats.join(" "));
            let tag = tag(&content);
            assert_eq!(first(&tag), checked(&tag), "{content}");
        }
    }

    #[test]
    fn a_duplicate_with_whitespace_in_its_value_is_skipped_whole() {
        let tag = tag(r#"m:e a="1" a="x b='evil'" c="3""#);
        assert_eq!(
            first(&tag),
            vec![
                Ok((b"a".to_vec(), b"1".to_vec())),
                Err(AttrError::Duplicated(10, 4)),
                Ok((b"c".to_vec(), b"3".to_vec())),
            ]
        );
    }

    #[test]
    fn a_tag_of_repeated_names_keeps_every_first_occurrence() {
        let distinct = 20_000;
        let mut content = String::from("m:e");
        for index in 0..distinct {
            content.push_str(&format!(" n{index:05}=\"{index}\""));
        }
        let last = format!(" n{:05}=\"late\"", distinct - 1);
        for _ in 0..distinct {
            content.push_str(&last);
        }
        let tag = tag(&content);
        let items = first(&tag);
        assert_eq!(items.len(), 2 * distinct);
        let kept = items
            .iter()
            .filter_map(|item| item.as_ref().ok())
            .collect::<Vec<_>>();
        assert_eq!(kept.len(), distinct);
        assert_eq!(
            kept.last().map(|(_, value)| value.as_slice()),
            Some(format!("{}", distinct - 1).as_bytes())
        );
    }
}
