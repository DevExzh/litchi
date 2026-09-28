//! Attribute iteration for readers that stop at a start tag's first attribute
//! error, with a worst case that stays bounded on hostile tags.
//!
//! This source-only 0800 candidate reads the first two attributes with
//! quick-xml's duplicate check disabled. After the second item, a third-key
//! preflight compares only the raw key prefixes already addressed by the
//! iterator; it never reparses either earlier value. A successful third item
//! seeds an ordered map from the raw key slices and continues with the same
//! bounded `O(n log n)` check. No transition reconstructs a quick-xml iterator
//! or replays an already yielded item.
//!
//! quick-xml 0.41 checks a start tag's attribute names for duplicates by
//! default. It compares each name with the names before it while it has seen
//! at most 32; from the 33rd name on it pre-filters with a hash set and, on
//! every pre-filter hit, scans all the names before it. The pre-filter's
//! hasher is not keyed (`DefaultHasher::new()`), so a name hashes to the same
//! value in every process built with the same standard library: whoever writes
//! the tag chooses how often the pre-filter hits, and each hit costs a scan of
//! the tag's earlier names. The check's worst case therefore grows with the
//! square of the tag, and nothing but the tag's size bounds it where no
//! per-element limit applies.
//!
//! [`BytesStartExt::checked_attributes`] yields what quick-xml's checked
//! iterator (`BytesStart::attributes`) yields up to and including its first
//! error, and nothing after it. The candidate checks the first two names with
//! raw key prefixes and starts its ordered map at the third successful item:
//! `O(log n)` name comparisons for each later one; no name is hashed. The
//! first error is the one quick-xml reports, at the same attribute, with the
//! same positions, so a reader that stops at an attribute's error (`?` on each
//! item) behaves as it did with quick-xml's iterator.
//!
//! A reader that skips attribute errors and reads on uses
//! `litchi_ooxml_common::xml::attributes::first_wins` instead (record 0764),
//! and a reader that needs no duplicate check uses
//! [`BytesStartExt::unchecked_attributes`]. The workspace's `clippy.toml`
//! disallows quick-xml's checked iteration in the crates that read untrusted
//! OOXML and OLE2 XML, so a new call site has to choose one of the three.

use std::borrow::Cow;
use std::collections::BTreeMap;
use std::collections::btree_map::Entry;

use quick_xml::events::BytesStart;
use quick_xml::events::attributes::{AttrError, Attribute, Attributes};

/// Kept for the shared bounded-cost test's historical comparison point.
#[cfg(test)]
const QUICK_XML_LINEAR_NAMES: usize = 32;

/// Bounded attribute iteration for quick-xml start tags.
pub trait BytesStartExt {
    /// Iterate the tag's attributes as `BytesStart::attributes` does, up to
    /// and including the first error; nothing is yielded after it.
    ///
    /// Items and errors are the ones quick-xml's checked iterator yields,
    /// including `AttrError::Duplicated(position, first_position)` at the
    /// first repeated name, and a repeated name whose value is malformed is
    /// reported as a duplicate, as quick-xml checks a name before reading its
    /// value. The first two names use raw key-prefix checks; after a third
    /// successful item the candidate uses `O(log n)` ordered-map comparisons
    /// per attribute and no name hashing.
    fn checked_attributes(&self) -> CheckedAttributes<'_>;

    /// Iterate the tag's attributes without any duplicate check: quick-xml's
    /// iterator with `with_checks(false)`. Every occurrence of a repeated name
    /// is yielded as `Ok`, and lexical errors are yielded as quick-xml reports
    /// them, with its recovery after each.
    fn unchecked_attributes(&self) -> Attributes<'_>;
}

impl BytesStartExt for BytesStart<'_> {
    #[inline]
    fn checked_attributes(&self) -> CheckedAttributes<'_> {
        CheckedAttributes::new(self)
    }

    // The two calls quick-xml offers for an unchecked iterator.
    #[allow(clippy::disallowed_methods)]
    #[inline]
    fn unchecked_attributes(&self) -> Attributes<'_> {
        let mut attributes = self.attributes();
        attributes.with_checks(false);
        attributes
    }
}

/// The iterator [`BytesStartExt::checked_attributes`] returns.
#[derive(Clone, Debug)]
pub struct CheckedAttributes<'a> {
    tag: &'a BytesStart<'a>,
    attributes: Attributes<'a>,
    phase: Phase<'a>,
}

#[derive(Clone, Debug)]
enum Phase<'a> {
    /// No attribute has been requested yet.
    First,
    /// The first attribute was yielded unchecked; this is its next raw offset.
    Second(usize),
    /// The first two attributes were yielded unchecked and a third may follow.
    /// `first_end` reconstructs the second key; `next_start` addresses the
    /// third key. Neither offset points into an earlier value for reparsing.
    ShortTwo {
        first_end: usize,
        next_start: usize,
    },
    /// The first three names have been seeded from raw key slices and this
    /// iterator checks every subsequent name without replay.
    Own(Box<OwnCheck<'a>>),
    /// An error or the end has been yielded.
    Done,
}

/// This iterator's bounded duplicate check after the short prefix.
#[derive(Clone, Debug)]
struct OwnCheck<'a> {
    /// Every name yielded so far, with its position. The map is allocated only
    /// after a third successful item may be followed by another item.
    names: BTreeMap<Name<'a>, usize>,
    /// Where quick-xml starts to look for the next attribute: just after the
    /// last one yielded.
    next_start: usize,
}

impl<'a> CheckedAttributes<'a> {
    // The short prefix uses quick-xml's parser without duplicate allocation.
    #[allow(clippy::disallowed_methods)]
    #[inline]
    fn new(tag: &'a BytesStart<'a>) -> Self {
        let mut attributes = tag.attributes();
        attributes.with_checks(false);
        Self {
            tag,
            attributes,
            phase: Phase::First,
        }
    }

    /// Yield the first item. No earlier name exists, so its lexical result is
    /// already the checked result.
    #[inline]
    fn next_first(&mut self) -> Option<Result<Attribute<'a>, AttrError>> {
        let item = self.attributes.next();
        match &item {
            Some(Ok(attribute)) if !trailing_whitespace(self.tag, end_of(self.tag, attribute)) => {
                self.phase = Phase::Second(end_of(self.tag, attribute));
            },
            Some(Ok(_)) | Some(Err(_)) | None => {
                self.phase = Phase::Done;
            },
        }
        item
    }

    /// Check the second key before asking quick-xml to scan its value. A
    /// repeated key without an equals sign remains quick-xml's `ExpectedEq`.
    #[inline]
    fn next_second(&mut self, next_start: usize) -> Option<Result<Attribute<'a>, AttrError>> {
        if let (Some((first_position, first_name, _)), Some((position, name, has_equals))) = (
            key_at(self.tag.as_ref(), self.tag.name().as_ref().len()),
            key_at(self.tag.as_ref(), next_start),
        ) && has_equals
            && name == first_name
        {
            self.phase = Phase::Done;
            return Some(Err(AttrError::Duplicated(position, first_position)));
        }

        let item = self.attributes.next();
        match &item {
            Some(Ok(attribute)) if !trailing_whitespace(self.tag, end_of(self.tag, attribute)) => {
                self.phase = Phase::ShortTwo {
                    first_end: next_start,
                    next_start: end_of(self.tag, attribute),
                };
            },
            Some(Ok(_)) | Some(Err(_)) | None => {
                self.phase = Phase::Done;
            },
        }
        item
    }

    /// Check the third key before parsing its value, then seed the ordered map
    /// from the three raw key slices. The first and second values are never
    /// parsed again; an error or end before a third successful item allocates
    /// no map.
    #[cold]
    #[inline(never)]
    fn next_short_two(
        &mut self,
        first_end: usize,
        next_start: usize,
    ) -> Option<Result<Attribute<'a>, AttrError>> {
        let first = key_at(self.tag.as_ref(), self.tag.name().as_ref().len());
        let second = key_at(self.tag.as_ref(), first_end);
        let third = key_at(self.tag.as_ref(), next_start);
        let Some((position, name, has_equals)) = third else {
            self.phase = Phase::Done;
            return None;
        };
        if has_equals {
            if let Some((first_position, _, _)) =
                first.filter(|(_, first_name, _)| *first_name == name)
            {
                self.phase = Phase::Done;
                return Some(Err(AttrError::Duplicated(position, first_position)));
            }
            if let Some((second_position, _, _)) =
                second.filter(|(_, second_name, _)| *second_name == name)
            {
                self.phase = Phase::Done;
                return Some(Err(AttrError::Duplicated(position, second_position)));
            }
        }

        let item = self.attributes.next();
        match item {
            Some(Ok(attribute)) => {
                let third_end = end_of(self.tag, &attribute);
                if trailing_whitespace(self.tag, third_end) {
                    self.phase = Phase::Done;
                } else {
                    let mut check = OwnCheck::from_prefix(
                        self.tag,
                        first_end,
                        next_start,
                        &attribute,
                    );
                    debug_assert_eq!(check.names.len(), 3);
                    check.next_start = third_end;
                    self.phase = Phase::Own(Box::new(check));
                }
                Some(Ok(attribute))
            },
            Some(Err(error)) => {
                self.phase = Phase::Done;
                Some(Err(error))
            },
            None => {
                self.phase = Phase::Done;
                None
            },
        }
    }

    /// Continue the ordered-map backend. Its raw-key preflight preserves
    /// quick-xml's duplicate-before-value boundary even though the underlying
    /// parser has duplicate checking disabled.
    #[cold]
    #[inline(never)]
    fn next_own(&mut self) -> Option<Result<Attribute<'a>, AttrError>> {
        let item = if let Phase::Own(check) = &mut self.phase {
            check.next(self.tag, &mut self.attributes)
        } else {
            unreachable!("ordered backend is called only from the Own phase")
        };
        if item.as_ref().map_or(true, Result::is_err) {
            self.phase = Phase::Done;
        }
        item
    }
}

impl<'a> OwnCheck<'a> {
    /// Seed the backend from raw key prefixes and the third borrowed key. No
    /// earlier attribute value is decoded or reparsed.
    fn from_prefix(
        tag: &'a BytesStart<'a>,
        first_end: usize,
        second_end: usize,
        third: &Attribute<'a>,
    ) -> Self {
        let mut names = BTreeMap::new();
        let (first_position, first_name, _) =
            key_at(tag.as_ref(), tag.name().as_ref().len()).expect("first prefix");
        let (second_position, second_name, _) =
            key_at(tag.as_ref(), first_end).expect("second prefix");
        names.insert(Name(first_name), first_position);
        names.insert(Name(second_name), second_position);
        let third_name = third.key.into_inner();
        names.insert(Name(third_name), offset_in(tag, third_name));
        Self {
            names,
            next_start: second_end,
        }
    }

    fn next(
        &mut self,
        tag: &'a BytesStart<'a>,
        attributes: &mut Attributes<'a>,
    ) -> Option<Result<Attribute<'a>, AttrError>> {
        if let Some((position, name, has_equals)) = key_at(tag.as_ref(), self.next_start)
            && has_equals
            && let Some(first) = self.names.get(&Name(name))
        {
            return Some(Err(AttrError::Duplicated(position, *first)));
        }
        let item = attributes.next()?;
        Some(match item {
            Ok(attribute) => self.check_attribute(tag, attribute),
            Err(error) => Err(self.duplicate_before_value(tag, error)),
        })
    }

    fn check_attribute(
        &mut self,
        tag: &'a BytesStart<'a>,
        attribute: Attribute<'a>,
    ) -> Result<Attribute<'a>, AttrError> {
        let name = attribute.key.into_inner();
        let position = offset_in(tag, name);
        match self.names.entry(Name(name)) {
            Entry::Occupied(first) => Err(AttrError::Duplicated(position, *first.get())),
            Entry::Vacant(entry) => {
                entry.insert(position);
                self.next_start = end_of(tag, &attribute);
                Ok(attribute)
            },
        }
    }

    /// quick-xml checks a name once it has read the name and the `=` after
    /// it, before the value: a repeated name followed by a malformed value is
    /// reported as a duplicate, not as the value's error.
    fn duplicate_before_value(&self, tag: &[u8], error: AttrError) -> AttrError {
        if !matches!(
            error,
            AttrError::UnquotedValue(_)
                | AttrError::ExpectedValue(_)
                | AttrError::ExpectedQuote(..)
        ) {
            return error;
        }
        let Some((position, name)) = name_at(tag, self.next_start) else {
            return error;
        };
        match self.names.get(&Name(name)) {
            Some(first) => AttrError::Duplicated(position, *first),
            None => error,
        }
    }
}

impl<'a> Iterator for CheckedAttributes<'a> {
    type Item = Result<Attribute<'a>, AttrError>;

    /// The first two names use the unchecked parser and later names use the
    /// lazily allocated ordered-map backend.
    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        if matches!(self.phase, Phase::First) {
            return self.next_first();
        }
        if let Phase::Second(next_start) = &self.phase {
            return self.next_second(*next_start);
        }
        if let Phase::ShortTwo {
            first_end,
            next_start,
        } = &self.phase
        {
            return self.next_short_two(*first_end, *next_start);
        }
        if matches!(self.phase, Phase::Own(_)) {
            return self.next_own();
        }
        None
    }
}

impl core::iter::FusedIterator for CheckedAttributes<'_> {}

/// The byte offset of `part` within `base`, of which it is a subslice.
fn offset_in(base: &[u8], part: &[u8]) -> usize {
    (part.as_ptr() as usize).wrapping_sub(base.as_ptr() as usize)
}

/// Where quick-xml looks for the attribute after `attribute`: just past the
/// closing quote of its value. quick-xml yields values borrowed from the tag.
fn end_of(base: &[u8], attribute: &Attribute<'_>) -> usize {
    debug_assert!(matches!(attribute.value, Cow::Borrowed(_)));
    let value = attribute.value.as_ref();
    offset_in(base, value)
        .saturating_add(value.len())
        .saturating_add(1)
}

/// Whether the bytes after a successfully decoded attribute contain no next
/// attribute. A non-whitespace byte deliberately enters the third-item path,
/// including a malformed tail or a missing separator.
fn trailing_whitespace(base: &[u8], from: usize) -> bool {
    base.get(from..)
        .is_some_and(|rest| rest.iter().all(|byte| is_whitespace(*byte)))
}

/// The name of the attribute that starts at or after `from`, read the way
/// quick-xml reads it: after any whitespace, up to `=` or whitespace; with
/// its position.
fn name_at(tag: &[u8], from: usize) -> Option<(usize, &[u8])> {
    let rest = tag.get(from..)?;
    let start = from + rest.iter().position(|byte| !is_whitespace(*byte))?;
    let name = &tag[start..];
    let length = name
        .iter()
        .position(|byte| *byte == b'=' || is_whitespace(*byte))
        .unwrap_or(name.len());
    Some((start, &name[..length]))
}

/// The key prefix quick-xml has parsed at `from`, together with whether the
/// prefix was followed by the `=` that makes duplicate checking precede value
/// parsing. This consumes the first key byte exactly as quick-xml does, so a
/// key beginning with `=` is retained rather than treated as an empty name.
fn key_at(tag: &[u8], from: usize) -> Option<(usize, &[u8], bool)> {
    let rest = tag.get(from..)?;
    let start = from + rest.iter().position(|byte| !is_whitespace(*byte))?;
    let after_first = tag.get(start + 1..)?;
    let length = after_first
        .iter()
        .position(|byte| *byte == b'=' || is_whitespace(*byte))
        .map_or(tag.len() - (start + 1), |length| length);
    let end = start + 1 + length;
    let has_equals = match tag.get(end) {
        Some(b'=') => true,
        Some(_) => tag[end..]
            .iter()
            .find(|byte| !is_whitespace(**byte))
            .is_some_and(|byte| *byte == b'='),
        None => false,
    };
    Some((start, &tag[start..end], has_equals))
}

/// quick-xml's whitespace: space, tab, carriage return and line feed.
fn is_whitespace(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\r' | b'\n')
}

/// A name ordered by its bytes, with every comparison counted in tests.
#[derive(Clone, Copy, Debug)]
struct Name<'a>(&'a [u8]);

impl PartialEq for Name<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.cmp(other).is_eq()
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

#[cfg(test)]
mod tests;
