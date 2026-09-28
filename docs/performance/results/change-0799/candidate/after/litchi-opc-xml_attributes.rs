//! Attribute iteration for readers that stop at a start tag's first attribute
//! error, with a worst case that stays bounded on hostile tags.
//!
//! The first two valid attributes are read with duplicate checking disabled.
//! A tag with at most two attributes can finish after checking that the
//! remaining bytes are whitespace without allocating quick-xml's name list.
//! Before reading the second value, the candidate mirrors quick-xml's key and
//! equals-sign prefix and refuses a repeated name at the same boundary as
//! quick-xml. If a third attribute may follow, the next call rebuilds
//! quick-xml's checked iterator, consumes the first two attributes as replay
//! seeds, and resumes the ordinary bounded handoff below.
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
//! error, and nothing after it. quick-xml checks the first 32 names with its
//! linear scan, exactly as before; from the 33rd on its check is off and this
//! iterator checks each name itself, in an ordered map: `O(log n)` name
//! comparisons per attribute and no hashing, so a tag of `n` attributes costs
//! `O(n log n)` comparisons whatever its names are. The first error is the one
//! quick-xml reports, at the same attribute, with the same positions, so a
//! reader that stops at an attribute's error (`?` on each item) behaves as it
//! did with quick-xml's iterator.
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

/// Names quick-xml checks with a linear scan of the names before them. From
/// the next name on it would hash names with its unkeyed hasher, so this
/// module turns its check off there and checks the names itself.
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
    /// value. The cost is quick-xml's linear check for the first 32 names and
    /// `O(log n)` name comparisons for each later one; no name is hashed.
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
    /// The first attribute has not been requested yet.
    First,
    /// The first attribute was yielded unchecked; this is the next raw offset.
    Second(usize),
    /// The first two attributes were yielded unchecked and another item may
    /// follow.
    ReplayFirstTwo,
    /// quick-xml checks the names; this many have been yielded.
    QuickXml(usize),
    /// quick-xml's check is off; this iterator checks the names.
    Own(Box<OwnCheck<'a>>),
    /// An error or the end has been yielded.
    Done,
}

/// This iterator's check of the names, from the 33rd on.
#[derive(Clone, Debug)]
struct OwnCheck<'a> {
    /// Every name yielded so far, with its position.
    names: BTreeMap<Name<'a>, usize>,
    /// Where quick-xml starts to look for the next attribute: just after the
    /// last one yielded.
    next_start: usize,
}

impl<'a> CheckedAttributes<'a> {
    // The first two items are initially read without duplicate checking. The
    // second item is checked for a duplicate name before its value is parsed.
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

    /// Yield the first item and avoid constructing quick-xml's duplicate-name
    /// list when its value is followed only by tag whitespace. A lexical error
    /// in the first item is unchanged because no earlier name exists.
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

    /// Check the second key before asking quick-xml to parse its value. The
    /// checked iterator compares a key only after it has found its `=`; a
    /// repeated key without an equals sign must therefore remain
    /// `ExpectedEq`. A nonduplicate second item is read by the unchecked
    /// iterator, whose lexical result and recovery are unchanged.
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
                self.phase = Phase::ReplayFirstTwo;
            },
            Some(Ok(_)) | Some(Err(_)) | None => {
                self.phase = Phase::Done;
            },
        }
        item
    }

    /// Replace the unchecked iterator with a fresh checked one and consume its
    /// first two items. Both were already yielded; replaying them seeds
    /// quick-xml's duplicate state without exposing either item twice.
    #[cold]
    #[inline(never)]
    #[allow(clippy::disallowed_methods)]
    fn replay_first_two(&mut self) {
        let mut attributes = self.tag.attributes();
        let first = attributes.next();
        let second = attributes.next();
        debug_assert!(matches!(first, Some(Ok(_))));
        debug_assert!(matches!(second, Some(Ok(_))));
        self.attributes = attributes;
        self.phase = Phase::QuickXml(2);
    }

    /// Everything after quick-xml's 32 names: the switch to this iterator's
    /// own check, its check of each later name, and the end.
    #[cold]
    #[inline(never)]
    fn next_after_quick_xml(&mut self) -> Option<Result<Attribute<'a>, AttrError>> {
        match self.phase {
            Phase::Done => return None,
            // quick-xml would hash the next name: turn its check off first.
            Phase::QuickXml(_) => self.take_over(),
            Phase::ReplayFirstTwo => {
                self.replay_first_two();
                return self.next();
            },
            Phase::Own(_) => {},
            Phase::First | Phase::Second(_) => {
                unreachable!("short prefix is handled by next")
            },
        }
        let Some(item) = self.attributes.next() else {
            self.phase = Phase::Done;
            return None;
        };
        let checked = if let Phase::Own(check) = &mut self.phase {
            check.check(self.tag, item)
        } else {
            let mut check = OwnCheck::after_quick_xml(self.tag);
            let checked = check.check(self.tag, item);
            self.phase = Phase::Own(Box::new(check));
            checked
        };
        if checked.is_err() {
            self.phase = Phase::Done;
        }
        Some(checked)
    }

    // Turning quick-xml's check off before its hashed check can run.
    #[allow(clippy::disallowed_methods)]
    fn take_over(&mut self) {
        self.attributes.with_checks(false);
    }
}

impl<'a> OwnCheck<'a> {
    /// The 32 names quick-xml has checked, with their positions, and where
    /// the attribute after them starts. Reading them again without the check
    /// yields the same attributes: the check never changes what is read.
    fn after_quick_xml(tag: &'a BytesStart<'a>) -> Self {
        let mut check = Self {
            names: BTreeMap::new(),
            next_start: 0,
        };
        for attribute in tag
            .unchecked_attributes()
            .take(QUICK_XML_LINEAR_NAMES)
            .flatten()
        {
            let name = attribute.key.into_inner();
            check.names.insert(Name(name), offset_in(tag, name));
            check.next_start = end_of(tag, &attribute);
        }
        debug_assert_eq!(check.names.len(), QUICK_XML_LINEAR_NAMES);
        check
    }

    fn check(
        &mut self,
        tag: &'a BytesStart<'a>,
        item: Result<Attribute<'a>, AttrError>,
    ) -> Result<Attribute<'a>, AttrError> {
        let attribute = match item {
            Ok(attribute) => attribute,
            Err(error) => return Err(self.duplicate_before_value(tag, error)),
        };
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

    /// The first 32 names are quick-xml's checked iterator plus a counter.
    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        if matches!(self.phase, Phase::First) {
            return self.next_first();
        }
        if let Phase::Second(next_start) = self.phase {
            return self.next_second(next_start);
        }
        if let Phase::QuickXml(yielded) = &mut self.phase
            && *yielded < QUICK_XML_LINEAR_NAMES
        {
            let item = self.attributes.next();
            if matches!(item, Some(Ok(_))) {
                *yielded += 1;
            } else {
                self.phase = Phase::Done;
            }
            return item;
        }
        self.next_after_quick_xml()
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
/// attribute. A non-whitespace byte deliberately takes the replay path,
/// including a malformed second attribute or a missing separator.
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
/// parsing. This mirrors `quick_xml::events::attributes::IterState::next`.
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
