//! Attribute iteration for readers that stop at a start tag's first attribute
//! error, with a worst case that stays bounded on hostile tags.
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
    fn checked_attributes(&self) -> CheckedAttributes<'_> {
        CheckedAttributes::new(self)
    }

    // The two calls quick-xml offers for an unchecked iterator.
    #[allow(clippy::disallowed_methods)]
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
    phase: Phase,
    /// From the 33rd name on: every name yielded so far, with its position.
    names: BTreeMap<Name<'a>, usize>,
    /// Where quick-xml starts to look for the next attribute: just after the
    /// last one yielded (once this iterator checks the names).
    next_start: usize,
}

#[derive(Clone, Copy, Debug)]
enum Phase {
    /// quick-xml checks the names; this many have been yielded.
    QuickXml(usize),
    /// quick-xml's check is off; this iterator checks the names.
    Own,
    /// An error or the end has been yielded.
    Done,
}

impl<'a> CheckedAttributes<'a> {
    // quick-xml's checked iterator, used for the first 32 names only.
    #[allow(clippy::disallowed_methods)]
    fn new(tag: &'a BytesStart<'a>) -> Self {
        Self {
            tag,
            attributes: tag.attributes(),
            phase: Phase::QuickXml(0),
            names: BTreeMap::new(),
            next_start: 0,
        }
    }

    // Turning quick-xml's check off before its hashed check can run.
    #[allow(clippy::disallowed_methods)]
    fn take_over(&mut self) {
        self.attributes.with_checks(false);
    }

    /// Record the 32 names quick-xml has checked, with their positions, and
    /// where the attribute after them starts. Reading them again without the
    /// check yields the same attributes: the check never changes what is read.
    fn record_checked_names(&mut self) {
        let tag = self.tag;
        for attribute in tag
            .unchecked_attributes()
            .take(QUICK_XML_LINEAR_NAMES)
            .flatten()
        {
            let name = attribute.key.into_inner();
            self.names.insert(Name(name), offset_in(tag, name));
            self.next_start = end_of(tag, &attribute);
        }
        debug_assert_eq!(self.names.len(), QUICK_XML_LINEAR_NAMES);
    }

    fn check(
        &mut self,
        item: Result<Attribute<'a>, AttrError>,
    ) -> Result<Attribute<'a>, AttrError> {
        let attribute = match item {
            Ok(attribute) => attribute,
            Err(error) => return Err(self.duplicate_before_value(error)),
        };
        let name = attribute.key.into_inner();
        let position = offset_in(self.tag, name);
        match self.names.entry(Name(name)) {
            Entry::Occupied(first) => Err(AttrError::Duplicated(position, *first.get())),
            Entry::Vacant(entry) => {
                entry.insert(position);
                self.next_start = end_of(self.tag, &attribute);
                Ok(attribute)
            },
        }
    }

    /// quick-xml checks a name once it has read the name and the `=` after
    /// it, before the value: a repeated name followed by a malformed value is
    /// reported as a duplicate, not as the value's error.
    fn duplicate_before_value(&self, error: AttrError) -> AttrError {
        if !matches!(
            error,
            AttrError::UnquotedValue(_)
                | AttrError::ExpectedValue(_)
                | AttrError::ExpectedQuote(..)
        ) {
            return error;
        }
        let Some((position, name)) = name_at(self.tag, self.next_start) else {
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

    fn next(&mut self) -> Option<Self::Item> {
        let item = match self.phase {
            Phase::Done => return None,
            Phase::QuickXml(yielded) if yielded < QUICK_XML_LINEAR_NAMES => {
                let item = self.attributes.next();
                self.phase = match item {
                    Some(Ok(_)) => Phase::QuickXml(yielded + 1),
                    _ => Phase::Done,
                };
                return item;
            },
            Phase::QuickXml(_) => {
                self.take_over();
                let item = self.attributes.next();
                if item.is_some() {
                    self.record_checked_names();
                    self.phase = Phase::Own;
                }
                item
            },
            Phase::Own => self.attributes.next(),
        };
        let Some(item) = item else {
            self.phase = Phase::Done;
            return None;
        };
        let checked = self.check(item);
        if checked.is_err() {
            self.phase = Phase::Done;
        }
        Some(checked)
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
