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
//! error, and nothing after it. This candidate performs the same lexical
//! parsing with quick-xml's duplicate check disabled. It keeps the first four
//! names inline, spills to a vector through the existing 32-name boundary,
//! and then moves to this module's ordered map: `O(log n)` name comparisons
//! per attribute and no hashing on the untrusted-name path. The first error
//! is the one quick-xml reports, at the same attribute, with the same
//! positions, so a reader that stops at an attribute's error (`?` on each
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

/// Number of names retained in the allocation-free prefix.
const INLINE_NAMES: usize = 4;

/// Number of names retained by the vector-like linear phase before the
/// ordered fallback takes over. This is also quick-xml's old linear boundary.
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
    /// quick-xml parses the bytes; this iterator owns duplicate checking.
    Own(OwnCheck<'a>),
    /// An error or the end has been yielded.
    Done,
}

/// This iterator's duplicate check over the whole tag.
#[derive(Clone, Debug)]
struct OwnCheck<'a> {
    /// Every name yielded so far, with an inline and bounded spill path.
    names: SeenNames<'a>,
    /// Where quick-xml starts to look for the next attribute: just after the
    /// last one yielded.
    next_start: usize,
}

impl<'a> CheckedAttributes<'a> {
    // quick-xml supplies lexical parsing; this iterator performs the check.
    #[inline]
    fn new(tag: &'a BytesStart<'a>) -> Self {
        let attributes = tag.unchecked_attributes();
        Self {
            tag,
            attributes,
            phase: Phase::Own(OwnCheck::new(tag)),
        }
    }
}

impl<'a> OwnCheck<'a> {
    fn new(tag: &'a BytesStart<'a>) -> Self {
        Self {
            names: SeenNames::new(),
            // Start after the element name so name_at cannot mistake that
            // name for an attribute when the first attribute is malformed.
            next_start: tag.name().as_ref().len(),
        }
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
        let name = Name(attribute.key.into_inner());
        let position = offset_in(tag, name.0);
        match self.names.check(tag, name, position) {
            Some(first) => Err(AttrError::Duplicated(position, first)),
            None => {
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
        match self.names.position(tag, name) {
            Some(first) => AttrError::Duplicated(position, first),
            None => error,
        }
    }
}

impl<'a> Iterator for CheckedAttributes<'a> {
    type Item = Result<Attribute<'a>, AttrError>;

    /// Lexical parsing comes from quick-xml; duplicate checking is local.
    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        let Phase::Own(check) = &mut self.phase else {
            return None;
        };
        let Some(item) = self.attributes.next() else {
            self.phase = Phase::Done;
            return None;
        };
        let checked = check.check(self.tag, item);
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

/// Duplicate names before the ordered fallback. The fixed array avoids the
/// first heap allocation for the common small-tag case; the vector retains
/// the old linear spill shape until the bounded ordered phase.
#[derive(Clone, Debug)]
enum SeenNames<'a> {
    Inline {
        names: [Name<'a>; INLINE_NAMES],
        len: usize,
    },
    Linear(Vec<Name<'a>>),
    Ordered(BTreeMap<Name<'a>, usize>),
}

impl<'a> SeenNames<'a> {
    fn new() -> Self {
        Self::Inline {
            names: [Name(&[]); INLINE_NAMES],
            len: 0,
        }
    }

    /// Returns the first position for a duplicate, or records a new name.
    fn check(&mut self, tag: &[u8], name: Name<'a>, position: usize) -> Option<usize> {
        match self {
            Self::Inline { names, len } => {
                if let Some(previous) = names[..*len].iter().find(|previous| previous.0 == name.0) {
                    return Some(offset_in(tag, previous.0));
                }
                if *len < INLINE_NAMES {
                    names[*len] = name;
                    *len += 1;
                    return None;
                }
                let mut linear = Vec::with_capacity(INLINE_NAMES * 2);
                linear.extend_from_slice(&names[..*len]);
                linear.push(name);
                *self = Self::Linear(linear);
                None
            },
            Self::Linear(names) => {
                if let Some(previous) = names.iter().find(|previous| previous.0 == name.0) {
                    return Some(offset_in(tag, previous.0));
                }
                if names.len() < QUICK_XML_LINEAR_NAMES {
                    names.push(name);
                    return None;
                }
                let previous = core::mem::take(names);
                let mut ordered = BTreeMap::new();
                for previous in previous {
                    ordered.insert(previous, offset_in(tag, previous.0));
                }
                ordered.insert(name, position);
                *self = Self::Ordered(ordered);
                None
            },
            Self::Ordered(names) => match names.entry(name) {
                Entry::Occupied(previous) => Some(*previous.get()),
                Entry::Vacant(entry) => {
                    entry.insert(position);
                    None
                },
            },
        }
    }

    fn position(&self, tag: &[u8], name: &[u8]) -> Option<usize> {
        match self {
            Self::Inline { names, len } => names[..*len]
                .iter()
                .find(|previous| previous.0 == name)
                .map(|previous| offset_in(tag, previous.0)),
            Self::Linear(names) => names
                .iter()
                .find(|previous| previous.0 == name)
                .map(|previous| offset_in(tag, previous.0)),
            Self::Ordered(names) => names.get(name).copied(),
        }
    }
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

impl std::borrow::Borrow<[u8]> for Name<'_> {
    fn borrow(&self) -> &[u8] {
        self.0
    }
}

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
mod candidate_tests {
    use super::{BytesStartExt, INLINE_NAMES, QUICK_XML_LINEAR_NAMES};
    use quick_xml::events::BytesStart;
    use quick_xml::events::attributes::AttrError;

    type Item = Result<(Vec<u8>, Vec<u8>), AttrError>;

    fn tag(content: &str) -> BytesStart<'static> {
        let name_len = content
            .find([' ', '\t', '\r', '\n'])
            .unwrap_or(content.len());
        BytesStart::from_content(content.to_owned(), name_len)
    }

    fn owned(item: Result<quick_xml::events::attributes::Attribute<'_>, AttrError>) -> Item {
        item.map(|attribute| (attribute.key.as_ref().to_vec(), attribute.value.to_vec()))
    }

    #[allow(clippy::disallowed_methods)]
    fn quick_xml_until_error(tag: &BytesStart<'_>) -> Vec<Item> {
        let mut items = Vec::new();
        for item in tag.attributes() {
            let error = item.is_err();
            items.push(owned(item));
            if error {
                break;
            }
        }
        items
    }

    fn checked(tag: &BytesStart<'_>) -> Vec<Item> {
        tag.checked_attributes().map(owned).collect()
    }

    fn names(count: usize) -> Vec<String> {
        (0..count)
            .map(|index| format!("n{index}=\"{index}\""))
            .collect()
    }

    #[test]
    fn matches_quick_xml_at_inline_spill_and_ordered_boundaries() {
        assert_eq!(INLINE_NAMES, 4);
        assert_eq!(QUICK_XML_LINEAR_NAMES, 32);
        for count in [0, 1, 2, 3, 4, 5, 8, 16, 31, 32, 33, 64] {
            let attributes = names(count);
            for separator in [" ", "", "\n\t "] {
                let content = format!("e {}", attributes.join(separator));
                let tag = tag(&content);
                assert_eq!(checked(&tag), quick_xml_until_error(&tag), "{content}");
            }
        }
    }

    #[test]
    fn preserves_duplicate_precedence_and_name_errors() {
        for count in [3, 4, 5, 32, 33] {
            let mut valid = names(count);
            valid.push("n1=\"again\"".to_owned());
            let content = format!("e {}", valid.join(" "));
            let tag = tag(&content);
            assert_eq!(checked(&tag), quick_xml_until_error(&tag), "{content}");

            let mut malformed = names(count);
            malformed.push("n1=x".to_owned());
            let content = format!("e {}", malformed.join("\n\t "));
            let tag = tag(&content);
            assert_eq!(checked(&tag), quick_xml_until_error(&tag), "{content}");
        }

        let long_value = "x".repeat(4096);
        for count in [1, 4, 5, 32, 33] {
            for value in [
                format!("\"{long_value}\""),
                format!("\"{long_value}"),
                format!("'{long_value}"),
                long_value.clone(),
                " ".repeat(4096),
            ] {
                let content = format!("e {} n0={value}", names(count).join(" "));
                let element = tag(&content);
                assert_eq!(
                    checked(&element),
                    quick_xml_until_error(&element),
                    "long duplicate at count={count}"
                );
            }
        }

        for content in [
            "e a=\"1\" a=x tail=\"2\"",
            "e a=\"1\" a= tail=\"2\"",
            "e a=\"1\" a=\"open tail=\"2\"",
            "e a=\"1\" a='open tail=\"2\"",
            "e a=\"1\" a \t=\"2\"",
            "e a=\"1\" a next=\"2\"",
            "e =\"1\"",
            "e flag next=\"2\"",
            "e \u{e9}=\"\u{503c}\" \u{e9}=\"again\"",
        ] {
            let tag = tag(content);
            assert_eq!(checked(&tag), quick_xml_until_error(&tag), "{content}");
            let mut iterator = tag.checked_attributes();
            while let Some(item) = iterator.next() {
                if item.is_err() {
                    assert!(iterator.next().is_none(), "{content}");
                    break;
                }
            }
        }
    }

    #[test]
    fn clones_preserve_inline_and_fallback_state() {
        let valid = tag(&format!("e {}", names(40).join(" ")));
        for advance in [0, 4, 5, 32, 33] {
            let mut iterator = valid.checked_attributes();
            for _ in 0..advance {
                assert!(iterator.next().is_some(), "advance={advance}");
            }
            let clone = iterator.clone();
            assert_eq!(
                iterator.map(owned).collect::<Vec<_>>(),
                clone.map(owned).collect::<Vec<_>>()
            );
        }

        let mut exhausted = valid.checked_attributes();
        assert!(exhausted.all(|item| item.is_ok()));
        let mut exhausted_clone = exhausted.clone();
        assert!(exhausted.next().is_none());
        assert!(exhausted_clone.next().is_none());

        let failed_tag = tag(&format!("e {} n3=\"again\"", names(33).join(" ")));
        let mut failed = failed_tag.checked_attributes();
        assert!(failed.any(|item| item.is_err()));
        let mut failed_clone = failed.clone();
        assert!(failed.next().is_none());
        assert!(failed_clone.next().is_none());
    }
}

#[cfg(test)]
#[path = "../../litchi-opc/src/xml_attributes/tests.rs"]
mod tests;
