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
//! `litchi_ooxml_common::xml::attributes::first_wins` instead, and a reader
//! that needs no duplicate check uses
//! [`BytesStartExt::unchecked_attributes`]. The workspace's `clippy.toml`
//! disallows quick-xml's checked iteration in the crates that read untrusted
//! OOXML and OLE2 XML, so a new call site has to choose one of the three.

// The code lives in `xml-minifier`, which this crate already depends on; the
// crates that may not depend on it keep checked copies.
pub use xml_minifier::xml_attributes::{BytesStartExt, CheckedAttributes};
