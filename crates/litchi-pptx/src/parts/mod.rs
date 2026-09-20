//! Low-level `PresentationML` part views.
//!
//! These wrappers own only vocabulary validation and relationship traversal.
//! They deliberately borrow the OPC graph so an opened package can retain
//! every unmodeled part and byte range until an explicit managed write.

mod presentation;
mod slide;

pub use presentation::{PresentationPart, SlideReference};
pub use slide::{SlideLayoutPart, SlideMasterPart, SlidePart};

use std::borrow::Cow;
use std::mem::size_of;
use std::sync::Arc;

use litchi_ooxml_common::mce::{Error as MceError, Limits as MceLimits, process_part};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, Part, Relationship};
use quick_xml::events::BytesStart;
use quick_xml::reader::NsReader;

use crate::{Error, Result};

pub(crate) const MAX_PART_XML_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_SLIDES: usize = 100_000;

pub(crate) type MceKey = (usize, usize);

/// One successful, default-profile transformed slide projection retained by an
/// immutable opened snapshot. The raw owner is held solely for the address/
/// length key's ABA proof; the package already owns that allocation.
#[derive(Clone)]
struct RetainedMceEntry {
    key: MceKey,
    raw: Arc<Vec<u8>>,
    processed: Arc<Vec<u8>>,
}

/// Capture-local default-profile MCE outputs. This is deliberately private to
/// the opened PPTX capture path; semantic readers and public part views keep
/// their existing processing APIs and ownership behavior.
pub(crate) struct RetainedMce {
    entries: Vec<RetainedMceEntry>,
    charged_bytes: usize,
}

impl RetainedMce {
    pub(crate) fn retained_bytes(&self) -> usize {
        self.charged_bytes
    }

    fn from_candidates(
        candidates: Vec<RetainedMceEntry>,
        maximum: usize,
    ) -> Option<Arc<RetainedMce>> {
        Self::from_candidates_with_reservation(candidates, maximum, |entries, count| {
            entries.try_reserve_exact(count).is_ok()
        })
    }

    fn from_candidates_with_reservation(
        candidates: Vec<RetainedMceEntry>,
        maximum: usize,
        reserve: impl FnOnce(&mut Vec<RetainedMceEntry>, usize) -> bool,
    ) -> Option<Arc<RetainedMce>> {
        debug_assert!(
            candidates
                .windows(2)
                .all(|entries| entries[0].key < entries[1].key)
        );
        let mut output_bytes = 0usize;
        let mut count = 0usize;
        for entry in &candidates {
            let Some(next) = retained_mce_output_charge(entry.processed.capacity())
                .and_then(|charge| output_bytes.checked_add(charge))
            else {
                break;
            };
            if retained_mce_total_charge(count + 1, next).is_none_or(|charge| charge > maximum) {
                break;
            }
            output_bytes = next;
            count += 1;
        }
        if count == 0 {
            return None;
        }
        let mut entries = Vec::new();
        if !reserve(&mut entries, count) || entries.capacity() < count {
            return None;
        }
        entries.extend(candidates.into_iter().take(count));
        // Charge actual capacity, including any allocator over-reservation.
        // Trim outputs without repeatedly allocating smaller candidate tables.
        loop {
            let charged = retained_mce_total_charge(entries.capacity(), output_bytes);
            if let Some(charged_bytes) = charged.filter(|charge| *charge <= maximum) {
                return Some(Arc::new(Self {
                    entries,
                    charged_bytes,
                }));
            }
            let removed = entries.pop()?;
            output_bytes = output_bytes
                .checked_sub(retained_mce_output_charge(removed.processed.capacity())?)?;
            if entries.is_empty() {
                return None;
            }
        }
    }

    pub(crate) fn lookup(&self, source: &[u8]) -> Option<Arc<Vec<u8>>> {
        let key = mce_key(source);
        let index = self
            .entries
            .binary_search_by_key(&key, |entry| entry.key)
            .ok()?;
        let entry = &self.entries[index];
        std::ptr::eq(entry.raw.as_slice().as_ptr(), source.as_ptr())
            .then(|| Arc::clone(&entry.processed))
    }

    pub(crate) fn project_with_owner(
        &self,
        owner: impl Fn(MceKey) -> Option<Arc<Vec<u8>>>,
    ) -> Option<Arc<RetainedMce>> {
        let mut candidates = Vec::new();
        candidates.try_reserve(self.entries.len()).ok()?;
        for entry in &self.entries {
            let Some(raw) = owner(entry.key) else {
                continue;
            };
            candidates.push(RetainedMceEntry {
                key: entry.key,
                raw,
                processed: Arc::clone(&entry.processed),
            });
        }
        Self::from_candidates(candidates, self.charged_bytes)
    }
}

/// Bytes retained by one transformed vector, excluding the final entry-vector
/// allocation. The raw owner is already part of the immutable package and is
/// not charged a second time merely because this table keeps a strong reference
/// to it.
fn retained_mce_entry_charge(processed_capacity: usize) -> Option<usize> {
    retained_mce_output_charge(processed_capacity)?.checked_add(size_of::<RetainedMceEntry>())
}

fn retained_mce_output_charge(processed_capacity: usize) -> Option<usize> {
    processed_capacity
        .checked_add(size_of::<Vec<u8>>())?
        .checked_add(2 * size_of::<usize>())
}

fn retained_mce_outer_charge() -> Option<usize> {
    size_of::<RetainedMce>().checked_add(2 * size_of::<usize>())
}

fn retained_mce_total_charge(entries_capacity: usize, output_bytes: usize) -> Option<usize> {
    let entries_bytes = entries_capacity.checked_mul(size_of::<RetainedMceEntry>())?;
    retained_mce_outer_charge()?
        .checked_add(entries_bytes)?
        .checked_add(output_bytes)
}

#[cfg(test)]
pub(crate) fn checked_mce_charge_for_test(
    entries_capacity: usize,
    output_capacity: usize,
) -> Option<usize> {
    retained_mce_total_charge(
        entries_capacity,
        retained_mce_output_charge(output_capacity)?,
    )
}

#[cfg(test)]
pub(crate) fn refused_mce_reservation_releases_for_test() -> (bool, bool) {
    let raw = Arc::new(vec![1]);
    let processed = Arc::new(vec![2]);
    let weak_raw = Arc::downgrade(&raw);
    let weak_processed = Arc::downgrade(&processed);
    let entry = RetainedMceEntry {
        key: mce_key(raw.as_slice()),
        raw,
        processed,
    };
    let refused =
        RetainedMce::from_candidates_with_reservation(vec![entry], 1024, |entries, _count| {
            entries.try_reserve_exact(usize::MAX).is_ok()
        });
    (
        refused.is_none(),
        weak_raw.upgrade().is_none() && weak_processed.upgrade().is_none(),
    )
}

fn mce_key(source: &[u8]) -> MceKey {
    (source.as_ptr() as usize, source.len())
}

struct PendingMce<'a> {
    source: &'a [u8],
    key: MceKey,
    processed: Arc<Vec<u8>>,
}

/// Capture-local handoff between slide processing and the owned package's
/// already-built [`PartDigests`]. No `blob` or `blob_arc` observation occurs
/// here: source identity comes from the exact second observation made by the
/// raw preflight, while final retention accepts only owned digest entries.
pub(crate) struct MceCapture<'source, 'parent> {
    parent: Option<&'parent RetainedMce>,
    maximum: usize,
    pending: Vec<PendingMce<'source>>,
    pending_charge: usize,
}

impl<'source, 'parent> MceCapture<'source, 'parent> {
    pub(crate) fn new(parent: Option<&'parent RetainedMce>, maximum: usize) -> Self {
        Self {
            parent,
            maximum,
            pending: Vec::new(),
            pending_charge: 0,
        }
    }

    // Provisional visits are conservatively charged individually, including
    // repeated references to the same output. Finalization deduplicates keys
    // and charges the unique final table's actual capacity. Capture-local
    // vector metadata is transient workspace, not part of the retained table.
    fn admit_visit(&mut self, capacity: usize) -> bool {
        let Some(next) = retained_mce_entry_charge(capacity)
            .and_then(|charge| self.pending_charge.checked_add(charge))
        else {
            return false;
        };
        if retained_mce_outer_charge()
            .and_then(|outer| outer.checked_add(next))
            .is_none_or(|total| total > self.maximum)
            || self.pending.try_reserve(1).is_err()
        {
            return false;
        }
        self.pending_charge = next;
        true
    }

    pub(crate) fn process(&mut self, source: &'source [u8]) -> Result<ProcessedBytes<'source>> {
        if let Some(parent) = self.parent
            && parent.retained_bytes() <= self.maximum
            && let Some(processed) = parent.lookup(source)
            && validate_cached_mce(source, processed.as_slice()).is_ok()
        {
            #[cfg(debug_assertions)]
            debug_assert_cached_mce(source, processed.as_slice());
            if self.admit_visit(processed.capacity()) {
                self.pending.push(PendingMce {
                    source,
                    key: mce_key(source),
                    processed: Arc::clone(&processed),
                });
            }
            return Ok(ProcessedBytes::Retained(processed));
        }
        let processed = litchi_ooxml_common::mce::process_ooxml(source)?;
        match processed {
            Cow::Borrowed(_) => Ok(ProcessedBytes::Borrowed(source)),
            Cow::Owned(processed) => {
                if !self.admit_visit(processed.capacity()) {
                    return Ok(ProcessedBytes::Owned(processed));
                }
                // Arc::new moves the existing output allocation. Global OOM
                // follows the surrounding library's infallible Arc behavior;
                // explicit vector reservations fall back to ordinary parsing.
                let processed = Arc::new(processed);
                self.pending.push(PendingMce {
                    source,
                    key: mce_key(source),
                    processed: Arc::clone(&processed),
                });
                Ok(ProcessedBytes::Retained(processed))
            },
        }
    }

    /// Finish only keys visited by this capture and proved to be owned by the
    /// captured package's digest memo. A foreign clone that returns a copied
    /// `blob_arc` therefore drops back to ordinary processing automatically.
    pub(crate) fn finish(
        self,
        owner: impl Fn(MceKey) -> Option<Arc<Vec<u8>>>,
    ) -> Option<Arc<RetainedMce>> {
        let Self {
            maximum,
            mut pending,
            ..
        } = self;
        if pending.is_empty() {
            return None;
        }
        pending.sort_unstable_by_key(|entry| entry.key);
        pending.dedup_by_key(|entry| entry.key);
        let mut candidates = Vec::new();
        candidates.try_reserve(pending.len()).ok()?;
        for entry in pending {
            if mce_key(entry.source) != entry.key {
                continue;
            }
            let Some(raw) = owner(entry.key) else {
                continue;
            };
            candidates.push(RetainedMceEntry {
                key: entry.key,
                raw,
                processed: entry.processed,
            });
        }
        RetainedMce::from_candidates(candidates, maximum)
    }
}

pub(crate) enum ProcessedBytes<'a> {
    Borrowed(&'a [u8]),
    Owned(Vec<u8>),
    Retained(Arc<Vec<u8>>),
}

impl AsRef<[u8]> for ProcessedBytes<'_> {
    fn as_ref(&self) -> &[u8] {
        match self {
            Self::Borrowed(bytes) => bytes,
            Self::Owned(bytes) => bytes,
            Self::Retained(bytes) => bytes.as_slice(),
        }
    }
}

/// The exact source slice and its temporary MCE projection used by one
/// capture-local reader.  Keeping the source slice beside the projection
/// prevents a later consumer from mistaking a projection of one foreign
/// `Part::blob()` result for a projection of another result.
pub(crate) struct ProcessedXml<'a> {
    pub(crate) source: &'a [u8],
    pub(crate) processed: ProcessedBytes<'a>,
}

fn raw_xml(part: &dyn Part) -> Result<&[u8]> {
    if part.blob().len() > MAX_PART_XML_BYTES {
        return Err(Error::Limit {
            resource: "PresentationML part XML",
            limit: MAX_PART_XML_BYTES,
        });
    }
    Ok(part.blob())
}

fn validate_cached_mce(source: &[u8], processed: &[u8]) -> Result<()> {
    let limits = MceLimits::default();
    if source.len() > limits.max_input_bytes {
        return Err(Error::MarkupCompatibility(MceError::LimitExceeded(
            "input bytes".to_owned(),
        )));
    }
    if processed.len() > limits.max_output_bytes {
        return Err(Error::MarkupCompatibility(MceError::LimitExceeded(
            "output bytes".to_owned(),
        )));
    }
    Ok(())
}

#[cfg(debug_assertions)]
fn debug_assert_cached_mce(source: &[u8], cached: &[u8]) {
    let recomputed = litchi_ooxml_common::mce::process_ooxml(source);
    debug_assert!(
        recomputed.is_ok(),
        "a retained MCE source must remain processable"
    );
    if let Ok(recomputed) = recomputed {
        debug_assert_eq!(recomputed.as_ref(), cached);
    }
}

/// Process one exact source observation while retaining that observation for
/// a capture-local proof.  The first `blob()` call preserves the established
/// raw-size preflight; the second call is the exact slice supplied to MCE.
pub(crate) fn processed_xml_with_source(part: &dyn Part) -> Result<ProcessedXml<'_>> {
    let source = raw_xml(part)?;
    let processed = litchi_ooxml_common::mce::process_ooxml(source)?;
    let processed = match processed {
        Cow::Borrowed(_) => ProcessedBytes::Borrowed(source),
        Cow::Owned(bytes) => ProcessedBytes::Owned(bytes),
    };
    Ok(ProcessedXml { source, processed })
}

pub(crate) fn processed_xml_with_capture<'a, 'parent>(
    part: &'a dyn Part,
    capture: &mut MceCapture<'a, 'parent>,
) -> Result<ProcessedXml<'a>> {
    let source = raw_xml(part)?;
    let processed = capture.process(source)?;
    Ok(ProcessedXml { source, processed })
}

pub(crate) fn processed_xml(part: &dyn Part) -> Result<Cow<'_, [u8]>> {
    if part.blob().len() > MAX_PART_XML_BYTES {
        return Err(Error::Limit {
            resource: "PresentationML part XML",
            limit: MAX_PART_XML_BYTES,
        });
    }
    Ok(process_part(part)?)
}

pub(crate) fn relationship_attribute(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> Result<Option<String>> {
    crate::namespace::relationship_attribute_value(
        element,
        b"id",
        reader.decoder(),
        reader.resolver(),
    )
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(crate) fn parse_u32(value: &str, field: &str) -> Result<u32> {
    value
        .parse()
        .map_err(|_err| invalid(format!("invalid {field} value '{value}'")))
}

pub(crate) fn parse_i64(value: &str, field: &str) -> Result<i64> {
    value
        .parse()
        .map_err(|_err| invalid(format!("invalid {field} value '{value}'")))
}

pub(crate) fn parse_bool(value: &str, field: &str) -> Result<bool> {
    match value {
        "1" | "true" | "on" => Ok(true),
        "0" | "false" | "off" => Ok(false),
        _ => Err(invalid(format!("invalid {field} value '{value}'"))),
    }
}

pub(crate) fn validate_content_type(part: &dyn Part, expected: &str) -> Result<()> {
    if part.content_type() == expected {
        return Ok(());
    }
    Err(Error::ContentType {
        expected: expected.to_string(),
        actual: part.content_type().to_string(),
    })
}

pub(crate) fn is_relationship_type(actual: &str, transitional: &str, local: &str) -> bool {
    actual == transitional
        || actual == format!("http://purl.oclc.org/ooxml/officeDocument/relationships/{local}")
}

pub(crate) fn related_part_by_type<'a>(
    package: &'a OpcPackage,
    source: &dyn Part,
    relationship_type: &str,
    local_type: &str,
    content_type: &str,
) -> Result<Option<&'a dyn Part>> {
    let mut matching = source.rels().iter().filter(|relationship| {
        is_relationship_type(relationship.reltype(), relationship_type, local_type)
    });
    let Some(relationship) = matching.next() else {
        return Ok(None);
    };
    if matching.next().is_some() {
        return Err(Error::Relationship(format!(
            "part '{}' has multiple '{local_type}' relationships",
            source.partname()
        )));
    }
    if relationship.is_external() {
        return Err(Error::Relationship(format!(
            "'{local_type}' relationship must be internal"
        )));
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    validate_content_type(part, content_type)?;
    Ok(Some(part))
}

pub(crate) fn expected_main_content_type(content_type: &str) -> bool {
    matches!(
        content_type,
        ct::PML_PRESENTATION_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    )
}

pub(crate) fn validate_slide_relationship<'a, T>(
    relationship: Option<&'a Relationship>,
    relationship_id: &str,
    resolve_target: impl FnOnce(&PackURI) -> Result<T>,
    content_type: impl FnOnce(&T) -> &str,
) -> Result<(&'a Relationship, PackURI, T)> {
    let relationship = relationship.ok_or_else(|| {
        Error::Relationship(format!(
            "presentation slide reference is missing relationship '{relationship_id}'"
        ))
    })?;
    if relationship.is_external() {
        return Err(Error::Relationship(format!(
            "slide relationship '{relationship_id}' must be internal"
        )));
    }
    if !is_relationship_type(relationship.reltype(), rt::SLIDE, "slide") {
        return Err(Error::Relationship(format!(
            "relationship '{relationship_id}' has unexpected type '{}'",
            relationship.reltype()
        )));
    }
    let target = relationship.target_partname()?;
    let part = resolve_target(&target)?;
    let actual = content_type(&part);
    if actual != ct::PML_SLIDE {
        return Err(Error::ContentType {
            expected: ct::PML_SLIDE.to_string(),
            actual: actual.to_string(),
        });
    }
    Ok((relationship, target, part))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::Relationships;
    use litchi_opc::part::BlobPart;
    use std::sync::Arc;
    use std::sync::atomic::{AtomicUsize, Ordering};

    #[derive(Clone)]
    struct AlternatingBlobPart {
        inner: BlobPart,
        first: &'static [u8],
        second: &'static [u8],
        calls: Arc<AtomicUsize>,
    }

    impl Part for AlternatingBlobPart {
        fn blob(&self) -> &[u8] {
            if self.calls.fetch_add(1, Ordering::SeqCst) % 2 == 0 {
                self.first
            } else {
                self.second
            }
        }

        fn blob_arc(&self) -> Arc<Vec<u8>> {
            Arc::new(self.blob().to_vec())
        }

        fn content_type(&self) -> &str {
            self.inner.content_type()
        }

        fn partname(&self) -> &PackURI {
            self.inner.partname()
        }

        fn rels(&self) -> &Relationships {
            self.inner.rels()
        }

        fn rels_mut(&mut self) -> &mut Relationships {
            self.inner.rels_mut()
        }

        fn set_blob(&mut self, blob: Vec<u8>) {
            self.inner.set_blob(blob);
        }

        fn set_content_type(&mut self, content_type: String) -> litchi_opc::Result<()> {
            self.inner.set_content_type(content_type)
        }
    }

    #[test]
    fn source_bound_processing_uses_the_second_blob_observation() {
        const FIRST: &[u8] = b"size-only source";
        const SECOND: &[u8] =
            br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
        let part = AlternatingBlobPart {
            inner: BlobPart::new(
                PackURI::new("/ppt/slides/alternate.xml").expect("valid part name"),
                crate::notes::SLIDE_CT.to_owned(),
                FIRST.to_vec(),
            ),
            first: FIRST,
            second: SECOND,
            calls: Arc::new(AtomicUsize::new(0)),
        };

        let processed = processed_xml_with_source(&part).expect("second source must be XML");
        assert_eq!(processed.source, SECOND);
        assert_eq!(processed.processed.as_ref(), SECOND);
        assert_eq!(part.calls.load(Ordering::SeqCst), 2);
    }
}
