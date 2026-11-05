//! Readback verification of shape-text publications (record 0264).
//!
//! A shape-text publication appends rewritten slide records and must prove,
//! before adoption, that the published package reads back the requested
//! texts. This module serves the batch publisher; the single-shape commit
//! uses [`readback_incremental`] for its target resolution in the same way
//! (`resolve_published`). The
//! reference verifier, [`readback_full`], reopens the candidate and parses the
//! whole drawing of every edited slide. Its verdict for a publication `P`
//! with targets `T` is *accept* exactly when:
//!
//! 1. the candidate opens as a presentation and lists its slides (done by the
//!    caller for both verifiers, so it is not repeated here);
//! 2. for every `t` in `T`, in order: the slide at `t.slide` exists and has
//!    `t`'s persist identifier; that slide's shapes parse (structural errors
//!    are fatal; content errors drop the shape and are recorded); the shape at
//!    `t.shape` exists; and its text equals `t`'s replacement.
//!
//! [`readback_incremental`] proves the same verdict without parsing unchanged
//! shapes. It answers *accept* (with the shapes that a parse returns) or
//! *cannot prove*; [`readback`] then runs the reference verifier, so
//! `readback(P) = accept` iff `readback_full(P) = accept`, with the same error
//! otherwise, as long as an incremental accept implies a full accept.
//!
//! The incremental accept rests on `Slide::parse_shapes` being a pure
//! function of the drawing payload and the slide's outline bodies, applied
//! shape by shape: `shapes(D) = filter_some(convert(e_i, i, outline))` over
//! the `OfficeArt` shapes `e_i` of `D`, where `e_i` depends only on its own
//! container's bytes and `OfficeArt` role, and the drawing-level checks
//! (container framing, record budget, nesting) depend only on the record
//! tree. The incremental verifier accepts only when all of these hold:
//!
//! - the source slide's shapes parsed with an empty tolerance report (so no
//!   shape was dropped, and every position equals its `OfficeArt` index);
//! - the new slide's outline bodies equal the source slide's;
//! - a lock-step walk of the source and published drawing payloads finds
//!   identical record trees: the same record count and header kind, version
//!   and instance at every position, byte-identical records everywhere except
//!   along the container chain of changed top-level shape containers, and,
//!   inside each changed shape container, byte-identical children except
//!   exactly one `ClientTextbox` record (whose header kind is unchanged);
//! - every changed container is the top-level shape of exactly one target
//!   (matched by `OfficeArt` index and shape identifier), and every target
//!   changed exactly one container;
//! - each changed container parses on its own, converts at its index without
//!   being dropped and without recording a repair, and reads back its
//!   replacement.
//!
//! Then the published record tree has the same shape, the same budget and the
//! same depth as the source tree, every unchanged shape converts to the value
//! the source parse produced, each changed shape converts to the value
//! computed here (the `OfficeArt` role checks read only the shape's `Sp`
//! atom and meta records, which are unchanged), and so a full parse returns
//! the source shapes with the targets replaced, no repairs, and the expected
//! texts at the target positions: the full verifier accepts.
//!
//! An accept also yields those shapes ([`PatchedShapes`]): the committed
//! package's slide memo is seeded with them, so the next edit of the same
//! slide resolves its targets and starts its own readback without parsing
//! the slide again. `src/text_edit/readback_tests.rs` runs both verifiers
//! on genuine, corrupted and adversarial publications and requires that an
//! incremental accept is always a full accept with exactly the parsed shapes.

use std::collections::BTreeMap;
use std::sync::Arc;

use super::{Error, Refusal, Result, Target, native_shape_id};
use crate::package::Error as PackageError;
use crate::shapes::ShapeEnum;
use crate::slide::{PatchedShapes, Slide};

const DG_CONTAINER: u16 = 0xF002;
const SPGR_CONTAINER: u16 = 0xF003;
const SP_CONTAINER: u16 = 0xF004;
const SP: u16 = 0xF00A;
const CLIENT_TEXTBOX: u16 = 0xF00D;
/// `fBackground` in the `OfficeArtFSP` flags (MS-ODRAW 2.2.40).
const BACKGROUND_FLAG: u32 = 0x0400;

/// What the source parse of one edited slide established.
pub(crate) struct SourceSlideView {
    shapes: Arc<[ShapeEnum<'static>]>,
    clean: bool,
    outline: Vec<crate::OutlineTextBody>,
    drawing: Option<Vec<u8>>,
}

impl std::fmt::Debug for SourceSlideView {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SourceSlideView")
            .field("shapes", &self.shapes.len())
            .field("clean", &self.clean)
            .finish_non_exhaustive()
    }
}

impl SourceSlideView {
    /// Capture an already parsed source slide.
    pub(super) fn capture(slide: &Slide<'_>) -> crate::package::Result<Self> {
        let shapes = slide.shared_shapes()?;
        let report = slide.tolerance_report()?;
        Ok(Self {
            shapes,
            clean: report.defects().is_empty() && report.unrecorded() == 0,
            outline: slide.outline_text_bodies().to_vec(),
            drawing: slide.drawing_payload().map(<[u8]>::to_vec),
        })
    }
}

/// One target the publication must read back.
pub(super) struct ReadbackTarget<'a> {
    pub(super) target: Target,
    pub(super) slide_persist_id: u32,
    pub(super) native_shape_id: u32,
    pub(super) replacement: &'a str,
}

/// Shapes a parse of each edited published slide returns, proven by the
/// incremental verifier, keyed by persist identifier.
pub(crate) type VerifiedSlides = Vec<(u32, PatchedShapes)>;

/// Verify a publication: incrementally when that can be proven, otherwise
/// with the reference verifier. Returns the proven shapes of each edited
/// slide when the incremental verifier accepted (empty otherwise).
pub(super) fn readback(
    slides: &[Slide<'_>],
    targets: &[ReadbackTarget<'_>],
    sources: &BTreeMap<u32, Arc<SourceSlideView>>,
) -> Result<VerifiedSlides> {
    if let Some(verified) = readback_incremental(slides, targets, sources) {
        return Ok(verified);
    }
    readback_full(slides, targets)?;
    Ok(Vec::new())
}

/// The reference verifier: parse every edited slide's whole drawing.
pub(super) fn readback_full(slides: &[Slide<'_>], targets: &[ReadbackTarget<'_>]) -> Result<()> {
    for target in targets {
        let slide = slides
            .get(target.target.slide().get())
            .ok_or(Error::Refused(Refusal::SlideNotFound {
                position: target.target.slide(),
            }))?;
        if slide.persist_id() != target.slide_persist_id {
            return Err(PackageError::Corrupted(
                "published PPT batch shape text resolved a different slide".into(),
            )
            .into());
        }
        let shape = slide
            .shapes()?
            .get(target.target.shape().get())
            .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
        if shape.text()? != target.replacement {
            return Err(PackageError::Corrupted(
                "published PPT batch shape text did not round-trip through its source shape".into(),
            )
            .into());
        }
    }
    Ok(())
}

/// The incremental verifier. `None` means "cannot prove", never "refuse".
pub(super) fn readback_incremental(
    slides: &[Slide<'_>],
    targets: &[ReadbackTarget<'_>],
    sources: &BTreeMap<u32, Arc<SourceSlideView>>,
) -> Option<VerifiedSlides> {
    let mut groups = BTreeMap::<u32, (&Slide<'_>, Vec<&ReadbackTarget<'_>>)>::new();
    for target in targets {
        let slide = slides.get(target.target.slide().get())?;
        if slide.persist_id() != target.slide_persist_id {
            return None;
        }
        let group = groups
            .entry(target.slide_persist_id)
            .or_insert_with(|| (slide, Vec::new()));
        if !std::ptr::eq(group.0, slide) {
            return None;
        }
        group.1.push(target);
    }
    let mut verified = Vec::new();
    for (persist_id, (slide, targets)) in groups {
        let source = sources.get(&persist_id)?;
        if !source.clean || slide.outline_text_bodies() != source.outline.as_slice() {
            return None;
        }
        let changed = drawing_delta(source.drawing.as_deref()?, slide.drawing_payload()?)?;
        if changed.len() != targets.len() {
            return None;
        }
        // The walk emits containers in strictly increasing index order.
        let indices = changed.iter().map(|shape| shape.index).collect::<Vec<_>>();
        if indices.windows(2).any(|pair| pair[0] >= pair[1]) {
            return None;
        }
        let mut by_index = changed.into_iter().map(Some).collect::<Vec<_>>();
        let mut replacements = Vec::with_capacity(targets.len());
        for target in targets {
            let position = target.target.shape().get();
            let before = source.shapes.get(position)?;
            if native_shape_id(before) != target.native_shape_id
                || before.source_index() != Some(position)
            {
                return None;
            }
            let slot = indices.binary_search(&position).ok()?;
            let container = by_index.get_mut(slot)?.take()?;
            if container.shape_id != target.native_shape_id {
                return None;
            }
            let after = convert_changed_shape(
                container.bytes,
                position,
                slide.outline_text_bodies(),
                slide.source_limits(),
                target,
            )?;
            replacements.push((position, after));
        }
        if by_index.iter().any(Option::is_some) {
            return None;
        }
        // Materialized only if the published slide is read again.
        replacements.sort_unstable_by_key(|(position, _)| *position);
        verified.push((
            persist_id,
            PatchedShapes::new(Arc::clone(&source.shapes), replacements)?,
        ));
    }
    Some(verified)
}

/// Convert one changed top-level shape container exactly as the slide parse
/// would at `index`, and check its readback.
fn convert_changed_shape(
    container: &[u8],
    index: usize,
    outline: &[crate::OutlineTextBody],
    source_limits: crate::SourceLimits,
    target: &ReadbackTarget<'_>,
) -> Option<ShapeEnum<'static>> {
    // The role checks a drawing applies (`PATRIARCH`/`CHILD` against the
    // topology) read only the unchanged `Sp` atom, which the source parse
    // already accepted at this position; a standalone parse builds the same
    // shape value.
    let escher =
        litchi_odraw::shape::parse(container, crate::odraw::limits_for_source(source_limits))
            .ok()?;
    let [escher] = escher.as_slice() else {
        return None;
    };
    let mut report = crate::leniency::ToleranceReport::default();
    let mut shape =
        Slide::convert_odraw_to_shape_enum(escher, index, outline, source_limits, &mut report)
            .ok()??;
    if !report.defects().is_empty() || report.unrecorded() != 0 {
        return None;
    }
    shape.mark_source_bound_recursive();
    if native_shape_id(&shape) != target.native_shape_id
        || shape.source_index() != Some(index)
        || shape.text().ok()? != target.replacement
    {
        return None;
    }
    Some(shape)
}

/// One changed top-level shape container of a published drawing.
pub(super) struct ChangedShape<'a> {
    /// Zero-based `OfficeArt` index among the drawing's top-level shapes.
    pub(super) index: usize,
    /// The shape identifier from the container's (unchanged) `Sp` atom.
    pub(super) shape_id: u32,
    /// The complete published container record.
    pub(super) bytes: &'a [u8],
}

/// One framed `OfficeArt` record.
#[derive(Clone, Copy)]
struct Raw<'a> {
    whole: &'a [u8],
    payload: &'a [u8],
}

impl Raw<'_> {
    fn kind(self) -> u16 {
        u16::from_le_bytes([self.whole[2], self.whole[3]])
    }

    fn is_container(self) -> bool {
        self.whole[0] & 0x0F == 0x0F
    }
}

/// Equal version, instance and kind: every header field except the length.
fn same_identity(left: Raw<'_>, right: Raw<'_>) -> bool {
    left.whole[..4] == right.whole[..4]
}

/// Iterate the records framed in `data`; yields `Some(None)` once on broken
/// framing.
struct Records<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> Records<'a> {
    const fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }
}

impl<'a> Iterator for Records<'a> {
    type Item = Option<Raw<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.offset == self.data.len() {
            return None;
        }
        let rest = self.data.get(self.offset..)?;
        let Some(header) = rest.get(..8) else {
            self.offset = self.data.len();
            return Some(None);
        };
        let length = u32::from_le_bytes([header[4], header[5], header[6], header[7]]);
        let end = usize::try_from(length)
            .ok()
            .and_then(|length| length.checked_add(8));
        let Some(whole) = end.and_then(|end| rest.get(..end)) else {
            self.offset = self.data.len();
            return Some(None);
        };
        self.offset += whole.len();
        Some(Some(Raw {
            whole,
            payload: &whole[8..],
        }))
    }
}

/// Pair the records of two payloads; `None` unless both frame cleanly into
/// the same number of records with the same identities.
fn paired<'a, 'b>(old: &'a [u8], new: &'b [u8]) -> Option<Vec<(Raw<'a>, Raw<'b>)>> {
    let mut left = Records::new(old);
    let mut right = Records::new(new);
    let mut pairs = Vec::new();
    loop {
        match (left.next(), right.next()) {
            (None, None) => return Some(pairs),
            (Some(Some(old)), Some(Some(new))) if same_identity(old, new) => {
                pairs.push((old, new));
            },
            _ => return None,
        }
    }
}

/// Lock-step walk of a source and a published `PPDrawing` payload. Returns
/// the changed top-level shape containers, or `None` when the two record
/// trees differ in any other way.
pub(super) fn drawing_delta<'b>(old: &[u8], new: &'b [u8]) -> Option<Vec<ChangedShape<'b>>> {
    let mut walk = Walk {
        index: 0,
        changed: Vec::new(),
    };
    for (old_root, new_root) in paired(old, new)? {
        // A PowerPoint drawing root is a `DgContainer`; other roots are
        // valid `OfficeArt` but are left to the reference verifier.
        if old_root.kind() != DG_CONTAINER || !old_root.is_container() {
            return None;
        }
        walk.drawing(old_root, new_root)?;
    }
    Some(walk.changed)
}

struct Walk<'b> {
    index: usize,
    changed: Vec<ChangedShape<'b>>,
}

impl<'b> Walk<'b> {
    /// `DgContainer` children, in the order the drawing parser emits shapes.
    fn drawing(&mut self, old: Raw<'_>, new: Raw<'b>) -> Option<()> {
        let mut root_groups = 0usize;
        for (old, new) in paired(old.payload, new.payload)? {
            match old.kind() {
                SP_CONTAINER if old.is_container() => self.top_level_shape(old, new)?,
                SPGR_CONTAINER if old.is_container() => {
                    root_groups += 1;
                    if root_groups > 1 {
                        return None;
                    }
                    self.root_group(old, new)?;
                },
                _ => equal(old, new)?,
            }
        }
        Some(())
    }

    /// The patriarch group: a group header, then top-level shapes and groups.
    fn root_group(&mut self, old: Raw<'_>, new: Raw<'b>) -> Option<()> {
        let pairs = paired(old.payload, new.payload)?;
        let ((old_header, new_header), children) = pairs.split_first()?;
        if old_header.kind() != SP_CONTAINER {
            return None;
        }
        equal(*old_header, *new_header)?;
        for &(old, new) in children {
            match old.kind() {
                SP_CONTAINER if old.is_container() => self.top_level_shape(old, new)?,
                SPGR_CONTAINER if old.is_container() => {
                    // A nested group is one top-level shape; edits inside it
                    // are left to the reference verifier.
                    equal(old, new)?;
                    self.index = self.index.checked_add(1)?;
                },
                _ => equal(old, new)?,
            }
        }
        Some(())
    }

    /// One top-level shape container; background shapes are not listed.
    fn top_level_shape(&mut self, old: Raw<'_>, new: Raw<'b>) -> Option<()> {
        let (shape_id, flags) = sp_atom(old)?;
        if flags & BACKGROUND_FLAG != 0 {
            return equal(old, new);
        }
        let index = self.index;
        self.index = self.index.checked_add(1)?;
        if old.whole == new.whole {
            return Some(());
        }
        let mut textboxes = 0usize;
        for (old, new) in paired(old.payload, new.payload)? {
            if old.whole == new.whole {
                continue;
            }
            if old.kind() != CLIENT_TEXTBOX {
                return None;
            }
            textboxes += 1;
        }
        if textboxes != 1 {
            return None;
        }
        self.changed.push(ChangedShape {
            index,
            shape_id,
            bytes: new.whole,
        });
        Some(())
    }
}

fn equal(old: Raw<'_>, new: Raw<'_>) -> Option<()> {
    (old.whole == new.whole).then_some(())
}

/// The shape identifier and flags of a shape container's first `Sp` atom.
fn sp_atom(container: Raw<'_>) -> Option<(u32, u32)> {
    for record in Records::new(container.payload) {
        let record = record?;
        if record.kind() == SP {
            let payload = record.payload;
            let id = u32::from_le_bytes(payload.get(..4)?.try_into().ok()?);
            let flags = u32::from_le_bytes(payload.get(4..8)?.try_into().ok()?);
            return Some((id, flags));
        }
    }
    None
}
