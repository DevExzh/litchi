//! Blank-slide insertion for the root slide-order transaction (record 0260).
//!
//! A blank slide is a `SlideContainer` with a `SlideAtom` (blank layout,
//! following its master's objects, scheme and background), a drawing that
//! holds only the patriarch group shape and the background shape, and a
//! `ColorSchemeAtom`, plus one `SlidePersistAtom` in the slide list. Apache
//! POI's `HSLFSlideShow.createSlide` builds the same records.
//!
//! The drawing needs an identifier and a shape-identifier cluster that no
//! other drawing uses. They are allocated after every drawing the
//! `OfficeArtFDGG` registers and after every drawing this transaction
//! already inserted: the next drawing identifier and the next cluster, whose
//! shape identifiers start at `cluster * 1024` (MS-ODRAW 2.2.46). At commit,
//! [`register_inserted_drawings`] appends one `OfficeArtIDCL` per new
//! cluster and updates the `OfficeArtFDGG` counts, as POI's
//! `EscherDggRecord.allocateShapeId` does; every other byte of the live
//! document is copied exactly.

use std::collections::BTreeMap;

use super::{Error, PackageError, Refusal, Result};

/// `PPDrawingGroup` (MS-PPT 2.4.3).
const PP_DRAWING_GROUP: u16 = 0x040B;
/// `PPDrawing` (MS-PPT 2.5.13).
const PP_DRAWING: u16 = 0x040C;
/// `ColorSchemeAtom` (MS-PPT 2.4.17).
const COLOR_SCHEME_ATOM: u16 = 0x07F0;
/// `SlideAtom` (MS-PPT 2.5.10).
const SLIDE_ATOM: u16 = 0x03EF;
/// `SlideContainer` (MS-PPT 2.5.1).
const SLIDE: u16 = 0x03EE;
/// `OfficeArtDggContainer`.
const DGG_CONTAINER: u16 = 0xF000;
/// `OfficeArtDgContainer`.
const DG_CONTAINER: u16 = 0xF002;
/// `OfficeArtFDGGBlock`.
const FDGG: u16 = 0xF006;
/// `OfficeArtFDG`.
const FDG: u16 = 0xF008;
/// `OfficeArtFSP`.
const FSP: u16 = 0xF00A;
/// `SL_Blank` (MS-PPT 2.13.43; POI `SlideLayoutType.BLANK_SLIDE`).
const SL_BLANK: u32 = 0x0000_0010;
/// `fMasterObjects | fMasterScheme | fMasterBackground` (MS-PPT 2.5.10).
const FOLLOW_MASTER: u16 = 0x0007;
/// Shape identifiers per cluster (MS-ODRAW 2.2.46).
const CLUSTER_SIZE: u32 = 1024;
/// The largest valid `spidMax` (MS-ODRAW 2.2.48).
const MAX_SPID: u32 = 0x03FF_D7FF;
/// Drawing identifiers are the 12-bit `OfficeArtFDG` record instance.
const MAX_DRAWING_ID: u32 = 0x0FFE;

/// The `OfficeArtFDGG` of a live document and its clusters.
#[derive(Debug, Clone, PartialEq, Eq)]
struct DrawingGroup {
    spid_max: u32,
    csp_saved: u32,
    cdg_saved: u32,
    /// `(dgid, cspidCur)` per cluster; cluster `n` (1-based) is entry `n-1`.
    clusters: Vec<(u32, u32)>,
}

impl DrawingGroup {
    fn parse(payload: &[u8]) -> Result<Self> {
        let field = |offset: usize| super::read_payload_u32(payload, offset);
        let cidcl = field(4)?;
        let count = usize::try_from(cidcl)
            .ok()
            .and_then(|cidcl| cidcl.checked_sub(1))
            .ok_or_else(|| corrupted("OfficeArtFDGG cidcl is zero"))?;
        if Some(payload.len()) != count.checked_mul(8).and_then(|size| size.checked_add(16)) {
            return Err(corrupted("OfficeArtFDGG length disagrees with its cidcl"));
        }
        let mut clusters = Vec::new();
        clusters
            .try_reserve_exact(count)
            .map_err(|_error| PackageError::AllocationFailed("PPT drawing clusters"))?;
        for index in 0..count {
            let offset = 16 + index * 8;
            clusters.push((field(offset)?, field(offset + 4)?));
        }
        Ok(Self {
            spid_max: field(0)?,
            csp_saved: field(8)?,
            cdg_saved: field(12)?,
            clusters,
        })
    }

    fn encode(&self) -> Result<Vec<u8>> {
        let cidcl = u32::try_from(self.clusters.len())
            .ok()
            .and_then(|count| count.checked_add(1))
            .ok_or_else(|| PackageError::ResourceLimit("PPT drawing cluster count".into()))?;
        let mut payload = Vec::with_capacity(16 + self.clusters.len() * 8);
        for value in [self.spid_max, cidcl, self.csp_saved, self.cdg_saved] {
            payload.extend_from_slice(&value.to_le_bytes());
        }
        for (dgid, used) in &self.clusters {
            payload.extend_from_slice(&dgid.to_le_bytes());
            payload.extend_from_slice(&used.to_le_bytes());
        }
        Ok(payload)
    }

    /// The 1-based number the next appended cluster receives.
    fn next_cluster(&self) -> u32 {
        u32::try_from(self.clusters.len())
            .unwrap_or(u32::MAX)
            .saturating_add(1)
    }
}

/// The drawing identifier and shape identifiers one inserted slide uses.
#[derive(Debug, Clone, PartialEq, Eq)]
struct InsertedDrawing {
    drawing_id: u32,
    shape_ids: Vec<u32>,
}

/// The identifiers allocated to a new blank slide's drawing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Allocation {
    pub(super) drawing_id: u32,
    pub(super) spid_base: u32,
}

/// Allocates a drawing identifier and a shape-identifier cluster after every
/// registered drawing and every drawing already inserted by this edit.
pub(super) fn allocate(document: &[u8], inserted: &BTreeMap<u32, Vec<u8>>) -> Result<Allocation> {
    let group = drawing_group(document)?;
    let mut drawing_id = group
        .clusters
        .iter()
        .map(|(dgid, _)| *dgid)
        .max()
        .unwrap_or(0);
    let mut cluster = group.next_cluster();
    for record in inserted.values() {
        let Some(drawing) = inserted_drawing(record)? else {
            continue;
        };
        drawing_id = drawing_id.max(drawing.drawing_id);
        for shape_id in drawing.shape_ids {
            cluster = cluster.max((shape_id / CLUSTER_SIZE).saturating_add(1));
        }
    }
    let drawing_id = drawing_id
        .checked_add(1)
        .filter(|identifier| *identifier <= MAX_DRAWING_ID)
        .ok_or_else(|| {
            PackageError::ResourceLimit("PPT drawing identifiers are exhausted".into())
        })?;
    let spid_base = cluster
        .checked_mul(CLUSTER_SIZE)
        .filter(|base| base.saturating_add(CLUSTER_SIZE) <= MAX_SPID)
        .ok_or_else(|| {
            PackageError::ResourceLimit("PPT shape-identifier clusters are exhausted".into())
        })?;
    Ok(Allocation {
        drawing_id,
        spid_base,
    })
}

/// Encodes a blank `SlideContainer` following `master_id`, with the given
/// color scheme record (or the builder's default scheme).
pub(super) fn slide_record(
    master_id: u32,
    color_scheme: Option<Vec<u8>>,
    allocation: Allocation,
) -> Result<Vec<u8>> {
    let mut atom = Vec::with_capacity(32);
    push_header(&mut atom, 2, 0, SLIDE_ATOM, 24)?;
    atom.extend_from_slice(&SL_BLANK.to_le_bytes());
    atom.extend_from_slice(&[0; 8]);
    atom.extend_from_slice(&master_id.to_le_bytes());
    atom.extend_from_slice(&0u32.to_le_bytes());
    atom.extend_from_slice(&FOLLOW_MASTER.to_le_bytes());
    atom.extend_from_slice(&0u16.to_le_bytes());

    let drawing = crate::writer::escher::create_blank_dg_container(
        allocation.drawing_id,
        allocation.spid_base,
    )
    .map_err(|error| PackageError::InvalidFormat(format!("blank PPT drawing: {error}")))?;
    let mut pp_drawing = Vec::with_capacity(drawing.len() + 8);
    push_header(&mut pp_drawing, 0xF, 0, PP_DRAWING, drawing.len())?;
    pp_drawing.extend_from_slice(&drawing);

    let color_scheme = match color_scheme {
        Some(record) => record,
        None => {
            let scheme = crate::writer::spec::ColorScheme::POI_DEFAULT.to_bytes();
            let mut record = Vec::with_capacity(scheme.len() + 8);
            push_header(&mut record, 0, 1, COLOR_SCHEME_ATOM, scheme.len())?;
            record.extend_from_slice(&scheme);
            record
        },
    };

    let length = atom.len() + pp_drawing.len() + color_scheme.len();
    let mut slide = Vec::with_capacity(length + 8);
    push_header(&mut slide, 0xF, 0, SLIDE, length)?;
    slide.extend_from_slice(&atom);
    slide.extend_from_slice(&pp_drawing);
    slide.extend_from_slice(&color_scheme);
    Ok(slide)
}

/// The blank slide's `SlidePersistAtom`; the caller assigns its persist and
/// slide identifiers.
pub(super) fn slide_list_group() -> Vec<crate::records::Record> {
    vec![crate::records::Record {
        record_type: crate::RecordType::SlidePersistAtom,
        record_type_raw: crate::RecordType::SlidePersistAtom as u16,
        version: 0,
        instance: 0,
        data_length: 20,
        data: vec![0u8; 20].into(),
        children: Vec::new(),
    }]
}

/// The master and color scheme a slide record follows.
pub(super) fn master_and_scheme(record: &[u8]) -> Result<(u32, Option<Vec<u8>>)> {
    let mut master = None;
    let mut scheme = None;
    for child in children(record)? {
        let child = child?;
        match record_type(child)? {
            SLIDE_ATOM => master = Some(super::read_payload_u32(payload(child)?, 12)?),
            COLOR_SCHEME_ATOM if scheme.is_none() => scheme = Some(child.to_vec()),
            _ => {},
        }
    }
    let master = master.ok_or_else(|| corrupted("PPT slide has no SlideAtom"))?;
    Ok((master, scheme))
}

/// Registers, in the live document's `OfficeArtFDGG`, each drawing that an
/// inserted slide allocated in a new cluster.
///
/// A drawing whose identifier the document already registers, or whose shape
/// identifiers lie in an existing cluster (a transferred slide reusing a
/// cluster), is left alone. New clusters must follow the registered ones
/// contiguously. Returns the document unchanged when nothing is registered.
pub(super) fn register_inserted_drawings(
    document: &[u8],
    inserted: &BTreeMap<u32, Vec<u8>>,
) -> Result<Vec<u8>> {
    if inserted.is_empty() {
        return Ok(document.to_vec());
    }
    let mut group = drawing_group(document)?;
    let first_new = group.next_cluster();
    let mut new = BTreeMap::<u32, (u32, u32, u32)>::new();
    for record in inserted.values() {
        let Some(drawing) = inserted_drawing(record)? else {
            continue;
        };
        if group
            .clusters
            .iter()
            .any(|(dgid, _)| *dgid == drawing.drawing_id)
        {
            continue;
        }
        let mut cluster = None;
        for shape_id in &drawing.shape_ids {
            let index = shape_id / CLUSTER_SIZE;
            if cluster.is_some_and(|cluster| cluster != index) {
                return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
                    dependency: super::TransferDependency::DrawingGroup,
                }));
            }
            cluster = Some(index);
        }
        let Some(cluster) = cluster.filter(|cluster| *cluster >= first_new) else {
            continue;
        };
        let used = drawing
            .shape_ids
            .iter()
            .map(|shape_id| shape_id % CLUSTER_SIZE)
            .max()
            .unwrap_or(0)
            .saturating_add(1);
        let shapes = u32::try_from(drawing.shape_ids.len()).unwrap_or(u32::MAX);
        if new
            .insert(cluster, (drawing.drawing_id, used, shapes))
            .is_some()
        {
            return Err(corrupted("two inserted PPT drawings share one cluster"));
        }
    }
    if new.is_empty() {
        return Ok(document.to_vec());
    }
    for (expected, (cluster, (drawing_id, used, shapes))) in (first_new..).zip(new) {
        if cluster != expected {
            return Err(corrupted(
                "inserted PPT drawing clusters are not contiguous",
            ));
        }
        group.clusters.push((drawing_id, used));
        group.csp_saved = group.csp_saved.saturating_add(shapes);
        group.cdg_saved = group.cdg_saved.saturating_add(1);
        group.spid_max = group
            .spid_max
            .max(cluster.saturating_mul(CLUSTER_SIZE).saturating_add(used));
    }
    let encoded = group.encode()?;
    replace_nested(document, &[PP_DRAWING_GROUP, DGG_CONTAINER, FDGG], &encoded)
}

fn drawing_group(document: &[u8]) -> Result<DrawingGroup> {
    let fdgg = find_nested(document, &[PP_DRAWING_GROUP, DGG_CONTAINER, FDGG])?
        .ok_or_else(|| corrupted("PPT document has no OfficeArtFDGG"))?;
    DrawingGroup::parse(payload(fdgg)?)
}

/// The drawing identifier and every shape identifier of a slide record.
fn inserted_drawing(record: &[u8]) -> Result<Option<InsertedDrawing>> {
    let Some(dg_container) = find_nested(record, &[PP_DRAWING, DG_CONTAINER])? else {
        return Ok(None);
    };
    let mut drawing_id = None;
    let mut shape_ids = Vec::new();
    collect_drawing(dg_container, 0, &mut drawing_id, &mut shape_ids)?;
    let drawing_id = drawing_id.ok_or_else(|| corrupted("PPT drawing has no OfficeArtFDG"))?;
    Ok(Some(InsertedDrawing {
        drawing_id,
        shape_ids,
    }))
}

fn collect_drawing(
    record: &[u8],
    depth: usize,
    drawing_id: &mut Option<u32>,
    shape_ids: &mut Vec<u32>,
) -> Result<()> {
    if depth > 64 {
        return Err(corrupted("PPT drawing nesting exceeds the insertion limit"));
    }
    for child in children(record)? {
        let child = child?;
        match record_type(child)? {
            FDG if depth == 0 => {
                *drawing_id = Some(u32::from(u16::from_le_bytes([child[0], child[1]]) >> 4));
            },
            FSP => shape_ids.push(super::read_payload_u32(payload(child)?, 0)?),
            _ if child[0] & 0x0F == 0x0F => {
                collect_drawing(child, depth + 1, drawing_id, shape_ids)?;
            },
            _ => {},
        }
    }
    Ok(())
}

fn find_nested<'a>(record: &'a [u8], path: &[u16]) -> Result<Option<&'a [u8]>> {
    let Some((first, rest)) = path.split_first() else {
        return Ok(Some(record));
    };
    let mut found = None;
    for child in children(record)? {
        let child = child?;
        if record_type(child)? == *first {
            if found.is_some() {
                return Err(corrupted("PPT record has an ambiguous nested record"));
            }
            found = Some(child);
        }
    }
    match found {
        Some(child) => find_nested(child, rest),
        None => Ok(None),
    }
}

/// Replaces the payload of the record at `path` below `record`, rebuilding
/// every owning container's length; every other byte is copied exactly.
fn replace_nested(record: &[u8], path: &[u16], replacement: &[u8]) -> Result<Vec<u8>> {
    let Some((first, rest)) = path.split_first() else {
        return rebuild(record, replacement);
    };
    let mut data = Vec::with_capacity(payload(record)?.len().saturating_add(replacement.len()));
    let mut replaced = 0usize;
    for child in children(record)? {
        let child = child?;
        if record_type(child)? == *first {
            data.extend_from_slice(&replace_nested(child, rest, replacement)?);
            replaced += 1;
        } else {
            data.extend_from_slice(child);
        }
    }
    if replaced != 1 {
        return Err(corrupted("PPT nested record replacement is ambiguous"));
    }
    rebuild(record, &data)
}

fn record_type(record: &[u8]) -> Result<u16> {
    record
        .get(2..4)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .ok_or_else(|| corrupted("truncated PPT record header"))
}

fn payload(record: &[u8]) -> Result<&[u8]> {
    let length = super::read_payload_u32(record, 4)?;
    let end = usize::try_from(length)
        .ok()
        .and_then(|length| length.checked_add(8))
        .ok_or_else(|| corrupted("PPT record length overflows"))?;
    record
        .get(8..end)
        .ok_or_else(|| corrupted("truncated PPT record payload"))
}

fn children(record: &[u8]) -> Result<impl Iterator<Item = Result<&[u8]>>> {
    let data = payload(record)?;
    let mut offset = 0usize;
    Ok(std::iter::from_fn(move || {
        if offset == data.len() {
            return None;
        }
        let rest = &data[offset..];
        let child = payload(rest).and_then(|child_payload| {
            rest.get(..child_payload.len() + 8)
                .ok_or_else(|| corrupted("truncated PPT child record"))
        });
        match child {
            Ok(child) => {
                offset += child.len();
                Some(Ok(child))
            },
            Err(error) => {
                offset = data.len();
                Some(Err(error))
            },
        }
    }))
}

fn rebuild(record: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    let length = u32::try_from(data.len())
        .map_err(|_error| PackageError::ResourceLimit("PPT record exceeds u32".into()))?;
    let header = record
        .get(..4)
        .ok_or_else(|| corrupted("truncated PPT record header"))?;
    let mut output = Vec::with_capacity(data.len() + 8);
    output.extend_from_slice(header);
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(data);
    Ok(output)
}

fn push_header(
    output: &mut Vec<u8>,
    version: u16,
    instance: u16,
    record_type: u16,
    length: usize,
) -> Result<()> {
    let length = u32::try_from(length)
        .map_err(|_error| PackageError::ResourceLimit("PPT record exceeds u32".into()))?;
    output.extend_from_slice(&(version | (instance << 4)).to_le_bytes());
    output.extend_from_slice(&record_type.to_le_bytes());
    output.extend_from_slice(&length.to_le_bytes());
    Ok(())
}

fn corrupted(message: &str) -> Error {
    PackageError::Corrupted(message.into()).into()
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn fdgg(spid_max: u32, clusters: &[(u32, u32)]) -> Vec<u8> {
        let group = DrawingGroup {
            spid_max,
            csp_saved: 9,
            cdg_saved: u32::try_from(clusters.len()).unwrap(),
            clusters: clusters.to_vec(),
        };
        let fdgg = group.encode().unwrap();
        let mut dgg_container = Vec::new();
        push_header(&mut dgg_container, 0xF, 0, DGG_CONTAINER, fdgg.len() + 8).unwrap();
        push_header(&mut dgg_container, 0, 0, FDGG, fdgg.len()).unwrap();
        dgg_container.extend_from_slice(&fdgg);
        let mut group = Vec::new();
        push_header(&mut group, 0xF, 0, PP_DRAWING_GROUP, dgg_container.len()).unwrap();
        group.extend_from_slice(&dgg_container);
        let mut document = Vec::new();
        push_header(&mut document, 0xF, 0, 0x03E8, group.len()).unwrap();
        document.extend_from_slice(&group);
        document
    }

    #[test]
    fn allocation_follows_registered_and_inserted_drawings() {
        let document = fdgg(3074, &[(1, 3), (2, 3)]);
        let mut inserted = BTreeMap::new();
        let first = allocate(&document, &inserted).unwrap();
        assert_eq!(
            first,
            Allocation {
                drawing_id: 3,
                spid_base: 3 * 1024
            }
        );
        inserted.insert(9, slide_record(0x8000_0000, None, first).unwrap());
        let second = allocate(&document, &inserted).unwrap();
        assert_eq!(
            second,
            Allocation {
                drawing_id: 4,
                spid_base: 4 * 1024
            }
        );
        inserted.insert(10, slide_record(0x8000_0000, None, second).unwrap());

        let registered = register_inserted_drawings(&document, &inserted).unwrap();
        let group = drawing_group(&registered).unwrap();
        assert_eq!(group.clusters, [(1, 3), (2, 3), (3, 2), (4, 2)]);
        assert_eq!(group.cdg_saved, 4);
        assert_eq!(group.csp_saved, 13);
        assert_eq!(group.spid_max, 4 * 1024 + 2);
        // Registration is idempotent: the drawings are now known.
        assert_eq!(
            register_inserted_drawings(&registered, &inserted).unwrap(),
            registered
        );
    }

    #[test]
    fn blank_slide_record_follows_its_master() {
        let record = slide_record(
            0x8000_0001,
            None,
            Allocation {
                drawing_id: 7,
                spid_base: 5 * 1024,
            },
        )
        .unwrap();
        let (master, scheme) = master_and_scheme(&record).unwrap();
        assert_eq!(master, 0x8000_0001);
        assert!(scheme.is_some());
        let drawing = inserted_drawing(&record).unwrap().unwrap();
        assert_eq!(drawing.drawing_id, 7);
        assert_eq!(drawing.shape_ids, [5 * 1024, 5 * 1024 + 1]);
    }
}
