//! Deferred resolution of Word floating picture anchors through OfficeArt.
use super::{Image, ImageError};
use litchi_odraw::{Container, Parser, Record, RecordKind};
use std::{collections::HashMap, ops::Range, sync::OnceLock};

pub(crate) struct DrawingSource {
    bytes: Result<Option<Vec<u8>>, String>,
    index: OnceLock<Result<Index, String>>,
    anchors: [OnceLock<Result<HashMap<u32, u32>, String>>; 2],
}

struct Index {
    group_end: usize,
    // None is a real non-picture shape, not a failed image lookup.
    shapes: HashMap<(bool, u32), Option<Range<usize>>>,
    // The shapes inside each top-level group, in drawing order. Anchors name
    // top-level shapes only, so every shape belongs to at most one list.
    groups: HashMap<(bool, u32), Vec<u32>>,
}

impl DrawingSource {
    pub(crate) fn new(bytes: Result<Option<Vec<u8>>, String>) -> Self {
        Self {
            bytes,
            index: OnceLock::new(),
            anchors: [OnceLock::new(), OnceLock::new()],
        }
    }
    fn bytes(&self) -> Result<&[u8], ImageError> {
        self.bytes
            .as_ref()
            .map_err(|e| ImageError::InvalidDrawing(e.clone()))?
            .as_deref()
            .ok_or_else(|| {
                ImageError::InvalidDrawing("floating anchor has no OfficeArtContent".into())
            })
    }
    fn index(&self) -> Result<&Index, ImageError> {
        let bytes = self.bytes()?;
        self.index
            .get_or_init(|| Index::parse(bytes))
            .as_ref()
            .map_err(|e| ImageError::InvalidDrawing(e.clone()))
    }
    /// Build each immutable story's CP lookup once, rather than scanning the
    /// complete Spa table for every floating run.
    pub(crate) fn at_anchor(
        &self,
        header: bool,
        cp: u32,
        hint: u32,
        anchors: &[crate::parts::spa::ShapeAnchor],
    ) -> Result<Option<Image>, ImageError> {
        match self.anchored_shape(header, cp, anchors)? {
            Some(id) => self.image(header, id, hint),
            None => Ok(None),
        }
    }
    /// Every picture the anchor at `cp` shows; see [`Self::pictures`].
    pub(crate) fn pictures_at_anchor(
        &self,
        header: bool,
        cp: u32,
        hint: u32,
        anchors: &[crate::parts::spa::ShapeAnchor],
    ) -> Result<Vec<Image>, ImageError> {
        match self.anchored_shape(header, cp, anchors)? {
            Some(id) => self.pictures(header, id, hint),
            None => Ok(Vec::new()),
        }
    }
    fn anchored_shape(
        &self,
        header: bool,
        cp: u32,
        anchors: &[crate::parts::spa::ShapeAnchor],
    ) -> Result<Option<u32>, ImageError> {
        let lookup = self.anchors[usize::from(header)]
            .get_or_init(|| {
                let mut lookup = HashMap::new();
                lookup
                    .try_reserve(anchors.len())
                    .map_err(|e| format!("floating anchor index allocation failed: {e}"))?;
                for anchor in anchors {
                    if let Some(previous) = lookup.insert(anchor.cp, anchor.spa.shape_id)
                        && previous != anchor.spa.shape_id
                    {
                        return Err(format!("ambiguous floating anchor at CP {}", anchor.cp));
                    }
                }
                Ok(lookup)
            })
            .as_ref()
            .map_err(|e| ImageError::InvalidDrawing(e.clone()))?;
        Ok(lookup.get(&cp).copied())
    }
    /// Every picture an anchored shape shows: the shape itself when it is a
    /// picture, or each picture inside it when it is a group, in drawing
    /// order.
    pub(crate) fn pictures(
        &self,
        header: bool,
        shape_id: u32,
        hint: u32,
    ) -> Result<Vec<Image>, ImageError> {
        if let Some(image) = self.image(header, shape_id, hint)? {
            return Ok(vec![image]);
        }
        let index = self.index()?;
        let Some(members) = index.groups.get(&(header, shape_id)) else {
            return Ok(Vec::new());
        };
        Ok(members
            .iter()
            .filter(|id| matches!(index.shapes.get(&(header, **id)), Some(Some(_))))
            .map(|id| Image::from_drawing(hint, header, *id))
            .collect())
    }
    pub(crate) fn image(
        &self,
        header: bool,
        shape_id: u32,
        hint: u32,
    ) -> Result<Option<Image>, ImageError> {
        let range = self
            .index()?
            .shapes
            .get(&(header, shape_id))
            .ok_or_else(|| {
                ImageError::InvalidDrawing(format!(
                    "floating anchor references missing shape {shape_id}"
                ))
            })?;
        Ok(range
            .as_ref()
            .map(|_| Image::from_drawing(hint, header, shape_id)))
    }
    pub(crate) fn data<'a>(
        &'a self,
        header: bool,
        shape_id: u32,
        word: &'a [u8],
    ) -> Result<litchi_odraw::image::File<'a>, ImageError> {
        let bytes = self.bytes()?;
        let index = self.index()?;
        let range = index
            .shapes
            .get(&(header, shape_id))
            .and_then(Option::as_ref)
            .ok_or(ImageError::NoPicture)?;
        let shape = Parser::new(&bytes[range.clone()], Default::default())
            .root()
            .map_err(ImageError::DecodeOfficeArtRecord)?
            .ok_or(ImageError::NoPicture)?;
        let props = litchi_odraw::prop::Props::from_container(&shape, Default::default())
            .map_err(ImageError::DecodeOfficeArtRecord)?;
        // A frame without pib stores no picture: a linked picture saved
        // without its data names its file through pibName only.
        let id = props
            .prop(litchi_odraw::prop::Id::BlipToDisplay)
            .ok_or(ImageError::NoPicture)?;
        let id = Some(id)
            .filter(|p| p.is_blip() && !p.is_complex() && p.raw_value() > 0)
            .ok_or_else(|| ImageError::InvalidDrawing("picture frame has no usable pib".into()))?;
        let raw_id = u32::try_from(id.raw_value())
            .map_err(|_| ImageError::InvalidDrawing("negative pib".into()))?;
        let id = litchi_odraw::image::Id::new(raw_id).map_err(ImageError::ExtractImageFailed)?;
        let store = litchi_odraw::image::store(&bytes[..index.group_end], Default::default())
            .map_err(ImageError::ExtractImageFailed)?
            .ok_or_else(|| ImageError::InvalidDrawing("picture frame has no image store".into()))?;
        litchi_odraw::image::get(&store, id, Some(word))
            .map_err(ImageError::ExtractImageFailed)?
            .ok_or(ImageError::NoPicture)
    }
}

impl Index {
    fn parse(bytes: &[u8]) -> Result<Self, String> {
        let limits = litchi_odraw::Limits::default();
        let (group, group_end) = Record::parse(bytes, 0, limits).map_err(|e| e.to_string())?;
        if group.kind() != RecordKind::DggContainer {
            return Err("OfficeArtContent has no DggContainer".into());
        }
        let mut shapes = HashMap::new();
        let mut groups = HashMap::new();
        let mut offset = group_end;
        let mut count = 0usize;
        while offset < bytes.len() {
            let label = bytes[offset];
            if label > 1 {
                return Err("invalid Word drawing story label".into());
            }
            let start = offset.checked_add(1).ok_or("drawing offset overflow")?;
            let (drawing, size) = Record::parse(bytes, start, limits).map_err(|e| e.to_string())?;
            if drawing.kind() != RecordKind::DgContainer {
                return Err("Word drawing has no DgContainer".into());
            }
            offset = start.checked_add(size).ok_or("drawing extent overflow")?;
            let records = Parser::new(&bytes[start..offset], limits)
                .shapes(limits)
                .map_err(|e| e.to_string())?;
            count = count
                .checked_add(records.len() + 1)
                .ok_or("drawing record count overflow")?;
            if u64::try_from(count).unwrap_or(u64::MAX) > u64::from(limits.max_records) {
                return Err("drawing image index exceeds record limit".into());
            }
            for record in records {
                let (id, atom) = shape_identity(&record)?;
                let range = if atom.instance() == litchi_odraw::shape::Native::PICTURE.raw() {
                    let begin = record
                        .data_offset(bytes)
                        .and_then(|v| v.checked_sub(8))
                        .ok_or("shape is outside drawing source")?;
                    let length =
                        usize::try_from(record.len()).map_err(|_| "shape length exceeds usize")?;
                    let end = begin
                        .checked_add(length)
                        .and_then(|v| v.checked_add(8))
                        .filter(|v| *v <= bytes.len())
                        .ok_or("invalid shape extent")?;
                    Some(begin..end)
                } else {
                    None
                };
                if shapes.insert((label == 1, id), range).is_some() {
                    return Err(format!("duplicate shape identity {id} in drawing story"));
                }
            }
            index_groups(&drawing, label == 1, limits, &mut groups)?;
        }
        Ok(Self {
            group_end,
            shapes,
            groups,
        })
    }
}

/// The `spid` and FSP atom of one shape container.
fn shape_identity<'data>(record: &Record<'data>) -> Result<(u32, Record<'data>), String> {
    let container = Container::try_new(record.clone()).map_err(|e| e.to_string())?;
    let atom = container
        .find(RecordKind::Sp)
        .map_err(|e| e.to_string())?
        .ok_or("shape has no FSP")?;
    let raw: [u8; 8] = atom.data().try_into().map_err(|_| "invalid FSP length")?;
    let id = u32::from_le_bytes(raw[..4].try_into().map_err(|_| "invalid shape identity")?);
    Ok((id, atom))
}

/// Record the shapes inside each top-level group of one drawing. A group is
/// an `SpgrContainer` whose first shape container describes the group itself.
fn index_groups(
    drawing: &Record<'_>,
    header: bool,
    limits: litchi_odraw::Limits,
    groups: &mut HashMap<(bool, u32), Vec<u32>>,
) -> Result<(), String> {
    let drawing = Container::try_new(drawing.clone()).map_err(|e| e.to_string())?;
    let Some(patriarch) = drawing
        .find(RecordKind::SpgrContainer)
        .map_err(|e| e.to_string())?
    else {
        return Ok(());
    };
    let patriarch = Container::try_new(patriarch).map_err(|e| e.to_string())?;
    for child in patriarch.children() {
        let child = child.map_err(|e| e.to_string())?;
        if child.kind() != RecordKind::SpgrContainer {
            continue;
        }
        let group = Container::try_new(child).map_err(|e| e.to_string())?;
        let mut shapes = group
            .find_recursive(RecordKind::SpContainer, limits)
            .map_err(|e| e.to_string())?
            .into_iter();
        let Some(own) = shapes.next() else {
            continue;
        };
        let (group_id, _) = shape_identity(&own)?;
        let members = shapes
            .map(|shape| shape_identity(&shape).map(|(id, _)| id))
            .collect::<Result<Vec<_>, _>>()?;
        groups.insert((header, group_id), members);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(kind: u16, version: u16, instance: u16, data: &[u8]) -> Vec<u8> {
        let mut bytes = ((instance << 4) | version).to_le_bytes().to_vec();
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&u32::try_from(data.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(data);
        bytes
    }
    fn shape(id: u32, kind: u16) -> Vec<u8> {
        let mut fsp = id.to_le_bytes().to_vec();
        fsp.extend_from_slice(&0xA00u32.to_le_bytes());
        let mut children = record(0xF00A, 2, kind, &fsp);
        let mut pib = 0x4104u16.to_le_bytes().to_vec();
        pib.extend_from_slice(&1u32.to_le_bytes());
        children.extend(record(0xF00B, 3, 1, &pib));
        record(0xF004, 15, 0, &children)
    }
    fn source(drawings: &[(u8, Vec<u8>)]) -> DrawingSource {
        let mut bytes = record(0xF000, 15, 0, &[]);
        for (label, shapes) in drawings {
            bytes.push(*label);
            bytes.extend(record(0xF002, 15, 0, shapes));
        }
        DrawingSource::new(Ok(Some(bytes)))
    }

    #[test]
    fn story_identity_distinguishes_picture_and_non_picture_anchors() {
        let source = source(&[(0, shape(7, 75)), (1, shape(7, 1))]);
        let image = source.image(false, 7, 0).unwrap().unwrap();
        assert_eq!(image.drawing_shape_id(), Some(7));
        assert!(source.image(true, 7, 0).unwrap().is_none());
        assert!(matches!(
            source.image(false, 8, 0),
            Err(ImageError::InvalidDrawing(_))
        ));
        assert!(matches!(
            image.data(&[0; 100], &[]),
            Err(ImageError::DrawingStoreRequired(7))
        ));
    }

    #[test]
    fn missing_store_is_an_error_instead_of_an_inline_picture_fallback() {
        let source = source(&[(0, shape(7, 75))]);
        assert!(source.image(false, 7, 42).unwrap().is_some());
        assert!(
            matches!(source.data(false,7,&[]),Err(ImageError::InvalidDrawing(message)) if message=="picture frame has no image store")
        );
    }

    #[test]
    fn group_anchor_shows_every_picture_inside_the_group() {
        let group =
            |own: u32, members: Vec<u8>| record(0xF003, 15, 0, &[shape(own, 0), members].concat());
        let inner = group(13, shape(14, 75));
        let outer = group(10, [shape(11, 75), shape(12, 1), inner].concat());
        let patriarch = record(0xF003, 15, 0, &[shape(1, 0), outer, shape(20, 75)].concat());
        let source = source(&[(0, patriarch)]);
        let ids = |images: Vec<Image>| {
            images
                .iter()
                .map(|image| image.drawing_shape_id().unwrap())
                .collect::<Vec<_>>()
        };
        assert_eq!(ids(source.pictures(false, 10, 0).unwrap()), [11, 14]);
        assert_eq!(ids(source.pictures(false, 20, 0).unwrap()), [20]);
        assert!(source.pictures(false, 12, 0).unwrap().is_empty());
        // Run::image keeps describing only a single anchored picture.
        assert!(source.image(false, 10, 0).unwrap().is_none());
    }

    #[test]
    fn linked_picture_frame_without_pib_stores_no_picture() {
        let mut fsp = 7u32.to_le_bytes().to_vec();
        fsp.extend_from_slice(&0xA00u32.to_le_bytes());
        let mut children = record(0xF00A, 2, 75, &fsp);
        // pibFlags = msoblipflagFile; the file name itself is not needed.
        let mut flags = 0x0106u16.to_le_bytes().to_vec();
        flags.extend_from_slice(&1u32.to_le_bytes());
        children.extend(record(0xF00B, 3, 1, &flags));
        let source = source(&[(0, record(0xF004, 15, 0, &children))]);
        assert!(source.image(false, 7, 0).unwrap().is_some());
        assert!(matches!(
            source.data(false, 7, &[]),
            Err(ImageError::NoPicture)
        ));
    }

    #[test]
    fn invalid_or_ambiguous_drawing_graphs_remain_errors() {
        let duplicate = [shape(7, 75), shape(7, 1)].concat();
        for source in [
            source(&[(0, duplicate)]),
            source(&[(2, shape(7, 75))]),
            DrawingSource::new(Ok(Some(vec![0; 8]))),
            DrawingSource::new(Err("range outside table".into())),
        ] {
            for _ in 0..2 {
                assert!(matches!(
                    source.image(false, 7, 0),
                    Err(ImageError::InvalidDrawing(_))
                ));
            }
        }
    }

    #[test]
    fn one_character_position_cannot_select_two_different_shapes() {
        let source = source(&[(0, shape(7, 75))]);
        let mut first = crate::parts::spa::Spa::parse(&[0; crate::parts::spa::SPA_LEN]).unwrap();
        first.shape_id = 7;
        let mut second = first;
        second.shape_id = 8;
        let anchors = [
            crate::parts::spa::ShapeAnchor { cp: 1, spa: first },
            crate::parts::spa::ShapeAnchor { cp: 1, spa: second },
        ];
        for _ in 0..2 {
            assert!(
                matches!(source.at_anchor(false,1,0,&anchors),Err(ImageError::InvalidDrawing(message)) if message=="ambiguous floating anchor at CP 1")
            );
        }
    }
}
