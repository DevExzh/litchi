//! Drawing-object operations of the worksheet-feature edit: notes,
//! pictures, shapes and auto-filter drop-downs (record 0327).
//!
//! An object is added after the last top-level shape of the worksheet
//! drawing, or as the first shape of a new drawing; it is removed with every
//! record it owns. The container lengths, the drawing's shape count and the
//! drawing group's cluster table follow; every other record keeps its bytes.

use litchi_odraw::builder::{PropertyBuilder, ShapeBuilder};
use litchi_odraw::write::{Atom, Container, atom, atom_header, container_header};

use super::super::Area;
use super::super::drawing::{self, Block, Bse, Cluster, Group, Object};
use super::super::encode;
use super::super::read;
use super::super::stream::{
    self, Globals, MSO_DRAWING, NOTE, OBJ, Raw, Sheet, TXO, u16_at, u32_at,
};
use super::{Edit, Error, LimitKind, Result, refused};
use litchi_core::edit::RefusalKind;

/// A cell anchor: the top-left and bottom-right cells (the shape ends at the
/// top-left corner of `last`) and the move/resize behavior.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Anchor {
    /// The first row.
    pub first_row: u16,
    /// The first column.
    pub first_column: u16,
    /// The row the shape ends in.
    pub last_row: u16,
    /// The column the shape ends in.
    pub last_column: u16,
    /// The `AnchorBehavior`.
    pub behavior: crate::AnchorBehavior,
}

impl Anchor {
    fn bytes(self) -> [u8; 18] {
        let flags: u16 = match self.behavior {
            crate::AnchorBehavior::MoveAndResize => 0,
            crate::AnchorBehavior::MoveWithoutResize => 2,
            _ => 3,
        };
        let mut out = [0u8; 18];
        for (index, value) in [
            flags,
            self.first_column,
            0,
            self.first_row,
            0,
            self.last_column,
            0,
            self.last_row,
            0,
        ]
        .into_iter()
        .enumerate()
        {
            out[index * 2..index * 2 + 2].copy_from_slice(&value.to_le_bytes());
        }
        out
    }

    fn check(self) -> Result<()> {
        if self.first_column > 255
            || self.last_column > 255
            || self.first_row > self.last_row
            || self.first_column > self.last_column
        {
            return Err(refused(
                RefusalKind::InvalidValue,
                "the anchor is outside the BIFF8 grid",
            ));
        }
        Ok(())
    }
}

/// A preset shape geometry the legacy drawing writes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeGeometry {
    /// A rectangle (`msosptRectangle`).
    Rectangle,
    /// A rounded rectangle.
    RoundedRectangle,
    /// An ellipse.
    Ellipse,
    /// A line.
    Line,
    /// A text box.
    TextBox,
}

impl ShapeGeometry {
    const fn spt(self) -> u16 {
        match self {
            Self::Rectangle => 1,
            Self::RoundedRectangle => 2,
            Self::Ellipse => 3,
            Self::Line => 20,
            Self::TextBox => 202,
        }
    }

    /// `ftCmo.ot` of the geometry.
    const fn object_type(self) -> u16 {
        match self {
            Self::Rectangle | Self::RoundedRectangle => 2,
            Self::Ellipse => 3,
            Self::Line => 1,
            Self::TextBox => 6,
        }
    }
}

/// A shape to add.
#[derive(Debug, Clone, PartialEq)]
pub struct ShapeSpec {
    /// The geometry.
    pub geometry: ShapeGeometry,
    /// The anchor.
    pub anchor: Anchor,
    /// The text, written into a `TXO`.
    pub text: Option<String>,
    /// The fill (`0xRRGGBB`); `None` leaves the default, `Some(None)`
    /// removes it.
    pub fill: Option<Option<u32>>,
    /// The outline color (`0xRRGGBB`) and width (EMU); `Some(None)` removes
    /// it.
    pub line: Option<Option<(u32, u32)>>,
    /// The shape name.
    pub name: Option<String>,
    /// The alternative text.
    pub description: Option<String>,
    /// The rotation in degrees.
    pub rotation: f64,
    /// Horizontal flip.
    pub flip_h: bool,
    /// Vertical flip.
    pub flip_v: bool,
}

/// A picture to add.
#[derive(Debug, Clone, PartialEq)]
pub struct PictureSpec {
    /// The encoded image (PNG or JPEG).
    pub data: Vec<u8>,
    /// The anchor.
    pub anchor: Anchor,
    /// The picture name.
    pub name: Option<String>,
    /// The alternative text.
    pub description: Option<String>,
    /// The rotation in degrees.
    pub rotation: f64,
    /// Horizontal flip.
    pub flip_h: bool,
    /// Vertical flip.
    pub flip_v: bool,
}

/// Changes to an existing picture or shape; `None` keeps a value.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct DrawingChange {
    /// A new anchor.
    pub anchor: Option<Anchor>,
    /// A new name (`Some(None)` removes it).
    pub name: Option<Option<String>>,
    /// A new alternative text (`Some(None)` removes it).
    pub description: Option<Option<String>>,
    /// A new rotation in degrees.
    pub rotation: Option<f64>,
    /// New flips.
    pub flip: Option<(bool, bool)>,
}

/// A new top-level object: its shape container up to its client data (the
/// container length counts a following text box atom), its `Obj` record and,
/// for a text shape, its text box `MsoDrawing` and `TXO` records.
struct NewObject {
    sp: Vec<u8>,
    obj: Vec<u8>,
    textbox: Option<Vec<u8>>,
}

impl NewObject {
    fn span(&self) -> usize {
        self.sp.len() + if self.textbox.is_some() { 8 } else { 0 }
    }

    fn records(&self) -> Result<Vec<u8>> {
        let mut out = stream::record(MSO_DRAWING, &self.sp)?;
        out.extend_from_slice(&self.obj);
        if let Some(textbox) = &self.textbox {
            out.extend_from_slice(textbox);
        }
        Ok(out)
    }
}

/// Where an object goes in the drawing group: its drawing and shape ids and
/// whether the drawing (and its cluster) is new.
struct Placement {
    drawing_id: u32,
    shape_id: u32,
    new_drawing: bool,
}

fn obj_record(ot: u16, id: u16, extra: &[u8]) -> Result<Vec<u8>> {
    let mut payload = Vec::with_capacity(26 + extra.len());
    payload.extend_from_slice(&0x0015u16.to_le_bytes());
    payload.extend_from_slice(&0x0012u16.to_le_bytes());
    payload.extend_from_slice(&ot.to_le_bytes());
    payload.extend_from_slice(&id.to_le_bytes());
    // fLocked | fPrint | autofill | autoline.
    payload.extend_from_slice(&0x6011u16.to_le_bytes());
    payload.extend_from_slice(&[0; 12]);
    payload.extend_from_slice(extra);
    payload.extend_from_slice(&[0; 4]);
    Ok(stream::record(OBJ, &payload)?)
}

fn complex_text(text: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(text.len() * 2 + 2);
    for unit in text.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out.extend_from_slice(&[0, 0]);
    out
}

fn fixed_rotation(degrees: f64) -> Result<i32> {
    if !degrees.is_finite() || degrees.abs() > 3_600.0 {
        return Err(refused(
            RefusalKind::InvalidValue,
            "the rotation is not a finite angle",
        ));
    }
    #[allow(clippy::cast_possible_truncation, reason = "bounded above")]
    Ok((degrees * 65_536.0).round() as i32)
}

/// OfficeArtCOLORREF of `0xRRGGBB`.
const fn colorref(rgb: u32) -> u32 {
    ((rgb & 0xff) << 16) | (rgb & 0xff00) | ((rgb >> 16) & 0xff)
}

/// A shape container: FSP, FOPT, client anchor and client data; `textbox`
/// counts a following client text box atom in the container length.
fn shape_container(
    spt: u16,
    shape_id: u32,
    flip: (bool, bool),
    properties: &PropertyBuilder<'_>,
    anchor: [u8; 18],
    textbox: bool,
) -> Result<Vec<u8>> {
    let mut body = Vec::with_capacity(128);
    let mut flags = litchi_odraw::shape::Flags::HAVE_ANCHOR | litchi_odraw::shape::Flags::HAVE_SPT;
    if flip.0 {
        flags |= litchi_odraw::shape::Flags::FLIP_H;
    }
    if flip.1 {
        flags |= litchi_odraw::shape::Flags::FLIP_V;
    }
    ShapeBuilder::new(litchi_odraw::shape::Native::from_raw(spt), shape_id)
        .with_flags(flags)
        .write(&mut body, Default::default())
        .map_err(io)?;
    properties
        .write(&mut body, Default::default())
        .map_err(io)?;
    atom(
        &mut body,
        0,
        Atom::ClientAnchor,
        &anchor,
        Default::default(),
    )
    .map_err(io)?;
    atom(&mut body, 0, Atom::ClientData, &[], Default::default()).map_err(io)?;
    let length = u32::try_from(body.len() + if textbox { 8 } else { 0 })
        .map_err(|_| refused(RefusalKind::InvalidValue, "a shape container is too large"))?;
    let mut out = Vec::with_capacity(body.len() + 8);
    container_header(&mut out, 0, Container::Sp, length, Default::default()).map_err(io)?;
    out.extend(body);
    Ok(out)
}

fn io(_: std::io::Error) -> Error {
    refused(
        RefusalKind::InvalidValue,
        "an OfficeArt record cannot be written",
    )
}

fn textbox_records(text: &str) -> Result<Vec<u8>> {
    let mut atom_bytes = Vec::with_capacity(8);
    atom(
        &mut atom_bytes,
        0,
        Atom::ClientTextbox,
        &[],
        Default::default(),
    )
    .map_err(io)?;
    let mut out = stream::record(MSO_DRAWING, &atom_bytes)?;
    out.extend(
        encode::txo(text)
            .map_err(|_| refused(RefusalKind::InvalidValue, "the text is too long"))?,
    );
    Ok(out)
}

/// The template of an auto-filter drop-down: an `Obj` (`ot` 0x14, `fUIObj`)
/// with its list-box data, and the shape's FOPT, as Excel writes them.
const DROPDOWN_OBJ: [u8; 70] = [
    0x15, 0x00, 0x12, 0x00, 0x14, 0x00, 0x01, 0x00, 0x01, 0x21, 0x00, 0x00, 0x00, 0x00, 0x9c, 0x06,
    0xd8, 0x00, 0x00, 0x00, 0x00, 0x00, 0x0c, 0x00, 0x14, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x0b, 0x00, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00, 0x13, 0x00,
    0xee, 0x1f, 0x00, 0x00, 0x0b, 0x00, 0x04, 0x00, 0x01, 0x03, 0x00, 0x00, 0x02, 0x00, 0x14, 0x00,
    0x6c, 0x00, 0x00, 0x00, 0x00, 0x00,
];

impl Edit {
    fn drawing(&self, sheet: usize) -> Result<(Globals, Sheet, Option<Block>, Option<Group>)> {
        let (globals, layout) = self.sheet(sheet)?;
        let block = Block::parse(&self.stream, &layout)?;
        let group = Group::parse(&self.stream, &globals)?;
        Ok((globals, layout, block, group))
    }

    /// A free `Obj` id on the worksheet.
    fn next_object_id(&self, layout: &Sheet, block: Option<&Block>) -> Result<u16> {
        let mut max = block.map_or(0, |block| block.object_ids().max().unwrap_or(0));
        for (_, record) in layout.tail() {
            if record.kind == NOTE {
                max = max.max(u16_at(record.payload(&self.stream), 6));
            }
        }
        max.checked_add(1)
            .filter(|id| *id < 0xffff)
            .ok_or(Error::Limit {
                kind: LimitKind::Shapes,
                observed: usize::from(max) + 1,
                maximum: 0xfffe,
            })
    }

    /// The drawing and shape ids for `count` new shapes.
    fn place(block: Option<&Block>, group: Option<&Group>) -> Result<Placement> {
        let clusters = match group {
            Some(group) => group.fdgg()?.clusters,
            None => Vec::new(),
        };
        if let Some(block) = block {
            let drawing_id = u32::from(block.drawing_id);
            let shape_id = block.max_shape_id() + 1;
            let cluster = (shape_id / 1024).checked_sub(1).map(|index| index as usize);
            let owned = cluster
                .and_then(|index| clusters.get(index))
                .is_some_and(|cluster| cluster.drawing == drawing_id);
            if !owned || shape_id % 1024 == 0 {
                return Err(Error::Limit {
                    kind: LimitKind::Shapes,
                    observed: (shape_id % 1024) as usize + 1,
                    maximum: 1023,
                });
            }
            return Ok(Placement {
                drawing_id,
                shape_id,
                new_drawing: false,
            });
        }
        let drawing_id = clusters
            .iter()
            .map(|cluster| cluster.drawing)
            .max()
            .unwrap_or(0)
            + 1;
        let index = u32::try_from(clusters.len()).unwrap_or(u32::MAX);
        let patriarch = (index + 1).checked_mul(1024).ok_or(refused(
            RefusalKind::UnsupportedOperation,
            "the drawing group has too many clusters",
        ))?;
        Ok(Placement {
            drawing_id,
            shape_id: patriarch + 1,
            new_drawing: true,
        })
    }

    /// The header bytes of a new drawing that holds `shapes` shapes after its
    /// patriarch, the first of which has id `first_shape`.
    fn drawing_header(drawing_id: u32, children: usize, first_shape: u32) -> Result<Vec<u8>> {
        let patriarch_id = first_shape - 1;
        let mut patriarch = Vec::with_capacity(40);
        atom_header(&mut patriarch, 0, Atom::Spgr, 16, Default::default()).map_err(io)?;
        patriarch.extend_from_slice(&[0; 16]);
        ShapeBuilder::new(litchi_odraw::shape::Native::FREEFORM, patriarch_id)
            .with_flags(litchi_odraw::shape::Flags::GROUP | litchi_odraw::shape::Flags::PATRIARCH)
            .write(&mut patriarch, Default::default())
            .map_err(io)?;
        let spgr_length = u32::try_from(8 + patriarch.len() + children)
            .map_err(|_| refused(RefusalKind::InvalidValue, "the drawing is too large"))?;
        let mut out = Vec::with_capacity(80);
        container_header(
            &mut out,
            0,
            Container::Dg,
            16 + 8 + spgr_length,
            Default::default(),
        )
        .map_err(io)?;
        let instance = u16::try_from(drawing_id).map_err(|_| {
            refused(
                RefusalKind::UnsupportedOperation,
                "the drawing id exceeds 12 bits",
            )
        })?;
        atom_header(&mut out, instance, Atom::Dg, 8, Default::default()).map_err(io)?;
        out.extend_from_slice(&2u32.to_le_bytes());
        out.extend_from_slice(&first_shape.to_le_bytes());
        container_header(
            &mut out,
            0,
            Container::Spgr,
            spgr_length,
            Default::default(),
        )
        .map_err(io)?;
        container_header(
            &mut out,
            0,
            Container::Sp,
            u32::try_from(patriarch.len()).unwrap_or(0),
            Default::default(),
        )
        .map_err(io)?;
        out.extend(patriarch);
        Ok(out)
    }

    /// Append one object to the worksheet drawing.
    fn append_object(
        &mut self,
        sheet: usize,
        build: impl FnOnce(u32, u16) -> Result<NewObject>,
    ) -> Result<(u16, u32)> {
        let (_, layout, block, group) = self.drawing(sheet)?;
        if block.as_ref().is_some_and(|block| !block.canonical) {
            return Err(refused(
                RefusalKind::UnsupportedOperation,
                "the worksheet drawing's record layout is not rewritten",
            ));
        }
        let placement = Self::place(block.as_ref(), group.as_ref())?;
        let object_id = self.next_object_id(&layout, block.as_ref())?;
        let object = build(placement.shape_id, object_id)?;
        let span = u32::try_from(object.span())
            .map_err(|_| refused(RefusalKind::InvalidValue, "the shape is too large"))?;
        match &block {
            Some(block) => {
                // Grow the containers and the drawing's shape count in the
                // header record, then append the object's records.
                let header = layout.records[block.first];
                let base = header.start + 4;
                let dg_length = u32_at(&block.ds, 4) + span;
                let spgr_at = block.children_start_of_spgr();
                let spgr_length = u32_at(&block.ds, spgr_at + 4) + span;
                let mut patches = vec![(base + 4, dg_length), (base + spgr_at + 4, spgr_length)];
                if let Some(fdg) = block.fdg_at() {
                    patches.push((base + fdg + 8, block.csp + 1));
                    patches.push((base + fdg + 12, placement.shape_id.max(block.spid_cur)));
                }
                for (at, value) in patches {
                    self.stream[at..at + 4].copy_from_slice(&value.to_le_bytes());
                }
                let bytes = object.records()?;
                let at = block.append_at;
                self.insert(&layout, at, &bytes)?;
            },
            None => {
                let mut first =
                    Self::drawing_header(placement.drawing_id, object.span(), placement.shape_id)?;
                first.extend_from_slice(&object.sp);
                let mut bytes = stream::record(MSO_DRAWING, &first)?;
                bytes.extend_from_slice(&object.obj);
                if let Some(textbox) = &object.textbox {
                    bytes.extend_from_slice(textbox);
                }
                let at = layout.objects_insert();
                self.insert(&layout, at, &bytes)?;
            },
        }
        self.update_group(|fdgg| {
            let cluster = (placement.shape_id / 1024 - 1) as usize;
            if placement.new_drawing {
                fdgg.clusters.push(Cluster {
                    drawing: placement.drawing_id,
                    used: 2,
                });
                fdgg.drawings_saved += 1;
                fdgg.shapes_saved += 2;
            } else {
                let entry = fdgg.clusters.get_mut(cluster).ok_or(refused(
                    RefusalKind::UnsupportedOperation,
                    "the drawing group lost its cluster",
                ))?;
                entry.used = entry.used.max(placement.shape_id % 1024 + 1);
                fdgg.shapes_saved += 1;
            }
            fdgg.spid_max = fdgg.spid_max.max(placement.shape_id + 1);
            Ok(())
        })?;
        Ok((object_id, placement.shape_id))
    }

    /// Rewrite the drawing group's `OfficeArtFDGGBlock` (creating the group
    /// when the workbook has none).
    fn update_group(
        &mut self,
        change: impl FnOnce(&mut drawing::Fdgg) -> Result<()>,
    ) -> Result<()> {
        let globals = Globals::parse(&self.stream)?;
        let group = Group::parse(&self.stream, &globals)?;
        let (mut bytes, range, continuation) = match &group {
            Some(group) => (
                group.bytes.clone(),
                Some((
                    globals.records[group.first].start,
                    globals.records[group.end - 1].end,
                )),
                group.continuation,
            ),
            None => (empty_group(), None, stream::CONTINUE),
        };
        let probe = Group {
            first: 0,
            end: 0,
            continuation,
            bytes: bytes.clone(),
        };
        let mut fdgg = probe.fdgg()?;
        let before = fdgg.clusters.len();
        change(&mut fdgg)?;
        let after = fdgg.clusters.len();
        // Rewrite the FDGG block.
        let mut block = Vec::with_capacity(16 + after * 8);
        block.extend_from_slice(&fdgg.spid_max.to_le_bytes());
        block.extend_from_slice(&u32::try_from(after + 1).unwrap_or(0).to_le_bytes());
        block.extend_from_slice(&fdgg.shapes_saved.to_le_bytes());
        block.extend_from_slice(&fdgg.drawings_saved.to_le_bytes());
        for cluster in &fdgg.clusters {
            block.extend_from_slice(&cluster.drawing.to_le_bytes());
            block.extend_from_slice(&cluster.used.to_le_bytes());
        }
        let old_length = 16 + before * 8;
        bytes.splice(16..16 + old_length, block.iter().copied());
        let delta = i64::try_from(after).unwrap_or(0) - i64::try_from(before).unwrap_or(0);
        let block_length = u32::try_from(16 + after * 8).unwrap_or(0);
        bytes[12..16].copy_from_slice(&block_length.to_le_bytes());
        let container = i64::from(u32_at(&bytes, 4)) + delta * 8;
        bytes[4..8].copy_from_slice(&u32::try_from(container).unwrap_or(0).to_le_bytes());
        self.put_group(bytes, range, continuation)
    }

    /// Replace (or insert) the drawing group's records.
    fn put_group(
        &mut self,
        bytes: Vec<u8>,
        range: Option<(usize, usize)>,
        continuation: u16,
    ) -> Result<()> {
        let group = Group {
            first: 0,
            end: 0,
            continuation,
            bytes,
        };
        let encoded = group.encode()?;
        match range {
            Some((start, end)) => stream::splice(&mut self.stream, start, end, &encoded)?,
            None => {
                let globals = Globals::parse(&self.stream)?;
                let following = [
                    stream::SST,
                    stream::EXT_SST,
                    0x0801, // WebPub
                    0x080b, // WOpt
                    0x0865, // CrErr
                    0x0863, // BookExt
                    0x0867, // FeatHdr
                    0x0876, // DConn
                    0x0896, // Theme
                    0x089b, // CompressPictures
                    0x088c, // Compat12
                    0x0897, // GUIDTypeLib
                    stream::EOF,
                ];
                let at = globals
                    .records
                    .iter()
                    .find(|record| following.contains(&record.kind))
                    .map(|record| record.start)
                    .ok_or(refused(
                        RefusalKind::UnsupportedOperation,
                        "the workbook globals have no EOF",
                    ))?;
                stream::splice(&mut self.stream, at, at, &encoded)?;
            },
        }
        Ok(())
    }

    /// Remove the drawing group when it holds no cluster and no blip.
    fn drop_empty_group(&mut self) -> Result<()> {
        let globals = Globals::parse(&self.stream)?;
        let Some(group) = Group::parse(&self.stream, &globals)? else {
            return Ok(());
        };
        let fdgg = group.fdgg()?;
        let blips = group.store()?.map_or(0, |(_, entries)| entries.len());
        if fdgg.clusters.is_empty() && blips == 0 {
            let start = globals.records[group.first].start;
            let end = globals.records[group.end - 1].end;
            stream::splice(&mut self.stream, start, end, &[])?;
        }
        Ok(())
    }

    /// Remove the top-level object at `position` of the worksheet drawing.
    fn remove_object(&mut self, sheet: usize, position: usize) -> Result<Object> {
        let (_, layout, block, _) = self.drawing(sheet)?;
        let block = block.ok_or(refused(
            RefusalKind::NotFound,
            "the worksheet has no drawing",
        ))?;
        if !block.canonical {
            return Err(refused(
                RefusalKind::UnsupportedOperation,
                "the worksheet drawing's record layout is not rewritten",
            ));
        }
        let object = block.objects.get(position).cloned().ok_or(refused(
            RefusalKind::NotFound,
            "the drawing object does not exist",
        ))?;
        if object.group {
            return Err(refused(
                RefusalKind::UnsupportedOperation,
                "a group shape is not removed",
            ));
        }
        if block.objects.len() == 1 {
            // The drawing's last shape: the whole block goes, with its
            // cluster when it is the last one.
            self.replace(&layout, block.first, block.end, &[])?;
            let drawing_id = u32::from(block.drawing_id);
            let csp = block.csp;
            self.update_group(|fdgg| {
                fdgg.drawings_saved = fdgg.drawings_saved.saturating_sub(1);
                fdgg.shapes_saved = fdgg.shapes_saved.saturating_sub(csp);
                while fdgg
                    .clusters
                    .last()
                    .is_some_and(|cluster| cluster.drawing == drawing_id)
                {
                    fdgg.clusters.pop();
                }
                Ok(())
            })?;
            return Ok(object);
        }
        let span = u32::try_from(object.span.1 - object.span.0).unwrap_or(0);
        // Shrink the containers and the shape count in the header record.
        let header = layout.records[block.first];
        let mut payload = header.payload(&self.stream).to_vec();
        let spgr_at = block.children_start_of_spgr();
        let dg_length = u32_at(&block.ds, 4) - span;
        let spgr_length = u32_at(&block.ds, spgr_at + 4) - span;
        payload[4..8].copy_from_slice(&dg_length.to_le_bytes());
        payload[spgr_at + 4..spgr_at + 8].copy_from_slice(&spgr_length.to_le_bytes());
        if let Some(fdg) = block.fdg_at() {
            payload[fdg + 8..fdg + 12]
                .copy_from_slice(&(block.csp.saturating_sub(1)).to_le_bytes());
        }
        if object.shares_header {
            payload.truncate(object.span.0);
        }
        let header_bytes = stream::record(MSO_DRAWING, &payload)?;
        // Records after the header that the object owns go first.
        let (first, end) = object.records;
        if end > first {
            self.replace(&layout, first, end, &[])?;
        }
        let (_, layout) = self.sheet(sheet)?;
        self.replace(&layout, block.first, block.first + 1, &header_bytes)?;
        self.update_group(|fdgg| {
            fdgg.shapes_saved = fdgg.shapes_saved.saturating_sub(1);
            Ok(())
        })?;
        Ok(object)
    }

    /// Replace the drawing-stream bytes `[start, end)` of the object at
    /// `position` (inside one of its `MsoDrawing` records) and grow its
    /// containers by the difference.
    fn rewrite_object(
        &mut self,
        sheet: usize,
        position: usize,
        start: usize,
        end: usize,
        bytes: &[u8],
    ) -> Result<()> {
        let (_, layout, block, _) = self.drawing(sheet)?;
        let block = block.ok_or(refused(
            RefusalKind::NotFound,
            "the worksheet has no drawing",
        ))?;
        if !block.canonical {
            return Err(refused(
                RefusalKind::UnsupportedOperation,
                "the worksheet drawing's record layout is not rewritten",
            ));
        }
        let object = block.objects.get(position).ok_or(refused(
            RefusalKind::NotFound,
            "the drawing object does not exist",
        ))?;
        let delta =
            i64::try_from(bytes.len()).unwrap_or(0) - i64::try_from(end - start).unwrap_or(0);
        let grow = |value: u32| u32::try_from(i64::from(value) + delta).unwrap_or(value);
        let mut ds = block.ds.clone();
        let spgr_at = block.children_start_of_spgr();
        let patches = [
            (4, grow(u32_at(&ds, 4))),
            (spgr_at + 4, grow(u32_at(&ds, spgr_at + 4))),
            (object.span.0 + 4, grow(u32_at(&ds, object.span.0 + 4))),
        ];
        for (at, value) in patches {
            ds[at..at + 4].copy_from_slice(&value.to_le_bytes());
        }
        ds.splice(start..end, bytes.iter().copied());
        // Re-cut the MsoDrawing records from the header through the record
        // holding the change; the records between keep their bytes.
        let mut out = Vec::new();
        let mut cursor = 0usize;
        let mut last = block.first;
        for index in block.first..block.end {
            let record = layout.records[index];
            if record.kind != MSO_DRAWING {
                out.extend_from_slice(record.bytes(&self.stream));
                continue;
            }
            let length = record.end - record.start - 4;
            let source_start = cursor;
            let source_end = cursor + length;
            cursor = source_end;
            let (new_start, new_end) = if source_end <= start {
                (source_start, source_end)
            } else if source_start <= start && end <= source_end {
                (
                    source_start,
                    usize::try_from(i64::try_from(source_end).unwrap_or(0) + delta).unwrap_or(0),
                )
            } else {
                // Past the change: stop; later records keep their bytes.
                break;
            };
            out.extend(stream::record(MSO_DRAWING, &ds[new_start..new_end])?);
            last = index + 1;
            if new_end >= start {
                break;
            }
        }
        self.replace(&layout, block.first, last, &out)
    }

    // -----------------------------------------------------------------------
    // Notes

    /// The worksheet's notes with their `Note` record indexes.
    fn note_records(&self, layout: &Sheet) -> Vec<(usize, u16, u16, u16)> {
        layout
            .tail()
            .filter(|(_, record)| record.kind == NOTE)
            .map(|(index, record)| {
                let payload = record.payload(&self.stream);
                (
                    index,
                    u16_at(payload, 0),
                    u16_at(payload, 2),
                    u16_at(payload, 6),
                )
            })
            .collect()
    }

    /// Add a hidden note on a cell; returns its row-major position, or
    /// `None` when the cell already has one.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] for a cell outside the grid or an author
    /// longer than 54 characters.
    pub fn add_note(
        &mut self,
        sheet: usize,
        row: u16,
        column: u16,
        author: &str,
        text: &str,
    ) -> Result<Option<usize>> {
        self.ensure_ready()?;
        self.step()?;
        if column > 255 {
            return Err(refused(
                RefusalKind::InvalidValue,
                "the note cell is outside the BIFF8 grid",
            ));
        }
        let (_, layout) = self.sheet(sheet)?;
        let notes = self.note_records(&layout);
        if notes.iter().any(|note| note.1 == row && note.2 == column) {
            return Ok(None);
        }
        // Validate the NOTE before staging the shape.
        encode::note(row, column, false, 1, author).map_err(|_| {
            refused(
                RefusalKind::InvalidValue,
                "the note author is longer than 54 characters",
            )
        })?;
        let column_u8 = u8::try_from(column).unwrap_or(u8::MAX);
        let (object_id, _) = self.append_object(sheet, |shape_id, object_id| {
            let anchor = crate::writer::shape::Anchor::default_for_cell(row, column_u8);
            let guid = crate::writer::deterministic_comment_guid(sheet, row, column_u8, object_id);
            let (sp, obj) = crate::writer::biff::note_shape(
                row, column_u8, author, text, anchor, guid, object_id, shape_id,
            )?;
            Ok(NewObject {
                sp,
                obj,
                textbox: Some(textbox_records(text)?),
            })
        })?;
        let note = encode::note(row, column, false, object_id, author).map_err(|_| {
            refused(
                RefusalKind::InvalidValue,
                "the note author is longer than 54 characters",
            )
        })?;
        let (_, layout) = self.sheet(sheet)?;
        let at = layout.notes_insert();
        self.insert(&layout, at, &note)?;
        let mut cells: Vec<(u16, u16)> = notes.iter().map(|note| (note.1, note.2)).collect();
        cells.push((row, column));
        cells.sort_unstable();
        Ok(cells.iter().position(|cell| *cell == (row, column)))
    }

    /// The drawing position of the note on a cell and its `Note` record.
    fn note_object(
        &self,
        sheet: usize,
        row: u16,
        column: u16,
    ) -> Result<(Sheet, Block, usize, usize)> {
        let (_, layout, block, _) = self.drawing(sheet)?;
        let (note_index, _, _, object_id) = *self
            .note_records(&layout)
            .iter()
            .find(|note| note.1 == row && note.2 == column)
            .ok_or(refused(RefusalKind::NotFound, "the cell has no note"))?;
        let block = block.ok_or(refused(RefusalKind::NotFound, "the note has no shape"))?;
        let position = block
            .objects
            .iter()
            .position(|object| {
                object
                    .obj
                    .is_some_and(|(ot, id)| ot == 0x19 && id == object_id)
            })
            .ok_or(refused(RefusalKind::NotFound, "the note has no shape"))?;
        Ok((layout, block, position, note_index))
    }

    /// Change the author, text and cell of the note on a cell.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the cell has no note and
    /// [`Error::Refused`] when the target cell has one.
    pub fn set_note(
        &mut self,
        sheet: usize,
        cell: (u16, u16),
        target: (u16, u16),
        author: &str,
        text: &str,
    ) -> Result<()> {
        self.ensure_ready()?;
        self.step()?;
        if target.1 > 255 {
            return Err(refused(
                RefusalKind::InvalidValue,
                "the note cell is outside the BIFF8 grid",
            ));
        }
        if cell != target {
            let (_, layout) = self.sheet(sheet)?;
            if self
                .note_records(&layout)
                .iter()
                .any(|note| (note.1, note.2) == target)
            {
                return Err(refused(
                    RefusalKind::InvalidValue,
                    "the target cell already has a note",
                ));
            }
        }
        let (layout, block, position, note_index) = self.note_object(sheet, cell.0, cell.1)?;
        let object = &block.objects[position];
        let note_payload = layout.records[note_index].payload(&self.stream).to_vec();
        let visible = u16_at(&note_payload, 4) & 2 != 0;
        let object_id = u16_at(&note_payload, 6);
        // The NOTE record (author, cell).
        let note = encode::note(target.0, target.1, visible, object_id, author).map_err(|_| {
            refused(
                RefusalKind::InvalidValue,
                "the note author is longer than 54 characters",
            )
        })?;
        // The TXO family of the shape.
        let (first, end) = object.records;
        let txo = (first..end).find(|index| layout.records[*index].kind == TXO);
        let txo_end = txo.map(|txo| {
            let mut stop = txo + 1;
            while stop < end && layout.records[stop].kind == stream::CONTINUE {
                stop += 1;
            }
            stop
        });
        let anchor_at = object.anchor_at;
        // Replace from the back so earlier indexes stay valid.
        self.replace(&layout, note_index, note_index + 1, &note)?;
        if let (Some(txo), Some(txo_end)) = (txo, txo_end) {
            let (_, layout) = self.sheet(sheet)?;
            let bytes = encode::txo(text)
                .map_err(|_| refused(RefusalKind::InvalidValue, "the note text is too long"))?;
            self.replace(&layout, txo, txo_end, &bytes)?;
        }
        if cell != target {
            if let Some(anchor_at) = anchor_at {
                let column = u8::try_from(target.1).unwrap_or(u8::MAX);
                let anchor = crate::writer::shape::Anchor::default_for_cell(target.0, column);
                let mut bytes = [0u8; 18];
                for (index, value) in anchor.fields().into_iter().enumerate() {
                    bytes[index * 2..index * 2 + 2].copy_from_slice(&value.to_le_bytes());
                }
                self.patch_drawing(sheet, anchor_at, &bytes)?;
            }
        }
        Ok(())
    }

    /// Overwrite drawing-stream bytes in place (same length).
    fn patch_drawing(&mut self, sheet: usize, at: usize, bytes: &[u8]) -> Result<()> {
        let (_, layout, block, _) = self.drawing(sheet)?;
        let block = block.ok_or(refused(
            RefusalKind::NotFound,
            "the worksheet has no drawing",
        ))?;
        let mut cursor = 0usize;
        for index in block.first..block.end {
            let record: Raw = layout.records[index];
            if record.kind != MSO_DRAWING {
                continue;
            }
            let length = record.end - record.start - 4;
            if at >= cursor && at + bytes.len() <= cursor + length {
                let offset = record.start + 4 + at - cursor;
                self.stream[offset..offset + bytes.len()].copy_from_slice(bytes);
                return Ok(());
            }
            cursor += length;
        }
        Err(refused(
            RefusalKind::UnsupportedOperation,
            "a drawing atom crosses an MsoDrawing record",
        ))
    }

    /// Remove the note on a cell with its shape.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the cell has no note.
    pub fn remove_note(&mut self, sheet: usize, row: u16, column: u16) -> Result<()> {
        self.ensure_ready()?;
        self.step()?;
        let (layout, _, position, note_index) = self.note_object(sheet, row, column)?;
        self.replace(&layout, note_index, note_index + 1, &[])?;
        self.remove_object(sheet, position)?;
        self.drop_empty_group()
    }

    // -----------------------------------------------------------------------
    // Pictures and shapes

    /// The drawing positions of the pictures (or shapes) of a worksheet.
    fn drawing_positions(&self, sheet: usize, picture: bool) -> Result<Vec<usize>> {
        let (_, _, block, _) = self.drawing(sheet)?;
        Ok(block.map_or_else(Vec::new, |block| {
            block
                .objects
                .iter()
                .enumerate()
                .filter_map(|(position, object)| {
                    let drawing = read::drawing_of(object, &[])?;
                    (drawing.kind
                        == if picture {
                            read::DrawingKind::Picture
                        } else {
                            read::DrawingKind::Shape
                        })
                    .then_some(position)
                })
                .collect()
        }))
    }

    /// Add a shape; returns its position among the worksheet's shapes.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] for an anchor outside the grid.
    pub fn add_shape(&mut self, sheet: usize, spec: &ShapeSpec) -> Result<usize> {
        self.ensure_ready()?;
        self.step()?;
        spec.anchor.check()?;
        let count = self.drawing_positions(sheet, false)?.len();
        let mut properties = PropertyBuilder::new();
        let rotation = fixed_rotation(spec.rotation)?;
        if rotation != 0 {
            properties.add_simple(litchi_odraw::prop::Id::from(0x0004), rotation);
        }
        if spec.text.is_some() || spec.geometry == ShapeGeometry::TextBox {
            // Text anchored middle-centered as Excel draws a new shape.
            properties.add_simple(litchi_odraw::prop::Id::from(0x0080), 0);
        }
        match spec.fill {
            Some(Some(rgb)) => {
                properties.add_simple(litchi_odraw::prop::Id::from(0x0181), colorref(rgb) as i32);
                properties.add_simple(litchi_odraw::prop::Id::from(0x01bf), 0x0010_0010);
            },
            Some(None) => {
                properties.add_simple(litchi_odraw::prop::Id::from(0x01bf), 0x0010_0000);
            },
            None => {},
        }
        match spec.line {
            Some(Some((rgb, width))) => {
                properties.add_simple(litchi_odraw::prop::Id::from(0x01c0), colorref(rgb) as i32);
                properties.add_simple(
                    litchi_odraw::prop::Id::from(0x01cb),
                    i32::try_from(width).unwrap_or(i32::MAX),
                );
                properties.add_simple(litchi_odraw::prop::Id::from(0x01ff), 0x0008_0008);
            },
            Some(None) => {
                properties.add_simple(litchi_odraw::prop::Id::from(0x01ff), 0x0008_0000);
            },
            None => {},
        }
        let name = spec.name.as_deref().map(complex_text);
        let description = spec.description.as_deref().map(complex_text);
        if let Some(name) = &name {
            properties.add_complex(litchi_odraw::prop::Id::from(0x0380), name.as_slice());
        }
        if let Some(description) = &description {
            properties.add_complex(litchi_odraw::prop::Id::from(0x0381), description.as_slice());
        }
        properties.add_simple(litchi_odraw::prop::Id::from(0x03bf), 0x0008_0000);
        let anchor = spec.anchor.bytes();
        let text = spec
            .text
            .clone()
            .or_else(|| (spec.geometry == ShapeGeometry::TextBox).then(String::new));
        let geometry = spec.geometry;
        let flip = (spec.flip_h, spec.flip_v);
        self.append_object(sheet, |shape_id, object_id| {
            let sp = shape_container(
                geometry.spt(),
                shape_id,
                flip,
                &properties,
                anchor,
                text.is_some(),
            )?;
            Ok(NewObject {
                sp,
                obj: obj_record(geometry.object_type(), object_id, &[])?,
                textbox: text.as_deref().map(textbox_records).transpose()?,
            })
        })?;
        Ok(count)
    }

    /// Add a picture; returns its position among the worksheet's pictures.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] for an image that is not PNG or JPEG
    /// or an anchor outside the grid.
    pub fn add_picture(&mut self, sheet: usize, spec: &PictureSpec) -> Result<usize> {
        self.ensure_ready()?;
        self.step()?;
        spec.anchor.check()?;
        let kind = image_kind(&spec.data).ok_or(refused(
            RefusalKind::InvalidValue,
            "the image is not PNG or JPEG",
        ))?;
        let count = self.drawing_positions(sheet, true)?.len();
        let pib = self.add_blip(kind, &spec.data)?;
        let mut properties = PropertyBuilder::new();
        let rotation = fixed_rotation(spec.rotation)?;
        if rotation != 0 {
            properties.add_simple(litchi_odraw::prop::Id::from(0x0004), rotation);
        }
        properties.add_blip_id(
            litchi_odraw::prop::Id::from(0x0104),
            i32::try_from(pib).unwrap_or(i32::MAX),
        );
        let name = spec.name.as_deref().map(complex_text);
        let description = spec.description.as_deref().map(complex_text);
        if let Some(name) = &name {
            properties.add_complex(litchi_odraw::prop::Id::from(0x0380), name.as_slice());
        }
        if let Some(description) = &description {
            properties.add_complex(litchi_odraw::prop::Id::from(0x0381), description.as_slice());
        }
        properties.add_simple(litchi_odraw::prop::Id::from(0x03bf), 0x0008_0000);
        let anchor = spec.anchor.bytes();
        let flip = (spec.flip_h, spec.flip_v);
        // ftCf (clipboard format: native) and ftPioGrbit.
        let extra = [
            0x07, 0x00, 0x02, 0x00, 0xff, 0xff, 0x08, 0x00, 0x02, 0x00, 0x01, 0x00,
        ];
        self.append_object(sheet, |shape_id, object_id| {
            Ok(NewObject {
                sp: shape_container(75, shape_id, flip, &properties, anchor, false)?,
                obj: obj_record(8, object_id, &extra)?,
                textbox: None,
            })
        })?;
        Ok(count)
    }

    /// Add (or share) a blip in the drawing group's store; returns its
    /// one-based index.
    fn add_blip(&mut self, kind: litchi_odraw::image::Kind, data: &[u8]) -> Result<u32> {
        let mut store = litchi_odraw::image::builder::StoreBuilder::new();
        let blip = litchi_odraw::image::builder::BlipBuilder::bitmap(kind, data)
            .map_err(|_| refused(RefusalKind::InvalidValue, "the image cannot be stored"))?;
        store
            .add_embedded(blip)
            .map_err(|_| refused(RefusalKind::InvalidValue, "the image cannot be stored"))?;
        let mut written = Vec::new();
        store
            .write_store(&mut written, litchi_odraw::image::Limits::default())
            .map_err(io)?;
        let entry = written
            .get(8..)
            .ok_or(refused(
                RefusalKind::InvalidValue,
                "the image cannot be stored",
            ))?
            .to_vec();
        let uid: [u8; 16] = entry
            .get(10..26)
            .and_then(|bytes| bytes.try_into().ok())
            .ok_or(refused(
                RefusalKind::InvalidValue,
                "the image cannot be stored",
            ))?;
        // Ensure a drawing group exists.
        let globals = Globals::parse(&self.stream)?;
        if Group::parse(&self.stream, &globals)?.is_none() {
            self.put_group(empty_group(), None, stream::CONTINUE)?;
        }
        let globals = Globals::parse(&self.stream)?;
        let group = Group::parse(&self.stream, &globals)?.ok_or(refused(
            RefusalKind::UnsupportedOperation,
            "the drawing group is missing",
        ))?;
        let range = (
            globals.records[group.first].start,
            globals.records[group.end - 1].end,
        );
        let mut bytes = group.bytes.clone();
        let (pib, grown) = match group.store()? {
            Some((store_at, entries)) => {
                if let Some(index) = entries
                    .iter()
                    .position(|bse| bse.uid == uid && bse.win != 0)
                {
                    let refs_at = entries[index].at + 8 + 24;
                    let refs = u32_at(&bytes, refs_at) + 1;
                    bytes[refs_at..refs_at + 4].copy_from_slice(&refs.to_le_bytes());
                    (index + 1, 0)
                } else {
                    let end = store_at + 8 + u32_at(&bytes, store_at + 4) as usize;
                    bytes.splice(end..end, entry.iter().copied());
                    let count = entries.len() + 1;
                    let options = 0x000f | (u16::try_from(count).unwrap_or(0) << 4);
                    bytes[store_at..store_at + 2].copy_from_slice(&options.to_le_bytes());
                    let length =
                        u32_at(&bytes, store_at + 4) + u32::try_from(entry.len()).unwrap_or(0);
                    bytes[store_at + 4..store_at + 8].copy_from_slice(&length.to_le_bytes());
                    (count, entry.len())
                }
            },
            None => {
                let fdgg_end = 16 + u32_at(&bytes, 12) as usize;
                let mut store_bytes = Vec::with_capacity(entry.len() + 8);
                container_header(
                    &mut store_bytes,
                    1,
                    Container::BStore,
                    u32::try_from(entry.len()).unwrap_or(0),
                    Default::default(),
                )
                .map_err(io)?;
                store_bytes.extend_from_slice(&entry);
                let grown = store_bytes.len();
                bytes.splice(fdgg_end..fdgg_end, store_bytes);
                (1, grown)
            },
        };
        let container = u32_at(&bytes, 4) + u32::try_from(grown).unwrap_or(0);
        bytes[4..8].copy_from_slice(&container.to_le_bytes());
        self.put_group(bytes, Some(range), group.continuation)?;
        Ok(u32::try_from(pib).unwrap_or(u32::MAX))
    }

    /// Release one reference to blip `pib`; the last blip of the store goes
    /// when nothing references it.
    fn release_blip(&mut self, pib: u32) -> Result<()> {
        let globals = Globals::parse(&self.stream)?;
        let Some(group) = Group::parse(&self.stream, &globals)? else {
            return Ok(());
        };
        let Some((store_at, entries)) = group.store()? else {
            return Ok(());
        };
        let Some(index) = (pib as usize)
            .checked_sub(1)
            .filter(|index| *index < entries.len())
        else {
            return Ok(());
        };
        let range = (
            globals.records[group.first].start,
            globals.records[group.end - 1].end,
        );
        let mut bytes = group.bytes.clone();
        let entry: &Bse = &entries[index];
        let refs_at = entry.at + 8 + 24;
        let refs = u32_at(&bytes, refs_at).saturating_sub(1);
        if refs == 0 && index + 1 == entries.len() {
            let removed = 8 + entry.length;
            bytes.drain(entry.at..entry.at + removed);
            let removed_u32 = u32::try_from(removed).unwrap_or(0);
            if entries.len() == 1 {
                bytes.drain(store_at..store_at + 8);
                let container = u32_at(&bytes, 4) - removed_u32 - 8;
                bytes[4..8].copy_from_slice(&container.to_le_bytes());
            } else {
                let options = 0x000f | (u16::try_from(entries.len() - 1).unwrap_or(0) << 4);
                bytes[store_at..store_at + 2].copy_from_slice(&options.to_le_bytes());
                let length = u32_at(&bytes, store_at + 4) - removed_u32;
                bytes[store_at + 4..store_at + 8].copy_from_slice(&length.to_le_bytes());
                let container = u32_at(&bytes, 4) - removed_u32;
                bytes[4..8].copy_from_slice(&container.to_le_bytes());
            }
        } else {
            bytes[refs_at..refs_at + 4].copy_from_slice(&refs.to_le_bytes());
        }
        self.put_group(bytes, Some(range), group.continuation)
    }

    /// Change a picture or shape.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when there is no such drawing.
    pub fn set_drawing(
        &mut self,
        sheet: usize,
        picture: bool,
        index: usize,
        change: &DrawingChange,
    ) -> Result<()> {
        self.ensure_ready()?;
        self.step()?;
        let position = *self
            .drawing_positions(sheet, picture)?
            .get(index)
            .ok_or(refused(RefusalKind::NotFound, "the drawing does not exist"))?;
        if let Some(anchor) = change.anchor {
            anchor.check()?;
            let (_, _, block, _) = self.drawing(sheet)?;
            let at = block
                .and_then(|block| {
                    block
                        .objects
                        .get(position)
                        .and_then(|object| object.anchor_at)
                })
                .ok_or(refused(
                    RefusalKind::UnsupportedOperation,
                    "the drawing has no cell anchor",
                ))?;
            self.patch_drawing(sheet, at, &anchor.bytes())?;
        }
        if let Some((flip_h, flip_v)) = change.flip {
            let (_, _, block, _) = self.drawing(sheet)?;
            let object = block
                .and_then(|block| block.objects.get(position).cloned())
                .ok_or(refused(RefusalKind::NotFound, "the drawing does not exist"))?;
            let at = object.fsp_at.ok_or(refused(
                RefusalKind::UnsupportedOperation,
                "the drawing has no FSP",
            ))?;
            let mut flags = object.flags & !0xc0;
            if flip_h {
                flags |= 0x40;
            }
            if flip_v {
                flags |= 0x80;
            }
            self.patch_drawing(sheet, at + 4, &flags.to_le_bytes())?;
        }
        if change.name.is_some() || change.description.is_some() || change.rotation.is_some() {
            let (_, _, block, _) = self.drawing(sheet)?;
            let block = block.ok_or(refused(
                RefusalKind::NotFound,
                "the worksheet has no drawing",
            ))?;
            let object = &block.objects[position];
            let (start, end, mut props) = fopt_of(&block.ds, object)?;
            let mut put = |id: u16, value: Option<Property>| {
                props.retain(|prop| prop.0 != id);
                if let Some(value) = value {
                    props.push((id, value));
                }
            };
            if let Some(name) = &change.name {
                put(
                    0x0380,
                    name.as_deref()
                        .map(|name| Property::Complex(complex_text(name))),
                );
            }
            if let Some(description) = &change.description {
                put(
                    0x0381,
                    description
                        .as_deref()
                        .map(|text| Property::Complex(complex_text(text))),
                );
            }
            if let Some(rotation) = change.rotation {
                let value = fixed_rotation(rotation)?;
                put(
                    0x0004,
                    (value != 0).then_some(Property::Simple(value as u32)),
                );
            }
            props.sort_by_key(|prop| prop.0);
            let fopt = encode_fopt(&props)?;
            self.rewrite_object(sheet, position, start, end, &fopt)?;
        }
        Ok(())
    }

    /// Remove a picture or shape.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when there is no such drawing.
    pub fn remove_drawing(&mut self, sheet: usize, picture: bool, index: usize) -> Result<()> {
        self.ensure_ready()?;
        self.step()?;
        let position = *self
            .drawing_positions(sheet, picture)?
            .get(index)
            .ok_or(refused(RefusalKind::NotFound, "the drawing does not exist"))?;
        let object = self.remove_object(sheet, position)?;
        if picture {
            if let Some(prop) = object.prop(0x0104) {
                self.release_blip(prop.value)?;
            }
        }
        self.drop_empty_group()
    }

    // -----------------------------------------------------------------------
    // Auto-filter drop-downs

    /// Remove the auto-filter drop-downs of a worksheet (UI combo objects
    /// that no `DVal` names).
    pub(super) fn remove_dropdowns(&mut self, sheet: usize) -> Result<()> {
        loop {
            let (_, layout, block, _) = self.drawing(sheet)?;
            let Some(block) = block else {
                return Ok(());
            };
            let validation = layout
                .tail()
                .find(|(_, record)| record.kind == stream::DVAL)
                .map(|(_, record)| u32_at(record.payload(&self.stream), 10));
            let Some(position) = block.objects.iter().position(|object| {
                let Some((ot, id)) = object.obj else {
                    return false;
                };
                let records = object.records.0..object.records.1;
                let ui = records
                    .clone()
                    .find(|index| layout.records[*index].kind == OBJ)
                    .is_some_and(|index| {
                        u16_at(layout.records[index].payload(&self.stream), 8) & 0x0100 != 0
                    });
                ot == 0x14 && ui && validation != Some(u32::from(id))
            }) else {
                return Ok(());
            };
            self.remove_object(sheet, position)?;
            self.drop_empty_group()?;
        }
    }

    /// Add one drop-down per column of the filter's header row.
    pub(super) fn add_dropdowns(&mut self, sheet: usize, area: Area) -> Result<()> {
        for column in area.first_column..=area.last_column {
            let anchor = {
                let mut bytes = [0u8; 18];
                for (index, value) in [
                    1u16,
                    column,
                    0,
                    area.first_row,
                    0,
                    column + 1,
                    0,
                    area.first_row + 1,
                    0,
                ]
                .into_iter()
                .enumerate()
                {
                    bytes[index * 2..index * 2 + 2].copy_from_slice(&value.to_le_bytes());
                }
                bytes
            };
            self.append_object(sheet, |shape_id, object_id| {
                let mut properties = PropertyBuilder::new();
                properties.add_simple(litchi_odraw::prop::Id::from(0x007f), 0x0104_0104);
                properties.add_simple(litchi_odraw::prop::Id::from(0x00bf), 0x0008_0008);
                properties.add_simple(litchi_odraw::prop::Id::from(0x01ff), 0x0008_0000);
                properties.add_simple(litchi_odraw::prop::Id::from(0x03bf), 0x0002_0000);
                let mut sp =
                    shape_container(201, shape_id, (false, false), &properties, anchor, false)?;
                // Excel's drop-down FSP flags are fHaveAnchor | fHaveSpt.
                let _ = &mut sp;
                let mut obj = DROPDOWN_OBJ;
                obj[6..8].copy_from_slice(&object_id.to_le_bytes());
                Ok(NewObject {
                    sp,
                    obj: stream::record(OBJ, &obj)?,
                    textbox: None,
                })
            })?;
        }
        Ok(())
    }
}

/// A FOPT property value.
#[derive(Debug, Clone)]
enum Property {
    Simple(u32),
    /// A complex property's data (its `opid` carries `fComplex`).
    Complex(Vec<u8>),
    /// A simple property whose `opid` carries `fBid`.
    Blip(u32),
}

/// The first FOPT of an object: its drawing-stream range and properties.
fn fopt_of(ds: &[u8], object: &Object) -> Result<(usize, usize, Vec<(u16, Property)>)> {
    let mut at = object.span.0 + 8;
    while at + 8 <= object.span.1 {
        let header = drawing::header(ds, at)?;
        if header.kind == 0xf00b {
            let count = usize::from(header.instance);
            let body = at + 8;
            let mut complex_at = body + count * 6;
            let mut props = Vec::with_capacity(count);
            for index in 0..count {
                let field = body + index * 6;
                let opid = u16_at(ds, field);
                let value = u32_at(ds, field + 2);
                let id = opid & 0x3fff;
                let prop = if opid & 0x8000 != 0 {
                    let length = value as usize;
                    let data = ds
                        .get(complex_at..complex_at + length)
                        .ok_or(refused(
                            RefusalKind::UnsupportedOperation,
                            "an OfficeArt complex property is truncated",
                        ))?
                        .to_vec();
                    complex_at += length;
                    Property::Complex(data)
                } else if opid & 0x4000 != 0 {
                    Property::Blip(value)
                } else {
                    Property::Simple(value)
                };
                props.push((id, prop));
            }
            return Ok((at, at + 8 + header.length, props));
        }
        if header.kind == 0xf010 || header.kind == 0xf011 {
            break;
        }
        at += 8 + header.length;
    }
    // No FOPT: an empty one goes after the FSP.
    let fsp = object.fsp_at.ok_or(refused(
        RefusalKind::UnsupportedOperation,
        "the drawing has no FSP",
    ))? + 8;
    Ok((fsp, fsp, Vec::new()))
}

fn encode_fopt(props: &[(u16, Property)]) -> Result<Vec<u8>> {
    let mut table = Vec::with_capacity(props.len() * 6);
    let mut complex = Vec::new();
    for (id, prop) in props {
        match prop {
            Property::Simple(value) => {
                table.extend_from_slice(&id.to_le_bytes());
                table.extend_from_slice(&value.to_le_bytes());
            },
            Property::Blip(value) => {
                table.extend_from_slice(&(id | 0x4000).to_le_bytes());
                table.extend_from_slice(&value.to_le_bytes());
            },
            Property::Complex(data) => {
                table.extend_from_slice(&(id | 0x8000).to_le_bytes());
                table.extend_from_slice(&u32::try_from(data.len()).unwrap_or(0).to_le_bytes());
                complex.extend_from_slice(data);
            },
        }
    }
    let count = u16::try_from(props.len())
        .map_err(|_| refused(RefusalKind::InvalidValue, "too many shape properties"))?;
    let mut out = Vec::with_capacity(8 + table.len() + complex.len());
    out.extend_from_slice(&(0x0003 | (count << 4)).to_le_bytes());
    out.extend_from_slice(&0xf00bu16.to_le_bytes());
    out.extend_from_slice(
        &u32::try_from(table.len() + complex.len())
            .unwrap_or(0)
            .to_le_bytes(),
    );
    out.extend(table);
    out.extend(complex);
    Ok(out)
}

/// A new drawing group: an `OfficeArtFDGGBlock` without clusters and the
/// default properties Excel writes.
fn empty_group() -> Vec<u8> {
    let mut opt = Vec::new();
    // fFitTextToShape..., fill color 0x08000041, line color 0x08000040,
    // fNoFillHitTest flags, as the XLS writer's group.
    for (id, value) in [
        (0x00bfu16, 0x0008_0008u32),
        (0x0181, 0x0800_0041),
        (0x01c0, 0x0800_0040),
    ] {
        opt.extend_from_slice(&id.to_le_bytes());
        opt.extend_from_slice(&value.to_le_bytes());
    }
    let split = [0x0800_000du32, 0x0800_000c, 0x0800_0017, 0x1000_00f7];
    let mut out = Vec::with_capacity(96);
    let body_length = 8 + 16 + 8 + opt.len() + 8 + 16;
    out.extend_from_slice(&0x000fu16.to_le_bytes());
    out.extend_from_slice(&0xf000u16.to_le_bytes());
    out.extend_from_slice(&u32::try_from(body_length).unwrap_or(0).to_le_bytes());
    out.extend_from_slice(&0x0000u16.to_le_bytes());
    out.extend_from_slice(&0xf006u16.to_le_bytes());
    out.extend_from_slice(&16u32.to_le_bytes());
    out.extend_from_slice(&1024u32.to_le_bytes());
    out.extend_from_slice(&1u32.to_le_bytes());
    out.extend_from_slice(&0u32.to_le_bytes());
    out.extend_from_slice(&0u32.to_le_bytes());
    out.extend_from_slice(&(0x0003u16 | (3 << 4)).to_le_bytes());
    out.extend_from_slice(&0xf00bu16.to_le_bytes());
    out.extend_from_slice(&u32::try_from(opt.len()).unwrap_or(0).to_le_bytes());
    out.extend(opt);
    out.extend_from_slice(&(4u16 << 4).to_le_bytes());
    out.extend_from_slice(&0xf11eu16.to_le_bytes());
    out.extend_from_slice(&16u32.to_le_bytes());
    for color in split {
        out.extend_from_slice(&color.to_le_bytes());
    }
    out
}

/// The blip kind of encoded image bytes.
fn image_kind(data: &[u8]) -> Option<litchi_odraw::image::Kind> {
    use litchi_odraw::image::Kind;
    if data.starts_with(b"\x89PNG\r\n\x1a\n") {
        Some(Kind::Png)
    } else if data.starts_with(&[0xff, 0xd8, 0xff]) {
        Some(Kind::Jpeg)
    } else {
        None
    }
}
