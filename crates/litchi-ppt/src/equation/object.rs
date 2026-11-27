//! The records an equation object is made of: its OLE storage, its preview
//! picture in the drawing group's picture store, and its slide shape.

use std::io::{Cursor, Write as _};

use litchi_cfb::{OleFile, OleWriter};
use litchi_equation::MtefTarget;

use super::Bounds;
use super::read::{
    CHILD_ANCHOR, CLIENT_ANCHOR, CLIENT_DATA, EQUATION_NATIVE, EX_OBJ_REF_ATOM, FOPT, FSP, PIB,
    SP_CONTAINER,
};
use super::tree::Node;
use crate::embedded::storage::{Kind as StorageKind, Storage};
use crate::package::{Error, Result};

const COMP_OBJ: &str = "\u{1}CompObj";
const OLE: &str = "\u{1}Ole";
const OBJ_INFO: &str = "\u{3}ObjInfo";
/// Cached presentations of an OLE object (`\u{2}OlePres000`, ...).
const OLE_PRES_PREFIX: &str = "\u{2}OlePres";

/// `\u{1}Ole` of an embedded object: version 0x02000001, no flags, no moniker.
const OLE_STREAM: [u8; 20] = [
    0x01, 0x00, 0x00, 0x02, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
];
/// `\u{3}ObjInfo` as the equation servers write it: `fDefHandler` clear,
/// `fLink` clear, `fIcon` clear, `fIsOle1` clear, `fManual` clear,
/// `fRecomposeOnResize` and `fOCX`-free view object; clipboard format 3
/// (metafile).
const OBJ_INFO_STREAM: [u8; 6] = [0x00, 0x82, 0x03, 0x00, 0x00, 0x00];

/// The `CompObj` stream ([MS-OLEDS] 2.3.8) of a new equation object.
fn comp_obj(target: &MtefTarget) -> Vec<u8> {
    let mut out = Vec::with_capacity(128);
    out.extend_from_slice(&[0x01, 0x00, 0xFE, 0xFF, 0x03, 0x0A, 0x00, 0x00]);
    out.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    out.extend_from_slice(&target.clsid());
    let mut ansi = |text: &str| {
        let len = u32::try_from(text.len() + 1).unwrap_or(u32::MAX);
        out.extend_from_slice(&len.to_le_bytes());
        out.extend_from_slice(text.as_bytes());
        out.push(0);
    };
    ansi(target.user_type());
    ansi(clipboard_format(target));
    ansi(target.prog_id());
    out.extend_from_slice(&0x71B2_39F4u32.to_le_bytes());
    out.extend_from_slice(&[0; 12]);
    out
}

/// The clipboard format name the equation servers register.
fn clipboard_format(target: &MtefTarget) -> &'static str {
    if target.prog_id() == "Equation.3" {
        "DS Equation"
    } else {
        "MathType EF"
    }
}

/// The compound file of a new equation object.
pub(super) fn new_storage(target: &MtefTarget, native: &[u8]) -> Result<Vec<u8>> {
    let mut writer = OleWriter::new();
    writer.set_root_clsid(target.clsid());
    writer.create_stream(&[COMP_OBJ], &comp_obj(target))?;
    writer.create_stream(&[OLE], &OLE_STREAM)?;
    writer.create_stream(&[OBJ_INFO], &OBJ_INFO_STREAM)?;
    writer.create_stream(&[EQUATION_NATIVE], native)?;
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out)?;
    Ok(out.into_inner())
}

/// `source` (an equation object's compound file) with its `Equation Native`
/// stream replaced by `native`. Cached presentations (`OlePres*`) are
/// dropped: they would show the old equation. Every other stream, storage
/// and class identifier is copied.
pub(super) fn replace_native(source: &[u8], native: &[u8]) -> Result<Vec<u8>> {
    let mut file = OleFile::open(Cursor::new(source), litchi_cfb::OleFileLimits::default())?;
    let root = file
        .root_entry()
        .ok_or_else(|| Error::Corrupted("an equation storage has no root entry".into()))?;
    let mut writer = OleWriter::new();
    if let Some(clsid) = parse_clsid(&root.clsid) {
        writer.set_root_clsid(clsid);
    }
    writer.set_root_state_bits(root.state_bits);
    // Every entry first (the directory borrows the reader), then the
    // stream contents.
    let mut entries = Vec::new();
    collect_entries(&file, &mut Vec::new(), &mut entries)?;
    let mut wrote_native = false;
    for (path, entry_type, clsid, state_bits) in entries {
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        let top = path.len() == 1;
        let name = path.last().map_or("", String::as_str);
        match entry_type {
            Entry::Storage => {
                writer.create_storage(&refs)?;
                if let Some(clsid) = parse_clsid(&clsid) {
                    writer.set_storage_clsid(&refs, clsid)?;
                }
                writer.set_storage_state_bits(&refs, state_bits)?;
            },
            Entry::Stream if top && name == EQUATION_NATIVE => {
                writer.create_stream(&refs, native)?;
                wrote_native = true;
            },
            Entry::Stream if top && name.starts_with(OLE_PRES_PREFIX) => {},
            Entry::Stream => {
                let data = file.open_stream(&refs)?;
                writer.create_stream(&refs, &data)?;
            },
        }
    }
    if !wrote_native {
        return Err(Error::Corrupted(
            "an equation storage has no Equation Native stream".into(),
        ));
    }
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out)?;
    Ok(out.into_inner())
}

#[derive(Clone, Copy)]
enum Entry {
    Storage,
    Stream,
}

/// Every storage and stream below `path`, parents before children: its
/// path, kind, class identifier and state bits.
fn collect_entries(
    file: &OleFile<Cursor<&[u8]>>,
    path: &mut Vec<String>,
    out: &mut Vec<(Vec<String>, Entry, String, u32)>,
) -> Result<()> {
    if path.len() > 32 {
        return Err(Error::Corrupted(
            "an equation storage nests too deeply".into(),
        ));
    }
    let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
    let children = file
        .list_directory_entries(&refs)?
        .into_iter()
        .map(|entry| {
            (
                entry.name.clone(),
                entry.entry_type,
                entry.clsid.clone(),
                entry.state_bits,
            )
        })
        .collect::<Vec<_>>();
    for (name, entry_type, clsid, state_bits) in children {
        path.push(name);
        match entry_type {
            // STGTY_STORAGE
            1 => {
                out.push((path.clone(), Entry::Storage, clsid, state_bits));
                collect_entries(file, path, out)?;
            },
            // STGTY_STREAM
            2 => out.push((path.clone(), Entry::Stream, clsid, state_bits)),
            _ => {},
        }
        path.pop();
    }
    Ok(())
}

/// The 16 on-disk bytes of a CLSID written as `XXXXXXXX-XXXX-XXXX-XXXX-
/// XXXXXXXXXXXX` (the first three fields little-endian); `None` for an
/// empty or malformed one.
fn parse_clsid(text: &str) -> Option<[u8; 16]> {
    let hex: String = text.chars().filter(|c| *c != '-').collect();
    if hex.len() != 32 {
        return None;
    }
    let mut bytes = [0u8; 16];
    for (index, byte) in bytes.iter_mut().enumerate() {
        *byte = u8::from_str_radix(hex.get(index * 2..index * 2 + 2)?, 16).ok()?;
    }
    bytes[..4].reverse();
    bytes[4..6].reverse();
    bytes[6..8].reverse();
    Some(bytes)
}

/// The complete `ExOleObjStg` record holding `storage`, zlib-compressed
/// when `compress` is set, as PowerPoint writes it.
pub(super) fn storage_record(storage: Vec<u8>, compress: bool) -> Result<Vec<u8>> {
    let stored = if compress {
        let len = u32::try_from(storage.len())
            .map_err(|_error| Error::ResourceLimit("an equation storage exceeds 4 GiB".into()))?;
        let mut encoder =
            flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
        encoder
            .write_all(&storage)
            .map_err(|error| Error::Corrupted(format!("equation storage compression: {error}")))?;
        let data = encoder
            .finish()
            .map_err(|error| Error::Corrupted(format!("equation storage compression: {error}")))?;
        Storage::compressed(StorageKind::OleObject, len, data)?
    } else {
        Storage::uncompressed(StorageKind::OleObject, storage)?
    };
    stored.to_record_bytes()
}

// ---------------------------------------------------------------------------
// The picture store.

const PP_DRAWING_GROUP: u16 = 0x040B;
const DGG_CONTAINER: u16 = 0xF000;
const BSTORE_CONTAINER: u16 = 0xF001;
const FDGG: u16 = 0xF006;
const FBSE: u16 = 0xF007;
/// `msoblipPNG` (MS-ODRAW 2.4.1).
const BLIP_PNG: u8 = 6;

/// The drawing group container (`OfficeArtDggContainer`) of a document.
pub(super) fn drawing_group(document: &mut Node) -> Result<&mut Node> {
    document
        .child_mut(PP_DRAWING_GROUP)?
        .ok_or_else(|| Error::Corrupted("the PPT document has no drawing group".into()))?
        .child_mut(DGG_CONTAINER)?
        .ok_or_else(|| {
            Error::Corrupted("the PPT drawing group has no OfficeArtDggContainer".into())
        })
}

/// The picture store, created after the `OfficeArtFDGGBlock` when the
/// drawing group has none.
fn store(group: &mut Node) -> Result<&mut Node> {
    group.expand()?;
    let children = group.children_mut().ok_or_else(not_container)?;
    let index = if let Some(index) = children
        .iter()
        .position(|child| child.record_type() == BSTORE_CONTAINER)
    {
        index
    } else {
        let after = children
            .iter()
            .position(|child| child.record_type() == FDGG)
            .ok_or_else(|| Error::Corrupted("the OfficeArtDggContainer has no FDGG".into()))?;
        children.insert(after + 1, Node::container(BSTORE_CONTAINER, 0, Vec::new()));
        after + 1
    };
    let store = &mut children[index];
    store.expand()?;
    Ok(store)
}

fn not_container() -> Error {
    Error::Corrupted("a PPT drawing record is not a container".into())
}

/// The PNG BLIP record of `png` and its picture-store entry pointing at
/// `offset` in the `Pictures` stream.
fn blip_entry(png: &[u8], offset: u32) -> Result<(Vec<u8>, Node)> {
    let builder =
        litchi_odraw::image::builder::BlipBuilder::bitmap(litchi_odraw::image::Kind::Png, png)
            .map_err(|error| Error::Corrupted(format!("equation preview: {error}")))?;
    let mut blip = Vec::new();
    builder
        .write(&mut blip, litchi_odraw::image::Limits::default())
        .map_err(|error| Error::Corrupted(format!("equation preview: {error}")))?;
    let uid = litchi_odraw::image::write::digest(png);
    let size = u32::try_from(blip.len())
        .map_err(|_error| Error::ResourceLimit("an equation preview exceeds 4 GiB".into()))?;
    let mut payload = Vec::with_capacity(36);
    payload.extend_from_slice(&[BLIP_PNG, BLIP_PNG]);
    payload.extend_from_slice(uid.as_bytes());
    payload.extend_from_slice(&0x00FFu16.to_le_bytes());
    payload.extend_from_slice(&size.to_le_bytes());
    payload.extend_from_slice(&1u32.to_le_bytes());
    payload.extend_from_slice(&offset.to_le_bytes());
    payload.extend_from_slice(&[0; 4]);
    Ok((blip, Node::atom(2, u16::from(BLIP_PNG), FBSE, payload)))
}

/// Append `png` to the picture store and the `Pictures` stream; returns
/// its one-based store index.
pub(super) fn add_picture(document: &mut Node, pictures: &mut Vec<u8>, png: &[u8]) -> Result<u32> {
    let offset = u32::try_from(pictures.len())
        .map_err(|_error| Error::ResourceLimit("the PPT Pictures stream exceeds 4 GiB".into()))?;
    let (blip, entry) = blip_entry(png, offset)?;
    let store = store(drawing_group(document)?)?;
    let children = store.children_mut().ok_or_else(not_container)?;
    if children.len() >= 0x0FFF {
        return Err(Error::ResourceLimit("the PPT picture store is full".into()));
    }
    children.push(entry);
    let count = u16::try_from(children.len()).unwrap_or(0x0FFF);
    store.set_instance(count);
    pictures.extend_from_slice(&blip);
    Ok(u32::from(count))
}

/// Drop one reference to picture-store entry `index` (one-based). The last
/// reference frees the entry as PowerPoint does (an empty entry keeps the
/// other entries' indexes) and removes its BLIP from the `Pictures`
/// stream, moving the later BLIPs' offsets.
pub(super) fn release_picture(
    document: &mut Node,
    pictures: &mut Vec<u8>,
    index: u32,
) -> Result<()> {
    // A frame whose `pib` names no picture-store entry has nothing to
    // release.
    let Some(store) = drawing_group(document)?.child_mut(BSTORE_CONTAINER)? else {
        return Ok(());
    };
    store.expand()?;
    let children = store.children_mut().ok_or_else(not_container)?;
    let Some(position) = usize::try_from(index)
        .ok()
        .and_then(|index| index.checked_sub(1))
        .filter(|position| *position < children.len())
    else {
        return Ok(());
    };
    let entry = &mut children[position];
    if entry.record_type() != FBSE {
        return Ok(());
    }
    let Some(payload) = entry.payload_mut().filter(|payload| payload.len() >= 36) else {
        return Ok(());
    };
    let refs = u32::from_le_bytes([payload[24], payload[25], payload[26], payload[27]]);
    if refs > 1 {
        payload[24..28].copy_from_slice(&(refs - 1).to_le_bytes());
        return Ok(());
    }
    let size = u32::from_le_bytes([payload[20], payload[21], payload[22], payload[23]]);
    let offset = u32::from_le_bytes([payload[28], payload[29], payload[30], payload[31]]);
    let delayed = payload.len() == 36 && size > 0;
    *entry = Node::atom(2, 0, FBSE, vec![0; 36]);
    if !delayed {
        return Ok(());
    }
    let start = offset as usize;
    let end = start.saturating_add(size as usize);
    if end > pictures.len() {
        // The BLIP is not where the entry says; leave the stream alone.
        return Ok(());
    }
    pictures.drain(start..end);
    for other in children.iter_mut() {
        if other.record_type() != FBSE {
            continue;
        }
        let Some(payload) = other.payload_mut().filter(|payload| payload.len() >= 36) else {
            continue;
        };
        let at = u32::from_le_bytes([payload[28], payload[29], payload[30], payload[31]]);
        let other_size = u32::from_le_bytes([payload[20], payload[21], payload[22], payload[23]]);
        if other_size > 0 && at >= offset.saturating_add(size) {
            payload[28..32].copy_from_slice(&(at - size).to_le_bytes());
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// The slide shape.

/// `msosptPictureFrame`.
const PICTURE_FRAME: u16 = 75;
/// `fHaveSpt | fHaveAnchor | fOleShape` (MS-ODRAW 2.2.40).
const FRAME_FLAGS: u32 = 0x0000_0A10;
/// `pictureId` (MS-ODRAW 2.3.23.20): the frame's external-object id.
const PICTURE_ID: u16 = 0x010B;
/// Picture-frame properties PowerPoint writes for an equation object:
/// blip booleans, no fill, no line, no shadow.
const FRAME_PROPERTIES: [(u16, u32); 4] = [
    (0x013F, 0x0001_0000),
    (0x01BF, 0x0010_0000),
    (0x01FF, 0x0008_0000),
    (0x023F, 0x0002_0000),
];

/// The `SpContainer` of a new equation frame.
pub(super) fn frame_shape(
    shape_id: u32,
    object_id: u32,
    picture: u32,
    bounds: Bounds,
) -> Result<Node> {
    let mut fsp = Vec::with_capacity(8);
    fsp.extend_from_slice(&shape_id.to_le_bytes());
    fsp.extend_from_slice(&FRAME_FLAGS.to_le_bytes());
    let mut options = vec![(PIB | 0x4000, picture), (PICTURE_ID, object_id)];
    options.extend_from_slice(&FRAME_PROPERTIES);
    let mut fopt = Vec::with_capacity(options.len() * 6);
    for (id, value) in &options {
        fopt.extend_from_slice(&id.to_le_bytes());
        fopt.extend_from_slice(&value.to_le_bytes());
    }
    let count = u16::try_from(options.len()).unwrap_or(0);
    Ok(Node::container(
        SP_CONTAINER,
        0,
        vec![
            Node::atom(2, PICTURE_FRAME, FSP, fsp),
            Node::atom(3, count, FOPT, fopt),
            Node::atom(0, 0, CLIENT_ANCHOR, small_rect(bounds)?),
            Node::container(
                CLIENT_DATA,
                0,
                vec![Node::atom(
                    0,
                    0,
                    EX_OBJ_REF_ATOM,
                    object_id.to_le_bytes().to_vec(),
                )],
            ),
        ],
    ))
}

/// A `SmallRectStruct` (top, left, right, bottom).
fn small_rect(bounds: Bounds) -> Result<Vec<u8>> {
    let mut out = Vec::with_capacity(8);
    for value in [bounds.top, bounds.left, bounds.right, bounds.bottom] {
        let value = i16::try_from(value).map_err(|_error| {
            Error::InvalidFormat("an equation frame lies outside the slide coordinate range".into())
        })?;
        out.extend_from_slice(&value.to_le_bytes());
    }
    Ok(out)
}

/// Write `bounds` into a frame's existing anchor, in that anchor's form.
pub(super) fn set_bounds(shape: &mut Node, bounds: Bounds) -> Result<()> {
    shape.expand()?;
    if let Some(anchor) = shape.child_mut(CLIENT_ANCHOR)? {
        let payload = anchor
            .payload_mut()
            .ok_or_else(|| Error::Corrupted("a client anchor is a container".into()))?;
        match payload.len() {
            8 => *payload = small_rect(bounds)?,
            16 => {
                payload.clear();
                for value in [bounds.top, bounds.left, bounds.right, bounds.bottom] {
                    payload.extend_from_slice(&value.to_le_bytes());
                }
            },
            _ => {},
        }
        return Ok(());
    }
    if let Some(anchor) = shape.child_mut(CHILD_ANCHOR)? {
        let payload = anchor
            .payload_mut()
            .ok_or_else(|| Error::Corrupted("a child anchor is a container".into()))?;
        if payload.len() == 16 {
            payload.clear();
            for value in [bounds.left, bounds.top, bounds.right, bounds.bottom] {
                payload.extend_from_slice(&value.to_le_bytes());
            }
        }
    }
    Ok(())
}

/// Point a frame's `pib` at picture-store entry `picture`, adding the
/// property when the frame has none.
pub(super) fn set_picture(shape: &mut Node, picture: u32) -> Result<()> {
    shape.expand()?;
    let options = shape
        .child_mut(FOPT)?
        .ok_or_else(|| Error::Corrupted("an equation frame has no property table".into()))?;
    let count = usize::from(options.instance());
    let payload = options
        .payload_mut()
        .ok_or_else(|| Error::Corrupted("a property table is a container".into()))?;
    if payload.len() < count * 6 {
        return Err(Error::Corrupted("a property table is truncated".into()));
    }
    for index in 0..count {
        let at = index * 6;
        let id = u16::from_le_bytes([payload[at], payload[at + 1]]);
        if id & 0x3FFF == PIB {
            payload[at + 2..at + 6].copy_from_slice(&picture.to_le_bytes());
            return Ok(());
        }
    }
    // Properties are sorted by identifier; `pib` (0x0104) goes before the
    // first property with a larger one.
    let at = (0..count)
        .find(|index| {
            let at = index * 6;
            u16::from_le_bytes([payload[at], payload[at + 1]]) & 0x3FFF > PIB
        })
        .unwrap_or(count)
        * 6;
    let mut entry = Vec::with_capacity(6);
    entry.extend_from_slice(&(PIB | 0x4000).to_le_bytes());
    entry.extend_from_slice(&picture.to_le_bytes());
    payload.splice(at..at, entry);
    let count = u16::try_from(count + 1)
        .map_err(|_error| Error::Corrupted("a property table is full".into()))?;
    options.set_instance(count);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn clsid_text_round_trips_to_disk_order() {
        assert_eq!(
            parse_clsid("0002CE02-0000-0000-C000-000000000046"),
            Some(MtefTarget::EquationEditor3.clsid())
        );
        assert_eq!(parse_clsid(""), None);
    }

    #[test]
    fn comp_obj_matches_the_equation_editor_stream() {
        // The `CompObj` stream of an Equation Editor 3.0 object PowerPoint
        // saved (Apache POI `slideshow/42520.ppt`).
        let mut expected = vec![
            0x01, 0x00, 0xFE, 0xFF, 0x03, 0x0A, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF,
        ];
        expected.extend_from_slice(&MtefTarget::EquationEditor3.clsid());
        expected.extend_from_slice(
            b"\x17\0\0\0Microsoft Equation 3.0\0\x0c\0\0\0DS Equation\0\x0b\0\0\0Equation.3\0",
        );
        expected.extend_from_slice(&[0xF4, 0x39, 0xB2, 0x71]);
        expected.extend_from_slice(&[0; 12]);
        assert_eq!(comp_obj(&MtefTarget::EquationEditor3), expected);
    }
}
