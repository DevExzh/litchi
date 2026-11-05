//! The OfficeArt drawing of one worksheet (`OBJECTS`, `[MS-XLS]`
//! 2.1.7.20.6) and the workbook's drawing group (`MsoDrawingGroup`).
//!
//! A worksheet's drawing is one `OfficeArtDgContainer` whose bytes are the
//! concatenated payloads of its `MsoDrawing` records; the `Obj`, `TXO` and
//! `Continue` records (and an embedded chart substream) of each shape sit
//! between them. The model maps every top-level child of the patriarch group
//! to the BIFF records it owns, so a shape is added or removed as whole
//! records and every other byte stays.
//!
//! Editing requires the canonical layout Excel writes: the first
//! `MsoDrawing` holds the container headers, the patriarch and the start of
//! the first shape, and no other `MsoDrawing` crosses from one top-level
//! shape into the next. Other layouts are read but refused for edits.

use super::stream::{
    self, MSO_DRAWING, MSO_DRAWING_SELECTION, OBJ, Raw, Sheet, TXO, u16_at, u32_at,
};
use crate::{Error, Result};

const DG_CONTAINER: u16 = 0xf002;
const FDG: u16 = 0xf008;
const SPGR_CONTAINER: u16 = 0xf003;
const SP_CONTAINER: u16 = 0xf004;
const FSP: u16 = 0xf00a;
const FOPT: u16 = 0xf00b;
const TERTIARY_FOPT: u16 = 0xf122;
const CLIENT_ANCHOR: u16 = 0xf010;
const CLIENT_TEXTBOX: u16 = 0xf00d;
pub(super) const DGG_CONTAINER: u16 = 0xf000;
pub(super) const FDGG: u16 = 0xf006;
pub(super) const BSTORE_CONTAINER: u16 = 0xf001;
pub(super) const FBSE: u16 = 0xf007;

/// One OfficeArt record header inside a drawing stream.
#[derive(Debug, Clone, Copy)]
pub(super) struct Header {
    pub(super) version: u8,
    pub(super) instance: u16,
    pub(super) kind: u16,
    pub(super) length: usize,
}

pub(super) fn header(bytes: &[u8], at: usize) -> Result<Header> {
    if at + 8 > bytes.len() {
        return Err(Error::InvalidData(
            "an OfficeArt record header is truncated".into(),
        ));
    }
    let options = u16_at(bytes, at);
    Ok(Header {
        version: (options & 0x000f) as u8,
        instance: options >> 4,
        kind: u16_at(bytes, at + 2),
        length: u32_at(bytes, at + 4) as usize,
    })
}

/// One OfficeArt property of an `OfficeArtFOPT`.
#[derive(Debug, Clone)]
pub(super) struct Property {
    pub(super) id: u16,
    pub(super) value: u32,
    pub(super) complex: Option<Vec<u8>>,
}

fn properties(bytes: &[u8], at: usize, header: Header) -> Vec<Property> {
    let count = usize::from(header.instance);
    let body = at + 8;
    let end = (body + header.length).min(bytes.len());
    let mut out = Vec::with_capacity(count);
    let mut complex_at = body + count * 6;
    for index in 0..count {
        let field = body + index * 6;
        if field + 6 > end {
            break;
        }
        let opid = u16_at(bytes, field);
        let value = u32_at(bytes, field + 2);
        let complex = if opid & 0x8000 != 0 {
            let len = value as usize;
            let data = bytes.get(complex_at..complex_at + len).map(<[u8]>::to_vec);
            complex_at += len;
            data
        } else {
            None
        };
        out.push(Property {
            id: opid & 0x3fff,
            value,
            complex,
        });
    }
    out
}

fn utf16(bytes: &[u8]) -> String {
    let units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .take_while(|unit| *unit != 0)
        .collect();
    String::from_utf16_lossy(&units)
}

/// One top-level shape of a worksheet drawing and the records it owns.
#[derive(Debug, Clone)]
pub(super) struct Object {
    /// The shape's byte range in the drawing stream.
    pub(super) span: (usize, usize),
    /// Record indices `[first, end)` the shape owns exclusively.
    pub(super) records: (usize, usize),
    /// Whether the shape starts in the header `MsoDrawing`.
    pub(super) shares_header: bool,
    /// A group shape (`OfficeArtSpgrContainer`).
    pub(super) group: bool,
    pub(super) shape_id: u32,
    pub(super) shape_type: u16,
    pub(super) flags: u32,
    pub(super) props: Vec<Property>,
    /// The `OfficeArtClientAnchorSheet` payload.
    pub(super) anchor: Option<[u8; 18]>,
    /// Drawing-stream offset of the anchor payload.
    pub(super) anchor_at: Option<usize>,
    /// Drawing-stream offset of the `OfficeArtFSP` payload.
    pub(super) fsp_at: Option<usize>,
    /// `ftCmo.ot` and `ftCmo.id` of the shape's `Obj`.
    pub(super) obj: Option<(u16, u16)>,
    /// The text of the shape's `TXO`.
    pub(super) text: Option<String>,
}

impl Object {
    pub(super) fn prop(&self, id: u16) -> Option<&Property> {
        self.props.iter().find(|prop| prop.id == id)
    }

    pub(super) fn complex_text(&self, id: u16) -> Option<String> {
        self.prop(id)
            .and_then(|prop| prop.complex.as_deref())
            .map(utf16)
            .filter(|text| !text.is_empty())
    }
}

/// A worksheet's drawing.
#[derive(Debug, Clone)]
pub(super) struct Block {
    /// Record indices `[first, end)` of the `OBJECTS` block.
    pub(super) first: usize,
    pub(super) end: usize,
    /// The concatenated `MsoDrawing` payloads.
    pub(super) ds: Vec<u8>,
    /// The drawing (`FDG` instance), its shape count and last shape id.
    pub(super) drawing_id: u16,
    pub(super) csp: u32,
    pub(super) spid_cur: u32,
    /// The drawing-stream offset of the patriarch group's header.
    pub(super) spgr_start: usize,
    /// The drawing-stream offset of the `OfficeArtFDG`.
    pub(super) fdg: Option<usize>,
    pub(super) objects: Vec<Object>,
    /// Whether the block has the layout edits require.
    pub(super) canonical: bool,
    /// The index of the record after the last object's records (before a
    /// trailing `MsoDrawingSelection` or solver container).
    pub(super) append_at: usize,
}

fn is_block_record(kind: u16) -> bool {
    matches!(
        kind,
        MSO_DRAWING | OBJ | TXO | stream::CONTINUE | stream::BOF | MSO_DRAWING_SELECTION
    ) || kind == 0x00e9 // ImData
}

/// Decode the text of a `TXO` record and its `Continue` records.
pub(super) fn txo_text(stream_bytes: &[u8], records: &[Raw], at: usize) -> Option<String> {
    let txo = records.get(at)?.payload(stream_bytes);
    let count = usize::from(u16_at(txo, 10));
    if count == 0 {
        return Some(String::new());
    }
    let mut text = String::new();
    let mut units = 0;
    let mut index = at + 1;
    while units < count {
        let record = records.get(index)?;
        if record.kind != stream::CONTINUE {
            return None;
        }
        let payload = record.payload(stream_bytes);
        let (&high, chars) = payload.split_first()?;
        if high & 1 == 1 {
            let mut decoded = Vec::new();
            for pair in chars.chunks_exact(2) {
                if units == count {
                    break;
                }
                decoded.push(u16::from_le_bytes([pair[0], pair[1]]));
                units += 1;
            }
            text.push_str(&String::from_utf16_lossy(&decoded));
        } else {
            for byte in chars {
                if units == count {
                    break;
                }
                text.push(char::from(*byte));
                units += 1;
            }
        }
        index += 1;
    }
    Some(text)
}

impl Block {
    /// The drawing of `sheet`, when it has one.
    pub(super) fn parse(stream_bytes: &[u8], sheet: &Sheet) -> Result<Option<Self>> {
        let Some(first) = sheet
            .tail()
            .find(|(_, record)| record.kind == MSO_DRAWING)
            .map(|(index, _)| index)
        else {
            return Ok(None);
        };
        let mut end = first;
        while end < sheet.records.len() && is_block_record(sheet.records[end].kind) {
            end += 1;
        }
        // Drawing stream and record positions within it.
        let mut ds = Vec::new();
        let mut segments = Vec::new();
        let mut positions = Vec::with_capacity(end - first);
        for index in first..end {
            let record = sheet.records[index];
            positions.push(ds.len());
            if record.kind == MSO_DRAWING {
                let start = ds.len();
                ds.extend_from_slice(record.payload(stream_bytes));
                segments.push((index, start, ds.len()));
            }
        }
        let dg = header(&ds, 0)?;
        if dg.kind != DG_CONTAINER {
            return Err(Error::InvalidData(
                "a worksheet drawing does not start with OfficeArtDgContainer".into(),
            ));
        }
        let dg_end = 8 + dg.length;
        let mut at = 8;
        let mut drawing_id = 0;
        let mut csp = 0;
        let mut spid_cur = 0;
        let mut spgr: Option<(usize, usize)> = None;
        let mut fdg = None;
        while at + 8 <= dg_end.min(ds.len()) {
            let child = header(&ds, at)?;
            match child.kind {
                FDG => {
                    fdg = Some(at);
                    drawing_id = child.instance;
                    csp = u32_at(&ds, at + 8);
                    spid_cur = u32_at(&ds, at + 12);
                },
                SPGR_CONTAINER if spgr.is_none() => spgr = Some((at, at + 8 + child.length)),
                _ => {},
            }
            at += 8 + child.length;
        }
        let (spgr_start, spgr_end) = spgr.ok_or_else(|| {
            Error::InvalidData("a worksheet drawing has no patriarch group".into())
        })?;
        // The patriarch shape, then the top-level children.
        let patriarch = header(&ds, spgr_start + 8)?;
        let children_start = spgr_start + 16 + patriarch.length;
        let mut spans = Vec::new();
        let mut at = children_start;
        while at + 8 <= spgr_end.min(ds.len()) {
            let child = header(&ds, at)?;
            let child_end = at + 8 + child.length;
            spans.push((at, child_end, child.kind == SPGR_CONTAINER));
            at = child_end;
        }
        let owner_of = |offset: usize, inclusive_end: bool| {
            spans.iter().position(|(start, end, _)| {
                if inclusive_end {
                    offset > *start && offset <= *end
                } else {
                    offset >= *start && offset < *end
                }
            })
        };
        let mut canonical = spgr_end == dg_end && dg_end == ds.len();
        let mut owned: Vec<Vec<usize>> = vec![Vec::new(); spans.len()];
        let mut shares_header = vec![false; spans.len()];
        for (offset, index) in (first..end).enumerate() {
            let record = sheet.records[index];
            let position = positions[offset];
            if record.kind == MSO_DRAWING {
                let (_, start, stop) = segments
                    .iter()
                    .find(|(record_index, _, _)| *record_index == index)
                    .copied()
                    .unwrap_or((index, position, position));
                if start == 0 {
                    // The header record: it must reach into the first shape
                    // without crossing it.
                    match spans.first() {
                        Some((first_start, first_end, _))
                            if stop > *first_start && stop <= *first_end =>
                        {
                            shares_header[0] = true;
                        },
                        // A header record that ends before the first shape.
                        Some((first_start, _, _)) if stop <= *first_start => {},
                        Some(_) => canonical = false,
                        None => {},
                    }
                    continue;
                }
                match owner_of(start, false) {
                    Some(owner) if stop <= spans[owner].1 => owned[owner].push(index),
                    _ => canonical = false,
                }
            } else if record.kind == MSO_DRAWING_SELECTION {
                // Belongs to the block, not to a shape.
            } else {
                match owner_of(position, true) {
                    Some(owner) => owned[owner].push(index),
                    None => canonical = false,
                }
            }
        }
        let mut objects = Vec::with_capacity(spans.len());
        let mut append_at = first + 1;
        for (position, (start, stop, group)) in spans.iter().copied().enumerate() {
            let records = &owned[position];
            let (record_first, record_end) = match (records.first(), records.last()) {
                (Some(first_index), Some(last_index)) => (*first_index, *last_index + 1),
                _ => {
                    let at = if shares_header[position] {
                        first + 1
                    } else {
                        first
                    };
                    (at, at)
                },
            };
            if record_end - record_first != records.len()
                || (!shares_header[position] && records.is_empty())
            {
                canonical = false;
            }
            append_at = append_at.max(record_end);
            let mut object = Object {
                span: (start, stop),
                records: (record_first, record_end),
                shares_header: shares_header[position],
                group,
                shape_id: 0,
                shape_type: 0,
                flags: 0,
                props: Vec::new(),
                anchor: None,
                anchor_at: None,
                fsp_at: None,
                obj: None,
                text: None,
            };
            // The shape's atoms (a group's own shape is its first child).
            let mut at = start + 8;
            if group {
                if let Ok(first_child) = header(&ds, at) {
                    if first_child.kind == SP_CONTAINER {
                        at += 8;
                    }
                }
            }
            let shape_end = if group {
                header(&ds, start + 8).map_or(stop, |child| start + 16 + child.length)
            } else {
                stop
            };
            while at + 8 <= shape_end.min(ds.len()) {
                let atom = header(&ds, at)?;
                match atom.kind {
                    FSP => {
                        object.shape_type = atom.instance;
                        object.shape_id = u32_at(&ds, at + 8);
                        object.flags = u32_at(&ds, at + 12);
                        object.fsp_at = Some(at + 8);
                    },
                    FOPT | TERTIARY_FOPT => {
                        object.props.extend(properties(&ds, at, atom));
                    },
                    CLIENT_ANCHOR if atom.length >= 18 => {
                        let mut anchor = [0; 18];
                        anchor.copy_from_slice(&ds[at + 8..at + 26]);
                        object.anchor = Some(anchor);
                        object.anchor_at = Some(at + 8);
                    },
                    _ => {},
                }
                at += 8 + atom.length;
            }
            for &index in records {
                let record = sheet.records[index];
                if record.kind == OBJ && object.obj.is_none() {
                    let payload = record.payload(stream_bytes);
                    if u16_at(payload, 0) == 0x0015 {
                        object.obj = Some((u16_at(payload, 4), u16_at(payload, 6)));
                    }
                } else if record.kind == TXO && object.text.is_none() {
                    object.text = txo_text(stream_bytes, &sheet.records, index);
                }
            }
            objects.push(object);
        }
        if objects.is_empty() {
            append_at = first + 1;
        }
        // A trailing MsoDrawingSelection stays last.
        if sheet.records[append_at..end]
            .iter()
            .any(|record| record.kind != MSO_DRAWING_SELECTION)
        {
            canonical = false;
        }
        let _ = CLIENT_TEXTBOX;
        Ok(Some(Self {
            first,
            end,
            ds,
            drawing_id,
            csp,
            spid_cur,
            spgr_start,
            fdg,
            objects,
            canonical,
            append_at,
        }))
    }

    /// The drawing-stream offset of the patriarch group's header.
    pub(super) const fn children_start_of_spgr(&self) -> usize {
        self.spgr_start
    }

    /// The drawing-stream offset of the `OfficeArtFDG`.
    pub(super) const fn fdg_at(&self) -> Option<usize> {
        self.fdg
    }

    /// Every `Obj` id the drawing uses.
    pub(super) fn object_ids(&self) -> impl Iterator<Item = u16> + '_ {
        self.objects
            .iter()
            .filter_map(|object| object.obj.map(|(_, id)| id))
    }

    /// The largest shape id the drawing uses.
    pub(super) fn max_shape_id(&self) -> u32 {
        let mut max = self.spid_cur;
        let mut at = 0;
        // Walk every FSP in the stream (nested group children included).
        while at + 8 <= self.ds.len() {
            let Ok(record) = header(&self.ds, at) else {
                break;
            };
            if record.kind == FSP {
                max = max.max(u32_at(&self.ds, at + 8));
            }
            if record.version == 0xf {
                at += 8;
            } else {
                at += 8 + record.length;
            }
        }
        max
    }
}

/// The workbook's drawing group: its concatenated payload and the records
/// that hold it.
#[derive(Debug, Clone)]
pub(super) struct Group {
    /// Globals record indices `[first, end)`.
    pub(super) first: usize,
    pub(super) end: usize,
    /// The record type continuation records use (`Continue` or another
    /// `MsoDrawingGroup`).
    pub(super) continuation: u16,
    pub(super) bytes: Vec<u8>,
}

/// One drawing cluster (`OfficeArtIDCL`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Cluster {
    pub(super) drawing: u32,
    pub(super) used: u32,
}

/// The `OfficeArtFDGG` of a drawing group.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Fdgg {
    /// Drawing-group offset of the `OfficeArtFDGGBlock` header.
    pub(super) at: usize,
    pub(super) spid_max: u32,
    pub(super) shapes_saved: u32,
    pub(super) drawings_saved: u32,
    pub(super) clusters: Vec<Cluster>,
}

/// One `OfficeArtFBSE` of the blip store.
#[derive(Debug, Clone)]
pub(super) struct Bse {
    /// Drawing-group offset of its header.
    pub(super) at: usize,
    pub(super) length: usize,
    pub(super) win: u8,
    pub(super) uid: [u8; 16],
    pub(super) size: u32,
}

impl Group {
    pub(super) fn parse(stream_bytes: &[u8], globals: &stream::Globals) -> Result<Option<Self>> {
        let Some(first) = globals.first(stream::MSO_DRAWING_GROUP) else {
            return Ok(None);
        };
        let mut bytes = globals.records[first].payload(stream_bytes).to_vec();
        let mut end = first + 1;
        let mut continuation = stream::CONTINUE;
        while let Some(record) = globals.records.get(end) {
            if record.kind != stream::CONTINUE && record.kind != stream::MSO_DRAWING_GROUP {
                break;
            }
            continuation = record.kind;
            bytes.extend_from_slice(record.payload(stream_bytes));
            end += 1;
        }
        Ok(Some(Self {
            first,
            end,
            continuation,
            bytes,
        }))
    }

    pub(super) fn fdgg(&self) -> Result<Fdgg> {
        let container = header(&self.bytes, 0)?;
        if container.kind != DGG_CONTAINER {
            return Err(Error::InvalidData(
                "MsoDrawingGroup does not start with OfficeArtDggContainer".into(),
            ));
        }
        let block = header(&self.bytes, 8)?;
        if block.kind != FDGG || block.length < 16 {
            return Err(Error::InvalidData(
                "OfficeArtDggContainer does not start with OfficeArtFDGGBlock".into(),
            ));
        }
        let body = 16;
        let cidcl = u32_at(&self.bytes, body + 4) as usize;
        let count = cidcl.saturating_sub(1);
        if 16 + count * 8 != block.length {
            return Err(Error::InvalidData(
                "OfficeArtFDGGBlock cluster count disagrees with its length".into(),
            ));
        }
        let clusters = (0..count)
            .map(|index| Cluster {
                drawing: u32_at(&self.bytes, body + 16 + index * 8),
                used: u32_at(&self.bytes, body + 20 + index * 8),
            })
            .collect();
        Ok(Fdgg {
            at: 8,
            spid_max: u32_at(&self.bytes, body),
            shapes_saved: u32_at(&self.bytes, body + 8),
            drawings_saved: u32_at(&self.bytes, body + 12),
            clusters,
        })
    }

    /// The blip store's offset and entries, when the group has one.
    pub(super) fn store(&self) -> Result<Option<(usize, Vec<Bse>)>> {
        let container = header(&self.bytes, 0)?;
        let end = (8 + container.length).min(self.bytes.len());
        let mut at = 8;
        while at + 8 <= end {
            let child = header(&self.bytes, at)?;
            if child.kind == BSTORE_CONTAINER {
                let mut entries = Vec::new();
                let mut entry_at = at + 8;
                let store_end = (at + 8 + child.length).min(end);
                while entry_at + 8 <= store_end {
                    let entry = header(&self.bytes, entry_at)?;
                    if entry.kind == FBSE && entry.length >= 36 {
                        let body = entry_at + 8;
                        let mut uid = [0; 16];
                        uid.copy_from_slice(&self.bytes[body + 2..body + 18]);
                        entries.push(Bse {
                            at: entry_at,
                            length: entry.length,
                            win: self.bytes[body],
                            uid,
                            size: u32_at(&self.bytes, body + 20),
                        });
                    } else {
                        // A blip outside an FBSE: keep it, but count it.
                        entries.push(Bse {
                            at: entry_at,
                            length: entry.length,
                            win: 0,
                            uid: [0; 16],
                            size: 0,
                        });
                    }
                    entry_at += 8 + entry.length;
                }
                return Ok(Some((at, entries)));
            }
            at += 8 + child.length;
        }
        Ok(None)
    }

    /// The group's records, split as the source split them.
    pub(super) fn encode(&self) -> Result<Vec<u8>> {
        let mut out = Vec::with_capacity(self.bytes.len() + 16);
        let mut chunks = self.bytes.chunks(8_224);
        out.extend(stream::record(
            stream::MSO_DRAWING_GROUP,
            chunks.next().unwrap_or(&[]),
        )?);
        for chunk in chunks {
            out.extend(stream::record(self.continuation, chunk)?);
        }
        Ok(out)
    }
}

/// The MIME type of an `OfficeArtFBSE` blip type.
pub(super) const fn content_type(win: u8) -> &'static str {
    match win {
        2 => "image/x-emf",
        3 => "image/x-wmf",
        4 => "image/x-pict",
        5 | 0x12 => "image/jpeg",
        6 => "image/png",
        7 => "image/bmp",
        0x11 => "image/tiff",
        _ => "application/octet-stream",
    }
}
