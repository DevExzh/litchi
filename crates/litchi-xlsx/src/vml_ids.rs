//! Workbook-wide allocation of legacy VML shape ids, shared by the note and
//! the OLE-object editors.
//!
//! Each VML drawing claims blocks of 1024 shape ids in its `o:idmap data`
//! (block `B` holds ids `B * 1024 + 1` to `B * 1024 + 1024`), and Excel
//! numbers the blocks in the order drawings are created, not by sheet. A
//! new shape takes the next id of its drawing's own block; a new drawing
//! takes a block past every block any drawing of the workbook claims. Either
//! way the id is one no shape of any VML drawing uses.

use std::collections::HashSet;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI};
use litchi_spreadsheet_drawing::vml;

use crate::error::Result;

/// The id a new shape takes, and the block a new drawing claims for it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Allocation {
    /// The shape id.
    pub(crate) id: u32,
    /// The `o:idmap data` block of a new drawing; the drawing's own block
    /// otherwise.
    pub(crate) block: u32,
}

fn find(haystack: &[u8], needle: &[u8], from: usize) -> Option<usize> {
    haystack
        .get(from..)?
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|at| at + from)
}

/// The value of the first `name="…"` (or `'…'`) after `from` and before the
/// next `>`.
fn quoted_after<'a>(xml: &'a [u8], name: &[u8], from: usize) -> Option<&'a [u8]> {
    let end = find(xml, b">", from)?;
    let tag = &xml[from..end];
    let at = find(tag, name, 0)? + name.len();
    let rest = &tag[at..];
    let rest = rest.strip_prefix(b"=")?;
    let quote = *rest.first()?;
    if quote != b'"' && quote != b'\'' {
        return None;
    }
    let close = rest[1..].iter().position(|byte| *byte == quote)?;
    Some(&rest[1..=close])
}

/// The shape-id blocks a VML drawing's `o:idmap data` claims: a
/// comma-separated list of block numbers.
pub(crate) fn idmap_blocks(xml: &[u8]) -> Vec<u32> {
    let Some(at) = find(xml, b"idmap", 0) else {
        return Vec::new();
    };
    let Some(value) = quoted_after(xml, b" data", at) else {
        return Vec::new();
    };
    String::from_utf8_lossy(value)
        .split(',')
        .filter_map(|block| block.trim().parse().ok())
        .collect()
}

/// The shape id `N` of a VML shape's `id="_x0000_sN"` (or `o:spid`).
pub(crate) fn shape_id(shape: &[u8]) -> Option<u32> {
    [b" o:spid".as_slice(), b" id"]
        .into_iter()
        .find_map(|name| {
            let value = quoted_after(shape, name, 0)?;
            std::str::from_utf8(value)
                .ok()?
                .trim()
                .strip_prefix("_x0000_s")?
                .parse()
                .ok()
        })
}

/// The shape ids one VML drawing's shapes use.
pub(crate) fn drawing_ids(opc: &OpcPackage, xml: &[u8]) -> Result<Vec<u32>> {
    let limits = vml::Limits::from_source_limits(opc.read_limits());
    Ok(vml::shapes(xml, &limits)?
        .into_iter()
        .filter_map(|shape| shape_id(&xml[shape.span]))
        .collect())
}

/// The first shape id of block `block`.
pub(crate) fn first_id(block: u32) -> u32 {
    block
        .checked_mul(1024)
        .and_then(|base| base.checked_add(1))
        .unwrap_or(u32::MAX)
}

/// The id a new shape of the worksheet at catalog position `sheet` takes:
/// in the VML drawing `drawing` when the worksheet has one, else in a new
/// block. `taken` holds further ids the shape must not use (an OLE
/// object's `shapeId`, a drawing shape's id).
pub(crate) fn allocate(
    opc: &OpcPackage,
    drawing: Option<&PackURI>,
    sheet: usize,
    taken: &HashSet<u32>,
) -> Result<Allocation> {
    let mut blocks = Vec::new();
    let mut used = taken.clone();
    for part in opc
        .iter_parts()
        .filter(|part| part.content_type() == ct::OFC_VML_DRAWING)
    {
        let xml = opc.get_part(part.partname())?.blob();
        blocks.extend(idmap_blocks(xml));
        used.extend(drawing_ids(opc, xml)?);
    }
    let (block, mut id) = match drawing {
        Some(name) => {
            let xml = opc.get_part(name)?.blob();
            let own = idmap_blocks(xml)
                .first()
                .copied()
                .unwrap_or_else(|| u32::try_from(sheet).unwrap_or(u32::MAX).saturating_add(1));
            let next = drawing_ids(opc, xml)?
                .into_iter()
                .filter(|id| id / 1024 == own)
                .max()
                .map_or_else(|| first_id(own), |id| id.saturating_add(1));
            (own, next)
        },
        None => {
            let block = blocks
                .iter()
                .copied()
                .max()
                .unwrap_or(0)
                .max(u32::try_from(sheet).unwrap_or(u32::MAX))
                .saturating_add(1);
            (block, first_id(block))
        },
    };
    while used.contains(&id) {
        id = id.saturating_add(1);
    }
    Ok(Allocation { id, block })
}

/// A new VML drawing claiming `block`, holding one note shape `id` for the
/// zero-based cell (`row`, `column`).
pub(crate) fn new_note_drawing(block: u32, id: u32, row: u32, column: u32) -> String {
    format!(
        "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"{block}\"/></o:shapelayout>{}{}</xml>",
        vml::SHAPE_TYPE,
        vml::note_shape(u64::from(id), row, column)
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn idmap_and_shape_ids_read_either_quote() {
        assert_eq!(
            idmap_blocks(b"<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1,3\"/>"),
            [1, 3]
        );
        assert_eq!(idmap_blocks(b"<o:idmap v:ext='edit' data='2'/>"), [2]);
        assert!(idmap_blocks(b"<xml/>").is_empty());
        assert_eq!(
            shape_id(b"<v:shape id=\"_x0000_s1027\" type=\"#_x0000_t75\">"),
            Some(1027)
        );
        assert_eq!(shape_id(b"<v:shape id='x' o:spid='_x0000_s9'>"), Some(9));
        assert_eq!(shape_id(b"<v:shape type=\"#_x0000_t75\">"), None);
    }
}
