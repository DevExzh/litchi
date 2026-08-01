//! OfficeArt SpgrContainer payloads for writable BIFF8 shape groups.
//!
//! A shape group serializes as one SpgrContainer whose first SpContainer is the
//! group header (Spgr + Sp + OPT + ClientAnchor + ClientData) followed by one
//! SpContainer per child anchored with an OfficeArtChildAnchor (MS-ODRAW 2.2.16).
//! Each OBJ-bearing SpContainer becomes its own MsoDrawing fragment so the BIFF
//! stream can interleave the matching OBJ records.

use crate::xls::XlsResult;
use crate::xls::writer::{XlsShapeGroupChild, XlsShapeGroupWrite, XlsShapeKind};
use litchi_odraw::{
    prop::Id,
    shape::{Flags, Native},
    write::{
        Atom as WriteAtom, Container as WriteContainer, PropertyBuilder, ShapeBuilder,
        atom as write_escher_atom, child_anchor as write_child_anchor,
        container as write_container, container_header as write_container_header,
        spgr as write_spgr,
    },
};

use super::drawing::{split_client_textbox, style_properties, write_xls_anchor};

/// OfficeArtFOPT "Protection Boolean Properties" property ID (MS-ODRAW 2.3.1.5).
const PROTECTION_BOOLEAN_PROPERTIES: Id = Id::LockAgainstGrouping;
/// OfficeArtFOPT "Group Shape Boolean Properties" property ID (MS-ODRAW 2.3.4.44).
const GROUP_SHAPE_BOOLEAN_PROPERTIES_RAW: u16 = 0x03BF;
/// `fLockAgainstGrouping` asserted together with its use bit.
const LOCK_AGAINST_GROUPING_ON: i32 = 0x0004_0004;
/// `fLockAgainstGrouping` cleared while keeping its use bit.
const LOCK_AGAINST_GROUPING_OFF: i32 = 0x0004_0000;
/// `fHidden` cleared while keeping its use bit.
const GROUP_SHAPE_VISIBLE: i32 = 0x0002_0000;
/// `fHidden` asserted together with its use bit.
const GROUP_SHAPE_HIDDEN: i32 = 0x0002_0002;

/// Serialization plan for one shape group with fully assigned OBJ identifiers.
pub(crate) struct GroupShapeConfig<'a> {
    pub group: &'a XlsShapeGroupWrite,
    pub object_id: u16,
    pub child_object_ids: Vec<u16>,
}

/// One MsoDrawing fragment of a group plus the OBJ payload that must follow it.
pub(crate) struct GroupFragment<'a> {
    pub escher: Vec<u8>,
    pub has_textbox: bool,
    pub obj: GroupFragmentObj<'a>,
}

/// The OBJ record kind owed after a group fragment's MsoDrawing record.
pub(crate) enum GroupFragmentObj<'a> {
    Header {
        object_id: u16,
        locked: bool,
        visible: bool,
    },
    Child {
        child: &'a XlsShapeGroupChild,
        object_id: u16,
    },
}

/// Whether a grouped child carries an OfficeArt ClientTextbox and a TXO record.
pub(crate) fn child_has_textbox(child: &XlsShapeGroupChild) -> bool {
    child.kind == XlsShapeKind::TextBox || child.text.is_some()
}

/// Split one group into ordered MsoDrawing fragments with sequential shape IDs.
///
/// The first fragment carries the SpgrContainer header (whose declared length
/// spans every fragment) plus the group-header SpContainer; each remaining
/// fragment is one child SpContainer.
pub(crate) fn group_fragments<'a>(
    config: &'a GroupShapeConfig<'a>,
    first_shape_id: u32,
) -> XlsResult<Vec<GroupFragment<'a>>> {
    let group = config.group;
    let header = group_header_shape(group, first_shape_id)?;
    let mut child_shapes = Vec::with_capacity(group.children.len());
    for (index, child) in group.children.iter().enumerate() {
        child_shapes.push(grouped_child_shape(
            child,
            first_shape_id + 1 + index as u32,
        )?);
    }
    let total = header.len()
        + child_shapes
            .iter()
            .map(|(escher, has_textbox)| escher.len() + usize::from(*has_textbox) * 8)
            .sum::<usize>();
    let mut first = Vec::with_capacity(8 + header.len());
    write_container_header(&mut first, 0, WriteContainer::Spgr, total as u32)?;
    first.extend_from_slice(&header);

    let mut fragments = Vec::with_capacity(1 + child_shapes.len());
    fragments.push(GroupFragment {
        escher: first,
        has_textbox: false,
        obj: GroupFragmentObj::Header {
            object_id: config.object_id,
            locked: group.locked,
            visible: group.visible,
        },
    });
    for ((child, (escher, has_textbox)), &object_id) in group
        .children
        .iter()
        .zip(child_shapes)
        .zip(&config.child_object_ids)
    {
        fragments.push(GroupFragment {
            escher,
            has_textbox,
            obj: GroupFragmentObj::Child { child, object_id },
        });
    }
    Ok(fragments)
}

/// Build the group-header SpContainer (Spgr + Sp + OPT + ClientAnchor + ClientData).
fn group_header_shape(group: &XlsShapeGroupWrite, shape_id: u32) -> XlsResult<Vec<u8>> {
    let mut children = Vec::with_capacity(104);
    write_spgr(
        &mut children,
        group.coordinates.left,
        group.coordinates.top,
        group.coordinates.right,
        group.coordinates.bottom,
    )?;
    ShapeBuilder::new(Native::FREEFORM, shape_id)
        .with_flags(Flags::GROUP | Flags::HAVE_ANCHOR)
        .write(&mut children)?;
    let mut properties = PropertyBuilder::new();
    properties.add_simple(
        PROTECTION_BOOLEAN_PROPERTIES,
        if group.locked {
            LOCK_AGAINST_GROUPING_ON
        } else {
            LOCK_AGAINST_GROUPING_OFF
        },
    );
    properties.add_simple(
        Id::from(GROUP_SHAPE_BOOLEAN_PROPERTIES_RAW),
        if group.visible {
            GROUP_SHAPE_VISIBLE
        } else {
            GROUP_SHAPE_HIDDEN
        },
    );
    properties.write(&mut children)?;
    write_xls_anchor(&mut children, &group.anchor)?;
    write_escher_atom(&mut children, 0, WriteAtom::ClientData, &[])?;
    let mut out = Vec::with_capacity(children.len() + 8);
    write_container(&mut out, 0, WriteContainer::Sp, &children)?;
    Ok(out)
}

/// Build one child SpContainer anchored with an OfficeArtChildAnchor.
fn grouped_child_shape(child: &XlsShapeGroupChild, shape_id: u32) -> XlsResult<(Vec<u8>, bool)> {
    let mut children = Vec::with_capacity(112);
    ShapeBuilder::new(Native::from_raw(child.kind.officeart_type()), shape_id)
        .with_flags(Flags::CHILD | Flags::HAVE_ANCHOR | Flags::HAVE_SPT)
        .write(&mut children)?;
    style_properties(child.locked, child.fill, child.line, child.visible).write(&mut children)?;
    write_child_anchor(
        &mut children,
        child.anchor.left,
        child.anchor.top,
        child.anchor.right,
        child.anchor.bottom,
    )?;
    write_escher_atom(&mut children, 0, WriteAtom::ClientData, &[])?;
    let has_textbox = child_has_textbox(child);
    if has_textbox {
        write_escher_atom(&mut children, 0, WriteAtom::ClientTextbox, &[])?;
    }
    let mut out = Vec::with_capacity(children.len() + 8);
    write_container(&mut out, 0, WriteContainer::Sp, &children)?;
    split_client_textbox(out, has_textbox)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xls::writer::{XlsGroupRect, XlsShapeAnchor};
    use litchi_odraw::RecordKind;

    fn anchor() -> XlsShapeAnchor {
        XlsShapeAnchor {
            move_with_cells: true,
            size_with_cells: true,
            first_column: 0,
            first_column_offset: 0,
            first_row: 0,
            first_row_offset: 0,
            last_column: 6,
            last_column_offset: 0,
            last_row: 10,
            last_row_offset: 0,
        }
    }

    fn config(group: &XlsShapeGroupWrite) -> GroupShapeConfig<'_> {
        GroupShapeConfig {
            group,
            object_id: 1,
            child_object_ids: (2..2 + group.children.len() as u16).collect(),
        }
    }

    fn read_escher_header(bytes: &[u8]) -> (u8, u16, u16, u32) {
        let ver_inst = u16::from_le_bytes(bytes[0..2].try_into().unwrap());
        let rec_type = u16::from_le_bytes(bytes[2..4].try_into().unwrap());
        let length = u32::from_le_bytes(bytes[4..8].try_into().unwrap());
        ((ver_inst & 0x0F) as u8, ver_inst >> 4, rec_type, length)
    }

    #[test]
    fn spgr_container_length_spans_every_fragment() {
        let mut group = XlsShapeGroupWrite::new(anchor());
        group.coordinates = XlsGroupRect::new(0, 0, 2000, 1000);
        group.children.push(XlsShapeGroupChild::new(
            XlsShapeKind::Rectangle,
            XlsGroupRect::new(0, 0, 900, 500),
        ));
        group.children.push(XlsShapeGroupChild::new(
            XlsShapeKind::Ellipse,
            XlsGroupRect::new(900, 400, 2000, 1000),
        ));
        let config = config(&group);
        let fragments = group_fragments(&config, 1026).unwrap();

        assert_eq!(fragments.len(), 3);
        let (version, instance, rec_type, length) = read_escher_header(&fragments[0].escher);
        assert_eq!(version, 0x0F);
        assert_eq!(instance, 0);
        assert_eq!(rec_type, RecordKind::SpgrContainer.raw());
        let payload = fragments
            .iter()
            .map(|fragment| fragment.escher.len())
            .sum::<usize>()
            - 8;
        assert_eq!(length as usize, payload);
    }

    #[test]
    fn group_header_holds_spgr_rect_group_flags_and_client_anchor() {
        let mut group = XlsShapeGroupWrite::new(anchor());
        group.coordinates = XlsGroupRect::new(10, 20, 1210, 820);
        group.children.push(XlsShapeGroupChild::new(
            XlsShapeKind::Rectangle,
            XlsGroupRect::new(10, 20, 400, 300),
        ));
        let config = config(&group);
        let fragments = group_fragments(&config, 2049).unwrap();
        let header = &fragments[0].escher;

        // SpgrContainer header, then SpContainer header, then the Spgr atom.
        let (_, _, sp_container, _) = read_escher_header(&header[8..16]);
        assert_eq!(sp_container, RecordKind::SpContainer.raw());
        let (version, _, spgr, spgr_len) = read_escher_header(&header[16..24]);
        assert_eq!(version, 1);
        assert_eq!(spgr, RecordKind::Spgr.raw());
        assert_eq!(spgr_len, 16);
        let rect = header[24..40]
            .chunks(4)
            .map(|chunk| i32::from_le_bytes(chunk.try_into().unwrap()))
            .collect::<Vec<_>>();
        assert_eq!(rect, vec![10, 20, 1210, 820]);

        // Sp atom: NotPrimitive shape type, group + anchor flags, first shape ID.
        let (_, sp_instance, sp, _) = read_escher_header(&header[40..48]);
        assert_eq!(sp, RecordKind::Sp.raw());
        assert_eq!(sp_instance, Native::FREEFORM.raw());
        assert_eq!(u32::from_le_bytes(header[48..52].try_into().unwrap()), 2049);
        assert_eq!(
            u32::from_le_bytes(header[52..56].try_into().unwrap()),
            (Flags::GROUP | Flags::HAVE_ANCHOR).bits()
        );

        // The header anchors to cells (18-byte ClientAnchor), never a ChildAnchor.
        assert!(
            header
                .windows(2)
                .any(|pair| pair == RecordKind::ClientAnchor.raw().to_le_bytes())
        );
        assert!(
            header
                .windows(2)
                .all(|pair| pair != RecordKind::ChildAnchor.raw().to_le_bytes())
        );
    }

    #[test]
    fn children_use_child_anchors_child_flags_and_sequential_shape_ids() {
        let mut group = XlsShapeGroupWrite::new(anchor());
        group.coordinates = XlsGroupRect::new(0, 0, 1000, 1000);
        let mut textbox =
            XlsShapeGroupChild::new(XlsShapeKind::TextBox, XlsGroupRect::new(-20, 0, 480, 480));
        textbox.text = Some(crate::xls::writer::XlsShapeText::new("grouped"));
        group.children.push(textbox);
        let config = config(&group);
        let fragments = group_fragments(&config, 3073).unwrap();
        let child = &fragments[1].escher;

        let (_, sp_instance, sp, _) = read_escher_header(&child[8..16]);
        assert_eq!(sp, RecordKind::Sp.raw());
        assert_eq!(sp_instance, XlsShapeKind::TextBox.officeart_type());
        assert_eq!(u32::from_le_bytes(child[16..20].try_into().unwrap()), 3074);
        assert_eq!(
            u32::from_le_bytes(child[20..24].try_into().unwrap()),
            (Flags::CHILD | Flags::HAVE_ANCHOR | Flags::HAVE_SPT).bits()
        );

        let anchor_offset = child
            .windows(2)
            .position(|pair| pair == RecordKind::ChildAnchor.raw().to_le_bytes())
            .unwrap()
            - 2;
        let rect = child[anchor_offset + 8..anchor_offset + 24]
            .chunks(4)
            .map(|chunk| i32::from_le_bytes(chunk.try_into().unwrap()))
            .collect::<Vec<_>>();
        assert_eq!(rect, vec![-20, 0, 480, 480]);
        // ClientTextbox is declared inside the SpContainer length but emitted
        // as its own MsoDrawing fragment after OBJ, per [MS-XLS].
        assert!(
            child
                .windows(2)
                .all(|pair| pair != RecordKind::ClientTextbox.raw().to_le_bytes())
        );
        assert!(fragments[1].has_textbox);
        assert!(matches!(
            fragments[1].obj,
            GroupFragmentObj::Child { object_id: 2, .. }
        ));
    }
}
