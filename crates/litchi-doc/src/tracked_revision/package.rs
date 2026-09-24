//! Package-layer transactional editor for DOC tracked revisions.

use super::Limits;
use super::codec::{
    CLX, DOP, FIB_CCP_TEXT, FIB_FC_LCB, MAX_AUTHORS, MAX_REVISIONS, MAX_TEXT_UNITS, PLCFANDREF,
    PLCFATNBKF, PLCFATNBKL, PLCFBKF, PLCFBKL, PLCFBTE_CHPX, PLCFBTE_PAPX, PLCFFLD_MOM,
    ParsedMetadata, STTBFRMARK, align2, align512, append_table_block, build_papx_pages, corrupted,
    delete_piece_range, encode_revision, fib_pair, infer_moves, insert_piece, kind_order,
    merge_adjacent, metadata_from_sprms, parse_authors, parse_chpx, parse_clx, parse_cp_table,
    parse_papx, property_metadata, put_fib_pair, put_u32, read_units, replace_papx_revision_sprms,
    replace_revision_sprms, restore_before_wall, retain_sprms, revision_opcodes, serialize_authors,
    serialize_clx, slice, split_transform_chpx, split_transform_papx, strict_sprms, u16_at, u32_at,
    validate_metadata, validate_range,
};
use super::model::{CpTable, FcRun, PapxRun, RawPiece, Revision, RevisionKind, RevisionMetadata};
use crate::package::{Error as PackageError, Result};
use crate::parts::dofr::{DofrArray, DofrPatch};
use crate::parts::fib::FileInformationBlock;
use crate::parts::protection::{EditProtection, ProtectionPolicy};
use crate::parts::saved_selection::{
    SavedSelection, SavedSelectionPatch, SavedSelectionSpliceError, table_range,
};
use crate::sprm_operations::{
    SPRM_C_DTTM_RMARK, SPRM_C_DTTM_RMARK_DEL, SPRM_C_F_BOLD, SPRM_C_F_ITALIC, SPRM_C_F_OBJ,
    SPRM_C_F_OLE2, SPRM_C_F_RMARK, SPRM_C_F_RMARK_DEL, SPRM_C_F_SPEC, SPRM_C_IBST_RMARK,
    SPRM_C_IBST_RMARK_DEL, SPRM_C_IDSL_RMARK, SPRM_C_IDSL_RMARK_DEL, SPRM_C_KUL,
    SPRM_C_PIC_LOCATION, SPRM_C_PROP_RMARK_CURRENT, SPRM_C_PROP_RMARK90, SPRM_C_RSID_PROP,
    SPRM_C_RSID_RM_DEL, SPRM_C_RSID_TEXT, SPRM_C_WALL, SPRM_P_PROP_RMARK,
    SPRM_P_PROP_RMARK_CURRENT, SPRM_P_PROP_RMARK90, SPRM_P_WALL, SPRM_T_PROP_RMARK, SPRM_T_RSID,
    SPRM_T_WALL,
};
use crate::writer::ChpxFkpBuilder;
use litchi_cfb::{OleFile, directory_names_equal};
#[cfg(feature = "performance-diagnostics")]
use litchi_ole_common::object::CfbParseEvent;
use litchi_ole_common::object::{Editor as ObjectEditor, Targets};
use litchi_ole_common::property_set::{
    DOCUMENT_SUMMARY_INFORMATION_FMTID, Stream as PropertySetStream,
    document_summary::DIGITAL_SIGNATURE,
};
use litchi_ole_common::vba_signature;
use smallvec::SmallVec;
use std::collections::HashSet;
use std::io::Cursor;
use std::sync::Arc;

/// Exact, deliberately narrow picture dependency closure shared with the
/// ordinary body transaction. The wire blocks are accepted only after they
/// match the crate's canonical singleton writer graph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct PictureGraph {
    pub(crate) floating: bool,
    pub(crate) picture_block: Vec<u8>,
    pub(crate) spa: Option<crate::parts::spa::Spa>,
    pub(crate) dgg_info: Vec<u8>,
    /// Receiver character formatting displaced by installation. Present only
    /// on an installed graph and carried by the reversible durable state.
    pub(crate) replaced_grpprl: Option<Vec<u8>>,
    /// Exact receiver Data-stream append offset. Present only after install.
    pub(crate) data_offset: Option<u32>,
}

impl PictureGraph {
    pub(crate) fn same_wire_graph(&self, other: &Self) -> bool {
        self.floating == other.floating
            && self.picture_block == other.picture_block
            && self.spa == other.spa
            && self.dgg_info == other.dgg_info
    }

    #[deny(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        clippy::cast_precision_loss,
        clippy::cast_sign_loss,
        clippy::expect_used,
        clippy::let_underscore_must_use,
        clippy::map_err_ignore,
        clippy::unwrap_used,
        reason = "durable picture graph validation is a checked legacy-codec boundary"
    )]
    pub(crate) fn validate_rehomed(&self) -> Result<()> {
        let (picture, width, height, shape_id) = canonical_picture_from_block(&self.picture_block)?;
        match self.spa {
            None if !self.floating && self.dgg_info.is_empty() => Ok(()),
            Some(spa) if self.floating && !self.dgg_info.is_empty() => {
                if spa.shape_id != shape_id
                    || spa.width()
                        != i32::try_from(width).map_err(|error| {
                            corrupted(format!("picture width exceeds i32: {error}"))
                        })?
                    || spa.height()
                        != i32::try_from(height).map_err(|error| {
                            corrupted(format!("picture height exceeds i32: {error}"))
                        })?
                {
                    return Err(corrupted("re-homed SPA does not match its PICF"));
                }
                let position = floating_position_from_spa(spa);
                let shape = crate::writer::images::FloatingShapeInfo {
                    anchor_cp: 0,
                    shape_id: spa.shape_id,
                    content: crate::writer::images::FloatingShapeContent::Picture(&picture),
                    width_twips: width,
                    height_twips: height,
                    position: &position,
                    text: None,
                };
                let expected =
                    crate::writer::images::build_dgg_info(std::slice::from_ref(&shape), &[], 1)
                        .map_err(|error| {
                            corrupted(format!("re-homed Dgg cannot be encoded: {error}"))
                        })?;
                if expected == self.dgg_info {
                    return Ok(());
                }
                let nested =
                    validate_floating_picture_identity(&self.dgg_info, shape_id, 0, &picture)?
                        .ok_or_else(|| corrupted("re-homed Dgg does not match its picture/SPA"))?;
                let expected = build_nested_picture_dgg_info(&picture, spa, &nested)?;
                if expected != self.dgg_info {
                    return Err(corrupted(
                        "re-homed nested image group is not a closed canonical dependency graph",
                    ));
                }
                Ok(())
            },
            _ => Err(corrupted("re-homed picture graph closure is inconsistent")),
        }
    }
}

fn floating_position_from_spa(spa: crate::parts::spa::Spa) -> crate::writer::FloatingPosition {
    crate::writer::FloatingPosition::new(spa.left, spa.top)
        .with_origins(spa.horizontal_origin, spa.vertical_origin)
        .with_text_wrap(spa.wrap)
        .with_wrap_side(spa.wrap_side)
        .behind_text(spa.below_text)
        .lock_anchor(spa.anchor_locked)
}

#[deny(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::expect_used,
    clippy::let_underscore_must_use,
    clippy::map_err_ignore,
    clippy::unwrap_used,
    reason = "durable picture blocks are a checked legacy-codec boundary"
)]
fn canonical_picture_from_block(block: &[u8]) -> Result<(crate::writer::Picture, u32, u32, u32)> {
    let fields = crate::image::PictureFields::try_parse(block, 0)
        .ok_or_else(|| corrupted("re-homed picture PICF is missing"))?;
    let lcb = fields.lcb;
    let goal_width = fields.dxa_goal;
    let goal_height = fields.dya_goal;
    if lcb <= 0
        || usize::try_from(lcb).ok() != Some(block.len())
        || fields.cb_header != 0x44
        || fields.mm != 0x64
        || goal_width <= 0
        || goal_height <= 0
        || fields.mx != 1000
        || fields.my != 1000
        || fields.dxa_reserved1 != 0
        || fields.dya_reserved1 != 0
        || fields.dxa_reserved2 != 0
        || fields.dya_reserved2 != 0
        || fields.dxa_reserved3 != 0
        || fields.dya_reserved3 != 0
        || fields.c_props != 0
    {
        return Err(corrupted("re-homed picture PICF is not canonical"));
    }
    let image = crate::image::Image::new(0)
        .data(block, &[])
        .map_err(|error| corrupted(format!("re-homed picture BLIP is invalid: {error}")))?;
    if !matches!(
        image.kind(),
        litchi_odraw::image::Kind::Jpeg
            | litchi_odraw::image::Kind::Png
            | litchi_odraw::image::Kind::Dib
            | litchi_odraw::image::Kind::Tiff
    ) {
        return Err(corrupted("re-homed picture BLIP kind is unsupported"));
    }
    let native = image
        .data()
        .map_err(|error| corrupted(format!("re-homed BLIP payload is invalid: {error}")))?;
    let width = u32::try_from(goal_width)
        .map_err(|error| corrupted(format!("picture width is invalid: {error}")))?;
    let height = u32::try_from(goal_height)
        .map_err(|error| corrupted(format!("picture height is invalid: {error}")))?;
    let picture =
        crate::writer::Picture::from_parts_as(native.to_vec(), image.kind(), width, height)
            .map_err(|error| corrupted(format!("re-homed picture is invalid: {error}")))?;
    let shape_id = picture_shape_id(block)?;
    let mut expected = Vec::new();
    crate::writer::images::write_picture_block(&picture, shape_id, &mut expected)
        .map_err(|error| corrupted(format!("re-homed picture cannot be encoded: {error}")))?;
    if expected != block {
        return Err(corrupted("re-homed picture block is not canonical"));
    }
    Ok((picture, width, height, shape_id))
}

#[deny(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::expect_used,
    clippy::let_underscore_must_use,
    clippy::map_err_ignore,
    clippy::unwrap_used,
    reason = "picture shape ownership is a checked OfficeArt boundary"
)]
fn picture_shape_id(block: &[u8]) -> Result<u32> {
    const PICF_HEADER_LEN: usize = 0x44;

    let (record, consumed) = litchi_odraw::Record::parse(block, PICF_HEADER_LEN)
        .map_err(|error| corrupted(format!("picture shape container is invalid: {error}")))?;
    if record.kind() != litchi_odraw::RecordKind::SpContainer {
        return Err(corrupted(
            "picture PICF is not followed by an OfficeArt shape container",
        ));
    }
    let end = PICF_HEADER_LEN
        .checked_add(consumed)
        .ok_or_else(|| corrupted("picture shape-container extent overflow"))?;
    if end > block.len() {
        return Err(corrupted(
            "picture shape container extends past its PICF block",
        ));
    }
    let container = litchi_odraw::Container::try_new(record)
        .map_err(|error| corrupted(format!("picture shape container is malformed: {error}")))?;
    let mut shape_id = None;
    for child in container.children() {
        let child = child.map_err(|error| {
            corrupted(format!("picture shape child record is malformed: {error}"))
        })?;
        if child.kind() != litchi_odraw::RecordKind::Sp {
            continue;
        }
        if shape_id.is_some() || child.version() != 2 || child.len() != 8 {
            return Err(corrupted(
                "picture block does not contain one canonical shape atom",
            ));
        }
        let bytes: [u8; 8] = child
            .data()
            .try_into()
            .map_err(|error| corrupted(format!("picture shape atom is invalid: {error}")))?;
        shape_id = Some(u32::from_le_bytes(bytes[..4].try_into().map_err(
            |error| corrupted(format!("picture shape identifier is invalid: {error}")),
        )?));
    }
    shape_id.ok_or_else(|| corrupted("picture block has no shape atom"))
}

#[deny(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::expect_used,
    clippy::let_underscore_must_use,
    clippy::map_err_ignore,
    clippy::unwrap_used,
    reason = "main/header OfficeArt traversal is bounded and checked"
)]
struct MainDrawing<'data> {
    bytes: &'data [u8],
    shapes: Vec<litchi_odraw::shape::Shape<'data>>,
}

fn main_drawing(dgg_info: &[u8]) -> Result<MainDrawing<'_>> {
    let (first, first_size) = litchi_odraw::Record::parse(dgg_info, 0)
        .map_err(|error| corrupted(format!("drawing-group root is invalid: {error}")))?;
    if first.kind() != litchi_odraw::RecordKind::DggContainer {
        return Err(corrupted(
            "OfficeArtContent does not start with a DggContainer",
        ));
    }
    let mut offset = first_size;
    let mut main = None;
    let mut header_seen = false;
    while offset < dgg_info.len() {
        let label = *dgg_info
            .get(offset)
            .ok_or_else(|| corrupted("OfficeArtWordDrawing label is truncated"))?;
        if label > 1 {
            return Err(corrupted(format!(
                "OfficeArtWordDrawing has invalid story label {label}"
            )));
        }
        let record_offset = offset
            .checked_add(1)
            .ok_or_else(|| corrupted("OfficeArtWordDrawing offset overflow"))?;
        let (drawing, drawing_size) = litchi_odraw::Record::parse(dgg_info, record_offset)
            .map_err(|error| corrupted(format!("main OfficeArt drawing is invalid: {error}")))?;
        if drawing.kind() != litchi_odraw::RecordKind::DgContainer {
            return Err(corrupted(
                "OfficeArtWordDrawing does not contain a DgContainer",
            ));
        }
        let drawing_end = record_offset
            .checked_add(drawing_size)
            .ok_or_else(|| corrupted("OfficeArt drawing extent overflow"))?;
        let drawing_bytes = dgg_info
            .get(record_offset..drawing_end)
            .ok_or_else(|| corrupted("OfficeArt drawing extends past DggInfo"))?;
        let shapes = litchi_odraw::shape::parse(drawing_bytes)
            .map_err(|error| corrupted(format!("OfficeArt shapes are invalid: {error}")))?;
        if label == 0 {
            if main
                .replace(MainDrawing {
                    bytes: drawing_bytes,
                    shapes,
                })
                .is_some()
            {
                return Err(corrupted(
                    "OfficeArtContent contains more than one main-story drawing",
                ));
            }
        } else if std::mem::replace(&mut header_seen, true) {
            return Err(corrupted(
                "OfficeArtContent contains more than one header drawing",
            ));
        }
        offset = drawing_end;
    }
    main.ok_or_else(|| corrupted("floating picture has no main-story OfficeArt drawing"))
}

#[deny(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::expect_used,
    clippy::let_underscore_must_use,
    clippy::map_err_ignore,
    clippy::unwrap_used,
    reason = "selected picture identity is a checked OfficeArt/BStore boundary"
)]
#[derive(Debug)]
struct NestedPictureGroup {
    bytes: Vec<u8>,
    shape_ids: Vec<u32>,
}

fn validate_floating_picture_identity(
    dgg_info: &[u8],
    shape_id: u32,
    anchor_index: usize,
    picture: &crate::writer::Picture,
) -> Result<Option<NestedPictureGroup>> {
    let drawing = main_drawing(dgg_info)?;
    let selected = drawing
        .shapes
        .iter()
        .find(|shape| shape.id() == shape_id)
        .ok_or_else(|| corrupted("floating picture shape is absent from the main drawing"))?;
    let mut matching_ids = 0usize;
    let mut all_ids = std::collections::BTreeSet::new();
    let mut pending = drawing.shapes.iter().collect::<Vec<_>>();
    while let Some(shape) = pending.pop() {
        if !all_ids.insert(shape.id()) {
            return Err(corrupted("OfficeArt drawing contains duplicate shape IDs"));
        }
        if shape.id() == shape_id {
            matching_ids = matching_ids
                .checked_add(1)
                .ok_or_else(|| corrupted("matching picture-shape count overflow"))?;
        }
        pending.extend(shape.children());
    }
    if matching_ids != 1 {
        return Err(corrupted(
            "floating picture shape ID is duplicated or ambiguously nested",
        ));
    }
    if selected.kind() == litchi_odraw::shape::Kind::Picture {
        validate_canonical_picture_shape(selected, anchor_index)?;
        validate_picture_store(dgg_info, selected, picture)?;
        return Ok(None);
    }
    if selected.kind() != litchi_odraw::shape::Kind::Group
        || !selected
            .flags()
            .contains(litchi_odraw::shape::Flags::GROUP | litchi_odraw::shape::Flags::HAVE_ANCHOR)
        || selected
            .flags()
            .intersects(litchi_odraw::shape::Flags::CHILD | litchi_odraw::shape::Flags::PATRIARCH)
        || selected.textbox().is_some()
    {
        return Err(corrupted(
            "selected floating shape is neither a picture nor a closed image group",
        ));
    }
    let client_anchor = selected
        .client_anchor()
        .ok_or_else(|| corrupted("selected image group has no Word client anchor"))?;
    let anchor_ordinal = u32::try_from(anchor_index)
        .map_err(|error| corrupted(format!("SPA anchor index exceeds u32: {error}")))?;
    if client_anchor.version() != 0
        || client_anchor.instance() != 0
        || client_anchor.data() != anchor_ordinal.to_le_bytes()
    {
        return Err(corrupted(
            "selected image group does not own the expected SPA client anchor",
        ));
    }

    let mut shape_ids = vec![selected.id()];
    let mut member = selected;
    loop {
        let [child] = member.children() else {
            return Err(corrupted(
                "selected image group is not a closed single-child chain",
            ));
        };
        shape_ids.push(child.id());
        if child.kind() == litchi_odraw::shape::Kind::Group {
            if child.textbox().is_some() {
                return Err(corrupted("nested image group owns a textbox"));
            }
            member = child;
            continue;
        }
        if child.kind() != litchi_odraw::shape::Kind::Picture
            || child.native_kind() != litchi_odraw::shape::Native::PICTURE
            || !child
                .flags()
                .contains(litchi_odraw::shape::Flags::CHILD | litchi_odraw::shape::Flags::HAVE_SPT)
            || child.textbox().is_some()
            || !child.children().is_empty()
        {
            return Err(corrupted(
                "selected image group has a non-picture or active terminal child",
            ));
        }
        validate_picture_store(dgg_info, child, picture)?;
        let pib = picture_pib(child)?;
        let record = selected.container().record();
        let payload_offset = record
            .data_offset(drawing.bytes)
            .ok_or_else(|| corrupted("selected image-group record is outside the main drawing"))?;
        let record_start = payload_offset
            .checked_sub(8)
            .ok_or_else(|| corrupted("selected image-group header offset underflows"))?;
        let record_len = usize::try_from(record.len())
            .map_err(|error| corrupted(format!("image-group length exceeds usize: {error}")))?
            .checked_add(8)
            .ok_or_else(|| corrupted("image-group record extent overflows"))?;
        let record_end = record_start
            .checked_add(record_len)
            .ok_or_else(|| corrupted("image-group record extent overflows"))?;
        let mut bytes = drawing
            .bytes
            .get(record_start..record_end)
            .ok_or_else(|| corrupted("selected image-group record is truncated"))?
            .to_vec();
        rewrite_nested_picture_group(&mut bytes, pib)?;
        return Ok(Some(NestedPictureGroup { bytes, shape_ids }));
    }
}

fn validate_canonical_picture_shape(
    selected: &litchi_odraw::shape::Shape<'_>,
    anchor_index: usize,
) -> Result<()> {
    if selected.kind() != litchi_odraw::shape::Kind::Picture
        || selected.native_kind() != litchi_odraw::shape::Native::PICTURE
        || selected.flags()
            != (litchi_odraw::shape::Flags::HAVE_ANCHOR | litchi_odraw::shape::Flags::HAVE_SPT)
        || selected.anchor().is_some()
        || !selected.children().is_empty()
        || selected.textbox().is_some()
    {
        return Err(corrupted(
            "selected floating shape is not a canonical top-level picture frame",
        ));
    }
    let _pib = selected
        .props()
        .prop(litchi_odraw::prop::Id::BlipToDisplay)
        .filter(|prop| {
            selected.props().len() == 1
                && prop.raw_opid() == 0x4104
                && prop.is_blip()
                && !prop.is_complex()
                && prop.raw_value() > 0
        })
        .ok_or_else(|| corrupted("selected picture frame has no canonical pib property"))?;
    let records = selected
        .meta()
        .children()
        .collect::<std::result::Result<Vec<_>, _>>()
        .map_err(|error| corrupted(format!("picture shape records are invalid: {error}")))?;
    let expected_kinds = [
        litchi_odraw::RecordKind::Sp,
        litchi_odraw::RecordKind::Opt,
        litchi_odraw::RecordKind::ClientAnchor,
        litchi_odraw::RecordKind::ClientData,
    ];
    let [sp, opt, client_anchor, client_data] = records.as_slice() else {
        return Err(corrupted(
            "selected picture frame has noncanonical OfficeArt records",
        ));
    };
    if records
        .iter()
        .zip(expected_kinds)
        .any(|(record, expected)| record.kind() != expected)
        || sp.version() != 2
        || sp.instance() != litchi_odraw::shape::Native::PICTURE.raw()
        || sp.len() != 8
        || opt.version() != 3
        || opt.instance() != 1
        || opt.len() != 6
        || client_anchor.version() != 0
        || client_anchor.instance() != 0
        || client_anchor.len() != 4
        || client_data.version() != 0
        || client_data.instance() != 0
        || client_data.len() != 4
    {
        return Err(corrupted(
            "selected picture frame has noncanonical OfficeArt records",
        ));
    }
    let anchor_ordinal = u32::try_from(anchor_index)
        .map_err(|error| corrupted(format!("SPA anchor index exceeds u32: {error}")))?;
    if client_anchor.data() != anchor_ordinal.to_le_bytes() || client_data.data() != [0; 4] {
        return Err(corrupted(
            "selected picture frame does not own the expected SPA client anchor",
        ));
    }

    Ok(())
}

fn picture_pib(selected: &litchi_odraw::shape::Shape<'_>) -> Result<u32> {
    let props = selected.props();
    let pib = props
        .prop(litchi_odraw::prop::Id::BlipToDisplay)
        .filter(|prop| {
            prop.raw_opid() == 0x4104
                && prop.is_blip()
                && !prop.is_complex()
                && prop.raw_value() > 0
        })
        .ok_or_else(|| corrupted("picture frame has no usable pib property"))?;
    u32::try_from(pib.raw_value())
        .map_err(|error| corrupted(format!("picture pib is invalid: {error}")))
}

fn validate_picture_store(
    dgg_info: &[u8],
    selected: &litchi_odraw::shape::Shape<'_>,
    picture: &crate::writer::Picture,
) -> Result<()> {
    let pib = picture_pib(selected)?;
    let (_drawing_group, drawing_group_size) = litchi_odraw::Record::parse(dgg_info, 0)
        .map_err(|error| corrupted(format!("drawing-group root is invalid: {error}")))?;
    let drawing_group = dgg_info
        .get(..drawing_group_size)
        .ok_or_else(|| corrupted("drawing-group root extends past DggInfo"))?;
    let store = litchi_odraw::image::store(drawing_group)
        .map_err(|error| corrupted(format!("drawing BStore is invalid: {error}")))?
        .ok_or_else(|| corrupted("floating picture drawing has no BStore"))?;
    let id = litchi_odraw::image::Id::new(pib)
        .map_err(|error| corrupted(format!("picture pib is outside the BStore range: {error}")))?;
    let file = litchi_odraw::image::get(&store, id, None)
        .map_err(|error| corrupted(format!("picture BStore entry is invalid: {error}")))?
        .ok_or_else(|| corrupted("picture pib does not resolve to an embedded BStore entry"))?;
    let native = file
        .data()
        .map_err(|error| corrupted(format!("picture BStore payload is invalid: {error}")))?;
    if file.kind() != picture.kind() || native != picture.data() {
        return Err(corrupted(
            "picture PICF and floating BStore entry have different binary identities",
        ));
    }
    Ok(())
}

#[derive(Default)]
struct NestedRewriteState {
    blip_references: usize,
    client_anchors: usize,
}

fn rewrite_nested_picture_group(bytes: &mut [u8], source_pib: u32) -> Result<()> {
    let (record, consumed) = litchi_odraw::Record::parse(bytes, 0)
        .map_err(|error| corrupted(format!("nested image-group record is invalid: {error}")))?;
    if consumed != bytes.len() || record.kind() != litchi_odraw::RecordKind::SpgrContainer {
        return Err(corrupted(
            "nested image-group closure is not one complete SpgrContainer",
        ));
    }
    let mut state = NestedRewriteState::default();
    rewrite_nested_records(bytes, 0, bytes.len(), source_pib, &mut state, 0)?;
    if state.blip_references == 0 || state.client_anchors != 1 {
        return Err(corrupted(
            "nested image-group closure lacks one BLIP reference or Word client anchor",
        ));
    }
    Ok(())
}

fn rewrite_nested_records(
    bytes: &mut [u8],
    start: usize,
    end: usize,
    source_pib: u32,
    state: &mut NestedRewriteState,
    depth: u16,
) -> Result<()> {
    if depth > 64 {
        return Err(corrupted("nested image-group depth exceeds 64"));
    }
    let mut offset = start;
    while offset < end {
        let header_end = offset
            .checked_add(8)
            .ok_or_else(|| corrupted("nested image-group record header overflows"))?;
        let header = bytes
            .get(offset..header_end)
            .ok_or_else(|| corrupted("nested image-group record header is truncated"))?;
        let ver_inst = u16::from_le_bytes([header[0], header[1]]);
        let version = ver_inst & 0x000f;
        let instance = usize::from(ver_inst >> 4);
        let kind = u16::from_le_bytes([header[2], header[3]]);
        let body_len = usize::try_from(u32::from_le_bytes([
            header[4], header[5], header[6], header[7],
        ]))
        .map_err(|error| corrupted(format!("nested image-group length exceeds usize: {error}")))?;
        let body_end = header_end
            .checked_add(body_len)
            .ok_or_else(|| corrupted("nested image-group record extent overflows"))?;
        if body_end > end {
            return Err(corrupted(
                "nested image-group child extends past its container",
            ));
        }
        if version == 0x000f {
            rewrite_nested_records(bytes, header_end, body_end, source_pib, state, depth + 1)?;
        } else if matches!(kind, 0xF00B | 0xF121 | 0xF122) {
            let descriptor_bytes = instance
                .checked_mul(6)
                .ok_or_else(|| corrupted("nested image-group property table overflows"))?;
            if descriptor_bytes > body_len {
                return Err(corrupted(
                    "nested image-group property descriptors are truncated",
                ));
            }
            for property in 0..instance {
                let property_offset = header_end
                    .checked_add(property * 6)
                    .ok_or_else(|| corrupted("nested image-group property offset overflows"))?;
                let opid = u16::from_le_bytes(
                    bytes[property_offset..property_offset + 2]
                        .try_into()
                        .map_err(|error| {
                            corrupted(format!("nested image-group opid is truncated: {error}"))
                        })?,
                );
                if opid & 0x4000 == 0 {
                    continue;
                }
                if opid & 0x8000 != 0 {
                    return Err(corrupted("nested image-group has a complex BLIP property"));
                }
                let value_offset = property_offset + 2;
                let value =
                    u32::from_le_bytes(bytes[value_offset..value_offset + 4].try_into().map_err(
                        |error| corrupted(format!("nested image-group pib is truncated: {error}")),
                    )?);
                if value != source_pib {
                    return Err(corrupted(
                        "nested image-group references more than one BLIP-store entry",
                    ));
                }
                bytes[value_offset..value_offset + 4].copy_from_slice(&1u32.to_le_bytes());
                state.blip_references = state
                    .blip_references
                    .checked_add(1)
                    .ok_or_else(|| corrupted("nested image-group BLIP count overflows"))?;
            }
        } else if kind == 0xF010 {
            if body_len != 4 {
                return Err(corrupted(
                    "nested image-group Word client anchor is not four bytes",
                ));
            }
            bytes[header_end..body_end].copy_from_slice(&0u32.to_le_bytes());
            state.client_anchors = state
                .client_anchors
                .checked_add(1)
                .ok_or_else(|| corrupted("nested image-group anchor count overflows"))?;
        }
        offset = body_end;
    }
    if offset != end {
        return Err(corrupted(
            "nested image-group record sequence has trailing bytes",
        ));
    }
    Ok(())
}

#[allow(
    clippy::drop_non_drop,
    reason = "explicit drops end zero-copy OfficeArt borrows before the backing buffer is spliced"
)]
fn build_nested_picture_dgg_info(
    picture: &crate::writer::Picture,
    spa: crate::parts::spa::Spa,
    group: &NestedPictureGroup,
) -> Result<Vec<u8>> {
    let first_shape_id = crate::writer::images::FIRST_PICTURE_SHAPE_ID;
    if group.shape_ids.is_empty()
        || group.shape_ids[0] != spa.shape_id
        || group
            .shape_ids
            .iter()
            .any(|id| *id < first_shape_id || *id >= 2_048)
    {
        return Err(corrupted(
            "nested image-group shape IDs are outside the main drawing cluster",
        ));
    }
    let max_shape_id = group
        .shape_ids
        .iter()
        .copied()
        .max()
        .ok_or_else(|| corrupted("nested image-group has no shape IDs"))?;
    let allocated = max_shape_id
        .checked_sub(first_shape_id)
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| corrupted("nested image-group allocation count overflows"))?;
    let position = floating_position_from_spa(spa);
    let placeholder = crate::writer::images::FloatingShapeInfo {
        anchor_cp: 0,
        shape_id: spa.shape_id,
        content: crate::writer::images::FloatingShapeContent::Picture(picture),
        width_twips: u32::try_from(spa.width())
            .map_err(|error| corrupted(format!("nested image-group width is invalid: {error}")))?,
        height_twips: u32::try_from(spa.height())
            .map_err(|error| corrupted(format!("nested image-group height is invalid: {error}")))?,
        position: &position,
        text: None,
    };
    let mut output = crate::writer::images::build_dgg_info(
        std::slice::from_ref(&placeholder),
        &[],
        allocated,
    )
    .map_err(|error| corrupted(format!("nested image-group Dgg cannot be encoded: {error}")))?;

    let (drawing_group_record, drawing_group_size) = litchi_odraw::Record::parse(&output, 0)
        .map_err(|error| corrupted(format!("generated Dgg root is invalid: {error}")))?;
    let dgg_meta = litchi_odraw::Container::try_new(drawing_group_record)
        .map_err(|error| corrupted(format!("generated Dgg root is invalid: {error}")))?
        .find(litchi_odraw::RecordKind::Dgg)
        .map_err(|error| corrupted(format!("generated Dgg atom is invalid: {error}")))?
        .ok_or_else(|| corrupted("generated Dgg root has no Dgg atom"))?;
    let dgg_meta_payload = dgg_meta
        .data_offset(&output)
        .ok_or_else(|| corrupted("generated Dgg atom is outside its root"))?;

    let dg_start = drawing_group_size
        .checked_add(1)
        .ok_or_else(|| corrupted("generated drawing offset overflows"))?;
    let (drawing_record, drawing_size) = litchi_odraw::Record::parse(&output, dg_start)
        .map_err(|error| corrupted(format!("generated drawing is invalid: {error}")))?;
    if drawing_record.kind() != litchi_odraw::RecordKind::DgContainer {
        return Err(corrupted("generated drawing has no DgContainer"));
    }
    let drawing_container = litchi_odraw::Container::try_new(drawing_record)
        .map_err(|error| corrupted(format!("generated drawing is invalid: {error}")))?;
    let dg_atom = drawing_container
        .find(litchi_odraw::RecordKind::Dg)
        .map_err(|error| corrupted(format!("generated Dg atom is invalid: {error}")))?
        .ok_or_else(|| corrupted("generated drawing has no Dg atom"))?;
    let dg_payload = dg_atom
        .data_offset(&output)
        .ok_or_else(|| corrupted("generated Dg atom is outside its drawing"))?;
    let root_group = drawing_container
        .find(litchi_odraw::RecordKind::SpgrContainer)
        .map_err(|error| corrupted(format!("generated root shape group is invalid: {error}")))?
        .ok_or_else(|| corrupted("generated drawing has no root shape group"))?;
    let root_group_len = usize::try_from(root_group.len())
        .map_err(|error| corrupted(format!("root shape-group length exceeds usize: {error}")))?;
    let root_group_payload = root_group
        .data_offset(&output)
        .ok_or_else(|| corrupted("generated root shape group is outside its drawing"))?;
    let root_group_start = root_group_payload
        .checked_sub(8)
        .ok_or_else(|| corrupted("generated root shape-group offset underflows"))?;
    let root_group_container = litchi_odraw::Container::try_new(root_group)
        .map_err(|error| corrupted(format!("generated root shape group is invalid: {error}")))?;
    let shape_records = root_group_container
        .children()
        .collect::<std::result::Result<Vec<_>, _>>()
        .map_err(|error| corrupted(format!("generated shape records are invalid: {error}")))?;
    let [_patriarch, selected] = shape_records.as_slice() else {
        return Err(corrupted(
            "generated root shape group is not a singleton drawing",
        ));
    };
    let selected_payload = selected
        .data_offset(&output)
        .ok_or_else(|| corrupted("generated selected shape is outside its drawing"))?;
    let selected_start = selected_payload
        .checked_sub(8)
        .ok_or_else(|| corrupted("generated selected-shape offset underflows"))?;
    let selected_end = selected_payload
        .checked_add(usize::try_from(selected.len()).map_err(|error| {
            corrupted(format!(
                "generated selected-shape length exceeds usize: {error}"
            ))
        })?)
        .ok_or_else(|| corrupted("generated selected-shape extent overflows"))?;
    let selected_len = selected_end - selected_start;
    let replacement_len = group.bytes.len();
    drop(shape_records);
    drop(root_group_container);
    drop(drawing_container);
    drop(dgg_meta);
    output.splice(selected_start..selected_end, group.bytes.iter().copied());
    let new_root_len = root_group_len
        .checked_sub(selected_len)
        .and_then(|value| value.checked_add(replacement_len))
        .ok_or_else(|| corrupted("root shape-group replacement length overflows"))?;
    let new_dg_len = drawing_size
        .checked_sub(8)
        .and_then(|value| value.checked_sub(selected_len))
        .and_then(|value| value.checked_add(replacement_len))
        .ok_or_else(|| corrupted("drawing replacement length overflows"))?;
    write_record_length(&mut output, root_group_start, new_root_len)?;
    write_record_length(&mut output, dg_start, new_dg_len)?;

    let shape_count = u32::try_from(group.shape_ids.len())
        .map_err(|error| corrupted(format!("nested shape count exceeds u32: {error}")))?
        .checked_add(1)
        .ok_or_else(|| corrupted("nested shape count overflows"))?;
    write_u32_at(&mut output, dg_payload, shape_count)?;
    write_u32_at(&mut output, dg_payload + 4, max_shape_id + 1)?;
    let spid_max = max_shape_id
        .checked_add(1_024 - max_shape_id % 1_024)
        .ok_or_else(|| corrupted("nested image-group spidMax overflows"))?;
    write_u32_at(&mut output, dgg_meta_payload, spid_max)?;
    write_u32_at(&mut output, dgg_meta_payload + 8, shape_count)?;
    write_u32_at(&mut output, dgg_meta_payload + 20, max_shape_id % 1_024 + 1)?;
    Ok(output)
}

fn write_record_length(bytes: &mut [u8], record_start: usize, length: usize) -> Result<()> {
    let length = u32::try_from(length)
        .map_err(|error| corrupted(format!("OfficeArt record length exceeds u32: {error}")))?;
    write_u32_at(bytes, record_start + 4, length)
}

fn write_u32_at(bytes: &mut [u8], offset: usize, value: u32) -> Result<()> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| corrupted("OfficeArt patch offset overflows"))?;
    bytes
        .get_mut(offset..end)
        .ok_or_else(|| corrupted("OfficeArt patch is outside generated bytes"))?
        .copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn containing_piece(pieces: &[RawPiece], cp: u32) -> Option<&RawPiece> {
    let index = pieces
        .partition_point(|piece| piece.start <= cp)
        .checked_sub(1)?;
    pieces.get(index).filter(|piece| cp < piece.end)
}

fn containing_papx(runs: &[PapxRun], fc: u32) -> Option<&PapxRun> {
    let index = runs.partition_point(|run| run.start <= fc).checked_sub(1)?;
    runs.get(index).filter(|run| fc < run.end)
}

/// A rendered CFB handoff retained by the owning DOC transaction.
///
/// Cloning a revision editor deliberately drops this value.  Every mutation
/// is clone-first, so a candidate can never inherit a render belonging to the
/// source state, and cloning an editor never copies or shares a caller-owned
/// retention token.
#[derive(Default)]
struct RetainedRender(Option<Vec<u8>>);

impl Clone for RetainedRender {
    fn clone(&self) -> Self {
        let _had_render = self.0.is_some();
        Self(None)
    }
}

#[derive(Clone)]
pub struct RevisionEditor {
    package: ObjectEditor,
    word_path: Vec<String>,
    table_path: Vec<String>,
    data_path: Vec<String>,
    word: Vec<u8>,
    table: Vec<u8>,
    data: Vec<u8>,
    pieces: Vec<RawPiece>,
    chpx: Vec<FcRun>,
    papx: Vec<PapxRun>,
    authors: Vec<String>,
    cp_tables: Vec<CpTable>,
    unmodeled_cp_tables: Vec<usize>,
    main_ccp: u32,
    /// Exact bytes admitted before this editor was exposed. The lower owner
    /// uses this source for publication-time signature checks, including a
    /// rendered candidate that happens to clear the sticky changed bit.
    source: Arc<Vec<u8>>,
    signed: bool,
    protection: EditProtection,
    protection_policy: ProtectionPolicy,
    data_changed: bool,
    changed: bool,
    retained_render_limit: usize,
    retained_render: RetainedRender,
}

impl RevisionEditor {
    pub fn open(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        Self::open_with_ole_file(bytes, limits).map(|(editor, _ole)| editor)
    }

    /// Opens a tracked-revision editor with an explicit protected-edit policy.
    ///
    /// The default [`Self::open`] path enforces document and range-level
    /// protection. Callers may opt into the explicit caller-granted capability
    /// carried by [`ProtectionPolicy`]. The capability records caller
    /// metadata; it does not authenticate a user or provide cryptographic
    /// audit.
    pub fn open_with_policy(
        bytes: Vec<u8>,
        limits: Limits,
        policy: ProtectionPolicy,
    ) -> Result<Self> {
        Self::open_with_ole_file_with_policy(bytes, limits, policy).map(|(editor, _ole)| editor)
    }

    pub(crate) fn open_with_ole_file(
        bytes: Vec<u8>,
        limits: Limits,
    ) -> Result<(Self, OleFile<Cursor<Vec<u8>>>)> {
        Self::open_with_ole_file_with_policy(bytes, limits, ProtectionPolicy::default())
    }

    pub(crate) fn open_with_ole_file_with_policy(
        bytes: Vec<u8>,
        limits: Limits,
        policy: ProtectionPolicy,
    ) -> Result<(Self, OleFile<Cursor<Vec<u8>>>)> {
        let (package, mut ole) =
            ObjectEditor::open_with_ole_file(bytes, Targets::default(), limits)
                .map_err(PackageError::from)?;
        let editor = Self::open_from_package(package, &mut ole, policy)?;
        Ok((editor, ole))
    }

    #[cfg(feature = "performance-diagnostics")]
    pub(crate) fn open_with_ole_file_profiled(
        bytes: Vec<u8>,
        limits: Limits,
        observer: impl FnMut(CfbParseEvent),
    ) -> Result<(Self, OleFile<Cursor<Vec<u8>>>)> {
        let (package, mut ole) =
            ObjectEditor::open_with_ole_file_profiled(bytes, Targets::default(), limits, observer)
                .map_err(PackageError::from)?;
        let editor = Self::open_from_package(package, &mut ole, ProtectionPolicy::default())?;
        Ok((editor, ole))
    }

    fn open_from_package(
        package: ObjectEditor,
        ole: &mut OleFile<Cursor<Vec<u8>>>,
        protection_policy: ProtectionPolicy,
    ) -> Result<Self> {
        let word_path = vec!["WordDocument".to_string()];
        let word = package
            .stream(&word_path)
            .ok_or_else(|| corrupted("WordDocument stream is missing"))?
            .to_vec();
        if word.len() < FIB_FC_LCB + (STTBFRMARK + 1) * 8 || u16_at(&word, 0)? != 0xA5EC {
            return Err(corrupted(
                "tracked-revision editing requires Word 97+ FIB data",
            ));
        }
        let flags = u16_at(&word, 10)?;
        if flags & 0x0100 != 0 || u32_at(&word, 14)? != 0 {
            return Err(corrupted("encrypted DOC cannot be edited"));
        }
        let table_path = vec![
            if flags & 0x0200 != 0 {
                "1Table"
            } else {
                "0Table"
            }
            .to_string(),
        ];
        let table = package
            .stream(&table_path)
            .ok_or_else(|| corrupted("selected Table stream is missing"))?
            .to_vec();
        let data_path = vec!["Data".to_string()];
        let data = package
            .stream(&data_path)
            .map_or_else(Vec::new, <[u8]>::to_vec);
        let signed = package_has_binary_signature(ole) || word_vba_signature_state(&word, &table);
        let fib = FileInformationBlock::parse(&word)?;
        let protection = crate::parts::protection::classify(&fib, &table)?;
        let main_ccp = u32_at(&word, FIB_CCP_TEXT)?;
        let pieces = parse_clx(&word, &table)?;
        if pieces.last().is_none_or(|piece| piece.end < main_ccp) {
            return Err(corrupted("piece table does not cover the main story"));
        }
        let chpx = parse_chpx(&word, &table)?;
        let papx = parse_papx(&word, &table)?;
        let authors = parse_authors(&word, &table)?;
        // Each modeled entry is a PLCF in the main/all-story CP coordinate
        // space. Fixed-size records stay opaque while their CP array moves.
        let modeled_cp_tables = [
            (2, 2),           // PlcffndRef / FRD
            (PLCFANDREF, 30), // PlcfandRef / ATRD
            (6, 12),          // PlcfSed / SED
            (PLCFFLD_MOM, 2), // PlcffldMom / FLD
            (PLCFBKF, 4),     // PlcfBkf / FBKF
            (PLCFBKL, 0),     // PlcfBkl
            (40, 26),         // PlcfSpaMom / SPA
            (PLCFATNBKF, 4),  // PlcfAtnBkf / FBKF
            (PLCFATNBKL, 0),  // PlcfAtnBkl
            (46, 2),          // PlcfendRef / FRD
            (54, 12),         // PlcfWKB / WKB
            (55, 2),          // PlcfSpl / SPLS
            (89, 4),          // PlcfAsumy / ASUMY
            (90, 2),          // PlcfGram / SPLS
            (93, 4),          // PlcfTch / TCH
            (98, 2),          // PlcfLvc / LSPD
            (115, 6),         // PlcfBkfFactoid / FBKFD
            (117, 4),         // PlcfBklFactoid / FBKLD
            (121, 6),         // PlcfBkfFcc / FBKFD
            (122, 4),         // PlcfBklFcc / FBKLD
            (124, 4),         // PlcfBkfBPRepairs / FBKF
            (125, 0),         // PlcfBklBPRepairs
            (132, 2),         // Plcffactoid / FactoidSpls
            (138, 6),         // PlcfBkfSdt / FBKFD
            (139, 4),         // PlcfBklSdt / FBKLD
            (142, 6),         // PlcfBkfProt / BKF
            (143, 0),         // PlcfBklProt
        ];
        let pair_count = usize::from(u16_at(&word, FIB_FC_LCB - 2)?);
        let cp_tables = modeled_cp_tables
            .into_iter()
            .filter(|(index, _size)| *index < pair_count)
            .filter_map(|(index, size)| parse_cp_table(&word, &table, index, size).transpose())
            .collect::<Result<Vec<_>>>()?;
        // Known CP-indexed tables whose coupled records are not owned here.
        // Length-changing edits refuse these instead of silently leaving stale
        // positions. Equal-length text and formatting edits remain safe.
        // PlcfSea has producer-private records. The cookie and UIM records
        // carry coupled character lengths that a CP-only splice cannot repair.
        let unmodeled_cp_tables = [14, 101, 110, 116]
            .into_iter()
            .filter(|index| *index < pair_count)
            .filter_map(|index| {
                fib_pair(&word, index)
                    .map(|(_offset, length)| (length != 0).then_some(index))
                    .transpose()
            })
            .collect::<Result<Vec<_>>>()?;
        let source = package.source_shared();
        let editor = Self {
            package,
            word_path,
            table_path,
            data_path,
            word,
            table,
            data,
            pieces,
            chpx,
            papx,
            authors,
            cp_tables,
            unmodeled_cp_tables,
            main_ccp,
            // The common owner already retains the admitted source allocation;
            // keep a handle to that same immutable bytes without copying it.
            source,
            signed,
            protection,
            protection_policy,
            data_changed: false,
            changed: false,
            retained_render_limit: 0,
            retained_render: RetainedRender::default(),
        };
        if editor.revisions()?.len() > MAX_REVISIONS {
            return Err(corrupted("revision count exceeds resource limit"));
        }
        Ok(editor)
    }

    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changed
    }

    /// Configures the finite owner-side ceiling for one already-rendered
    /// handoff. A zero ceiling releases any held render and disables
    /// retention. The ordinary tracked-revision API leaves this at zero.
    pub(crate) fn set_retained_render_limit(&mut self, limit: usize) {
        self.retained_render_limit = limit;
        if limit == 0
            || self
                .retained_render
                .0
                .as_ref()
                .is_some_and(|rendered| rendered.capacity() > limit)
        {
            self.retained_render.0 = None;
        }
    }

    /// Returns the allocation capacity held by the validated render, if any.
    #[must_use]
    pub(crate) fn retained_render_capacity(&self) -> Option<usize> {
        self.retained_render.0.as_ref().map(Vec::capacity)
    }

    /// Releases the retained render without changing the editor's meaning.
    pub(crate) fn release_retained_render(&mut self) {
        self.retained_render.0 = None;
    }

    /// Moves the retained validated render out of this editor, if any.
    ///
    /// The editor's meaning is unchanged; a later finish renders again.
    pub(crate) fn take_retained_render(&mut self) -> Option<Vec<u8>> {
        self.retained_render.0.take()
    }

    #[must_use]
    pub fn authors(&self) -> &[String] {
        &self.authors
    }

    /// Parses the optional saved-selection cache from this exact editor state.
    ///
    /// The FIB and table stream remain owned by the DOC editor; callers only
    /// receive the inert, source-retaining semantic value.
    pub(crate) fn saved_selection(&self) -> Result<Option<SavedSelection>> {
        let fib = FileInformationBlock::parse(&self.word)?;
        let selection = SavedSelection::parse(&fib, &self.table)?;
        if let Some(selection) = &selection {
            selection.validate_main_story_bound(self.main_ccp)?;
        }
        Ok(selection)
    }

    /// Parses the optional frame-set/list record array from this exact editor
    /// state without exposing raw FIB offsets or mutable streams.
    pub(crate) fn dofr_records(&self) -> Result<Option<DofrArray>> {
        let fib = FileInformationBlock::parse(&self.word)?;
        DofrArray::parse(&fib, &self.table)
    }

    /// Validate a DOFR patch against the current FIB-selected range before a
    /// candidate table stream is cloned. Exact byte no-ops are accepted even
    /// when the source carries a VBA signature.
    pub(crate) fn preflight_dofr_patch(&self, patch: &DofrPatch) -> Result<bool> {
        let fib = FileInformationBlock::parse(&self.word)?;
        patch.preflight_table_stream(&fib, &self.table)?;
        Ok(patch.changed())
    }

    /// Whether this source carries inert binary VBA signature metadata in
    /// PIDDSI or `StwUser`. The marker is read once at editor admission and
    /// is never interpreted as executable or trusted signature content.
    #[must_use]
    pub(crate) const fn dofr_source_is_signed(&self) -> bool {
        self.signed
    }

    /// Validate a saved-selection patch against this exact editor state before
    /// cloning the table stream. Exact byte no-ops are accepted even when the
    /// source carries a binary signature.
    pub(crate) fn preflight_saved_selection_patch(
        &self,
        patch: &SavedSelectionPatch,
    ) -> Result<bool> {
        let fib = FileInformationBlock::parse(&self.word)?;
        patch.preflight_table_stream(&fib, &self.table)?;
        patch.validate_main_story_bound(self.main_ccp)?;
        Ok(patch.changed())
    }

    /// Applies a source-checked, same-length saved-selection patch and
    /// publishes it through the existing failure-atomic package owner.
    pub(crate) fn apply_saved_selection_patch(
        &mut self,
        patch: &SavedSelectionPatch,
    ) -> Result<bool> {
        let fib = FileInformationBlock::parse(&self.word)?;
        patch.preflight_table_stream(&fib, &self.table)?;
        patch.validate_main_story_bound(self.main_ccp)?;
        if !patch.changed() {
            return Ok(false);
        }
        if self.signed {
            return Err(corrupted("signed DOC cannot accept a changed Selsf patch"));
        }
        let mut table = self.table.clone();
        patch.apply_to_table_stream(&fib, &mut table)?;
        self.replace_table_stream(table)
    }

    /// Applies a source-checked, same-length `RgDofr` patch through the DOC
    /// package owner. The caller never handles a mutable table stream.
    pub(crate) fn apply_dofr_patch(&mut self, patch: &DofrPatch) -> Result<bool> {
        let fib = FileInformationBlock::parse(&self.word)?;
        patch.preflight_table_stream(&fib, &self.table)?;
        if !patch.changed() {
            return Ok(false);
        }
        if self.signed {
            return Err(corrupted("signed DOC cannot accept a changed RgDofr patch"));
        }
        let mut table = self.table.clone();
        patch.apply_to_table_stream(&fib, &mut table)?;
        self.replace_table_stream(table)
    }

    fn replace_table_stream(&mut self, replacement: Vec<u8>) -> Result<bool> {
        if replacement == self.table {
            return Ok(false);
        }
        if replacement.len() != self.table.len() {
            return Err(corrupted(
                "DOC auxiliary table replacement must preserve stream length",
            ));
        }
        let mut candidate = self.clone();
        candidate.table = replacement;
        candidate.commit()?;
        *self = candidate;
        Ok(true)
    }

    /// Returns the exact main-story text after strict piece-table decoding.
    ///
    /// This deliberately differs from the permissive reader projection: a
    /// transaction must not treat malformed UTF-16 or a discontinuous piece
    /// table as editable text.
    pub(crate) fn main_story_text(&self) -> Result<String> {
        self.decode_text_range(0, self.main_ccp)
    }

    /// Returns one of the seven FIB story ranges in global piece-table CPs.
    pub(crate) fn story_range(&self, story_index: usize) -> Result<(u32, u32)> {
        if story_index >= 7 {
            return Err(corrupted("story index is out of range"));
        }
        let mut start = 0u32;
        for index in 0..story_index {
            start = start
                .checked_add(self.story_length(index)?)
                .ok_or_else(|| corrupted("story CP range overflow"))?;
        }
        let end = start
            .checked_add(self.story_length(story_index)?)
            .ok_or_else(|| corrupted("story CP range overflow"))?;
        if self.pieces.last().is_none_or(|piece| piece.end < end) {
            return Err(corrupted("piece table does not cover selected story"));
        }
        Ok((start, end))
    }

    /// Strict text for one FIB story, plus its global CP origin.
    pub(crate) fn story_text(&self, story_index: usize) -> Result<(u32, String)> {
        let (start, end) = self.story_range(story_index)?;
        self.decode_text_range(start, end).map(|text| (start, text))
    }

    fn decode_text_range(&self, range_start: u32, range_end: u32) -> Result<String> {
        let mut output = String::new();
        let mut covered = 0u32;
        for piece in &self.pieces {
            if piece.start >= range_end {
                break;
            }
            let start = piece.start.max(range_start);
            let end = piece.end.min(range_end);
            if end <= start {
                continue;
            }
            let count = end - start;
            let relative = start - piece.start;
            if piece.unicode {
                let fc = piece
                    .fc
                    .checked_add(
                        relative
                            .checked_mul(2)
                            .ok_or_else(|| corrupted("Unicode piece relative offset overflow"))?,
                    )
                    .ok_or_else(|| corrupted("Unicode piece offset overflow"))?;
                let offset = usize::try_from(fc)
                    .map_err(|_| corrupted("Unicode piece offset exceeds usize"))?;
                let byte_count = usize::try_from(count)
                    .ok()
                    .and_then(|value| value.checked_mul(2))
                    .ok_or_else(|| corrupted("Unicode piece byte count overflow"))?;
                let bytes = self
                    .word
                    .get(offset..offset + byte_count)
                    .ok_or_else(|| corrupted("Unicode piece exceeds WordDocument"))?;
                let units = bytes
                    .as_chunks::<2>()
                    .0
                    .iter()
                    .map(|value| u16::from_le_bytes([value[0], value[1]]))
                    .collect::<Vec<_>>();
                output.push_str(
                    &String::from_utf16(&units)
                        .map_err(|_| corrupted("main-story text contains invalid UTF-16"))?,
                );
            } else {
                let fc = piece
                    .fc
                    .checked_add(relative)
                    .ok_or_else(|| corrupted("compressed piece offset overflow"))?;
                let offset = usize::try_from(fc)
                    .map_err(|_| corrupted("compressed piece offset exceeds usize"))?;
                let byte_count = usize::try_from(count)
                    .map_err(|_| corrupted("compressed piece byte count exceeds usize"))?;
                let bytes = self
                    .word
                    .get(offset..offset + byte_count)
                    .ok_or_else(|| corrupted("compressed piece exceeds WordDocument"))?;
                let (decoded, _, had_errors) = encoding_rs::WINDOWS_1252.decode(bytes);
                if had_errors || decoded.encode_utf16().count() != bytes.len() {
                    return Err(corrupted("compressed piece cannot be decoded losslessly"));
                }
                output.push_str(&decoded);
            }
            covered = covered
                .checked_add(count)
                .ok_or_else(|| corrupted("main-story CP count overflow"))?;
        }
        let expected = range_end - range_start;
        if covered != expected || output.encode_utf16().count() != expected as usize {
            return Err(corrupted(
                "piece table does not exactly cover the selected story",
            ));
        }
        Ok(output)
    }

    fn story_length(&self, story_index: usize) -> Result<u32> {
        if story_index == 0 {
            Ok(self.main_ccp)
        } else {
            u32_at(&self.word, FIB_CCP_TEXT + story_index * 4)
        }
    }

    /// Known CP-indexed FIB tables outside the length-changing splice model.
    #[must_use]
    pub(crate) fn unmodeled_length_dependencies(&self) -> &[usize] {
        &self.unmodeled_cp_tables
    }

    /// Whether all character runs in a non-empty range have byte-identical
    /// direct formatting. Length-changing replacement can preserve only this
    /// unambiguous dependency closure.
    pub(crate) fn has_uniform_character_format(&self, start: u32, end: u32) -> Result<bool> {
        let groups = self.character_groups(start, end)?;
        Ok(groups.windows(2).all(|pair| pair[0] == pair[1]))
    }

    #[must_use]
    pub(crate) fn is_unicode_range(&self, start: u32, end: u32) -> bool {
        start < end
            && self
                .pieces
                .iter()
                .filter(|piece| piece.start < end && start < piece.end)
                .all(|piece| piece.unicode)
    }

    /// Returns a uniform direct bold override, or `None` when the selected
    /// runs disagree or use a non-literal toggle value.
    pub(crate) fn uniform_bold_override(
        &self,
        start: u32,
        end: u32,
    ) -> Result<Option<Option<bool>>> {
        self.uniform_character_override(start, end, SPRM_C_F_BOLD)
    }

    pub(crate) fn uniform_italic_override(
        &self,
        start: u32,
        end: u32,
    ) -> Result<Option<Option<bool>>> {
        self.uniform_character_override(start, end, SPRM_C_F_ITALIC)
    }

    pub(crate) fn uniform_underline_override(
        &self,
        start: u32,
        end: u32,
    ) -> Result<Option<Option<bool>>> {
        self.uniform_character_override(start, end, SPRM_C_KUL)
    }

    fn uniform_character_override(
        &self,
        start: u32,
        end: u32,
        opcode: u16,
    ) -> Result<Option<Option<bool>>> {
        let groups = self.character_groups(start, end)?;
        let mut uniform = None;
        for group in groups {
            let value = strict_sprms(group)?
                .iter()
                .rev()
                .find(|sprm| sprm.opcode == opcode)
                .and_then(super::super::sprm::Sprm::operand_byte);
            let value = match value {
                Some(0) => Some(false),
                Some(1) => Some(true),
                Some(_) => return Ok(None),
                None => None,
            };
            match uniform {
                Some(previous) if previous != value => return Ok(None),
                Some(_) => {},
                None => uniform = Some(value),
            }
        }
        Ok(uniform)
    }

    /// Replaces a non-empty main-story range by appending one Unicode piece,
    /// shifting modeled CP tables, rebuilding CHPX FKPs, and publishing a new
    /// CLX. Callers must first prove a uniform character-format closure.
    pub(crate) fn replace_plain_text(
        &mut self,
        start: u32,
        end: u32,
        replacement: &str,
    ) -> Result<()> {
        validate_range(start, end, self.main_ccp)?;
        self.reject_destructive_interactions(start, end)?;
        if !self.has_uniform_character_format(start, end)? {
            return Err(corrupted(
                "length-changing body replacement crosses character formatting runs",
            ));
        }
        let groups = self.character_groups(start, end)?;
        let formatting = groups
            .first()
            .ok_or_else(|| corrupted("body replacement has no character formatting"))?
            .to_vec();
        let units = replacement.encode_utf16().collect::<Vec<_>>();
        if units.len() > MAX_TEXT_UNITS {
            return Err(corrupted("body replacement exceeds text resource limit"));
        }
        let length_changed = units.len() != (end - start) as usize;
        let added = u32::try_from(units.len())
            .map_err(|_error| corrupted("body replacement length exceeds u32"))?;
        let saved_selection = if length_changed {
            self.prepared_saved_selection_splice(start, end, added)?
        } else {
            None
        };
        if length_changed && !self.unmodeled_cp_tables.is_empty() {
            return Err(corrupted(
                "length-changing body replacement has unmodeled CP-indexed dependencies",
            ));
        }

        let mut candidate = self.clone();
        let removed = end - start;
        delete_piece_range(&mut candidate.pieces, start, end)?;
        if added != 0 {
            let fc = align2(candidate.word.len())?;
            candidate.word.resize(fc, 0);
            for unit in units {
                candidate.word.extend_from_slice(&unit.to_le_bytes());
            }
            insert_piece(
                &mut candidate.pieces,
                start,
                added,
                u32::try_from(fc).map_err(|_error| corrupted("body replacement FC exceeds u32"))?,
            )?;
            candidate.chpx.push(FcRun {
                start: u32::try_from(fc)
                    .map_err(|_error| corrupted("body replacement FC exceeds u32"))?,
                end: u32::try_from(candidate.word.len())
                    .map_err(|_error| corrupted("body replacement FC exceeds u32"))?,
                grpprl: formatting,
            });
        }
        candidate.shift_cp_tables(start, removed, added)?;
        candidate.main_ccp = candidate
            .main_ccp
            .checked_sub(removed)
            .and_then(|value| value.checked_add(added))
            .ok_or_else(|| corrupted("main story CP replacement overflow"))?;
        candidate.rewrite_chpx()?;
        candidate.append_clx_and_cp_tables()?;
        candidate.patch_sizes()?;
        if let Some(replacement) = saved_selection {
            candidate.apply_saved_selection_splice(replacement)?;
        }
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// Sets or clears one direct bold override while retaining every other
    /// character SPRM and rebuilding the affected CHPX FKPs.
    pub(crate) fn set_character_bold_override(
        &mut self,
        start: u32,
        end: u32,
        value: Option<bool>,
    ) -> Result<()> {
        self.set_character_override(start, end, SPRM_C_F_BOLD, value)
    }

    pub(crate) fn set_character_italic_override(
        &mut self,
        start: u32,
        end: u32,
        value: Option<bool>,
    ) -> Result<()> {
        self.set_character_override(start, end, SPRM_C_F_ITALIC, value)
    }

    pub(crate) fn set_character_underline_override(
        &mut self,
        start: u32,
        end: u32,
        value: Option<bool>,
    ) -> Result<()> {
        self.set_character_override(start, end, SPRM_C_KUL, value)
    }

    fn set_character_override(
        &mut self,
        start: u32,
        end: u32,
        opcode: u16,
        value: Option<bool>,
    ) -> Result<()> {
        validate_range(
            start,
            end,
            self.pieces.last().map_or(self.main_ccp, |piece| piece.end),
        )?;
        let intervals = self.fc_intervals(start, end)?;
        let mut candidate = self.clone();
        split_transform_chpx(&mut candidate.chpx, &intervals, |group| {
            let mut output = retain_sprms(group, &[opcode])?;
            if let Some(enabled) = value {
                output.extend_from_slice(&opcode.to_le_bytes());
                output.push(u8::from(enabled));
            }
            if output.len() > 255 {
                return Err(corrupted("edited CHPX exceeds one-byte FKP limit"));
            }
            Ok(output)
        })?;
        candidate.rewrite_chpx()?;
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// Same-length overwrite for a non-main Unicode story range. No CP, CLX,
    /// FKP, PLCF, or FIB length changes, so story-relative dependencies remain
    /// byte-identical.
    pub(crate) fn replace_unicode_text_same_length(
        &mut self,
        start: u32,
        end: u32,
        replacement: &str,
    ) -> Result<()> {
        if start >= end || end > self.pieces.last().map_or(0, |piece| piece.end) {
            return Err(corrupted("story text range is outside the piece table"));
        }
        let units = replacement.encode_utf16().collect::<Vec<_>>();
        if units.len() != (end - start) as usize {
            return Err(corrupted("story replacement changes UTF-16 length"));
        }
        let mut candidate = self.clone();
        let mut copied = 0usize;
        for piece in &candidate.pieces {
            let left = start.max(piece.start);
            let right = end.min(piece.end);
            if left >= right {
                continue;
            }
            if !piece.unicode {
                return Err(corrupted("story replacement intersects a compressed piece"));
            }
            let count = (right - left) as usize;
            let fc = piece
                .fc
                .checked_add(
                    (left - piece.start)
                        .checked_mul(2)
                        .ok_or_else(|| corrupted("story replacement FC offset overflow"))?,
                )
                .ok_or_else(|| corrupted("story replacement FC overflow"))?;
            let offset =
                usize::try_from(fc).map_err(|_| corrupted("story replacement FC exceeds usize"))?;
            let bytes = candidate
                .word
                .get_mut(offset..offset + count * 2)
                .ok_or_else(|| corrupted("story replacement exceeds WordDocument"))?;
            for (slot, unit) in bytes
                .as_chunks_mut::<2>()
                .0
                .iter_mut()
                .zip(units[copied..copied + count].iter().copied())
            {
                slot.copy_from_slice(&unit.to_le_bytes());
            }
            copied += count;
        }
        if copied != units.len() {
            return Err(corrupted("story replacement is not fully piece-covered"));
        }
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// Whether the paragraph ending at `cp` has the MS-DOC in-table flag.
    pub(crate) fn is_in_table_at_cp(&self, cp: u32) -> Result<bool> {
        let piece = containing_piece(&self.pieces, cp)
            .ok_or_else(|| corrupted("paragraph terminator has no text piece"))?;
        let width = if piece.unicode { 2 } else { 1 };
        let fc = piece
            .fc
            .checked_add(
                cp.checked_sub(piece.start)
                    .ok_or_else(|| corrupted("paragraph CP underflow"))?
                    .checked_mul(width)
                    .ok_or_else(|| corrupted("paragraph FC overflow"))?,
            )
            .ok_or_else(|| corrupted("paragraph FC overflow"))?;
        let run = containing_papx(&self.papx, fc)
            .ok_or_else(|| corrupted("paragraph terminator has no PAPX run"))?;
        Ok(run.in_table())
    }

    /// Resolves the exact Data-stream offset carried by
    /// `sprmCPicLocation` on one special picture/object character.
    #[deny(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        clippy::cast_precision_loss,
        clippy::cast_sign_loss,
        clippy::expect_used,
        clippy::let_underscore_must_use,
        clippy::map_err_ignore,
        clippy::unwrap_used,
        reason = "picture-location resolution is a checked legacy-codec boundary"
    )]
    pub(crate) fn picture_location_at_cp(&self, cp: u32) -> Result<u32> {
        let end = cp
            .checked_add(1)
            .ok_or_else(|| corrupted("picture CP overflow"))?;
        let intervals = self.fc_intervals(cp, end)?;
        if intervals.len() != 1 {
            return Err(corrupted(
                "picture character crosses physical text intervals",
            ));
        }
        let (start, finish) = intervals[0];
        let run = self
            .chpx
            .iter()
            .find(|run| run.start <= start && finish <= run.end)
            .ok_or_else(|| corrupted("picture character has no CHPX run"))?;
        let sprm = strict_sprms(&run.grpprl)?
            .into_iter()
            .rev()
            .find(|sprm| sprm.opcode == SPRM_C_PIC_LOCATION)
            .ok_or_else(|| corrupted("picture character has no sprmCPicLocation"))?;
        let bytes: [u8; 4] = sprm
            .operand_bytes()
            .try_into()
            .map_err(|error| corrupted(format!("sprmCPicLocation operand is invalid: {error}")))?;
        Ok(u32::from_le_bytes(bytes))
    }

    /// Extracts a canonical singleton picture graph rooted at one special
    /// character. The selected PICF/Data block is proved independently. A
    /// floating picture additionally proves its exact SPA, top-level shape,
    /// `pib`, and native `BStore` entry while leaving unrelated drawing nodes
    /// outside the transferred closure.
    #[deny(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        clippy::cast_precision_loss,
        clippy::cast_sign_loss,
        clippy::expect_used,
        clippy::let_underscore_must_use,
        clippy::map_err_ignore,
        clippy::unwrap_used,
        reason = "picture graph transfer is a checked legacy-codec boundary"
    )]
    pub(crate) fn picture_graph_at_cp(&self, cp: u32, floating: bool) -> Result<PictureGraph> {
        let expected_marker = if floating { '\u{0008}' } else { '\u{0001}' };
        let cp_end = cp
            .checked_add(1)
            .ok_or_else(|| corrupted("picture CP overflow"))?;
        if self.decode_text_range(cp, cp_end)? != expected_marker.to_string() {
            return Err(corrupted("selected CP is not the requested picture marker"));
        }
        if self.picture_character_is_object(cp)? {
            return Err(corrupted(
                "embedded-object previews must use embedded-object transfer",
            ));
        }
        let (picture, width, height, shape_id) = self.canonical_picture_at_cp(cp)?;
        let (selected_spa, nested_group) = if floating {
            let (spa_offset, spa_length) = fib_pair(&self.word, 40)?;
            let (header_spa_offset, header_spa_length) = fib_pair(&self.word, 41)?;
            if header_spa_length != 0 {
                let header_spa = slice(
                    &self.table,
                    header_spa_offset,
                    header_spa_length,
                    "PlcfSpaHdr",
                )?;
                drop(crate::parts::spa::parse_plcf_spa(header_spa)?);
            }
            if spa_length == 0 {
                return Err(corrupted("floating picture has no PlcfSpaMom"));
            }
            let spa_bytes = slice(&self.table, spa_offset, spa_length, "PlcfSpaMom")?;
            let anchors = crate::parts::spa::parse_plcf_spa(spa_bytes)?;
            let mut matches = anchors
                .iter()
                .enumerate()
                .filter(|(_, anchor)| anchor.cp == cp);
            let (anchor_index, anchor) = matches
                .next()
                .ok_or_else(|| corrupted("floating picture has no matching SPA anchor"))?;
            if matches.next().is_some() {
                return Err(corrupted("floating picture has ambiguous SPA ownership"));
            }
            if anchor.spa.shape_id != shape_id
                || anchor.spa.width()
                    != i32::try_from(width)
                        .map_err(|error| corrupted(format!("picture width exceeds i32: {error}")))?
                || anchor.spa.height()
                    != i32::try_from(height).map_err(|error| {
                        corrupted(format!("picture height exceeds i32: {error}"))
                    })?
            {
                return Err(corrupted("floating picture SPA does not match its PICF"));
            }
            let (dgg_offset, dgg_length) = fib_pair(&self.word, crate::shape::FIB_INDEX_DGG_INFO)?;
            if dgg_length == 0 {
                return Err(corrupted("floating picture has no DggInfo"));
            }
            let dgg_info = slice(&self.table, dgg_offset, dgg_length, "DggInfo")?;
            let nested =
                validate_floating_picture_identity(dgg_info, shape_id, anchor_index, &picture)?;
            (Some(anchor.spa), nested)
        } else {
            (None, None)
        };
        let rehomed_shape_id = if nested_group.is_some() {
            shape_id
        } else {
            crate::writer::images::FIRST_PICTURE_SHAPE_ID
        };
        let mut picture_block = Vec::new();
        crate::writer::images::write_picture_block(&picture, rehomed_shape_id, &mut picture_block)
            .map_err(|error| corrupted(format!("selected picture cannot be re-homed: {error}")))?;
        let (spa, dgg_info) = if let Some(mut spa) = selected_spa {
            spa.shape_id = rehomed_shape_id;
            let position = floating_position_from_spa(spa);
            let shape = crate::writer::images::FloatingShapeInfo {
                anchor_cp: cp,
                shape_id: spa.shape_id,
                content: crate::writer::images::FloatingShapeContent::Picture(&picture),
                width_twips: width,
                height_twips: height,
                position: &position,
                text: None,
            };
            let dgg = if let Some(group) = nested_group.as_ref() {
                build_nested_picture_dgg_info(&picture, spa, group)?
            } else {
                crate::writer::images::build_dgg_info(std::slice::from_ref(&shape), &[], 1)
                    .map_err(|error| {
                        corrupted(format!(
                            "selected drawing graph cannot be re-homed: {error}"
                        ))
                    })?
            };
            (Some(spa), dgg)
        } else {
            (None, Vec::new())
        };
        Ok(PictureGraph {
            floating,
            picture_block,
            spa,
            dgg_info,
            replaced_grpprl: None,
            data_offset: None,
        })
    }

    #[deny(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        clippy::cast_precision_loss,
        clippy::cast_sign_loss,
        clippy::expect_used,
        clippy::let_underscore_must_use,
        clippy::map_err_ignore,
        clippy::unwrap_used,
        reason = "PICF/Data canonicalization is a checked legacy-codec boundary"
    )]
    fn canonical_picture_at_cp(&self, cp: u32) -> Result<(crate::writer::Picture, u32, u32, u32)> {
        let picture_offset = self.picture_location_at_cp(cp)?;
        let block_start = usize::try_from(picture_offset)
            .map_err(|error| corrupted(format!("picture offset exceeds usize: {error}")))?;
        let fields = crate::image::PictureFields::try_parse(&self.data, block_start)
            .ok_or_else(|| corrupted("picture PICF is missing or truncated"))?;
        let picture_lcb = fields.lcb;
        let goal_width = fields.dxa_goal;
        let goal_height = fields.dya_goal;
        if picture_lcb <= 0
            || fields.cb_header != 0x44
            || fields.mm != 0x64
            || goal_width <= 0
            || goal_height <= 0
            || fields.mx != 1000
            || fields.my != 1000
            || fields.dxa_reserved1 != 0
            || fields.dya_reserved1 != 0
            || fields.dxa_reserved2 != 0
            || fields.dya_reserved2 != 0
            || fields.dxa_reserved3 != 0
            || fields.dya_reserved3 != 0
            || fields.c_props != 0
        {
            return Err(corrupted(
                "picture PICF has scaling, cropping, or producer extensions outside the transfer model",
            ));
        }
        let block_len = usize::try_from(picture_lcb)
            .map_err(|error| corrupted(format!("picture length exceeds usize: {error}")))?;
        let block_end = block_start
            .checked_add(block_len)
            .ok_or_else(|| corrupted("picture block extent overflow"))?;
        let picture_block = self
            .data
            .get(block_start..block_end)
            .ok_or_else(|| corrupted("picture block extends past the Data stream"))?;
        let shape_id = picture_shape_id(picture_block)?;
        let image = crate::image::Image::new(picture_offset)
            .data(&self.data, &self.word)
            .map_err(|error| corrupted(format!("picture BLIP is invalid: {error}")))?;
        if !matches!(
            image.kind(),
            litchi_odraw::image::Kind::Jpeg
                | litchi_odraw::image::Kind::Png
                | litchi_odraw::image::Kind::Dib
                | litchi_odraw::image::Kind::Tiff
        ) {
            return Err(corrupted(
                "picture BLIP kind is outside the native bitmap transfer model",
            ));
        }
        let native = image
            .data()
            .map_err(|error| corrupted(format!("picture BLIP payload is invalid: {error}")))?;
        if native.is_empty()
            || !picture_block
                .windows(native.len())
                .any(|candidate| candidate == native)
        {
            return Err(corrupted(
                "delay-loaded or external picture BLIPs cannot be transferred",
            ));
        }
        let width = u32::try_from(goal_width)
            .map_err(|error| corrupted(format!("picture width is invalid: {error}")))?;
        let height = u32::try_from(goal_height)
            .map_err(|error| corrupted(format!("picture height is invalid: {error}")))?;
        let picture =
            crate::writer::Picture::from_parts_as(native.to_vec(), image.kind(), width, height)
                .map_err(|error| {
                    corrupted(format!(
                        "picture cannot be represented canonically: {error}"
                    ))
                })?;
        let mut canonical_block = Vec::new();
        crate::writer::images::write_picture_block(&picture, shape_id, &mut canonical_block)
            .map_err(|error| corrupted(format!("picture graph cannot be encoded: {error}")))?;
        if canonical_block != picture_block {
            return Err(corrupted(
                "picture block is not the canonical OfficeArt graph for its shape ID",
            ));
        }
        Ok((picture, width, height, shape_id))
    }

    #[deny(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        clippy::cast_precision_loss,
        clippy::cast_sign_loss,
        clippy::expect_used,
        clippy::let_underscore_must_use,
        clippy::map_err_ignore,
        clippy::unwrap_used,
        reason = "picture object-flag inspection is a checked legacy-codec boundary"
    )]
    fn picture_character_is_object(&self, cp: u32) -> Result<bool> {
        let end = cp
            .checked_add(1)
            .ok_or_else(|| corrupted("picture CP overflow"))?;
        let intervals = self.fc_intervals(cp, end)?;
        let Some(&(start, finish)) = intervals.first().filter(|_| intervals.len() == 1) else {
            return Err(corrupted(
                "picture character crosses physical text intervals",
            ));
        };
        let run = self
            .chpx
            .iter()
            .find(|run| run.start <= start && finish <= run.end)
            .ok_or_else(|| corrupted("picture character has no CHPX run"))?;
        Ok(strict_sprms(&run.grpprl)?.iter().any(|sprm| {
            matches!(sprm.opcode, SPRM_C_F_OBJ | SPRM_C_F_OLE2) && sprm.operand_byte() == Some(1)
        }))
    }

    /// Whether the receiving artifact has no picture marker, SPA owner, or
    /// drawing group. This is the collision-free precondition for installing
    /// the singleton graph without rewriting arbitrary shape/BStore ids.
    #[deny(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        clippy::cast_precision_loss,
        clippy::cast_sign_loss,
        clippy::expect_used,
        clippy::let_underscore_must_use,
        clippy::map_err_ignore,
        clippy::unwrap_used,
        reason = "picture collision checks are a checked legacy-codec boundary"
    )]
    pub(crate) fn has_empty_picture_graph(&self) -> Result<bool> {
        for story in 0..7 {
            let (_origin, text) = self.story_text(story)?;
            if text
                .chars()
                .any(|character| matches!(character, '\u{0001}' | '\u{0008}'))
            {
                return Ok(false);
            }
        }
        Ok([40, 41, crate::shape::FIB_INDEX_DGG_INFO]
            .into_iter()
            .all(|index| fib_pair(&self.word, index).is_ok_and(|(_offset, length)| length == 0)))
    }

    /// Replaces one non-empty main-story placeholder with a canonical
    /// singleton inline or floating picture graph.
    #[deny(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        clippy::cast_precision_loss,
        clippy::cast_sign_loss,
        clippy::expect_used,
        clippy::let_underscore_must_use,
        clippy::map_err_ignore,
        clippy::unwrap_used,
        reason = "picture graph publication is a checked legacy-codec boundary"
    )]
    pub(crate) fn replace_with_picture_graph(
        &mut self,
        start: u32,
        end: u32,
        graph: &PictureGraph,
    ) -> Result<Option<PictureGraph>> {
        validate_range(start, end, self.main_ccp)?;
        self.reject_destructive_interactions(start, end)?;
        if !self.has_empty_picture_graph()? {
            return Err(corrupted("receiver picture graph is not empty"));
        }
        if !self.has_uniform_character_format(start, end)? {
            return Err(corrupted(
                "picture placeholder crosses character formatting runs",
            ));
        }
        graph.validate_rehomed()?;
        let replaced_grpprl = self
            .character_groups(start, end)?
            .first()
            .ok_or_else(|| corrupted("picture placeholder has no character formatting"))?
            .to_vec();
        let length_changed = end - start != 1;
        let saved_selection = if length_changed {
            self.prepared_saved_selection_splice(start, end, 1)?
        } else {
            None
        };
        if !self.unmodeled_cp_tables.is_empty() && length_changed {
            return Err(corrupted(
                "picture insertion changes length with unmodeled CP dependencies",
            ));
        }
        if graph.floating != graph.spa.is_some() || graph.floating == graph.dgg_info.is_empty() {
            return Err(corrupted(
                "picture graph closure is internally inconsistent",
            ));
        }

        let mut candidate = self.clone();
        let picture_offset = u32::try_from(candidate.data.len())
            .map_err(|error| corrupted(format!("Data stream exceeds u32: {error}")))?;
        candidate.data.extend_from_slice(&graph.picture_block);
        candidate.data_changed = true;
        let removed = end - start;
        delete_piece_range(&mut candidate.pieces, start, end)?;
        let fc = align2(candidate.word.len())?;
        candidate.word.resize(fc, 0);
        candidate
            .word
            .extend_from_slice(&if graph.floating { 8u16 } else { 1u16 }.to_le_bytes());
        insert_piece(
            &mut candidate.pieces,
            start,
            1,
            u32::try_from(fc)
                .map_err(|error| corrupted(format!("picture marker FC exceeds u32: {error}")))?,
        )?;
        let mut grpprl = Vec::with_capacity(10);
        grpprl.extend_from_slice(&SPRM_C_PIC_LOCATION.to_le_bytes());
        grpprl.extend_from_slice(&picture_offset.to_le_bytes());
        grpprl.extend_from_slice(&SPRM_C_F_SPEC.to_le_bytes());
        grpprl.push(1);
        candidate.chpx.push(FcRun {
            start: u32::try_from(fc)
                .map_err(|error| corrupted(format!("picture marker FC exceeds u32: {error}")))?,
            end: u32::try_from(candidate.word.len()).map_err(|error| {
                corrupted(format!("picture marker end FC exceeds u32: {error}"))
            })?,
            grpprl,
        });
        candidate.shift_cp_tables(start, removed, 1)?;
        candidate.main_ccp = candidate
            .main_ccp
            .checked_sub(removed)
            .and_then(|value| value.checked_add(1))
            .ok_or_else(|| corrupted("picture replacement CP overflow"))?;
        candidate.rewrite_chpx()?;
        candidate.append_clx_and_cp_tables()?;
        if let Some(spa) = graph.spa {
            let mut plcf = Vec::with_capacity(34);
            plcf.extend_from_slice(&start.to_le_bytes());
            plcf.extend_from_slice(&candidate.main_ccp.to_le_bytes());
            plcf.extend_from_slice(&spa.to_bytes());
            append_table_block(&mut candidate.word, &mut candidate.table, 40, &plcf)?;
            append_table_block(
                &mut candidate.word,
                &mut candidate.table,
                crate::shape::FIB_INDEX_DGG_INFO,
                &graph.dgg_info,
            )?;
        }
        let mut installed = graph.clone();
        installed.replaced_grpprl = Some(replaced_grpprl);
        installed.data_offset = Some(picture_offset);
        if graph.data_offset.is_some() && installed != *graph {
            return Ok(None);
        }
        candidate.patch_sizes()?;
        if let Some(replacement) = saved_selection {
            candidate.apply_saved_selection_splice(replacement)?;
        }
        candidate.commit()?;
        *self = candidate;
        Ok(Some(installed))
    }

    /// Removes one graph installed by [`Self::replace_with_picture_graph`] and
    /// restores its displaced inert text and character formatting. The Data
    /// block must still be the exact append-only tail and the SPA/Dgg graph
    /// must still be the exclusive canonical singleton.
    #[deny(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        clippy::cast_precision_loss,
        clippy::cast_sign_loss,
        clippy::expect_used,
        clippy::let_underscore_must_use,
        clippy::map_err_ignore,
        clippy::unwrap_used,
        reason = "durable picture reversal is a checked legacy-codec boundary"
    )]
    pub(crate) fn replace_picture_graph_with_text(
        &mut self,
        start: u32,
        end: u32,
        graph: &PictureGraph,
        replacement: &str,
    ) -> Result<()> {
        validate_range(start, end, self.main_ccp)?;
        if end - start != 1 {
            return Err(corrupted("installed picture marker is not one CP"));
        }
        let observed = self.picture_graph_at_cp(start, graph.floating)?;
        if !observed.same_wire_graph(graph) {
            return Err(corrupted("installed picture graph precondition changed"));
        }
        let data_offset = graph
            .data_offset
            .ok_or_else(|| corrupted("installed picture has no Data offset"))?;
        if self.picture_location_at_cp(start)? != data_offset {
            return Err(corrupted("installed picture Data offset changed"));
        }
        let data_start = usize::try_from(data_offset)
            .map_err(|error| corrupted(format!("picture Data offset exceeds usize: {error}")))?;
        let data_end = data_start
            .checked_add(graph.picture_block.len())
            .ok_or_else(|| corrupted("installed picture Data extent overflow"))?;
        if data_end != self.data.len()
            || self.data.get(data_start..data_end) != Some(graph.picture_block.as_slice())
        {
            return Err(corrupted(
                "installed picture is no longer the append-only Data tail",
            ));
        }
        let formatting = graph
            .replaced_grpprl
            .as_ref()
            .ok_or_else(|| corrupted("installed picture has no displaced formatting"))?;
        if formatting.len() > 255 {
            return Err(corrupted("restored picture formatting exceeds CHPX limit"));
        }
        let units = replacement.encode_utf16().collect::<Vec<_>>();
        if units.is_empty() || units.len() > MAX_TEXT_UNITS {
            return Err(corrupted("restored picture text length is invalid"));
        }
        let length_changed = units.len() != 1;
        let added = u32::try_from(units.len())
            .map_err(|error| corrupted(format!("restored text length exceeds u32: {error}")))?;
        let saved_selection = if length_changed {
            self.prepared_saved_selection_splice(start, end, added)?
        } else {
            None
        };
        if length_changed && !self.unmodeled_cp_tables.is_empty() {
            return Err(corrupted(
                "picture reversal changes length with unmodeled CP dependencies",
            ));
        }

        let mut candidate = self.clone();
        candidate.data.truncate(data_start);
        candidate.data_changed = true;
        delete_piece_range(&mut candidate.pieces, start, end)?;
        let fc = align2(candidate.word.len())?;
        candidate.word.resize(fc, 0);
        for unit in units {
            candidate.word.extend_from_slice(&unit.to_le_bytes());
        }
        let fc_u32 = u32::try_from(fc)
            .map_err(|error| corrupted(format!("restored text FC exceeds u32: {error}")))?;
        insert_piece(&mut candidate.pieces, start, added, fc_u32)?;
        candidate.chpx.push(FcRun {
            start: fc_u32,
            end: u32::try_from(candidate.word.len())
                .map_err(|error| corrupted(format!("restored text end FC exceeds u32: {error}")))?,
            grpprl: formatting.clone(),
        });
        if graph.floating {
            candidate.cp_tables.retain(|table| table.index != 40);
            put_fib_pair(&mut candidate.word, 40, 0, 0)?;
            put_fib_pair(&mut candidate.word, crate::shape::FIB_INDEX_DGG_INFO, 0, 0)?;
        }
        candidate.shift_cp_tables(start, 1, added)?;
        candidate.main_ccp = candidate
            .main_ccp
            .checked_sub(1)
            .and_then(|value| value.checked_add(added))
            .ok_or_else(|| corrupted("picture reversal CP overflow"))?;
        candidate.rewrite_chpx()?;
        candidate.append_clx_and_cp_tables()?;
        candidate.patch_sizes()?;
        if let Some(replacement) = saved_selection {
            candidate.apply_saved_selection_splice(replacement)?;
        }
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    /// Main-story length in MS-DOC CP (UTF-16 code-unit) coordinates.
    #[must_use]
    pub(crate) const fn main_story_cp_len(&self) -> u32 {
        self.main_ccp
    }

    /// Prepare the passive `Selsf` record for one main-story splice before a
    /// candidate is cloned. The record is remapped when each CP has a proven
    /// boundary mapping, including a CP at the point of a pure insertion, which
    /// keeps its value unless the record selects an inline picture or shape
    /// whose character the text is inserted in front of; positions strictly
    /// inside replaced or removed text, and text inserted at the first CP of a
    /// non-empty list-prefix selection, are returned as an ambiguity so the
    /// body facade can expose its typed dependency refusal.
    pub(crate) fn saved_selection_splice(
        &self,
        start: u32,
        end: u32,
        added: u32,
    ) -> std::result::Result<
        Option<[u8; crate::parts::saved_selection::SELSF_SIZE]>,
        SavedSelectionSpliceError,
    > {
        let fib =
            FileInformationBlock::parse(&self.word).map_err(SavedSelectionSpliceError::Invalid)?;
        let selection =
            SavedSelection::parse(&fib, &self.table).map_err(SavedSelectionSpliceError::Invalid)?;
        let Some(selection) = selection else {
            return Ok(None);
        };
        selection
            .validate_main_story_bound(self.main_ccp)
            .map_err(SavedSelectionSpliceError::Invalid)?;
        let removed = end.checked_sub(start).ok_or_else(|| {
            SavedSelectionSpliceError::Invalid(corrupted("Selsf splice range is reversed"))
        })?;
        let new_ccp = self
            .main_ccp
            .checked_sub(removed)
            .and_then(|value| value.checked_add(added))
            .ok_or_else(|| {
                SavedSelectionSpliceError::Invalid(corrupted("Selsf splice CP count overflows"))
            })?;
        let replacement = selection.remap_for_splice(start, end, added, new_ccp)?;
        if replacement == selection.bytes() {
            Ok(None)
        } else {
            Ok(Some(replacement))
        }
    }

    fn prepared_saved_selection_splice(
        &self,
        start: u32,
        end: u32,
        added: u32,
    ) -> Result<Option<[u8; crate::parts::saved_selection::SELSF_SIZE]>> {
        self.saved_selection_splice(start, end, added)
            .map_err(|error| match error {
                SavedSelectionSpliceError::Invalid(error) => error,
                SavedSelectionSpliceError::Ambiguous => {
                    corrupted("length-changing edit intersects an ambiguous Selsf CP")
                },
            })
    }

    fn apply_saved_selection_splice(
        &mut self,
        replacement: [u8; crate::parts::saved_selection::SELSF_SIZE],
    ) -> Result<()> {
        let fib = FileInformationBlock::parse(&self.word)?;
        let (start, end) = table_range(&fib, self.table.len())?;
        self.table[start..end].copy_from_slice(&replacement);
        Ok(())
    }

    fn character_groups(&self, start: u32, end: u32) -> Result<Vec<&[u8]>> {
        let intervals = self.fc_intervals(start, end)?;
        let mut output = Vec::new();
        for (interval_start, interval_end) in intervals {
            let mut cursor = interval_start;
            for run in &self.chpx {
                let left = interval_start.max(run.start);
                let right = interval_end.min(run.end);
                if left >= right {
                    continue;
                }
                if left > cursor {
                    return Err(corrupted("CHPX formatting has a physical FC gap"));
                }
                output.push(run.grpprl.as_slice());
                cursor = cursor.max(right);
            }
            if cursor < interval_end {
                return Err(corrupted("CHPX formatting does not cover body text"));
            }
        }
        if output.is_empty() {
            return Err(corrupted("body text has no CHPX formatting"));
        }
        Ok(output)
    }

    /// Lists character and PAPX property revisions, merging adjacent runs with
    /// identical metadata even when the range crosses piece boundaries.
    pub fn revisions(&self) -> Result<Vec<Revision>> {
        let mut output = Vec::new();
        for run in &self.chpx {
            let sprms = strict_sprms(&run.grpprl)?;
            for (kind, flag, author_op, time_op, reason_op, rsid_op) in [
                (
                    RevisionKind::Insertion,
                    SPRM_C_F_RMARK,
                    SPRM_C_IBST_RMARK,
                    SPRM_C_DTTM_RMARK,
                    SPRM_C_IDSL_RMARK,
                    SPRM_C_RSID_TEXT,
                ),
                (
                    RevisionKind::Deletion,
                    SPRM_C_F_RMARK_DEL,
                    SPRM_C_IBST_RMARK_DEL,
                    SPRM_C_DTTM_RMARK_DEL,
                    SPRM_C_IDSL_RMARK_DEL,
                    SPRM_C_RSID_RM_DEL,
                ),
            ] {
                if sprms
                    .iter()
                    .any(|s| s.opcode == flag && s.operand_byte() == Some(1))
                {
                    let metadata = metadata_from_sprms(
                        &sprms,
                        author_op,
                        time_op,
                        reason_op,
                        rsid_op,
                        &self.authors,
                    )?;
                    self.push_fc_revision(&mut output, run.start, run.end, kind, metadata)?;
                }
            }
            for opcode in [SPRM_C_PROP_RMARK90, SPRM_C_PROP_RMARK_CURRENT] {
                if let Some(mark) = sprms.iter().rev().find(|s| s.opcode == opcode) {
                    if mark.operand_bytes().first() == Some(&1) {
                        let metadata = property_metadata(
                            mark.operand_bytes(),
                            &sprms,
                            SPRM_C_RSID_PROP,
                            &self.authors,
                        )?;
                        self.push_fc_revision(
                            &mut output,
                            run.start,
                            run.end,
                            RevisionKind::CharacterFormatting,
                            metadata,
                        )?;
                    }
                    break;
                }
            }
        }
        for run in &self.papx {
            let body = run
                .grpprl
                .get(2..)
                .ok_or_else(|| corrupted("PAPX has no style index"))?;
            let sprms = strict_sprms(body)?;
            for (kind, op, rsid) in [
                (
                    RevisionKind::ParagraphFormatting,
                    [
                        SPRM_P_PROP_RMARK,
                        SPRM_P_PROP_RMARK90,
                        SPRM_P_PROP_RMARK_CURRENT,
                    ]
                    .as_slice(),
                    None,
                ),
                (
                    RevisionKind::TableRowFormatting,
                    [SPRM_T_PROP_RMARK].as_slice(),
                    Some(SPRM_T_RSID),
                ),
            ] {
                if let Some(mark) = sprms.iter().rev().find(|s| op.contains(&s.opcode))
                    && mark.operand_bytes().first() == Some(&1)
                {
                    let metadata = property_metadata(
                        mark.operand_bytes(),
                        &sprms,
                        rsid.unwrap_or(0),
                        &self.authors,
                    )?;
                    self.push_fc_revision(&mut output, run.start, run.end, kind, metadata)?;
                }
            }
        }
        output.sort_by_key(|r| (r.start_cp, r.end_cp, kind_order(r.kind)));
        merge_adjacent(&mut output);
        infer_moves(&mut output);
        Ok(output)
    }

    /// Inserts inert plain text and marks it as an insertion or move destination.
    /// Field delimiters, object markers, paragraph marks, and macro characters
    /// are rejected rather than interpreted.
    ///
    /// `cp` must lie before the main story's final paragraph mark, that is
    /// below `ccpText`: MS-DOC 2.3.1 requires the last character of the main
    /// document to be a paragraph mark, so text cannot follow it. To extend the
    /// last paragraph, insert at `ccpText - 1`, in front of its mark.
    pub fn add_text(
        &mut self,
        cp: u32,
        text: &str,
        kind: RevisionKind,
        metadata: RevisionMetadata,
    ) -> Result<Revision> {
        if !matches!(kind, RevisionKind::Insertion | RevisionKind::MoveTo) {
            return Err(corrupted(
                "add_text requires an insertion or move-to revision",
            ));
        }
        let units = text.encode_utf16().collect::<Vec<_>>();
        if units.is_empty()
            || units.len() > MAX_TEXT_UNITS
            || units.iter().any(|u| matches!(*u, 0..=8 | 11..=31 | 0xFFFC))
        {
            return Err(corrupted(
                "tracked text is empty, oversized, or contains an active control character",
            ));
        }
        if cp > self.main_ccp {
            return Err(corrupted("tracked insertion CP exceeds main story"));
        }
        if cp == self.main_ccp {
            return Err(corrupted(
                "tracked insertion CP follows the main story's final paragraph mark (MS-DOC 2.3.1)",
            ));
        }
        validate_metadata(kind, &metadata)?;
        let length =
            u32::try_from(units.len()).map_err(|_| corrupted("tracked text length exceeds u32"))?;
        let saved_selection = self.prepared_saved_selection_splice(cp, cp, length)?;
        let mut candidate = self.clone();
        let author = candidate.author_index(&metadata.author)?;
        let fc = align2(candidate.word.len())?;
        candidate.word.resize(fc, 0);
        for unit in &units {
            candidate.word.extend_from_slice(&unit.to_le_bytes());
        }
        insert_piece(
            &mut candidate.pieces,
            cp,
            length,
            u32::try_from(fc).map_err(|_| corrupted("FC exceeds u32"))?,
        )?;
        candidate.shift_cp_tables(cp, 0, length)?;
        candidate.main_ccp = candidate
            .main_ccp
            .checked_add(length)
            .ok_or_else(|| corrupted("main story CP overflow"))?;
        let grpprl = encode_revision(kind, author, &metadata)?;
        let fc_end =
            u32::try_from(candidate.word.len()).map_err(|_| corrupted("FC exceeds u32"))?;
        candidate.chpx.push(FcRun {
            start: fc as u32,
            end: fc_end,
            grpprl,
        });
        candidate.chpx.sort_by_key(|run| run.start);
        candidate.rewrite_chpx()?;
        candidate.enable_tracking()?;
        candidate.append_clx_and_cp_tables()?;
        candidate.patch_sizes()?;
        if let Some(replacement) = saved_selection {
            candidate.apply_saved_selection_splice(replacement)?;
        }
        candidate.commit()?;
        *self = candidate;
        self.find_exact(cp, cp + length, kind)
    }

    /// Adds a revision mark to an existing main-story range.
    pub fn add(
        &mut self,
        start_cp: u32,
        end_cp: u32,
        kind: RevisionKind,
        metadata: RevisionMetadata,
    ) -> Result<Revision> {
        validate_range(start_cp, end_cp, self.main_ccp)?;
        validate_metadata(kind, &metadata)?;
        let mut candidate = self.clone();
        let author = candidate.author_index(&metadata.author)?;
        candidate.mutate_mark(start_cp, end_cp, kind, Some((author, &metadata)))?;
        candidate.enable_tracking()?;
        candidate.commit()?;
        *self = candidate;
        self.find_exact(start_cp, end_cp, kind)
    }

    /// Replaces revision metadata without touching unrelated formatting SPRMs.
    pub fn update(&mut self, index: usize, metadata: RevisionMetadata) -> Result<Revision> {
        let revision = self
            .revisions()?
            .get(index)
            .cloned()
            .ok_or_else(|| corrupted("revision index is out of range"))?;
        validate_metadata(revision.kind, &metadata)?;
        let mut candidate = self.clone();
        let author = candidate.author_index(&metadata.author)?;
        candidate.mutate_mark(
            revision.start_cp,
            revision.end_cp,
            revision.kind,
            Some((author, &metadata)),
        )?;
        candidate.commit()?;
        *self = candidate;
        self.find_exact(revision.start_cp, revision.end_cp, revision.kind)
    }

    /// Removes a mark while retaining its text/current formatting.
    pub fn remove(&mut self, index: usize) -> Result<Revision> {
        let revision = self
            .revisions()?
            .get(index)
            .cloned()
            .ok_or_else(|| corrupted("revision index is out of range"))?;
        let mut candidate = self.clone();
        candidate.mutate_mark(revision.start_cp, revision.end_cp, revision.kind, None)?;
        candidate.commit()?;
        *self = candidate;
        Ok(revision)
    }

    /// Accepts a revision using Word redline semantics.
    pub fn accept(&mut self, index: usize) -> Result<Revision> {
        let revision = self
            .revisions()?
            .get(index)
            .cloned()
            .ok_or_else(|| corrupted("revision index is out of range"))?;
        if matches!(
            revision.kind,
            RevisionKind::Deletion | RevisionKind::MoveFrom
        ) {
            self.delete_revision_text(&revision)?;
        } else {
            self.remove(index)?;
        }
        Ok(revision)
    }

    /// Rejects a revision using Word redline semantics.
    pub fn reject(&mut self, index: usize) -> Result<Revision> {
        let revision = self
            .revisions()?
            .get(index)
            .cloned()
            .ok_or_else(|| corrupted("revision index is out of range"))?;
        if matches!(
            revision.kind,
            RevisionKind::Insertion | RevisionKind::MoveTo
        ) {
            self.delete_revision_text(&revision)?;
        } else if matches!(
            revision.kind,
            RevisionKind::CharacterFormatting
                | RevisionKind::ParagraphFormatting
                | RevisionKind::TableRowFormatting
        ) {
            let mut candidate = self.clone();
            candidate.reject_formatting_revision(&revision)?;
            candidate.commit()?;
            *self = candidate;
        } else {
            self.remove(index)?;
        }
        Ok(revision)
    }

    /// Finishes the source-backed DOC edit.
    ///
    /// The embedded [`ObjectEditor`] retains the exact opened CFB bytes and
    /// renders changed saves with its default
    /// [`litchi_cfb::SectorLayoutPolicy::Reuse`] policy. This is the ordinary tracked
    /// revision save route; callers do not need to opt into the CFB policy.
    /// A changed edit is authorized under the editor's protection policy, and
    /// a signed source cannot publish changed bytes.
    pub fn finish(self) -> Result<Vec<u8>> {
        if self.changed {
            self.protection_policy.authorize(self.protection)?;
        }
        let source = Arc::clone(&self.source);
        let signed = self.signed;
        let bytes = self.finish_unchecked()?;
        if signed && bytes.as_slice() != source.as_slice() {
            return Err(corrupted(
                "signed DOC cannot publish changed bytes through RevisionEditor",
            ));
        }
        Ok(bytes)
    }

    /// Renders the current package without the public signed-source guard.
    ///
    /// The body transaction uses this narrow seam so it can compare the
    /// rendered bytes with its retained source and preserve its typed
    /// [`crate::body_text::Refusal::SignedSource`] diagnostic. Callers of the
    /// public editor use [`Self::finish`], which always applies the guard.
    /// A validated render retained under the owner's ceiling is moved out
    /// instead of rendering the package again.
    pub(crate) fn finish_unchecked(mut self) -> Result<Vec<u8>> {
        if let Some(rendered) = self.retained_render.0.take() {
            return Ok(rendered);
        }
        self.package.finish().map_err(PackageError::from)
    }

    /// Apply this editor's protection policy to a rendered candidate while
    /// preserving exact final-byte no-ops for higher-level facades.
    pub(crate) fn authorize_rendered(&self, bytes: &[u8]) -> Result<()> {
        if bytes != self.source.as_slice() {
            self.protection_policy.authorize(self.protection)?;
        }
        Ok(())
    }

    pub(crate) fn protection_state(&self) -> EditProtection {
        self.protection
    }

    fn delete_revision_text(&mut self, revision: &Revision) -> Result<()> {
        self.reject_destructive_interactions(revision.start_cp, revision.end_cp)?;
        let saved_selection =
            self.prepared_saved_selection_splice(revision.start_cp, revision.end_cp, 0)?;
        let mut candidate = self.clone();
        delete_piece_range(&mut candidate.pieces, revision.start_cp, revision.end_cp)?;
        let removed = revision.end_cp - revision.start_cp;
        candidate.shift_cp_tables(revision.start_cp, removed, 0)?;
        candidate.main_ccp = candidate
            .main_ccp
            .checked_sub(removed)
            .ok_or_else(|| corrupted("main story CP underflow"))?;
        candidate.append_clx_and_cp_tables()?;
        candidate.patch_sizes()?;
        if let Some(replacement) = saved_selection {
            candidate.apply_saved_selection_splice(replacement)?;
        }
        candidate.commit()?;
        *self = candidate;
        Ok(())
    }

    fn reject_destructive_interactions(&self, start: u32, end: u32) -> Result<()> {
        let text = read_units(&self.word, &self.pieces, start, end)?;
        if text.iter().any(|u| matches!(*u, 0x13..=0x15)) {
            return Err(corrupted("accept/reject would delete a field boundary"));
        }
        for table in &self.cp_tables {
            for cp in &table.cps {
                if *cp > start && *cp < end {
                    return Err(corrupted(
                        "accept/reject would split a field, bookmark, or comment range",
                    ));
                }
            }
        }
        Ok(())
    }

    fn mutate_mark(
        &mut self,
        start: u32,
        end: u32,
        kind: RevisionKind,
        replacement: Option<(u16, &RevisionMetadata)>,
    ) -> Result<()> {
        let intervals = self.fc_intervals(start, end)?;
        match kind {
            RevisionKind::Insertion
            | RevisionKind::Deletion
            | RevisionKind::MoveFrom
            | RevisionKind::MoveTo
            | RevisionKind::CharacterFormatting => {
                split_transform_chpx(&mut self.chpx, &intervals, |grp| {
                    replace_revision_sprms(grp, kind, replacement)
                })?;
                self.rewrite_chpx()?;
            },
            RevisionKind::ParagraphFormatting | RevisionKind::TableRowFormatting => {
                split_transform_papx(&mut self.papx, &intervals, |grp| {
                    replace_papx_revision_sprms(grp, kind, replacement)
                })?;
                self.rewrite_papx()?;
            },
        }
        if let Some((_, metadata)) = replacement
            && !self.authors.iter().any(|a| a == &metadata.author)
        {
            return Err(corrupted("revision author indexing failed"));
        }
        Ok(())
    }

    fn reject_formatting_revision(&mut self, revision: &Revision) -> Result<()> {
        let intervals = self.fc_intervals(revision.start_cp, revision.end_cp)?;
        match revision.kind {
            RevisionKind::CharacterFormatting => {
                split_transform_chpx(&mut self.chpx, &intervals, |grp| {
                    restore_before_wall(
                        grp,
                        SPRM_C_WALL,
                        &revision_opcodes(RevisionKind::CharacterFormatting, true),
                    )
                })?;
                self.rewrite_chpx()
            },
            RevisionKind::ParagraphFormatting => {
                split_transform_papx(&mut self.papx, &intervals, |grp| {
                    let style = grp
                        .get(..2)
                        .ok_or_else(|| corrupted("PAPX style index is truncated"))?;
                    let mut restored = style.to_vec();
                    restored.extend_from_slice(&restore_before_wall(
                        &grp[2..],
                        SPRM_P_WALL,
                        &revision_opcodes(RevisionKind::ParagraphFormatting, true),
                    )?);
                    Ok(restored)
                })?;
                self.rewrite_papx()
            },
            RevisionKind::TableRowFormatting => {
                split_transform_papx(&mut self.papx, &intervals, |grp| {
                    let style = grp
                        .get(..2)
                        .ok_or_else(|| corrupted("PAPX style index is truncated"))?;
                    let mut restored = style.to_vec();
                    restored.extend_from_slice(&restore_before_wall(
                        &grp[2..],
                        SPRM_T_WALL,
                        &revision_opcodes(RevisionKind::TableRowFormatting, true),
                    )?);
                    Ok(restored)
                })?;
                self.rewrite_papx()
            },
            _ => Err(corrupted("revision is not a formatting revision")),
        }
    }

    fn fc_intervals(&self, start: u32, end: u32) -> Result<Vec<(u32, u32)>> {
        let mut output = Vec::new();
        for piece in &self.pieces {
            let left = start.max(piece.start);
            let right = end.min(piece.end);
            if left >= right {
                continue;
            }
            let scale = if piece.unicode { 2 } else { 1 };
            let fc_start = piece
                .fc
                .checked_add((left - piece.start) * scale)
                .ok_or_else(|| corrupted("FC overflow"))?;
            let fc_end = piece
                .fc
                .checked_add((right - piece.start) * scale)
                .ok_or_else(|| corrupted("FC overflow"))?;
            output.push((fc_start, fc_end));
        }
        if output.is_empty() {
            return Err(corrupted("revision range has no text pieces"));
        }
        Ok(output)
    }

    fn push_fc_revision(
        &self,
        output: &mut Vec<Revision>,
        fc_start: u32,
        fc_end: u32,
        kind: RevisionKind,
        metadata: ParsedMetadata,
    ) -> Result<()> {
        for piece in &self.pieces {
            let width = if piece.unicode { 2 } else { 1 };
            let piece_fc_end = piece
                .fc
                .checked_add((piece.end - piece.start) * width)
                .ok_or_else(|| corrupted("piece FC overflow"))?;
            let left = fc_start.max(piece.fc);
            let right = fc_end.min(piece_fc_end);
            if left >= right {
                continue;
            }
            if (left - piece.fc) % width != 0 || (right - piece.fc) % width != 0 {
                return Err(corrupted("CHPX boundary splits a text character"));
            }
            let start_cp = piece.start + (left - piece.fc) / width;
            let end_cp = piece.start + (right - piece.fc) / width;
            if start_cp < self.main_ccp {
                output.push(metadata.to_revision(kind, start_cp, end_cp.min(self.main_ccp)));
            }
        }
        Ok(())
    }

    fn find_exact(&self, start: u32, end: u32, kind: RevisionKind) -> Result<Revision> {
        self.revisions()?
            .into_iter()
            .find(|r| {
                r.start_cp == start
                    && r.end_cp == end
                    && (r.kind == kind
                        || matches!(
                            (r.kind, kind),
                            (RevisionKind::Insertion, RevisionKind::MoveTo)
                                | (RevisionKind::Deletion, RevisionKind::MoveFrom)
                        ))
            })
            .ok_or_else(|| corrupted("authored revision was not discoverable"))
    }

    fn author_index(&mut self, author: &str) -> Result<u16> {
        if author.is_empty() || author.encode_utf16().count() > u16::MAX as usize {
            return Err(corrupted("revision author is empty or too long"));
        }
        if self.authors.is_empty() {
            self.authors.push("Unknown".to_string());
        }
        if let Some(index) = self.authors.iter().position(|value| value == author) {
            return u16::try_from(index)
                .map_err(|_| corrupted("revision author index exceeds u16"));
        }
        if self.authors.len() >= MAX_AUTHORS {
            return Err(corrupted("revision author limit exceeded"));
        }
        self.authors.push(author.to_string());
        let bytes = serialize_authors(&self.authors)?;
        append_table_block(&mut self.word, &mut self.table, STTBFRMARK, &bytes)?;
        Ok((self.authors.len() - 1) as u16)
    }

    fn rewrite_chpx(&mut self) -> Result<()> {
        if self.chpx.is_empty() {
            return Err(corrupted("CHPX table has no runs"));
        }
        self.chpx.sort_by_key(|r| r.start);
        if self
            .chpx
            .iter()
            .any(|r| r.start >= r.end || r.grpprl.len() > 255)
        {
            return Err(corrupted("CHPX run is invalid or exceeds FKP limits"));
        }
        let mut builder = ChpxFkpBuilder::new();
        for run in &self.chpx {
            builder.add_entry(run.start, run.end, run.grpprl.clone());
        }
        let pages = builder.generate_pages().map_err(PackageError::from)?;
        let base = align512(self.word.len())?;
        self.word.resize(base, 0);
        let mut pns = Vec::new();
        for page in &pages.pages {
            pns.push(
                u32::try_from(self.word.len() / 512)
                    .map_err(|_| corrupted("CHPX page number exceeds u32"))?,
            );
            self.word.extend_from_slice(page);
        }
        let mut plc = Vec::new();
        for (start, _) in &pages.ranges {
            plc.extend_from_slice(&start.to_le_bytes());
        }
        plc.extend_from_slice(
            &pages
                .ranges
                .last()
                .ok_or_else(|| corrupted("CHPX page list is empty"))?
                .1
                .to_le_bytes(),
        );
        for pn in pns {
            plc.extend_from_slice(&pn.to_le_bytes());
        }
        append_table_block(&mut self.word, &mut self.table, PLCFBTE_CHPX, &plc)
    }

    fn rewrite_papx(&mut self) -> Result<()> {
        self.papx.sort_by_key(|r| r.start);
        let pages = build_papx_pages(&self.papx)?;
        let base = align512(self.word.len())?;
        self.word.resize(base, 0);
        let mut plc = Vec::new();
        for page in &pages {
            plc.extend_from_slice(&page.start.to_le_bytes());
        }
        plc.extend_from_slice(
            &pages
                .last()
                .ok_or_else(|| corrupted("PAPX page list is empty"))?
                .end
                .to_le_bytes(),
        );
        for page in pages {
            let pn = u32::try_from(self.word.len() / 512)
                .map_err(|_| corrupted("PAPX page number exceeds u32"))?;
            plc.extend_from_slice(&pn.to_le_bytes());
            self.word.extend_from_slice(&page.bytes);
        }
        append_table_block(&mut self.word, &mut self.table, PLCFBTE_PAPX, &plc)
    }

    fn shift_cp_tables(&mut self, start: u32, removed: u32, added: u32) -> Result<()> {
        let end = start
            .checked_add(removed)
            .ok_or_else(|| corrupted("CP range overflow"))?;
        for table in &mut self.cp_tables {
            for cp in &mut table.cps {
                *cp = if removed == 0 {
                    if *cp >= start {
                        cp.checked_add(added)
                            .ok_or_else(|| corrupted("PLCF CP overflow"))?
                    } else {
                        *cp
                    }
                } else if *cp <= start {
                    *cp
                } else if *cp >= end {
                    cp.checked_sub(removed)
                        .and_then(|v| v.checked_add(added))
                        .ok_or_else(|| corrupted("PLCF CP shift overflow"))?
                } else {
                    start
                        .checked_add(added)
                        .ok_or_else(|| corrupted("PLCF CP overflow"))?
                };
            }
            if table.cps.windows(2).any(|v| v[0] > v[1]) {
                return Err(corrupted("PLCF CPs became non-monotonic"));
            }
        }
        Ok(())
    }

    fn append_clx_and_cp_tables(&mut self) -> Result<()> {
        let clx = serialize_clx(&self.pieces)?;
        append_table_block(&mut self.word, &mut self.table, CLX, &clx)?;
        for table in &self.cp_tables {
            let mut bytes = Vec::new();
            for cp in &table.cps {
                bytes.extend_from_slice(&cp.to_le_bytes());
            }
            bytes.extend_from_slice(&table.records);
            append_table_block(&mut self.word, &mut self.table, table.index, &bytes)?;
        }
        Ok(())
    }

    fn enable_tracking(&mut self) -> Result<()> {
        let (offset, length) = fib_pair(&self.word, DOP)?;
        if length < 84 {
            return Err(corrupted("DOP is too short to enable revision tracking"));
        }
        let mut dop = slice(&self.table, offset, length, "DOP")?.to_vec();
        dop[5] |= 0x80;
        append_table_block(&mut self.word, &mut self.table, DOP, &dop)
    }

    fn patch_sizes(&mut self) -> Result<()> {
        let word_len =
            u32::try_from(self.word.len()).map_err(|_| corrupted("WordDocument exceeds u32"))?;
        put_u32(&mut self.word, FIB_CCP_TEXT, self.main_ccp)?;
        put_u32(&mut self.word, 28, word_len)?;
        put_u32(&mut self.word, 64, word_len)
    }

    fn commit(&mut self) -> Result<()> {
        self.protection_policy.authorize(self.protection)?;
        // The package mutation below invalidates every render for the prior
        // editor state, including the Data-stream add path. Mutation callers
        // are clone-first, but clearing here keeps the ownership rule local
        // to the sole package publication boundary as well.
        self.release_retained_render();
        let mut replacements: SmallVec<[(&[String], Arc<[u8]>); 3]> = SmallVec::new();
        replacements.push((
            self.word_path.as_slice(),
            Arc::<[u8]>::from(self.word.clone()),
        ));
        replacements.push((
            self.table_path.as_slice(),
            Arc::<[u8]>::from(self.table.clone()),
        ));
        if self.data_changed {
            if self.package.stream(&self.data_path).is_some() {
                replacements.push((
                    self.data_path.as_slice(),
                    Arc::<[u8]>::from(self.data.clone()),
                ));
            } else {
                self.package
                    .add_stream(self.data_path.clone(), self.data.clone())
                    .map_err(PackageError::from)?;
            }
        }
        let rendered = self
            .package
            .put_streams_shared_with_rendered(replacements)
            .map_err(PackageError::from)?;
        self.retained_render.0 = rendered.filter(|bytes| {
            self.retained_render_limit != 0 && bytes.capacity() <= self.retained_render_limit
        });
        self.changed = true;
        Ok(())
    }
}

fn package_has_binary_signature(ole: &mut OleFile<Cursor<Vec<u8>>>) -> bool {
    let Some(stream_name) = ole.list_directory_entries(&[]).ok().and_then(|entries| {
        entries
            .into_iter()
            .find(|entry| directory_names_equal(&entry.name, "\u{0005}DocumentSummaryInformation"))
            .map(|entry| entry.name.clone())
    }) else {
        return false;
    };
    let Ok(bytes) = ole.open_stream(&[stream_name.as_str()]) else {
        return true;
    };
    let Ok(stream) = PropertySetStream::parse(&bytes) else {
        // Preserve the established permissive reader admission. An opaque or
        // malformed PIDDSI cannot prove that a changed DOFR edit is safe, so
        // classify it conservatively as signed for the mutation gate.
        return true;
    };
    stream
        .section(DOCUMENT_SUMMARY_INFORMATION_FMTID)
        .is_some_and(|section| section.property(DIGITAL_SIGNATURE).is_some())
}

const FIB_INDEX_STW_USER: usize = 60;
const MAX_STW_USER_BYTES: usize = 16 * 1024 * 1024;

/// Detect a valid legacy Word VBA signature variable without making optional
/// `StwUser` metadata part of ordinary DOC admission. A malformed selected
/// table is classified conservatively so a changed DOFR edit cannot silently
/// invalidate an opaque signature value; exact no-op edits still pass the
/// normal source checks before this marker is consulted.
fn word_vba_signature_state(word: &[u8], table: &[u8]) -> bool {
    let Ok(fib) = FileInformationBlock::parse(word) else {
        return true;
    };
    let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX_STW_USER) else {
        return false;
    };
    if length == 0 {
        return false;
    }
    let Ok(offset) = usize::try_from(offset) else {
        return true;
    };
    let Ok(length) = usize::try_from(length) else {
        return true;
    };
    let Some(end) = offset.checked_add(length) else {
        return true;
    };
    let Some(data) = table.get(offset..end) else {
        return true;
    };
    stw_user_contains_signature(data).unwrap_or(true)
}

fn stw_user_contains_signature(data: &[u8]) -> Result<bool> {
    if data.len() > MAX_STW_USER_BYTES {
        return Err(corrupted("StwUser exceeds its bounded table size"));
    }
    if u16_at(data, 0)? != 0xFFFF {
        return Err(corrupted("StwUser fExtend is not 0xFFFF"));
    }
    let count = usize::from(u16_at(data, 2)?);
    if u16_at(data, 4)? != 4 {
        return Err(corrupted("StwUser cbExtra is not 4"));
    }

    let mut signature_names = Vec::new();
    signature_names
        .try_reserve_exact(count)
        .map_err(|error| corrupted(format!("StwUser name allocation failed: {error}")))?;
    let mut seen_names = HashSet::<&[u8]>::new();
    seen_names
        .try_reserve(count)
        .map_err(|error| corrupted(format!("StwUser name allocation failed: {error}")))?;
    let mut cursor = 6usize;
    for _ in 0..count {
        let cch = usize::from(u16_at(data, cursor)?);
        let chars_start = cursor
            .checked_add(2)
            .ok_or_else(|| corrupted("StwUser name offset overflows"))?;
        let chars_len = cch
            .checked_mul(2)
            .ok_or_else(|| corrupted("StwUser name length overflows"))?;
        let chars_end = chars_start
            .checked_add(chars_len)
            .ok_or_else(|| corrupted("StwUser name range overflows"))?;
        let chars = data
            .get(chars_start..chars_end)
            .ok_or_else(|| corrupted("StwUser name is truncated"))?;
        if !seen_names.insert(chars) {
            return Err(corrupted("StwUser names are not unique"));
        }
        let is_signature = ["Sign", "SigAgile", "SigV3"]
            .into_iter()
            .any(|name| utf16_name_equals(chars, name));
        signature_names.push(is_signature);
        cursor = chars_end
            .checked_add(4)
            .ok_or_else(|| corrupted("StwUser name extra offset overflows"))?;
        if cursor > data.len() {
            return Err(corrupted("StwUser name extra is truncated"));
        }
    }

    let mut signed = false;
    for is_signature in signature_names {
        let cch = usize::from(u16_at(data, cursor)?);
        let chars_start = cursor
            .checked_add(2)
            .ok_or_else(|| corrupted("StwUser value offset overflows"))?;
        let chars_len = cch
            .checked_mul(2)
            .ok_or_else(|| corrupted("StwUser value length overflows"))?;
        let chars_end = chars_start
            .checked_add(chars_len)
            .ok_or_else(|| corrupted("StwUser value range overflows"))?;
        let value = data
            .get(cursor..chars_end)
            .ok_or_else(|| corrupted("StwUser value is truncated"))?;
        if is_signature {
            vba_signature::Snapshot::parse_word(value)
                .map_err(|error| corrupted(format!("invalid StwUser WordSigBlob: {error}")))?;
            signed = true;
        }
        cursor = chars_end;
    }
    if cursor != data.len() {
        return Err(corrupted("StwUser has trailing bytes"));
    }
    Ok(signed)
}

fn utf16_name_equals(bytes: &[u8], expected: &str) -> bool {
    let mut units = expected.encode_utf16();
    let mut chunks = bytes.chunks_exact(2);
    chunks.all(|bytes| {
        units
            .next()
            .is_some_and(|unit| u16::from_le_bytes([bytes[0], bytes[1]]) == unit)
    }) && chunks.remainder().is_empty()
        && units.next().is_none()
}

#[cfg(test)]
mod containment_tests {
    use super::*;

    fn piece(start: u32, end: u32) -> RawPiece {
        RawPiece {
            start,
            end,
            fc: start,
            unicode: false,
            prefix: [0; 2],
            prm: [0; 2],
        }
    }

    fn papx(start: u32, end: u32) -> PapxRun {
        PapxRun::new(
            start,
            end,
            vec![0; 2],
            false,
            crate::parts::fkp::ParagraphHeight {
                info_field: 0,
                reserved: 0,
                dxa_col: 0,
                dym_line_or_height: 0,
            },
        )
    }

    #[test]
    fn indexed_piece_containment_matches_scalar_boundaries() {
        let cases = [
            Vec::new(),
            vec![piece(0, 1)],
            vec![piece(0, 4), piece(4, 8), piece(8, 12)],
            vec![piece(2, 5), piece(9, 11), piece(u32::MAX - 2, u32::MAX)],
        ];
        let points = [
            0,
            1,
            2,
            3,
            4,
            5,
            7,
            8,
            9,
            10,
            11,
            12,
            u32::MAX - 2,
            u32::MAX - 1,
            u32::MAX,
        ];
        for pieces in &cases {
            for cp in points {
                let scalar = pieces
                    .iter()
                    .position(|piece| piece.start <= cp && cp < piece.end);
                let indexed = containing_piece(pieces, cp).and_then(|candidate| {
                    pieces
                        .iter()
                        .position(|piece| std::ptr::eq(piece, candidate))
                });
                assert_eq!(indexed, scalar, "piece containment differs at CP {cp}");
            }
        }
    }

    #[test]
    fn indexed_papx_containment_matches_scalar_boundaries() {
        let cases = [
            Vec::new(),
            vec![papx(0, 1)],
            vec![papx(0, 4), papx(4, 8), papx(8, 12)],
            vec![papx(2, 5), papx(9, 11), papx(u32::MAX - 2, u32::MAX)],
        ];
        let points = [
            0,
            1,
            2,
            3,
            4,
            5,
            7,
            8,
            9,
            10,
            11,
            12,
            u32::MAX - 2,
            u32::MAX - 1,
            u32::MAX,
        ];
        for runs in &cases {
            for fc in points {
                let scalar = runs.iter().position(|run| run.start <= fc && fc < run.end);
                let indexed = containing_papx(runs, fc)
                    .and_then(|candidate| runs.iter().position(|run| std::ptr::eq(run, candidate)));
                assert_eq!(indexed, scalar, "PAPX containment differs at FC {fc}");
            }
        }
    }
}

#[cfg(test)]
mod papx_cache_tests {
    use super::*;
    use crate::sprm_operations::SPRM_P_F_IN_TABLE;
    use crate::writer::{CharacterFormatting, ParagraphFormatting, Writer};
    use std::io::Cursor;

    fn table_sprm(value: u8) -> [u8; 3] {
        let opcode = SPRM_P_F_IN_TABLE.to_le_bytes();
        [opcode[0], opcode[1], value]
    }

    fn doc_with_table_state(in_table: bool) -> Vec<u8> {
        let mut writer = Writer::new();
        writer
            .add_paragraph_runs(
                vec![("cell".to_owned(), CharacterFormatting::default())],
                ParagraphFormatting {
                    in_table: Some(in_table),
                    ..ParagraphFormatting::default()
                },
            )
            .expect("table-state fixture paragraph should be valid");
        let mut output = Cursor::new(Vec::new());
        writer
            .write_to(&mut output)
            .expect("table-state fixture should serialize");
        output.into_inner()
    }

    fn simple_table_doc() -> Vec<u8> {
        let mut writer = Writer::new();
        let table = writer
            .add_table(1, 1)
            .expect("table fixture should be accepted");
        writer
            .set_table_cell_text(table, 0, 0, "cached cell")
            .expect("table cell fixture should be accepted");
        let mut output = Cursor::new(Vec::new());
        writer
            .write_to(&mut output)
            .expect("table fixture should serialize");
        output.into_inner()
    }

    #[test]
    fn retained_render_clone_recomputes_and_finish_moves_the_allocation() {
        let mut editor = RevisionEditor::open(doc_with_table_state(false), Limits::default())
            .expect("source opens");
        assert_eq!(editor.retained_render_capacity(), None);
        editor.set_retained_render_limit(8 * 1024 * 1024);
        editor
            .replace_plain_text(0, 4, "replacement")
            .expect("edit");
        let retained = editor.retained_render.0.as_ref().expect("held render");
        let pointer = retained.as_ptr();
        let capacity = retained.capacity();
        let clone = editor.clone();
        assert_eq!(clone.retained_render_capacity(), None);
        assert_eq!(clone.retained_render_limit, editor.retained_render_limit);
        let recomputed = clone.finish().expect("clone recomputes");
        let moved = editor.finish().expect("original consumes handoff");
        assert_eq!(moved.as_ptr(), pointer);
        assert_eq!(moved.capacity(), capacity);
        assert_eq!(moved, recomputed);
    }

    #[test]
    fn editor_table_queries_use_cached_state_and_track_edits() {
        let mut editor = RevisionEditor::open(doc_with_table_state(true), Limits::default())
            .expect("table-state fixture should open");
        let terminator_cp = editor.main_story_cp_len() - 1;
        assert!(editor.is_in_table_at_cp(terminator_cp).unwrap());

        let original = editor.papx[0].grpprl.clone();
        let intervals = editor.fc_intervals(0, terminator_cp + 1).unwrap();
        split_transform_papx(&mut editor.papx, &intervals, |grpprl| {
            let mut transformed = grpprl.to_vec();
            transformed.extend_from_slice(&table_sprm(0));
            Ok(transformed)
        })
        .unwrap();
        assert!(!editor.is_in_table_at_cp(terminator_cp).unwrap());
        assert_eq!(
            editor.papx[0].grpprl,
            [original, table_sprm(0).to_vec()].concat()
        );
        assert!(!editor.papx[0].in_table());
    }

    #[test]
    fn editor_source_shares_the_common_admitted_allocation() {
        let editor = RevisionEditor::open(doc_with_table_state(true), Limits::default())
            .expect("table-state fixture should open");
        assert!(Arc::ptr_eq(&editor.source, &editor.package.source_shared()));
    }

    #[test]
    fn editor_table_queries_preserve_absent_state_and_malformed_open_errors() {
        let mut editor = RevisionEditor::open(doc_with_table_state(false), Limits::default())
            .expect("non-table fixture should open");
        let terminator_cp = editor.main_story_cp_len() - 1;
        assert!(!editor.is_in_table_at_cp(terminator_cp).unwrap());

        let before = editor.papx[0].clone();
        let intervals = editor.fc_intervals(0, terminator_cp + 1).unwrap();
        assert!(
            split_transform_papx(&mut editor.papx, &intervals, |_| {
                Ok(vec![0, 0, table_sprm(1)[0]])
            })
            .is_err()
        );
        assert_eq!(editor.papx[0].grpprl, before.grpprl);
        assert_eq!(editor.papx[0].in_table(), before.in_table());
    }

    #[test]
    fn public_table_cell_extraction_remains_indirectly_regressed() {
        let snapshot = crate::body_text::Snapshot::parse(&simple_table_doc())
            .expect("table fixture should be readable");
        let cells = snapshot
            .table_cells()
            .expect("table cell extraction should succeed");
        assert!(cells.iter().any(|cell| cell.text() == "cached cell"));
    }
}

#[cfg(test)]
mod picture_group_tests {
    use super::*;
    use crate::parts::spa::{
        ShapeHorizontalOrigin, ShapeTextWrap, ShapeVerticalOrigin, ShapeWrapSide,
    };

    fn normalize_word2002_dop(bytes: Vec<u8>) -> Vec<u8> {
        let mut package = ObjectEditor::open(bytes, Targets::default(), Limits::default()).unwrap();
        let word_path = ["WordDocument".to_string()];
        let word = package.stream(&word_path).unwrap();
        let fib = FileInformationBlock::parse(word).unwrap();
        let table_name = if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        };
        let table_path = [table_name.to_string()];
        let mut word = word.to_vec();
        let mut table = package.stream(&table_path).unwrap().to_vec();
        let (offset, length) = fib.get_table_pointer(31).unwrap();
        let offset = usize::try_from(offset).unwrap();
        let length = usize::try_from(length).unwrap();
        let dop = crate::parts::document_properties::DocumentProperties::writer_bytes(
            false, false, false, true,
        );
        if length < dop.len() {
            let insertion = offset + length;
            let extra = dop.len() - length;
            table.splice(insertion..insertion, std::iter::repeat_n(0, extra));
            let count = fib.table_pointer_count().unwrap();
            for index in 0..count {
                let pointer = 154 + index * 8;
                let current = usize::try_from(u32::from_le_bytes(
                    word[pointer..pointer + 4].try_into().unwrap(),
                ))
                .unwrap();
                if current >= insertion {
                    let shifted = u32::try_from(current + extra).unwrap();
                    word[pointer..pointer + 4].copy_from_slice(&shifted.to_le_bytes());
                }
            }
        }
        table[offset..offset + dop.len()].copy_from_slice(&dop);
        let pointer = 154 + 31 * 8;
        word[pointer + 4..pointer + 8].copy_from_slice(&(dop.len() as u32).to_le_bytes());
        package.put_stream(&word_path, word).unwrap();
        package.put_stream(&table_path, table).unwrap();
        package.finish().unwrap()
    }

    fn record(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        crate::writer::images::write_record_header(
            &mut output,
            version,
            instance,
            kind,
            u32::try_from(payload.len()).unwrap(),
        );
        output.extend_from_slice(payload);
        output
    }

    fn shape_container(children: &[Vec<u8>]) -> Vec<u8> {
        let payload = children.concat();
        record(0x0f, 0, 0xF004, &payload)
    }

    fn nested_group(shape_id: u32, picture_id: u32, width: i32, height: i32) -> Vec<u8> {
        let mut bounds = Vec::new();
        for value in [0, 0, width, height] {
            bounds.extend_from_slice(&value.to_le_bytes());
        }
        let mut group_shape = Vec::new();
        group_shape.extend_from_slice(&shape_id.to_le_bytes());
        group_shape.extend_from_slice(&0x0201u32.to_le_bytes());
        let group_meta = shape_container(&[
            record(1, 0, 0xF009, &bounds),
            record(2, 0, 0xF00A, &group_shape),
            record(0, 0, 0xF010, &0u32.to_le_bytes()),
            record(0, 0, 0xF011, &0u32.to_le_bytes()),
        ]);

        let mut picture_shape = Vec::new();
        picture_shape.extend_from_slice(&picture_id.to_le_bytes());
        picture_shape.extend_from_slice(&0x0A02u32.to_le_bytes());
        let mut opt = Vec::new();
        crate::writer::images::write_opt_record(&mut opt, &[(0x4104, 1)]);
        let picture = shape_container(&[
            record(2, 75, 0xF00A, &picture_shape),
            opt,
            record(0, 0, 0xF00F, &bounds),
            record(0, 0, 0xF011, &0u32.to_le_bytes()),
        ]);
        let payload = [group_meta, picture].concat();
        record(0x0f, 0, 0xF003, &payload)
    }

    #[test]
    fn nested_single_image_group_is_a_closed_rehomed_graph() {
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/images/png/lena.png"),
        )
        .unwrap();
        let picture = crate::writer::Picture::new(bytes).unwrap();
        let shape_id = crate::writer::images::FIRST_PICTURE_SHAPE_ID;
        let spa = crate::parts::spa::Spa {
            shape_id,
            left: 120,
            top: 240,
            right: 120 + i32::try_from(picture.width_twips()).unwrap(),
            bottom: 240 + i32::try_from(picture.height_twips()).unwrap(),
            horizontal_origin: ShapeHorizontalOrigin::Page,
            vertical_origin: ShapeVerticalOrigin::Page,
            wrap: ShapeTextWrap::Square,
            wrap_side: ShapeWrapSide::Both,
            below_text: false,
            anchor_locked: true,
        };
        let group = NestedPictureGroup {
            bytes: nested_group(shape_id, shape_id + 1, spa.width(), spa.height()),
            shape_ids: vec![shape_id, shape_id + 1],
        };
        let dgg_info = build_nested_picture_dgg_info(&picture, spa, &group).unwrap();
        let normalized =
            validate_floating_picture_identity(&dgg_info, shape_id, 0, &picture).unwrap();
        assert!(normalized.is_some());

        let mut picture_block = Vec::new();
        crate::writer::images::write_picture_block(&picture, shape_id, &mut picture_block).unwrap();
        PictureGraph {
            floating: true,
            picture_block,
            spa: Some(spa),
            dgg_info,
            replaced_grpprl: None,
            data_offset: None,
        }
        .validate_rehomed()
        .unwrap();
    }

    #[test]
    fn nested_image_group_transfers_through_the_body_transaction() {
        use crate::body_text::{Projection, Snapshot as BodySnapshot, TextTarget};
        use crate::writer::{CharacterFormatting, FloatingPosition, ParagraphFormatting, Writer};
        use litchi_core::Position;
        use std::io::Cursor;

        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/images/png/lena.png"),
        )
        .unwrap();
        let picture = crate::writer::Picture::new(bytes).unwrap();
        let mut writer = Writer::new();
        writer
            .insert_floating_picture(picture, FloatingPosition::new(1440, 720))
            .unwrap();
        let mut donor_bytes = Cursor::new(Vec::new());
        writer.write_to(&mut donor_bytes).unwrap();

        let donor_bytes = normalize_word2002_dop(donor_bytes.into_inner());
        let mut editor = RevisionEditor::open(donor_bytes, Limits::default()).unwrap();
        let (picture, width, height, shape_id) = editor.canonical_picture_at_cp(0).unwrap();
        let (spa_offset, spa_length) = fib_pair(&editor.word, 40).unwrap();
        let anchors = crate::parts::spa::parse_plcf_spa(
            slice(&editor.table, spa_offset, spa_length, "PlcfSpaMom").unwrap(),
        )
        .unwrap();
        let spa = anchors[0].spa;
        let group = NestedPictureGroup {
            bytes: nested_group(
                shape_id,
                shape_id + 1,
                i32::try_from(width).unwrap(),
                i32::try_from(height).unwrap(),
            ),
            shape_ids: vec![shape_id, shape_id + 1],
        };
        let dgg_info = build_nested_picture_dgg_info(&picture, spa, &group).unwrap();
        append_table_block(
            &mut editor.word,
            &mut editor.table,
            crate::shape::FIB_INDEX_DGG_INFO,
            &dgg_info,
        )
        .unwrap();
        editor.commit().unwrap();
        let donor = BodySnapshot::parse(&editor.finish().unwrap()).unwrap();

        let mut receiver_writer = Writer::new();
        receiver_writer
            .add_paragraph_runs(
                vec![("placeholder".to_owned(), CharacterFormatting::default())],
                ParagraphFormatting::default(),
            )
            .unwrap();
        let mut receiver_bytes = Cursor::new(Vec::new());
        receiver_writer.write_to(&mut receiver_bytes).unwrap();
        let receiver_bytes = normalize_word2002_dop(receiver_bytes.into_inner());
        let receiver = BodySnapshot::parse(&receiver_bytes).unwrap();
        let plan = receiver
            .plan_picture_transfer_from(
                &donor,
                TextTarget::body_paragraph(Position::new(0)),
                TextTarget::body_paragraph(Position::new(0)),
            )
            .unwrap();
        assert!(plan.is_floating());
        let mut edit = receiver.edit().unwrap();
        edit.apply_picture_transfer(&plan).unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(
            commit.snapshot().paragraphs(Projection::All).unwrap()[0].text(),
            "\u{0008}"
        );
        assert_eq!(
            commit.patch().inverse().apply(commit.snapshot()).unwrap(),
            receiver
        );
    }
}
