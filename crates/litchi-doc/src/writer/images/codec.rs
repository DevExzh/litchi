//! OfficeArt/DOC wire encoding for image blocks and floating drawings.

use super::super::core::WriteError;
use super::model::{FloatingShapeContent, FloatingShapeInfo, Picture};
use super::validation::metafile_bounds;
use crate::parts::spa::SPA_LEN;
use litchi_odraw::image::Point;
use litchi_odraw::image::write::{BlipBuilder, digest};

/// Size of the PICF picture descriptor header ([MS-DOC] 2.9.161).
pub(super) const PICF_HEADER_LEN: usize = 0x44;
/// PICF `cbHeader` value, matching [`PICF_HEADER_LEN`].
pub(super) const PICF_CB_HEADER: i16 = 0x44;
/// PICF `mfp.mm` value Word writes for `OfficeArt` shape pictures.
pub(super) const PICF_MM_SHAPE: i16 = 0x64;
/// Unscaled picture factor in permille (1000 = 100%).
pub(super) const SCALE_100_PERCENT: i16 = 1000;

// Escher record types ([MS-ODRAW] 2.3).
pub(super) const RECORD_DGG_CONTAINER: u16 = 0xF000;
pub(super) const RECORD_BSTORE_CONTAINER: u16 = 0xF001;
pub(super) const RECORD_DG_CONTAINER: u16 = 0xF002;
pub(super) const RECORD_SPGR_CONTAINER: u16 = 0xF003;
pub(super) const RECORD_SP_CONTAINER: u16 = 0xF004;
pub(super) const RECORD_DGG: u16 = 0xF006;
pub(super) const RECORD_BSE: u16 = 0xF007;
pub(super) const RECORD_DG: u16 = 0xF008;
pub(super) const RECORD_SPGR: u16 = 0xF009;
pub(super) const RECORD_SP: u16 = 0xF00A;
pub(super) const RECORD_OPT: u16 = 0xF00B;
pub(super) const RECORD_CLIENT_ANCHOR: u16 = 0xF010;
pub(super) const RECORD_CLIENT_DATA: u16 = 0xF011;

// Escher record versions.
pub(super) const VERSION_CONTAINER: u16 = 0xF;
pub(super) const VERSION_DG: u16 = 0x0;
pub(super) const VERSION_SPGR: u16 = 0x1;
pub(super) const VERSION_SP: u16 = 0x2;
pub(super) const VERSION_OPT: u16 = 0x3;
pub(super) const VERSION_ATOM: u16 = 0x0;

/// MSOSHAPETYPE value for a picture frame shape.
pub(super) const SHAPE_TYPE_PICTURE_FRAME: u16 = 0x4B;
/// `OfficeArtFSP` `fHaveAnchor` flag.
pub(super) const SP_FLAG_HAVE_ANCHOR: u32 = 0x0200;
/// `OfficeArtFSP` `fHaveShapeType` flag.
pub(super) const SP_FLAG_HAVE_SHAPE_TYPE: u32 = 0x0800;
/// `OfficeArtFSP` `fGroup` flag, set on the group shape of a drawing.
pub(super) const SP_FLAG_GROUP: u32 = 0x0001;
/// `OfficeArtFSP` `fPatriarch` flag, set on the topmost group shape.
pub(super) const SP_FLAG_PATRIARCH: u32 = 0x0004;
/// `OfficeArt` `pib` property (0x0104) with the fBid bit set, meaning the
/// value is a 1-based index of the BSE within the same drawing block.
pub(super) const OPT_PIB_BLIP_INDEX: u16 = 0x4104;
/// The single BSE stored inside each `OfficeArtWordDrawing` block.
pub(super) const OPT_PIB_FIRST_BSE: u32 = 1;
/// Payload length of the empty `ClientAnchor` record used for inline pictures.
pub(super) const CLIENT_ANCHOR_PAYLOAD_LEN: u32 = 4;
/// Payload length of the `OfficeArtFSP` record (spid + flags).
pub(super) const SP_PAYLOAD_LEN: u32 = 8;
/// Payload length of the `OfficeArtOPT` record holding one simple property.
pub(super) const OPT_PAYLOAD_LEN: u32 = 6;
/// Total length of the `OfficeArtSpContainer` record including its header.
pub(super) const SHAPE_CONTAINER_LEN: u32 = (RECORD_HEADER_LEN as u32 + SP_PAYLOAD_LEN)
    + (RECORD_HEADER_LEN as u32 + OPT_PAYLOAD_LEN)
    + (RECORD_HEADER_LEN as u32 + CLIENT_ANCHOR_PAYLOAD_LEN)
    + RECORD_HEADER_LEN as u32;

/// Length of an Escher record header in bytes.
pub(super) const RECORD_HEADER_LEN: usize = 8;
/// Fixed `OfficeArtFBSE` payload length when no name is stored.
pub(super) const BSE_HEADER_LEN: usize = 36;
/// `OfficeArtFBSE` `foDelay` value when the BLIP is embedded in the BSE record
/// (no delay-stream position).
pub(super) const BSE_NO_DELAY_STREAM: u32 = u32::MAX;

/// Shape id assigned to the first inline picture. Word numbers inline shapes
/// starting at 1025 in documents without an `OfficeArtDgContainer`.
pub(crate) const FIRST_PICTURE_SHAPE_ID: u32 = 1025;
/// Shape id of the group shape that parents the shapes of a drawing.
pub(super) const GROUP_SHAPE_ID: u32 = 1024;
/// Number of shape ids in one `OfficeArt` drawing cluster.
pub(super) const SHAPE_IDS_PER_CLUSTER: u32 = 1024;

/// Twips per inch; the writer assumes 96 DPI when converting pixel sizes.
pub(super) const TWIPS_PER_INCH: u32 = 1440;
/// Assumed screen resolution for pixel-sized images.
pub(super) const ASSUMED_DPI: u32 = 96;
/// Largest dimension expressible in the signed 16-bit PICF goal fields.
pub(super) const MAX_PICF_DIMENSION_TWIPS: u32 = i16::MAX as u32;

pub(super) fn picture_blip(picture: &Picture) -> Result<BlipBuilder<'_>, WriteError> {
    if picture.kind.is_meta() {
        let width = i32::try_from(picture.width_twips)
            .ok()
            .and_then(|value| value.checked_mul(635))
            .ok_or_else(|| {
                WriteError::InvalidData("metafile width extent exceeds i32".to_string())
            })?;
        let height = i32::try_from(picture.height_twips)
            .ok()
            .and_then(|value| value.checked_mul(635))
            .ok_or_else(|| {
                WriteError::InvalidData("metafile height extent exceeds i32".to_string())
            })?;
        Ok(BlipBuilder::meta(
            picture.kind,
            picture.data(),
            metafile_bounds(picture)?,
            Point {
                x: width,
                y: height,
            },
        )?)
    } else {
        Ok(BlipBuilder::bitmap(picture.kind, picture.data())?.tag(0xFF))
    }
}

/// Append an Escher record header ([MS-ODRAW] 2.2.1).
pub(crate) fn write_record_header(
    out: &mut Vec<u8>,
    version: u16,
    instance: u16,
    record_type: u16,
    payload_len: u32,
) {
    let ver_inst = (instance << 4) | version;
    out.extend_from_slice(&ver_inst.to_le_bytes());
    out.extend_from_slice(&record_type.to_le_bytes());
    out.extend_from_slice(&payload_len.to_le_bytes());
}

/// Append an `OfficeArtOPT` record holding simple (non-complex) properties as
/// (opid, value) pairs.
pub(crate) fn write_opt_record(out: &mut Vec<u8>, properties: &[(u16, u32)]) {
    write_record_header(
        out,
        VERSION_OPT,
        properties.len() as u16,
        RECORD_OPT,
        properties.len() as u32 * 6,
    );
    for &(opid, value) in properties {
        out.extend_from_slice(&opid.to_le_bytes());
        out.extend_from_slice(&value.to_le_bytes());
    }
}

/// Append the PICF picture descriptor ([MS-DOC] 2.9.161).
///
/// The 68-byte header is followed by the `OfficeArt` shape data; `lcb` covers
/// both, matching the layout Word writes for inline pictures.
fn write_picf(out: &mut Vec<u8>, picture: &Picture, lcb: u32) {
    out.extend_from_slice(&(lcb as i32).to_le_bytes());
    out.extend_from_slice(&PICF_CB_HEADER.to_le_bytes());
    out.extend_from_slice(&PICF_MM_SHAPE.to_le_bytes());
    out.extend_from_slice(&(picture.width_twips as i16).to_le_bytes()); // xExt
    out.extend_from_slice(&(picture.height_twips as i16).to_le_bytes()); // yExt
    out.extend_from_slice(&0i16.to_le_bytes()); // swHMF
    out.extend_from_slice(&0i32.to_le_bytes()); // grf: low byte 0x00 = Word 2000 picture block
    out.extend_from_slice(&0i32.to_le_bytes()); // padding
    out.extend_from_slice(&0i16.to_le_bytes()); // mmPM
    out.extend_from_slice(&0i32.to_le_bytes()); // padding2
    out.extend_from_slice(&(picture.width_twips as i16).to_le_bytes()); // dxaGoal
    out.extend_from_slice(&(picture.height_twips as i16).to_le_bytes()); // dyaGoal
    out.extend_from_slice(&SCALE_100_PERCENT.to_le_bytes()); // mx
    out.extend_from_slice(&SCALE_100_PERCENT.to_le_bytes()); // my
    out.extend_from_slice(&[0; 8]); // crop left/top/right/bottom
    out.push(0); // fReserved
    out.push(0); // bpp
    out.extend_from_slice(&[0; 16]); // brcTop/brcLeft/brcBottom/brcRight
    out.extend_from_slice(&0i16.to_le_bytes()); // dxaReserved3
    out.extend_from_slice(&0i16.to_le_bytes()); // dyaReserved3
    out.extend_from_slice(&0i16.to_le_bytes()); // cProps
}

/// Append the `OfficeArtSpContainer` describing the inline picture shape.
///
/// The shape is a picture frame whose single `pib` property references the
/// BSE that follows inside the same drawing block. Inline pictures carry an
/// empty `ClientAnchor` record, as Word writes for 0x0001-anchored pictures.
fn write_shape_container(out: &mut Vec<u8>, shape_id: u32) {
    let container_payload_len = SHAPE_CONTAINER_LEN - RECORD_HEADER_LEN as u32;

    write_record_header(
        out,
        VERSION_CONTAINER,
        0,
        RECORD_SP_CONTAINER,
        container_payload_len,
    );
    // OfficeArtFSP: picture frame shape with an anchor and explicit type.
    write_record_header(
        out,
        VERSION_SP,
        SHAPE_TYPE_PICTURE_FRAME,
        RECORD_SP,
        SP_PAYLOAD_LEN,
    );
    out.extend_from_slice(&shape_id.to_le_bytes());
    out.extend_from_slice(&(SP_FLAG_HAVE_ANCHOR | SP_FLAG_HAVE_SHAPE_TYPE).to_le_bytes());
    // OfficeArtOPT: pib referencing the adjacent BSE (1-based index).
    write_record_header(out, VERSION_OPT, 1, RECORD_OPT, OPT_PAYLOAD_LEN);
    out.extend_from_slice(&OPT_PIB_BLIP_INDEX.to_le_bytes());
    out.extend_from_slice(&OPT_PIB_FIRST_BSE.to_le_bytes());
    // OfficeArtClientAnchor: empty for inline pictures.
    write_record_header(
        out,
        VERSION_ATOM,
        0,
        RECORD_CLIENT_ANCHOR,
        CLIENT_ANCHOR_PAYLOAD_LEN,
    );
    out.extend_from_slice(&[0; CLIENT_ANCHOR_PAYLOAD_LEN as usize]);
}

/// Append an `OfficeArtFBSE` record with the embedded `OfficeArtBlip` record for
/// a picture. Used both for the Data-stream picture blocks and for the
/// `BStoreContainer` inside the drawing group of floating pictures.
fn write_bse_with_embedded_blip(out: &mut Vec<u8>, picture: &Picture) -> Result<(), WriteError> {
    let blip = picture_blip(picture)?;
    let blip_record_len = blip.wire_len()?;
    let bse_payload_len = u32::try_from(BSE_HEADER_LEN)
        .ok()
        .and_then(|header| header.checked_add(blip_record_len))
        .ok_or_else(|| WriteError::InvalidData("DOC FBSE length exceeds u32".to_string()))?;

    write_record_header(
        out,
        2,
        u16::from(picture.kind.raw()),
        RECORD_BSE,
        bse_payload_len,
    );
    out.push(picture.kind.raw()); // btWin32
    out.push(picture.kind.raw()); // btMacOS
    out.extend_from_slice(digest(picture.data()).as_bytes()); // rgbUid
    out.extend_from_slice(&0u16.to_le_bytes()); // tag
    out.extend_from_slice(&blip_record_len.to_le_bytes()); // size
    out.extend_from_slice(&1u32.to_le_bytes()); // cRef
    out.extend_from_slice(&BSE_NO_DELAY_STREAM.to_le_bytes()); // foDelay: BLIP is embedded
    out.push(0); // usage
    out.push(0); // cbName
    out.push(0); // unused2
    out.push(0); // unused3

    blip.write(out)?;
    Ok(())
}

/// Append an `OfficeArtWordDrawing` block (PICF + shape container + BSE with an
/// embedded BLIP) to the Data stream.
pub(crate) fn write_picture_block(
    picture: &Picture,
    shape_id: u32,
    out: &mut Vec<u8>,
) -> Result<(), WriteError> {
    let blip_record_len = picture_blip(picture)?.wire_len()?;
    let bse_payload_len = u32::try_from(BSE_HEADER_LEN)
        .ok()
        .and_then(|header| header.checked_add(blip_record_len))
        .ok_or_else(|| WriteError::InvalidData("DOC FBSE length exceeds u32".to_string()))?;

    let block_start = out.len();
    // lcb covers the PICF header plus everything that follows it.
    let lcb = u32::try_from(PICF_HEADER_LEN)
        .ok()
        .and_then(|value| value.checked_add(SHAPE_CONTAINER_LEN))
        .and_then(|value| value.checked_add(RECORD_HEADER_LEN as u32))
        .and_then(|value| value.checked_add(bse_payload_len))
        .ok_or_else(|| WriteError::InvalidData("DOC picture block exceeds u32".to_string()))?;
    write_picf(out, picture, lcb);
    write_shape_container(out, shape_id);
    write_bse_with_embedded_blip(out, picture)?;

    let actual = out
        .len()
        .checked_sub(block_start)
        .ok_or_else(|| WriteError::InvalidData("DOC picture block length underflow".to_string()))?;
    if actual != usize::try_from(lcb).unwrap_or(usize::MAX) {
        return Err(WriteError::InvalidData(
            "DOC picture block length mismatch".to_string(),
        ));
    }
    Ok(())
}

// ============================================================================
// Floating pictures: PlcfSpa and OfficeArtContent (DggInfo)
// ============================================================================

/// Build the `PlcfSpa` for the Main Document ([MS-DOC] 2.8.27).
///
/// `shapes` must be ordered by ascending anchor CP (guaranteed by the
/// writer's append-only story). `final_cp` is the document's ccpText; the
/// final CP entry is undefined per spec but must exceed all anchor CPs.
pub(crate) fn build_plcf_spa(shapes: &[FloatingShapeInfo<'_>], final_cp: u32) -> Vec<u8> {
    let mut out = Vec::with_capacity((shapes.len() + 1) * 4 + shapes.len() * SPA_LEN);
    for shape in shapes {
        out.extend_from_slice(&shape.anchor_cp.to_le_bytes());
    }
    out.extend_from_slice(&final_cp.to_le_bytes());
    for shape in shapes {
        out.extend_from_slice(&shape.spa().to_bytes());
    }
    out
}

/// `OfficeArtWordDrawing` `dgglbl` value for the Main Document drawing.
pub(super) const DGGLBL_MAIN_DOCUMENT: u8 = 0x00;
/// `OfficeArtWordDrawing` `dgglbl` value for the Header Document drawing.
const DGGLBL_HEADER_DOCUMENT: u8 = 0x01;
/// Shape id assigned to the first header-story shape. Each drawing owns a
/// cluster of shape ids ([MS-ODRAW] OfficeArtIDCL): the Main Document
/// drawing uses the cluster starting at 1024, the Header Document drawing
/// the next one.
pub(crate) const HEADER_FIRST_SHAPE_ID: u32 = FIRST_PICTURE_SHAPE_ID + SHAPE_IDS_PER_CLUSTER;
/// `OfficeArtFDG` `csp` counting mode: shapes plus the group shape.
const DG_GROUP_SHAPE_COUNT: u32 = 1;
/// `OfficeArtFSP` payload length (spid + flags).
const FSP_PAYLOAD_LEN: u32 = 8;
/// `OfficeArtSpgr` payload length (empty group bounds rectangle).
const SPGR_PAYLOAD_LEN: u32 = 16;
/// Word's `ClientAnchor` payload: a 4-byte index into the `PlcfSpa` aCP array.
const WORD_CLIENT_ANCHOR_LEN: u32 = 4;
/// `OfficeArtClientData` payload length used by Word for shapes.
const CLIENT_DATA_LEN: u32 = 4;
/// `OfficeArtIDCL` `dgid` of the Main Document drawing.
const DGID_MAIN_DOCUMENT: u32 = 1;
/// `OfficeArtIDCL` `dgid` of the Header Document drawing.
const DGID_HEADER_DOCUMENT: u32 = 2;

/// Compute the `OfficeArtFDGG` `spidMax`: the start of the next shape-id
/// cluster beyond the highest allocated shape id.
pub(super) fn spid_max(highest_shape_id: u32) -> u32 {
    highest_shape_id
        .saturating_add(SHAPE_IDS_PER_CLUSTER - highest_shape_id % SHAPE_IDS_PER_CLUSTER)
}

/// Append one `OfficeArtWordDrawing` element (dgglbl byte + `OfficeArtDgContainer`
/// holding the drawing's group shape and one shape per floating picture or
/// primitive). `dgid` is the drawing identifier (1 = Main, 2 = Header).
fn write_dg_container(
    out: &mut Vec<u8>,
    dgglbl: u8,
    dgid: u16,
    first_shape_id: u32,
    shapes: &[FloatingShapeInfo<'_>],
    bse_index_start: u32,
) {
    out.push(dgglbl);
    let dg_container_start = out.len();
    write_record_header(out, VERSION_CONTAINER, 0, RECORD_DG_CONTAINER, 0);

    // OfficeArtFDG: shape count (including the group shape) and next free spid.
    let shape_count = shapes.len() as u32 + DG_GROUP_SHAPE_COUNT;
    write_record_header(out, VERSION_DG, dgid, RECORD_DG, 8);
    out.extend_from_slice(&shape_count.to_le_bytes()); // csp
    out.extend_from_slice(&(first_shape_id + shapes.len() as u32).to_le_bytes()); // spidCur

    // OfficeArtSpgrContainer: the drawing's group shape plus all shapes.
    let spgr_container_start = out.len();
    write_record_header(out, VERSION_CONTAINER, 0, RECORD_SPGR_CONTAINER, 0);

    // Group shape: empty bounds rectangle and a group/patriarch FSP.
    let group_container_start = out.len();
    write_record_header(out, VERSION_CONTAINER, 0, RECORD_SP_CONTAINER, 0);
    write_record_header(out, VERSION_SPGR, 0, RECORD_SPGR, SPGR_PAYLOAD_LEN);
    out.extend_from_slice(&[0; SPGR_PAYLOAD_LEN as usize]);
    write_record_header(out, VERSION_SP, 0, RECORD_SP, FSP_PAYLOAD_LEN);
    out.extend_from_slice(&GROUP_SHAPE_ID.to_le_bytes());
    out.extend_from_slice(&(SP_FLAG_GROUP | SP_FLAG_PATRIARCH).to_le_bytes());
    patch_record_len(out, group_container_start);

    // One shape per floating picture or primitive. Pictures reference their
    // BSE through a 1-based pib index assigned in document order; text boxes
    // reference their FTXBXS entry the same way through the TXID (each
    // drawing's textbox PLC is indexed independently).
    let mut bse_index = bse_index_start;
    let mut ftxbxs_index = 0;
    for (index, shape) in shapes.iter().enumerate() {
        let shape_start = out.len();
        write_record_header(out, VERSION_CONTAINER, 0, RECORD_SP_CONTAINER, 0);
        write_record_header(
            out,
            VERSION_SP,
            shape.shape_type(),
            RECORD_SP,
            FSP_PAYLOAD_LEN,
        );
        out.extend_from_slice(&shape.shape_id.to_le_bytes());
        out.extend_from_slice(&(SP_FLAG_HAVE_ANCHOR | SP_FLAG_HAVE_SHAPE_TYPE).to_le_bytes());
        match &shape.content {
            // OfficeArtOPT: pib referencing this picture's BSE.
            FloatingShapeContent::Picture(_) => {
                write_opt_record(out, &[(OPT_PIB_BLIP_INDEX, bse_index)]);
                bse_index += 1;
            },
            // OfficeArtOPT: fill/line colors and boolean style properties.
            FloatingShapeContent::Primitive(primitive) => {
                super::super::shapes::write_shape_opt(out, primitive);
            },
        }
        // ClientAnchor: index of this shape's anchor CP in the PlcfSpa.
        write_record_header(
            out,
            VERSION_ATOM,
            0,
            RECORD_CLIENT_ANCHOR,
            WORD_CLIENT_ANCHOR_LEN,
        );
        out.extend_from_slice(&(index as u32).to_le_bytes());
        // OfficeArtClientData: present but unused.
        write_record_header(out, VERSION_ATOM, 0, RECORD_CLIENT_DATA, CLIENT_DATA_LEN);
        out.extend_from_slice(&0u32.to_le_bytes());
        // OfficeArtClientTextbox: links a text box to its FTXBXS entry.
        if shape.text.is_some() {
            super::super::shapes::write_client_textbox(out, ftxbxs_index);
            ftxbxs_index += 1;
        }
        patch_record_len(out, shape_start);
    }

    patch_record_len(out, spgr_container_start);
    patch_record_len(out, dg_container_start);
}

/// Build the `OfficeArtContent` referenced by `fcDggInfo` ([MS-DOC] 2.9.171):
/// an `OfficeArtDggContainer` (drawing defaults plus the blip store) followed
/// by one `OfficeArtWordDrawing` per non-empty drawing — the Main Document
/// drawing first, then the Header Document drawing.
///
/// `allocated_main_shapes` counts every shape id allocated in the Main
/// Document cluster (inline and floating pictures plus main-story shapes);
/// it only feeds the advisory spidMax/cluster bookkeeping.
pub(crate) fn build_dgg_info(
    main_shapes: &[FloatingShapeInfo<'_>],
    header_shapes: &[FloatingShapeInfo<'_>],
    allocated_main_shapes: u32,
) -> Result<Vec<u8>, WriteError> {
    let mut out = Vec::new();

    // ── OfficeArtDggContainer ──
    let dgg_container_start = out.len();
    write_record_header(&mut out, VERSION_CONTAINER, 0, RECORD_DGG_CONTAINER, 0);

    // OfficeArtFDGG: spidMax, cluster table, saved shape/drawing counts.
    let has_main_drawing = !main_shapes.is_empty();
    let has_header_drawing = !header_shapes.is_empty();
    let drawing_count = u32::from(has_main_drawing) + u32::from(has_header_drawing);
    // [MS-ODRAW] 2.2.47 defines cidcl as the number of OfficeArtIDCL
    // records plus one. There can be at most the main and header drawings.
    let cluster_count = drawing_count + 1;
    let highest_shape_id = if has_header_drawing {
        HEADER_FIRST_SHAPE_ID + header_shapes.len() as u32
    } else {
        FIRST_PICTURE_SHAPE_ID + allocated_main_shapes
    };
    let dgg_payload_len = 16 + 8 * drawing_count;
    write_record_header(&mut out, VERSION_ATOM, 0, RECORD_DGG, dgg_payload_len);
    out.extend_from_slice(&spid_max(highest_shape_id).to_le_bytes()); // spidMax
    out.extend_from_slice(&cluster_count.to_le_bytes()); // cidcl
    let saved_shapes = main_shapes.len() as u32
        + header_shapes.len() as u32
        + drawing_count * DG_GROUP_SHAPE_COUNT;
    out.extend_from_slice(&saved_shapes.to_le_bytes()); // cspSaved
    out.extend_from_slice(&drawing_count.to_le_bytes()); // cdgSaved
    // rgidcl: one OfficeArtIDCL per drawing, each owning a spid cluster.
    if has_main_drawing {
        out.extend_from_slice(&DGID_MAIN_DOCUMENT.to_le_bytes());
        out.extend_from_slice(&(DG_GROUP_SHAPE_COUNT + allocated_main_shapes).to_le_bytes());
    }
    if has_header_drawing {
        out.extend_from_slice(&DGID_HEADER_DOCUMENT.to_le_bytes());
        out.extend_from_slice(&(DG_GROUP_SHAPE_COUNT + header_shapes.len() as u32).to_le_bytes());
    }

    // OfficeArtBStoreContainer: one FBSE (with embedded BLIP) per picture in
    // either drawing. Main-drawing pictures are indexed first; the header
    // drawing's pib indices continue after them.
    let bstore_start = out.len();
    let main_picture_count = main_shapes
        .iter()
        .filter(|shape| matches!(shape.content, FloatingShapeContent::Picture(_)))
        .count();
    let picture_count = main_picture_count
        + header_shapes
            .iter()
            .filter(|shape| matches!(shape.content, FloatingShapeContent::Picture(_)))
            .count();
    let picture_count = u16::try_from(picture_count)
        .map_err(|_| WriteError::InvalidData("DOC BStore picture count exceeds u16".to_string()))?;
    if picture_count > 0x0fff {
        return Err(WriteError::InvalidData(
            "DOC BStore contains more than 4095 pictures".to_string(),
        ));
    }
    write_record_header(
        &mut out,
        VERSION_CONTAINER,
        picture_count,
        RECORD_BSTORE_CONTAINER,
        0,
    );
    for shape in main_shapes.iter().chain(header_shapes.iter()) {
        if let FloatingShapeContent::Picture(picture) = shape.content {
            write_bse_with_embedded_blip(&mut out, picture)?;
        }
    }
    patch_record_len(&mut out, bstore_start);

    patch_record_len(&mut out, dgg_container_start);

    // ── OfficeArtWordDrawing elements, main drawing first ──
    if has_main_drawing {
        write_dg_container(
            &mut out,
            DGGLBL_MAIN_DOCUMENT,
            DGID_MAIN_DOCUMENT as u16,
            FIRST_PICTURE_SHAPE_ID,
            main_shapes,
            OPT_PIB_FIRST_BSE,
        );
    }
    if has_header_drawing {
        write_dg_container(
            &mut out,
            DGGLBL_HEADER_DOCUMENT,
            DGID_HEADER_DOCUMENT as u16,
            HEADER_FIRST_SHAPE_ID,
            header_shapes,
            OPT_PIB_FIRST_BSE + main_picture_count as u32,
        );
    }

    Ok(out)
}

/// Back-patch the payload length of a container record header written with a
/// zero placeholder length.
fn patch_record_len(out: &mut [u8], record_start: usize) {
    let payload_len = (out.len() - record_start - RECORD_HEADER_LEN) as u32;
    out[record_start + 4..record_start + 8].copy_from_slice(&payload_len.to_le_bytes());
}
