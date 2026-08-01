//! Primitive drawing-shape writer for DOC files
//!
//! Builds the OfficeArt shape descriptions ([MS-ODRAW]) for floating
//! primitive shapes (rectangles, rounded rectangles, ellipses). Shapes anchor
//! in the text stream exactly like floating pictures ([MS-DOC] 1.3): a 0x0008
//! anchor character, a PlcfSpa entry for position and wrapping, and an
//! OfficeArtSpContainer in the document's drawing group (fcDggInfo).
//!
//! Shape text (text boxes) is not supported: it requires the textbox story
//! machinery ([MS-DOC] PlcftxbxTxt/PlcffldTxbx and the OfficeArtClientTextbox
//! record) which is out of scope for this writer.

use super::core::DocWriteError;
use super::images::{write_opt_record, write_record_header};

// MSOSPT shape types ([MS-ODRAW] 2.4.24).
const MSOSPT_RECTANGLE: u16 = 0x0001;
const MSOSPT_ROUND_RECTANGLE: u16 = 0x0002;
const MSOSPT_ELLIPSE: u16 = 0x0003;
/// msosptTextBox, used for shapes carrying a textbox story.
pub(crate) const MSOSPT_TEXT_BOX: u16 = 0x00CA;

// OfficeArt property identifiers ([MS-ODRAW] 2.3).
/// `fillColor` property ([MS-ODRAW] 2.3.7.2).
const OPT_FILL_COLOR: u16 = 0x0181;
/// `lineColor` property ([MS-ODRAW] 2.3.8.1).
const OPT_LINE_COLOR: u16 = 0x01C0;
/// Fill Style Boolean Properties ([MS-ODRAW] 2.3.7.43).
const OPT_FILL_STYLE_BOOLEAN: u16 = 0x01BF;
/// Line Style Boolean Properties ([MS-ODRAW] 2.3.8.38).
const OPT_LINE_STYLE_BOOLEAN: u16 = 0x01FF;

/// Bit of `fUsefFilled` in the Fill Style Boolean Properties value.
const FILL_FLAG_USE_FILLED: u32 = 1 << 2;
/// Bit of `fFilled` in the Fill Style Boolean Properties value.
const FILL_FLAG_FILLED: u32 = 1 << 9;
/// Bit of `fUsefLine` in the Line Style Boolean Properties value.
const LINE_FLAG_USE_LINE: u32 = 1 << 6;
/// Bit of `fLine` in the Line Style Boolean Properties value.
const LINE_FLAG_LINE: u32 = 1 << 16;

/// Largest shape dimension the writer accepts (matches the DOC picture
/// dimension limit; about 22.7 inches at 100% scale).
const MAX_SHAPE_DIMENSION_TWIPS: u32 = i16::MAX as u32;

/// Primitive shape kinds supported by the DOC drawing writer.
///
/// The values are the MSOSPT enumeration ([MS-ODRAW] 2.4.24) that the
/// OfficeArtFSP record instance carries.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocShapeKind {
    /// msosptRectangle.
    Rectangle,
    /// msosptRoundRectangle.
    RoundRectangle,
    /// msosptEllipse.
    Ellipse,
}

impl DocShapeKind {
    /// The MSOSPT value written as the OfficeArtFSP record instance.
    pub(crate) fn shape_type(self) -> u16 {
        match self {
            Self::Rectangle => MSOSPT_RECTANGLE,
            Self::RoundRectangle => MSOSPT_ROUND_RECTANGLE,
            Self::Ellipse => MSOSPT_ELLIPSE,
        }
    }
}

/// A floating primitive drawing shape.
///
/// The shape uses its preset geometry ([MS-ODRAW] MSOSPT); position and
/// wrapping come from the `FloatingPosition` passed to
/// [`super::core::DocWriter::insert_floating_shape`]. By default a shape is
/// drawn as an outline: no fill, black line.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocDrawingShape {
    /// Preset geometry kind.
    kind: DocShapeKind,
    /// Width in twips (1/1440 inch).
    width_twips: u32,
    /// Height in twips.
    height_twips: u32,
    /// Fill color as (R, G, B); `None` disables the fill.
    fill_color: Option<(u8, u8, u8)>,
    /// Line color as (R, G, B); `None` disables the outline.
    line_color: Option<(u8, u8, u8)>,
}

impl DocDrawingShape {
    /// Create a shape with the given size in twips (1/1440 inch).
    ///
    /// The shape starts as an outline (no fill, black line); use
    /// [`Self::with_fill`] / [`Self::with_line`] / [`Self::without_fill`] /
    /// [`Self::without_line`] to adjust.
    pub fn new(
        kind: DocShapeKind,
        width_twips: u32,
        height_twips: u32,
    ) -> Result<Self, DocWriteError> {
        for dimension in [width_twips, height_twips] {
            if !(1..=MAX_SHAPE_DIMENSION_TWIPS).contains(&dimension) {
                return Err(DocWriteError::InvalidData(format!(
                    "DOC shape dimension {dimension} twips is outside 1..={MAX_SHAPE_DIMENSION_TWIPS}"
                )));
            }
        }
        Ok(Self {
            kind,
            width_twips,
            height_twips,
            fill_color: None,
            line_color: Some((0, 0, 0)),
        })
    }

    /// Set the fill color (enables filling).
    pub fn with_fill(mut self, red: u8, green: u8, blue: u8) -> Self {
        self.fill_color = Some((red, green, blue));
        self
    }

    /// Disable the fill.
    pub fn without_fill(mut self) -> Self {
        self.fill_color = None;
        self
    }

    /// Set the outline color (enables the outline).
    pub fn with_line(mut self, red: u8, green: u8, blue: u8) -> Self {
        self.line_color = Some((red, green, blue));
        self
    }

    /// Disable the outline.
    pub fn without_line(mut self) -> Self {
        self.line_color = None;
        self
    }

    /// The preset geometry kind.
    pub fn kind(&self) -> DocShapeKind {
        self.kind
    }

    /// Width in twips.
    pub fn width_twips(&self) -> u32 {
        self.width_twips
    }

    /// Height in twips.
    pub fn height_twips(&self) -> u32 {
        self.height_twips
    }
}

/// Encode an (R, G, B) color as an OfficeArtCOLORREF value
/// ([MS-ODRAW] 2.2.2): red, green, blue bytes followed by zero flags.
fn color_ref(red: u8, green: u8, blue: u8) -> u32 {
    u32::from(red) | (u32::from(green) << 8) | (u32::from(blue) << 16)
}

/// Append the OfficeArtOPT record for a primitive shape: fill and line
/// colors, and the boolean properties that explicitly disable the fill or
/// the line when unset. Properties are emitted in ascending opid order as
/// Word writes them.
pub(crate) fn write_shape_opt(out: &mut Vec<u8>, shape: &DocDrawingShape) {
    let mut properties: Vec<(u16, u32)> = Vec::with_capacity(4);
    if let Some((red, green, blue)) = shape.fill_color {
        properties.push((OPT_FILL_COLOR, color_ref(red, green, blue)));
    } else {
        properties.push((
            OPT_FILL_STYLE_BOOLEAN,
            FILL_FLAG_USE_FILLED & !FILL_FLAG_FILLED,
        ));
    }
    if let Some((red, green, blue)) = shape.line_color {
        properties.push((OPT_LINE_COLOR, color_ref(red, green, blue)));
    } else {
        properties.push((OPT_LINE_STYLE_BOOLEAN, LINE_FLAG_USE_LINE & !LINE_FLAG_LINE));
    }
    properties.sort_by_key(|&(opid, _)| opid);
    write_opt_record(out, &properties);
}

// ============================================================================
// Text boxes: OfficeArtClientTextbox record and PlcftxbxTxt
// ============================================================================

/// OfficeArtClientTextbox record type ([MS-DOC] 2.9.170, msofbtClientTextbox).
const RECORD_CLIENT_TEXTBOX: u16 = 0xF00D;
/// Payload length of the OfficeArtClientTextbox record (one TXID).
const CLIENT_TEXTBOX_LEN: u32 = 4;
/// Size of one FTXBXS structure in bytes ([MS-DOC] 2.9.106).
const FTXBXS_LEN: usize = 22;
/// Number of shapes in an unlinked textbox chain (FTXBXS `cTxbx`).
const SINGLE_SHAPE_CHAIN: i32 = 1;
/// FTXBXS `fReusable` flag marking a structure available for reuse.
const FTXBXS_REUSABLE: u16 = 1;
/// FTXBXSReusable `iNextReuse` for the last reusable structure.
const NO_NEXT_REUSE: i32 = -1;
/// Ignored FTXBXS `itxbxsDest` value, as written by LibreOffice.
const ITXBXS_DEST_IGNORED: i32 = -1;

/// Append an OfficeArtClientTextbox record linking a shape to its textbox
/// story entry. The TXID encodes the 1-based FTXBXS index in its high two
/// bytes and the zero-based textbox chain index (always 0 here) in its low
/// two bytes ([MS-DOC] 2.9.170).
pub(crate) fn write_client_textbox(out: &mut Vec<u8>, ftxbxs_index: u32) {
    write_record_header(out, 0, 0, RECORD_CLIENT_TEXTBOX, CLIENT_TEXTBOX_LEN);
    let txid = (ftxbxs_index + 1) << 16;
    out.extend_from_slice(&txid.to_le_bytes());
}

/// Build the PlcftxbxTxt ([MS-DOC] 2.8.32): `n + 2` story-relative CPs
/// followed by `n + 1` FTXBXS records, the last of which is a reusable spare
/// ([MS-DOC] 2.9.106: "The last FTXBXS in the PLC MUST be a reusable
/// structure").
///
/// * `shape_ids` - spid of each text box, in story order
/// * `start_cps` - story-relative start CP of each text box's text
/// * `ccp_txbx` - total textbox story length (including the story-final CR)
pub(crate) fn build_plcf_txbx_txt(shape_ids: &[u32], start_cps: &[u32], ccp_txbx: u32) -> Vec<u8> {
    debug_assert_eq!(shape_ids.len(), start_cps.len());
    let count = shape_ids.len();
    let mut out = Vec::with_capacity((count + 2) * 4 + (count + 1) * FTXBXS_LEN);

    for &cp in start_cps {
        out.extend_from_slice(&cp.to_le_bytes());
    }
    // The spare entry's range covers just the story-final CR; the final CP is
    // the story length itself.
    out.extend_from_slice(&(ccp_txbx - 1).to_le_bytes());
    out.extend_from_slice(&ccp_txbx.to_le_bytes());

    for &shape_id in shape_ids {
        // FTXBXNonReusable: chain of a single shape, no edits.
        out.extend_from_slice(&SINGLE_SHAPE_CHAIN.to_le_bytes()); // cTxbx
        out.extend_from_slice(&0i32.to_le_bytes()); // cTxbxEdit
        out.extend_from_slice(&0u16.to_le_bytes()); // fReusable
        out.extend_from_slice(&ITXBXS_DEST_IGNORED.to_le_bytes()); // itxbxsDest
        out.extend_from_slice(&(shape_id as i32).to_le_bytes()); // lid
        out.extend_from_slice(&0i32.to_le_bytes()); // txidUndo
    }

    // Final reusable spare FTXBXS.
    out.extend_from_slice(&NO_NEXT_REUSE.to_le_bytes()); // iNextReuse
    out.extend_from_slice(&0i32.to_le_bytes()); // cReusable
    out.extend_from_slice(&FTXBXS_REUSABLE.to_le_bytes()); // fReusable
    out.extend_from_slice(&ITXBXS_DEST_IGNORED.to_le_bytes()); // itxbxsDest
    out.extend_from_slice(&0i32.to_le_bytes()); // lid
    out.extend_from_slice(&0i32.to_le_bytes()); // txidUndo

    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shape_kind_msospt_values_match_spec() {
        assert_eq!(DocShapeKind::Rectangle.shape_type(), 0x0001);
        assert_eq!(DocShapeKind::RoundRectangle.shape_type(), 0x0002);
        assert_eq!(DocShapeKind::Ellipse.shape_type(), 0x0003);
    }

    #[test]
    fn new_shape_is_an_outline_by_default() {
        let shape = DocDrawingShape::new(DocShapeKind::Rectangle, 1440, 720).unwrap();
        assert_eq!(shape.fill_color, None);
        assert_eq!(shape.line_color, Some((0, 0, 0)));
        assert!(DocDrawingShape::new(DocShapeKind::Rectangle, 0, 720).is_err());
        assert!(
            DocDrawingShape::new(DocShapeKind::Ellipse, 1440, MAX_SHAPE_DIMENSION_TWIPS + 1)
                .is_err()
        );
    }

    #[test]
    fn color_ref_uses_office_byte_order() {
        // OfficeArtCOLORREF: red in the low byte, blue in byte 2.
        assert_eq!(color_ref(0xFF, 0x00, 0x00), 0x0000_00FF);
        assert_eq!(color_ref(0x00, 0x00, 0xFF), 0x00FF_0000);
        assert_eq!(color_ref(0x12, 0x34, 0x56), 0x0056_3412);
    }

    #[test]
    fn write_shape_opt_emits_colors_and_disable_flags() {
        // Filled + lined: two color properties in ascending opid order.
        let shape = DocDrawingShape::new(DocShapeKind::Rectangle, 1440, 720)
            .unwrap()
            .with_fill(0xFF, 0x00, 0x00)
            .with_line(0x00, 0x00, 0xFF);
        let mut out = Vec::new();
        write_shape_opt(&mut out, &shape);
        let (_ver, inst, record_type, len) = parse_opt_header(&out);
        assert_eq!((inst, record_type), (2, 0xF00B));
        assert_eq!(len as usize, out.len() - 8);
        let (opid0, val0) = read_property(&out, 0);
        let (opid1, val1) = read_property(&out, 1);
        assert_eq!((opid0, val0), (OPT_FILL_COLOR, 0x0000_00FF));
        assert_eq!((opid1, val1), (OPT_LINE_COLOR, 0x00FF_0000));

        // No fill, no line: the two boolean property sets with the "use" bits
        // set and the value bits clear.
        let shape = shape.without_fill().without_line();
        let mut out = Vec::new();
        write_shape_opt(&mut out, &shape);
        let (_ver, inst, _record_type, _len) = parse_opt_header(&out);
        assert_eq!(inst, 2);
        let (opid0, val0) = read_property(&out, 0);
        let (opid1, val1) = read_property(&out, 1);
        assert_eq!(
            (opid0, val0),
            (OPT_FILL_STYLE_BOOLEAN, FILL_FLAG_USE_FILLED)
        );
        assert_eq!((opid1, val1), (OPT_LINE_STYLE_BOOLEAN, LINE_FLAG_USE_LINE));
        assert_eq!(val0 & FILL_FLAG_FILLED, 0);
        assert_eq!(val1 & LINE_FLAG_LINE, 0);
    }

    #[test]
    fn client_textbox_txid_encodes_ftxbxs_index() {
        let mut out = Vec::new();
        write_client_textbox(&mut out, 0);
        let (ver, _inst, record_type, len) = parse_opt_header(&out);
        assert_eq!((ver, record_type, len), (0, RECORD_CLIENT_TEXTBOX, 4));
        let txid = u32::from_le_bytes(out[8..12].try_into().unwrap());
        assert_eq!(txid, 0x0001_0000);

        let mut out = Vec::new();
        write_client_textbox(&mut out, 4);
        let txid = u32::from_le_bytes(out[8..12].try_into().unwrap());
        assert_eq!(txid, 0x0005_0000);
    }

    #[test]
    fn plcf_txbx_txt_layout_matches_ms_doc() {
        // Two text boxes: "Hi" (3 story CPs incl. trailing CR) and "Yo\rHo"
        // (6 story CPs incl. trailing CR), plus the story-final CR.
        let shape_ids = [1027, 1028];
        let start_cps = [0, 3];
        let ccp_txbx = 3 + 6 + 1;

        let plcf = build_plcf_txbx_txt(&shape_ids, &start_cps, ccp_txbx);
        // n+2 CPs and n+1 FTXBXS records (the last one reusable).
        assert_eq!(plcf.len(), 4 * 4 + 3 * FTXBXS_LEN);
        let cps: Vec<u32> = (0..4)
            .map(|i| u32::from_le_bytes(plcf[i * 4..i * 4 + 4].try_into().unwrap()))
            .collect();
        assert_eq!(cps, vec![0, 3, ccp_txbx - 1, ccp_txbx]);

        let entry = |index: usize| -> &[u8] {
            let start = 16 + index * FTXBXS_LEN;
            &plcf[start..start + FTXBXS_LEN]
        };
        // First box: chain length 1, not reusable, lid = spid.
        let first = entry(0);
        assert_eq!(i32_at(first, 0), 1); // cTxbx
        assert_eq!(i32_at(first, 4), 0); // cTxbxEdit
        assert_eq!(u16_at(first, 8), 0); // fReusable
        assert_eq!(i32_at(first, 14), 1027); // lid
        assert_eq!(i32_at(first, 18), 0); // txidUndo
        let second = entry(1);
        assert_eq!(i32_at(second, 14), 1028);
        // Final spare FTXBXS: reusable chain terminator, lid = 0.
        let spare = entry(2);
        assert_eq!(i32_at(spare, 0), -1); // iNextReuse
        assert_eq!(i32_at(spare, 4), 0); // cReusable
        assert_eq!(u16_at(spare, 8), 1); // fReusable
        assert_eq!(i32_at(spare, 14), 0); // lid
    }

    fn i32_at(data: &[u8], offset: usize) -> i32 {
        i32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
    }

    fn u16_at(data: &[u8], offset: usize) -> u16 {
        u16::from_le_bytes(data[offset..offset + 2].try_into().unwrap())
    }

    fn parse_opt_header(data: &[u8]) -> (u16, u16, u16, u32) {
        let ver_inst = u16::from_le_bytes(data[0..2].try_into().unwrap());
        let record_type = u16::from_le_bytes(data[2..4].try_into().unwrap());
        let length = u32::from_le_bytes(data[4..8].try_into().unwrap());
        (ver_inst & 0xF, ver_inst >> 4, record_type, length)
    }

    fn read_property(data: &[u8], index: usize) -> (u16, u32) {
        let offset = 8 + index * 6;
        (
            u16::from_le_bytes(data[offset..offset + 2].try_into().unwrap()),
            u32::from_le_bytes(data[offset + 2..offset + 6].try_into().unwrap()),
        )
    }
}
