//! Master slide PPDrawing builder (MS-ODRAW/MS-PPT)
//!
//! Programmatically constructs the MainMaster PPDrawing record structure
//! including all placeholder shapes and their properties.

use super::escher::{ShapeFlags, record_type as escher_rt};

// =============================================================================
// PPT Record Types
// =============================================================================

/// PPT record types used in PPDrawing
pub mod ppt_record_type {
    /// PPDrawing container
    pub const PPDRAWING: u16 = 0x040C;
    /// OEPlaceholderAtom (MS-PPT 2.9.39)
    pub const OE_PLACEHOLDER_ATOM: u16 = 0x0BC3;
    /// RoundTripProgIdCString
    pub const ROUND_TRIP_PROGID: u16 = 0x0F9F;
    /// RoundTripTextSpecAtom
    pub const ROUND_TRIP_TEXT_SPEC: u16 = 0x0FA8;
    /// RoundTripStyleTextPropAtom
    pub const ROUND_TRIP_STYLE_TEXT_PROP: u16 = 0x0FA2;
    /// RoundTripHFPlaceholder
    pub const ROUND_TRIP_HF_PLACEHOLDER: u16 = 0x0FAA;
}

// =============================================================================
// Placeholder Types (MS-PPT 2.9.39)
// =============================================================================

/// PPT placeholder types for OEPlaceholderAtom
pub mod placeholder_type {
    pub const TITLE: u8 = 0x00;
    pub const BODY: u8 = 0x01;
    pub const DATE: u8 = 0x07;
    pub const SLIDE_NUMBER: u8 = 0x08;
    pub const FOOTER: u8 = 0x09;
}

// Shape flags are imported from super::escher::ShapeFlags

// =============================================================================
// Escher Property IDs (MS-ODRAW 2.3.1)
// =============================================================================

/// Escher property IDs
pub mod escher_prop {
    pub const LOCK_AGGR: u16 = 0x007F;
    pub const ADJUST_VALUE: u16 = 0x0080;
    pub const FILL_COLOR: u16 = 0x0181;
    pub const FILL_BACK_COLOR: u16 = 0x0183;
    pub const FILL_RECT_RIGHT: u16 = 0x0188;
    pub const FILL_RECT_BOTTOM: u16 = 0x0189;
    pub const NO_FILL_HIT_TEST: u16 = 0x01BF;
    pub const LINE_COLOR: u16 = 0x01C0;
    pub const LINE_NO_DRAW_DASH: u16 = 0x01FF; // lineStyleBooleanProperties
    pub const SHAPE_BOOL: u16 = 0x01FF;
    pub const SHADOW_COLOR: u16 = 0x0201;
    pub const BW_MODE: u16 = 0x0304; // blackAndWhiteMode
    pub const BACKGROUND_SHAPE: u16 = 0x017F; // fBackground
    // Complex property flag
    pub const COMPLEX_FLAG: u16 = 0x4000;
}

// =============================================================================
// Drawing Counts and IDs
// =============================================================================

/// Drawing structure constants
pub mod drawing_counts {
    pub const MASTER_DRAWING_ID: u32 = 1;
    pub const MASTER_SHAPE_COUNT: u32 = 7; // 1 group + 5 placeholders + 1 background
    pub const MASTER_BASE_SPID: u32 = 0x0400;
    pub const MASTER_BG_SPID: u32 = 0x0401; // Background shape ID (group + 1)
    /// spidCur - next available shape ID (base + shape_count)
    pub const MASTER_SPID_CUR: u32 = 0x0407; // 0x0400 + 7 = 1031
}

// =============================================================================
// Placeholder Text
// =============================================================================

/// Master placeholder text content
pub mod placeholder_text {
    pub const TITLE: &[u8] = b"Click to edit Master title style";
    pub const BODY: &[u8] =
        b"Click to edit Master text styles\rSecond level\rThird level\rFourth level\rFifth level";
}

// =============================================================================
// Client Anchor positions (in master units)
// =============================================================================

/// Placeholder anchor positions
pub mod anchor {
    /// Title placeholder: left=173, top=288, right=5493, bottom=893
    pub const TITLE: (u16, u16, u16, u16) = (0x00AD, 0x0120, 0x1575, 0x037D);
    /// Body placeholder: left=1008, top=288, right=5493, bottom=3859
    pub const BODY: (u16, u16, u16, u16) = (0x03F0, 0x0120, 0x1575, 0x0F13);
    /// Date placeholder
    pub const DATE: (u16, u16, u16, u16) = (0x00AD, 0x0F80, 0x0CAD, 0x1020);
    /// Footer placeholder
    pub const FOOTER: (u16, u16, u16, u16) = (0x0E4D, 0x0F80, 0x1575, 0x1020);
    /// Slide number placeholder
    pub const SLIDE_NUMBER: (u16, u16, u16, u16) = (0x15F5, 0x0F80, 0x1775, 0x1020);
}

// =============================================================================
// MasterPPDrawingBuilder
// =============================================================================

/// Builder for the master slide PPDrawing record
pub struct MasterPPDrawingBuilder {
    data: Vec<u8>,
}

impl MasterPPDrawingBuilder {
    pub fn new() -> Self {
        Self {
            data: Vec::with_capacity(1300),
        }
    }

    /// Write an Escher header
    fn write_escher_header(&mut self, version: u8, instance: u16, rec_type: u16, length: u32) {
        let ver_inst = (version as u16) | ((instance & 0x0FFF) << 4);
        self.data.extend_from_slice(&ver_inst.to_le_bytes());
        self.data.extend_from_slice(&rec_type.to_le_bytes());
        self.data.extend_from_slice(&length.to_le_bytes());
    }

    /// Write a PPT record header
    fn write_ppt_header(&mut self, rec_type: u16, length: u32) {
        // PPT records use version=0, instance=0
        self.data.extend_from_slice(&[0x00, 0x00]);
        self.data.extend_from_slice(&rec_type.to_le_bytes());
        self.data.extend_from_slice(&length.to_le_bytes());
    }

    /// Build EscherDg record (drawing info)
    fn build_escher_dg(&mut self, shape_count: u32, last_spid: u32) {
        self.write_escher_header(0x10, 0, escher_rt::DG, 8);
        self.data.extend_from_slice(&shape_count.to_le_bytes());
        self.data.extend_from_slice(&last_spid.to_le_bytes());
    }

    /// Build EscherSpgr record (shape group coordinates)
    fn build_escher_spgr(&mut self) {
        self.write_escher_header(0x01, 0, escher_rt::SPGR, 16);
        // Bounding rect: left, top, right, bottom (all zeros for patriarch)
        self.data.extend_from_slice(&[0u8; 16]);
    }

    /// Build EscherSp record (shape)
    fn build_escher_sp(&mut self, instance: u16, spid: u32, flags: u32) {
        self.write_escher_header(0x02, instance, escher_rt::SP, 8);
        self.data.extend_from_slice(&spid.to_le_bytes());
        self.data.extend_from_slice(&flags.to_le_bytes());
    }

    /// Build EscherOpt record with properties
    fn build_escher_opt(&mut self, instance: u16, props: &[(u16, u32)]) {
        let length = (props.len() * 6) as u32;
        self.write_escher_header(0x03, instance, escher_rt::OPT, length);
        for &(prop_id, value) in props {
            self.data.extend_from_slice(&prop_id.to_le_bytes());
            self.data.extend_from_slice(&value.to_le_bytes());
        }
    }

    /// Build ClientAnchor record
    fn build_client_anchor(&mut self, left: u16, top: u16, right: u16, bottom: u16) {
        self.write_escher_header(0x00, 0, escher_rt::CLIENT_ANCHOR, 8);
        self.data.extend_from_slice(&left.to_le_bytes());
        self.data.extend_from_slice(&top.to_le_bytes());
        self.data.extend_from_slice(&right.to_le_bytes());
        self.data.extend_from_slice(&bottom.to_le_bytes());
    }

    /// Build ClientData record (empty wrapper)
    fn build_client_data(&mut self, content_len: u32) {
        self.write_escher_header(0x0F, 0, escher_rt::CLIENT_DATA, content_len);
    }

    /// Build OEPlaceholderAtom
    fn build_oe_placeholder(&mut self, position: u32, placeholder_type: u8, size: u8) {
        self.write_ppt_header(ppt_record_type::OE_PLACEHOLDER_ATOM, 8);
        self.data.extend_from_slice(&position.to_le_bytes());
        self.data.push(placeholder_type);
        self.data.push(size);
        self.data.extend_from_slice(&[0x00, 0x00]); // unused
    }

    /// Build RoundTrip text records for a placeholder
    fn build_roundtrip_text(&mut self, text: &[u8], text_spec_value: u8, style_level_count: u32) {
        // RoundTripProgIdCString (empty)
        self.write_ppt_header(ppt_record_type::ROUND_TRIP_PROGID, 4);
        self.data.extend_from_slice(&[0x00, 0x00, 0x00, 0x00]);

        // RoundTripTextSpecAtom with text
        let text_len = text.len() as u32;
        let padded_len = (text_len + 1 + 1) & !1; // +1 for null, align to 2
        self.write_ppt_header(ppt_record_type::ROUND_TRIP_TEXT_SPEC, padded_len);
        self.data.extend_from_slice(text);
        self.data.push(0x00); // null terminator
        if (text_len + 1) & 1 != 0 {
            self.data.push(0x00); // padding
        }

        // RoundTripStyleTextPropAtom
        let style_len = 6 + style_level_count * 6;
        self.write_ppt_header(ppt_record_type::ROUND_TRIP_STYLE_TEXT_PROP, style_len);
        // Character count
        self.data.extend_from_slice(&(text_len + 1).to_le_bytes());
        self.data.push(0x00);
        self.data.push(0x00);
        // Style runs
        for i in 0..style_level_count {
            if i == style_level_count - 1 {
                self.data.extend_from_slice(&(text_len + 1).to_le_bytes());
            } else {
                self.data.extend_from_slice(&0u32.to_le_bytes());
            }
            self.data.push(text_spec_value);
            self.data.push(0x00);
        }
    }

    /// Build the group patriarch SpContainer
    fn build_group_patriarch(&mut self) -> usize {
        let start = self.data.len();
        // SpContainer header (will be patched)
        self.write_escher_header(0x0F, 0, escher_rt::SP_CONTAINER, 0);
        let content_start = self.data.len();

        self.build_escher_spgr();
        self.build_escher_sp(
            0,
            drawing_counts::MASTER_BASE_SPID,
            (ShapeFlags::GROUP | ShapeFlags::PATRIARCH).bits(),
        );

        // Patch container length
        let content_len = (self.data.len() - content_start) as u32;
        let len_bytes = content_len.to_le_bytes();
        self.data[start + 4..start + 8].copy_from_slice(&len_bytes);
        start
    }

    /// Build background SpContainer (outside SpgrContainer, per POI)
    fn build_background_sp_container(&mut self) {
        let start = self.data.len();
        // SpContainer header (will be patched)
        self.write_escher_header(0x0F, 0, escher_rt::SP_CONTAINER, 0);
        let content_start = self.data.len();

        // Background Sp record: rectangle shape type, BACKGROUND | HAVE_SPT flags
        self.build_escher_sp(
            1, // RECTANGLE shape type
            drawing_counts::MASTER_BG_SPID,
            (ShapeFlags::BACKGROUND | ShapeFlags::HAVE_SPT).bits(),
        );

        // Background EscherOpt properties (per POI PPDrawing.create())
        let bg_props = [
            (escher_prop::FILL_COLOR, 0x08000000),        // fillColor
            (escher_prop::FILL_BACK_COLOR, 0x08000005),   // fillBackColor
            (escher_prop::FILL_RECT_RIGHT, 0x0099A040),   // fillRectRight (10064960)
            (escher_prop::FILL_RECT_BOTTOM, 0x0076BE60),  // fillRectBottom (7782016)
            (escher_prop::NO_FILL_HIT_TEST, 0x00120012),  // noFillHitTest
            (escher_prop::LINE_NO_DRAW_DASH, 0x00080000), // lineNoDrawDash
            (escher_prop::BW_MODE, 0x00000009),           // bwMode
            (escher_prop::BACKGROUND_SHAPE, 0x00010001),  // fBackground
        ];
        self.build_escher_opt(bg_props.len() as u16, &bg_props);

        // Patch container length
        let content_len = (self.data.len() - content_start) as u32;
        let len_bytes = content_len.to_le_bytes();
        self.data[start + 4..start + 8].copy_from_slice(&len_bytes);
    }

    /// Build a placeholder SpContainer
    fn build_placeholder_container(
        &mut self,
        spid: u32,
        anchor: (u16, u16, u16, u16),
        ph_type: u8,
        ph_position: u32,
        text: Option<&[u8]>,
        opt_props: &[(u16, u32)],
    ) -> usize {
        let start = self.data.len();
        // SpContainer header (will be patched)
        self.write_escher_header(0x0F, 0, escher_rt::SP_CONTAINER, 0);
        let content_start = self.data.len();

        // EscherSp - TextBox shape (0xCA = 202)
        let shape_type = 0x00CA; // TextBox
        self.build_escher_sp(
            shape_type,
            spid,
            (ShapeFlags::HAVE_ANCHOR | ShapeFlags::HAVE_SPT).bits(),
        );

        // EscherOpt with properties
        self.build_escher_opt(opt_props.len() as u16, opt_props);

        // ClientAnchor
        self.build_client_anchor(anchor.0, anchor.1, anchor.2, anchor.3);

        // ClientData with nested records
        let client_data_start = self.data.len();
        self.build_client_data(0); // length will be patched
        let client_content_start = self.data.len();

        // OEPlaceholderAtom
        self.build_oe_placeholder(ph_position, ph_type, 0x01);

        // RoundTrip records if text provided
        if let Some(t) = text {
            let level_count = if ph_type == placeholder_type::BODY {
                5
            } else {
                1
            };
            self.build_roundtrip_text(t, 0x21, level_count);
        }

        // Patch ClientData length
        let client_content_len = (self.data.len() - client_content_start) as u32;
        let len_bytes = client_content_len.to_le_bytes();
        self.data[client_data_start + 4..client_data_start + 8].copy_from_slice(&len_bytes);

        // Patch SpContainer length
        let content_len = (self.data.len() - content_start) as u32;
        let len_bytes = content_len.to_le_bytes();
        self.data[start + 4..start + 8].copy_from_slice(&len_bytes);
        start
    }

    /// Build the complete master PPDrawing
    pub fn build(mut self) -> Vec<u8> {
        // PPDrawing record header (will be patched)
        self.write_escher_header(0x0F, 0, ppt_record_type::PPDRAWING, 0);
        let ppdrawing_start = self.data.len();

        // DgContainer (will be patched)
        self.write_escher_header(0x0F, 0, escher_rt::DG_CONTAINER, 0);
        let dg_container_start = self.data.len();

        // EscherDg
        self.build_escher_dg(
            drawing_counts::MASTER_SHAPE_COUNT,
            drawing_counts::MASTER_SPID_CUR,
        );

        // SpgrContainer (will be patched)
        self.write_escher_header(0x0F, 0, escher_rt::SPGR_CONTAINER, 0);
        let spgr_container_start = self.data.len();

        // Group patriarch
        self.build_group_patriarch();

        // Title placeholder properties
        let title_props = [
            (escher_prop::LOCK_AGGR, 0x00010005),
            (escher_prop::ADJUST_VALUE, 0x067C94EC),
            (escher_prop::ADJUST_VALUE | 0x0007, 0x00010000),
            (escher_prop::FILL_COLOR, 0x08000004),
            (escher_prop::FILL_BACK_COLOR, 0x08000000),
            (escher_prop::NO_FILL_HIT_TEST, 0x00110001),
            (escher_prop::LINE_COLOR, 0x08000001),
            (escher_prop::SHAPE_BOOL, 0x00090001),
            (escher_prop::SHADOW_COLOR, 0x08000002),
        ];

        // Title placeholder
        self.build_placeholder_container(
            0x0402,
            anchor::TITLE,
            placeholder_type::TITLE,
            0,
            Some(placeholder_text::TITLE),
            &title_props,
        );

        // Body placeholder properties (similar but different lock value)
        let body_props = [
            (escher_prop::LOCK_AGGR, 0x00010005),
            (escher_prop::ADJUST_VALUE, 0x067C9784),
            (escher_prop::FILL_COLOR, 0x08000004),
            (escher_prop::FILL_BACK_COLOR, 0x08000000),
            (escher_prop::NO_FILL_HIT_TEST, 0x00110001),
            (escher_prop::LINE_COLOR, 0x08000001),
            (escher_prop::SHAPE_BOOL, 0x00090001),
            (escher_prop::SHADOW_COLOR, 0x08000002),
        ];

        // Body placeholder
        self.build_placeholder_container(
            0x0403,
            anchor::BODY,
            placeholder_type::BODY,
            1,
            Some(placeholder_text::BODY),
            &body_props,
        );

        // Date/Footer/SlideNumber placeholders (simpler, no text)
        let simple_props = [
            (escher_prop::LOCK_AGGR, 0x00010005),
            (escher_prop::ADJUST_VALUE, 0x067C9784),
            (escher_prop::FILL_COLOR, 0x08000004),
            (escher_prop::FILL_BACK_COLOR, 0x08000000),
            (escher_prop::NO_FILL_HIT_TEST, 0x00110001),
            (escher_prop::LINE_COLOR, 0x08000001),
            (escher_prop::SHAPE_BOOL, 0x00090001),
            (escher_prop::SHADOW_COLOR, 0x08000002),
        ];

        self.build_placeholder_container(
            0x0404,
            anchor::DATE,
            placeholder_type::DATE,
            2,
            None,
            &simple_props,
        );
        self.build_placeholder_container(
            0x0405,
            anchor::FOOTER,
            placeholder_type::FOOTER,
            3,
            None,
            &simple_props,
        );
        self.build_placeholder_container(
            0x0406,
            anchor::SLIDE_NUMBER,
            placeholder_type::SLIDE_NUMBER,
            4,
            None,
            &simple_props,
        );

        // Patch SpgrContainer length
        let spgr_len = (self.data.len() - spgr_container_start) as u32;
        let len_bytes = spgr_len.to_le_bytes();
        self.data[spgr_container_start - 4..spgr_container_start].copy_from_slice(&len_bytes);

        // Background SpContainer (outside SpgrContainer, per POI PPDrawing.create())
        self.build_background_sp_container();

        // Patch DgContainer length
        let dg_len = (self.data.len() - dg_container_start) as u32;
        let len_bytes = dg_len.to_le_bytes();
        self.data[dg_container_start - 4..dg_container_start].copy_from_slice(&len_bytes);

        // Patch PPDrawing length
        let ppdrawing_len = (self.data.len() - ppdrawing_start) as u32;
        let len_bytes = ppdrawing_len.to_le_bytes();
        self.data[ppdrawing_start - 4..ppdrawing_start].copy_from_slice(&len_bytes);

        self.data
    }
}

impl Default for MasterPPDrawingBuilder {
    fn default() -> Self {
        Self::new()
    }
}

/// Build the master PPDrawing bytes programmatically
pub fn build_master_ppdrawing() -> Vec<u8> {
    MasterPPDrawingBuilder::new().build()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_master_ppdrawing_structure() {
        let built = build_master_ppdrawing();
        // Verify structure starts correctly
        assert!(built.len() > 100, "Built PPDrawing should be substantial");
        // PPDrawing record type at bytes 2-3
        assert_eq!(&built[2..4], &[0x0C, 0x04]); // 0x040C
        // DgContainer at bytes 10-11
        assert_eq!(&built[10..12], &[0x02, 0xF0]); // 0xF002
    }
}
