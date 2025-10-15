// MTEF Binary Parser - Main parsing logic
//
// Based on rtf2latex2e Eqn_GetObjectList and related parsing functions

use crate::formula::mtef::constants::*;
use super::headers::*;
use super::objects::*;
use crate::formula::mtef::MtefError;

/// Binary MTEF parser
pub struct MtefBinaryParser<'arena> {
    arena: &'arena bumpalo::Bump,
    data: &'arena [u8],
    pos: usize,
    pub mtef_version: u8,
    pub platform: u8,
    pub product: u8,
    pub version: u8,
    pub version_sub: u8,
    pub inline: u8,
    pub mode: i32, // Current math/text mode (EQN_MODE_TEXT, EQN_MODE_INLINE, EQN_MODE_DISPLAY)
}

impl<'arena> MtefBinaryParser<'arena> {
    /// Get attribute byte(s) according to MTEF version (matches rtf2latex2e GetAttribute)
    fn get_attribute(&mut self) -> Result<u8, MtefError> {
        if self.mtef_version < 5 {
            // For MTEF < 5, attribute is in high nibble of current byte
            let byte = self.read_u8()?;
            Ok((byte & 0xF0) >> 4) // HiNibble equivalent
        } else {
            // For MTEF >= 5, attribute is the next byte
            self.read_u8()
        }
    }

    /// Get nudge values (matches rtf2latex2e GetNudge)
    fn get_nudge(&mut self) -> Result<(i16, i16), MtefError> {
        let b1 = self.read_u8()?;
        let b2 = self.read_u8()?;

        if b1 == 128 && b2 == 128 {
            // Extended nudge format
            let x = self.read_i16()?;
            let y = self.read_i16()?;
            Ok((x, y))
        } else {
            // Simple nudge format
            Ok((b1 as i16, b2 as i16))
        }
    }

    /// Create a new MTEF binary parser
    pub fn new(arena: &'arena bumpalo::Bump, data: &'arena [u8]) -> Result<Self, MtefError> {
        if data.len() < 28 {
            return Err(MtefError::InvalidFormat("Data too short for OLE header".to_string()));
        }

        // Parse OLE header
        let ole_header = OleFileHeader {
            cb_hdr: u16::from_le_bytes([data[0], data[1]]),
            version: u32::from_le_bytes([data[2], data[3], data[4], data[5]]),
            format: u16::from_le_bytes([data[6], data[7]]),
            size: u32::from_le_bytes([data[8], data[9], data[10], data[11]]),
            reserved: [
                u32::from_le_bytes([data[12], data[13], data[14], data[15]]),
                u32::from_le_bytes([data[16], data[17], data[18], data[19]]),
                u32::from_le_bytes([data[20], data[21], data[22], data[23]]),
                u32::from_le_bytes([data[24], data[25], data[26], data[27]]),
            ],
        };

        if ole_header.cb_hdr != 28 {
            return Err(MtefError::InvalidFormat("Invalid OLE header length".to_string()));
        }

        // Accept both 0x00020000 and 0x00000200 as valid versions (observed in real files)
        if ole_header.version != 0x00020000 && ole_header.version != 0x00000200 {
            return Err(MtefError::InvalidFormat(format!("Invalid OLE version: 0x{:08X}", ole_header.version)));
        }

        // Note: The clipboard format can vary (0xC2D3, 0xC1B0, 0xC1E1, 0xC1AE, etc.)
        // so we don't validate it strictly. The MTEF signature check below is sufficient.

        let mut parser = Self {
            arena,
            data,
            pos: 28,
            mtef_version: 0,
            platform: 0,
            product: 0,
            version: 0,
            version_sub: 0,
            inline: 0,
            mode: EQN_MODE_DISPLAY, // Default to display mode
        };

        parser.read_mtef_header()?;
        Ok(parser)
    }

    fn read_mtef_header(&mut self) -> Result<(), MtefError> {
        if self.data.len() < self.pos + 5 {
            return Err(MtefError::UnexpectedEof);
        }

        // Check if we have the full MTEF signature "(\x04mt" (0x28 0x04 0x6D 0x74)
        // or if this is a headerless/embedded format that starts directly with the version
        let has_signature = self.pos + 4 <= self.data.len() &&
                           self.data[self.pos] == 0x28 &&
                           self.data[self.pos + 1] == 0x04 &&
                           self.data[self.pos + 2] == 0x6D &&
                           self.data[self.pos + 3] == 0x74;

        if has_signature {
            // Full format with signature
            self.pos += 4;
            self.mtef_version = self.read_u8()?;
        } else {
            // Headerless/embedded format - starts directly with version byte
            // This format is used in some embedded equations
            self.mtef_version = self.read_u8()?;
        }

        // Handle different MTEF versions
        match self.mtef_version {
            0 => {
                self.mtef_version = 5;
                self.platform = 0;
                self.product = 0;
                self.version = 0;
                self.version_sub = 0;
            }
            1 | 101 => {
                self.platform = if self.mtef_version == 101 { 1 } else { 0 };
                self.product = 0;
                self.version = 1;
                self.version_sub = 0;
            }
            2..=4 => {
                self.platform = self.read_u8()?;
                self.product = self.read_u8()?;
                self.version = self.read_u8()?;
                self.version_sub = self.read_u8()?;
            }
            5 => {
                self.platform = self.read_u8()?;
                self.product = self.read_u8()?;
                self.version = self.read_u8()?;
                self.version_sub = self.read_u8()?;

                // Application key (null-terminated string)
                while self.pos < self.data.len() && self.data[self.pos] != 0 {
                    self.pos += 1;
                }
                if self.pos >= self.data.len() {
                    return Err(MtefError::UnexpectedEof);
                }
                self.pos += 1; // Skip null terminator

                self.inline = self.read_u8()?;
            }
            _ => {
                return Err(MtefError::InvalidFormat(format!("Unsupported MTEF version: {}", self.mtef_version)));
            }
        }

        Ok(())
    }

    /// Parse the MTEF equation into AST nodes
    pub fn parse(&mut self) -> Result<Vec<crate::formula::ast::MathNode<'arena>>, MtefError> {
        let object_list = self.parse_object_list(2)?; // Expect at least 2 objects (SIZE + LINE/PILE)

        if let Some(obj_list) = object_list {
            self.convert_objects_to_ast(&obj_list)
        } else {
            Ok(Vec::new())
        }
    }

    fn parse_object_list(&mut self, num_objs: usize) -> Result<Option<Box<MtefObjectList>>, MtefError> {
        let mut head: Option<Box<MtefObjectList>> = None;
        let mut curr: Option<*mut MtefObjectList> = None;
        let mut tally = 0;
        let start_pos = self.pos; // For error reporting

        // Prevent infinite loops by limiting iterations
        let mut iterations = 0;
        const MAX_ITERATIONS: usize = 10000;

        loop {
            if self.pos >= self.data.len() {
                break;
            }

            // Prevent infinite loops
            iterations += 1;
            if iterations > MAX_ITERATIONS {
                return Err(MtefError::ParseError(format!(
                    "Too many objects parsed (>{}), possible infinite loop at position {}",
                    MAX_ITERATIONS, start_pos
                )));
            }

            // Read tag byte with bounds check
            if self.pos >= self.data.len() {
                break;
            }

            // Get current tag based on MTEF version
            let curr_tag = if self.mtef_version == 5 {
                self.data[self.pos]
            } else {
                self.data[self.pos] & 0x0F
            };

            // END tag handling - return immediately
            if curr_tag == crate::formula::mtef::constants::END {
                self.pos += 1;
                break;
            }

            let record_type = match curr_tag {
                crate::formula::mtef::constants::LINE => MtefRecordType::Line,
                crate::formula::mtef::constants::CHAR => MtefRecordType::Char,
                crate::formula::mtef::constants::TMPL => MtefRecordType::Tmpl,
                crate::formula::mtef::constants::PILE => MtefRecordType::Pile,
                crate::formula::mtef::constants::MATRIX => MtefRecordType::Matrix,
                crate::formula::mtef::constants::EMBELL => MtefRecordType::Embell,
                crate::formula::mtef::constants::RULER => MtefRecordType::Ruler,
                crate::formula::mtef::constants::FONT => MtefRecordType::Font,
                crate::formula::mtef::constants::SIZE => MtefRecordType::Size,
                crate::formula::mtef::constants::FULL => MtefRecordType::Full,
                crate::formula::mtef::constants::SUB => MtefRecordType::Sub,
                crate::formula::mtef::constants::SUB2 => MtefRecordType::Sub2,
                crate::formula::mtef::constants::SYM => MtefRecordType::Sym,
                crate::formula::mtef::constants::SUBSYM => MtefRecordType::SubSym,
                crate::formula::mtef::constants::COLOR => MtefRecordType::Color,
                crate::formula::mtef::constants::COLOR_DEF => MtefRecordType::ColorDef,
                crate::formula::mtef::constants::FONT_DEF => MtefRecordType::FontDef,
                crate::formula::mtef::constants::EQN_PREFS => MtefRecordType::EqnPrefs,
                crate::formula::mtef::constants::ENCODING_DEF => MtefRecordType::EncodingDef,
                _ => MtefRecordType::Future,
            };

            // Parse the object based on its type
            let obj_ptr: Option<Box<dyn MtefObject>> = match record_type {
                MtefRecordType::Char => Some(Box::new(self.parse_char()?)),
                MtefRecordType::Tmpl => Some(Box::new(self.parse_template()?)),
                MtefRecordType::Line => Some(Box::new(self.parse_line()?)),
                MtefRecordType::Pile => Some(Box::new(self.parse_pile()?)),
                MtefRecordType::Matrix => Some(Box::new(self.parse_matrix()?)),
                MtefRecordType::Embell => Some(Box::new(self.parse_embell()?)),
                MtefRecordType::Ruler => Some(Box::new(self.parse_ruler()?)),
                MtefRecordType::Font => Some(Box::new(self.parse_font()?)),
                MtefRecordType::Size | MtefRecordType::Full | MtefRecordType::Sub |
                MtefRecordType::Sub2 | MtefRecordType::Sym | MtefRecordType::SubSym => {
                    Some(Box::new(self.parse_size()?))
                }
                MtefRecordType::ColorDef => {
                    // Skip color definition - just skip the tag
                    self.pos += 1;
                    None
                }
                MtefRecordType::FontDef => {
                    self.skip_font_def()?;
                    None
                }
                MtefRecordType::EqnPrefs => {
                    self.skip_eqn_prefs()?;
                    None
                }
                MtefRecordType::EncodingDef => {
                    self.skip_encoding_def()?;
                    None
                }
                MtefRecordType::Future => {
                    self.skip_future_record()?;
                    None
                }
                _ => {
                    // Unknown record type - skip it
                    self.skip_unknown_record()?;
                    None
                }
            };

            // Only create a node if we have an object
            if let Some(obj) = obj_ptr {
                // Create object list node
                let new_node = Box::new(MtefObjectList {
                    tag: record_type,
                    obj_ptr: obj,
                    next: None,
                });

                // Link into the list
                match curr {
                    Some(curr_ptr) => unsafe {
                        (*curr_ptr).next = Some(new_node);
                        curr = (*curr_ptr).next.as_mut().map(|n| n.as_mut() as *mut _);
                    },
                    None => {
                        head = Some(new_node);
                        curr = head.as_mut().map(|n| n.as_mut() as *mut _);
                    }
                }

                tally += 1;

                if num_objs > 0 && tally == num_objs {
                    break;
                }
            }
        }

        Ok(head)
    }

    fn parse_char(&mut self) -> Result<MtefChar, MtefError> {
        let attrs = self.get_attribute()?;

        let mut nudge_x = 0i16;
        let mut nudge_y = 0i16;
        if attrs & CHAR_NUDGE != 0 {
            let nudge_result = self.get_nudge()?;
            nudge_x = nudge_result.0;
            nudge_y = nudge_result.1;
        }

        let typeface = self.read_u8()?;

        let mut character = 0u16;
        let mut bits16 = 0u16;

        if self.mtef_version < 5 {
            character = self.read_u8()? as u16;
            if self.platform == 1 { // PLATFORM_WIN
                character |= (self.read_u8()? as u16) << 8;
            }
        } else {
            // Nearly always have a 16 bit MT character
            if attrs & CHAR_ENC_NO_MTCODE == 0 {
                character = self.read_u16()?;
            }

            if attrs & CHAR_ENC_CHAR_8 != 0 {
                character = self.read_u8()? as u16;
            }

            if attrs & CHAR_ENC_CHAR_16 != 0 {
                bits16 = self.read_u16()?;
            }
        }

        let embellishment_list = if self.mtef_version == 5 {
            if attrs & CHAR_EMBELL != 0 {
                Some(Box::new(self.parse_embell()?))
            } else {
                None
            }
        } else if attrs & crate::formula::mtef::constants::XF_EMBELL != 0 {
            Some(Box::new(self.parse_embell()?))
        } else {
            None
        };

        Ok(MtefChar {
            nudge_x,
            nudge_y,
            atts: attrs,
            typeface,
            character,
            bits16,
            embellishment_list,
        })
    }

    fn parse_template(&mut self) -> Result<MtefTemplate, MtefError> {
        let attrs = self.get_attribute()?;

        let mut nudge_x = 0i16;
        let mut nudge_y = 0i16;
        if attrs & XF_LMOVE != 0 {
            let nudge_result = self.get_nudge()?;
            nudge_x = nudge_result.0;
            nudge_y = nudge_result.1;
        }

        let selector = self.read_u8()?;
        let mut variation = self.read_u8()? as u16;

        if self.mtef_version == 5 && (variation & 0x80) != 0 {
            variation &= 0x7F;
            variation |= (self.read_u8()? as u16) << 7;
        }

        let options = self.read_u8()?;

        let subobject_list = if attrs & XF_NULL != 0 {
            None
        } else {
            self.parse_object_list(0)?
        };

        Ok(MtefTemplate {
            nudge_x,
            nudge_y,
            selector,
            variation,
            options,
            subobject_list,
        })
    }

    fn parse_line(&mut self) -> Result<MtefLine, MtefError> {
        let attrs = self.get_attribute()?;

        let mut nudge_x = 0i16;
        let mut nudge_y = 0i16;
        if attrs & XF_LMOVE != 0 {
            let nudge_result = self.get_nudge()?;
            nudge_x = nudge_result.0;
            nudge_y = nudge_result.1;
        }

        let line_spacing = if attrs & XF_LSPACE != 0 {
            self.read_u8()?
        } else {
            0
        };

        let ruler = if attrs & XF_RULER != 0 {
            Some(Box::new(self.parse_ruler()?))
        } else {
            None
        };

        let object_list = self.parse_object_list(0)?;

        Ok(MtefLine {
            nudge_x,
            nudge_y,
            line_spacing,
            ruler,
            object_list,
        })
    }

    fn parse_pile(&mut self) -> Result<MtefPile, MtefError> {
        let attrs = self.get_attribute()?;

        let mut nudge_x = 0i16;
        let mut nudge_y = 0i16;
        if attrs & XF_LMOVE != 0 {
            let nudge_result = self.get_nudge()?;
            nudge_x = nudge_result.0;
            nudge_y = nudge_result.1;
        }

        let halign = self.read_u8()?;
        let valign = self.read_u8()?;

        let ruler = if attrs & XF_RULER != 0 {
            Some(Box::new(self.parse_ruler()?))
        } else {
            None
        };

        let line_list = self.parse_object_list(0)?;

        Ok(MtefPile {
            nudge_x,
            nudge_y,
            halign,
            valign,
            ruler,
            line_list,
        })
    }

    fn parse_matrix(&mut self) -> Result<MtefMatrix, MtefError> {
        let attrs = self.get_attribute()?;

        let mut nudge_x = 0i16;
        let mut nudge_y = 0i16;
        if attrs & XF_LMOVE != 0 {
            let nudge_result = self.get_nudge()?;
            nudge_x = nudge_result.0;
            nudge_y = nudge_result.1;
        }

        let valign = self.read_u8()?;
        let h_just = self.read_u8()?;
        let v_just = self.read_u8()?;
        let rows = self.read_u8()?;
        let cols = self.read_u8()?;

        // Read row and column partitions
        let mut row_parts = [0u8; 16];
        let mut col_parts = [0u8; 16];

        // Row partition consists of (rows+1) two-bit values
        let row_bytes = (2 * (rows as usize + 1)).div_ceil(8);
        for i in 0..row_bytes {
            if i < row_parts.len() {
                row_parts[i] = self.read_u8()?;
            }
        }

        // Col partition consists of (cols+1) two-bit values
        let col_bytes = (2 * (cols as usize + 1)).div_ceil(8);
        for i in 0..col_bytes {
            if i < col_parts.len() {
                col_parts[i] = self.read_u8()?;
            }
        }

        let element_list = self.parse_object_list(0)?;

        Ok(MtefMatrix {
            nudge_x,
            nudge_y,
            valign,
            h_just,
            v_just,
            rows,
            cols,
            row_parts,
            col_parts,
            element_list,
        })
    }

    fn parse_embell(&mut self) -> Result<MtefEmbell, MtefError> {
        let attrs = self.get_attribute()?;

        let mut nudge_x = 0i16;
        let mut nudge_y = 0i16;
        if attrs & XF_LMOVE != 0 {
            let nudge_result = self.get_nudge()?;
            nudge_x = nudge_result.0;
            nudge_y = nudge_result.1;
        }

        let embell = self.read_u8()?;

        Ok(MtefEmbell {
            nudge_x,
            nudge_y,
            embell,
            next: None, // Chaining is handled at a higher level
        })
    }

    fn parse_ruler(&mut self) -> Result<MtefRuler, MtefError> {
        // If we arrived here from LINE, skip the RULER tag if present
        let tag = if self.mtef_version == 5 {
            self.data[self.pos]
        } else {
            self.data[self.pos] & 0x0F
        };
        if tag == crate::formula::mtef::constants::RULER {
            self.pos += 1; // Skip the ruler tag
        }

        let n_stops = self.read_u8()? as i16;
        let mut head: Option<Box<MtefTabstop>> = None;
        let mut curr: Option<*mut MtefTabstop> = None;

        for _ in 0..n_stops {
            let r#type = self.read_u8()? as i16;
            let offset = self.read_i16()?;

            let new_tabstop = Box::new(MtefTabstop {
                r#type,
                offset,
                next: None,
            });

            match curr {
                Some(curr_ptr) => unsafe {
                    (*curr_ptr).next = Some(new_tabstop);
                    curr = Some((*curr_ptr).next.as_mut().unwrap().as_mut() as *mut _);
                },
                None => {
                    head = Some(new_tabstop);
                    curr = head.as_mut().map(|n| n.as_mut() as *mut _);
                }
            }
        }

        Ok(MtefRuler {
            n_stops,
            tabstop_list: head,
        })
    }

    fn parse_font(&mut self) -> Result<MtefFont, MtefError> {
        let tface = self.read_u8()? as i32;
        let style = self.read_u8()? as i32;

        // Read null-terminated font name
        let start_pos = self.pos;
        while self.pos < self.data.len() && self.data[self.pos] != 0 {
            self.pos += 1;
        }
        if self.pos >= self.data.len() {
            return Err(MtefError::UnexpectedEof);
        }

        let font_name = std::str::from_utf8(&self.data[start_pos..self.pos])
            .map_err(|_| MtefError::ParseError("Invalid font name encoding".to_string()))?
            .to_string();

        self.pos += 1; // Skip null terminator

        Ok(MtefFont {
            tface,
            style,
            zname: font_name,
        })
    }

    fn parse_size(&mut self) -> Result<MtefSize, MtefError> {
        // Also works in MTEF5 because all supported tags are less than 16
        let tag = self.read_u8()? & 0x0F;

        // FULL or SUB or SUB2 or SYM or SUBSYM
        if (FULL..=SUBSYM).contains(&tag) {
            return Ok(MtefSize {
                r#type: tag as i32,
                lsize: (tag - FULL) as i32,
                dsize: 0,
            });
        }

        let option = self.read_u8()?;

        // Large dsize
        if option == 100 {
            let lsize = self.read_u8()? as i32;
            let mut dsize = self.read_u8()? as i32;
            dsize += (self.read_u8()? as i32) << 8;
            return Ok(MtefSize {
                r#type: option as i32,
                lsize,
                dsize,
            });
        }

        // Explicit point size
        if option == 101 {
            let mut lsize = self.read_u8()? as i32;
            lsize += (self.read_u8()? as i32) << 8;
            return Ok(MtefSize {
                r#type: option as i32,
                lsize,
                dsize: 0,
            });
        }

        // -128 < dsize < 128
        let dsize = (self.read_u8()? as i32) - 128;
        Ok(MtefSize {
            r#type: 0,
            lsize: option as i32,
            dsize,
        })
    }

    fn skip_font_def(&mut self) -> Result<(), MtefError> {
        self.pos += 1; // Skip tag
        let _id = self.read_u8()?;
        while self.pos < self.data.len() && self.data[self.pos] != 0 {
            self.pos += 1;
        }
        self.pos += 1; // Skip null terminator
        Ok(())
    }

    fn skip_eqn_prefs(&mut self) -> Result<(), MtefError> {
        self.pos += 1; // Skip tag
        let _options = self.read_u8()?; // Options byte

        let size_count = self.read_u8()? as usize;
        self.pos += self.skip_nibbles(size_count)?; // Skip size array

        let space_count = self.read_u8()? as usize;
        self.pos += self.skip_nibbles(space_count)?; // Skip space array

        let style_count = self.read_u8()? as usize;
        for _ in 0..style_count {
            let c = self.read_u8()?;
            if c != 0 {
                self.pos += 1; // Skip style data
            }
        }

        Ok(())
    }

    fn skip_encoding_def(&mut self) -> Result<(), MtefError> {
        self.pos += 1; // Skip tag
        while self.pos < self.data.len() && self.data[self.pos] != 0 {
            self.pos += 1;
        }
        self.pos += 1; // Skip null terminator
        Ok(())
    }

    fn skip_future_record(&mut self) -> Result<(), MtefError> {
        self.pos += 1; // Skip tag
        let size = self.read_u16()? as usize;
        self.pos += size;
        Ok(())
    }

    fn skip_unknown_record(&mut self) -> Result<(), MtefError> {
        self.pos += 1; // Skip tag
        let size = self.read_u16()? as usize;
        self.pos += size;
        Ok(())
    }

    fn skip_nibbles(&mut self, count: usize) -> Result<usize, MtefError> {
        let bytes = count.div_ceil(2); // 2 nibbles per byte
        for _ in 0..bytes {
            self.read_u8()?;
        }
        Ok(bytes)
    }

    // Helper methods for reading binary data with bounds checking
    #[inline]
    fn read_u8(&mut self) -> Result<u8, MtefError> {
        if self.pos >= self.data.len() {
            return Err(MtefError::UnexpectedEof);
        }
        let val = unsafe { *self.data.get_unchecked(self.pos) };
        self.pos += 1;
        Ok(val)
    }

    #[inline]
    fn read_i16(&mut self) -> Result<i16, MtefError> {
        if self.pos + 2 > self.data.len() {
            return Err(MtefError::UnexpectedEof);
        }
        let val = i16::from_le_bytes([
            unsafe { *self.data.get_unchecked(self.pos) },
            unsafe { *self.data.get_unchecked(self.pos + 1) }
        ]);
        self.pos += 2;
        Ok(val)
    }

    #[inline]
    fn read_u16(&mut self) -> Result<u16, MtefError> {
        if self.pos + 2 > self.data.len() {
            return Err(MtefError::UnexpectedEof);
        }
        let val = u16::from_le_bytes([
            unsafe { *self.data.get_unchecked(self.pos) },
            unsafe { *self.data.get_unchecked(self.pos + 1) }
        ]);
        self.pos += 2;
        Ok(val)
    }
}
