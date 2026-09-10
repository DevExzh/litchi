//! Obj/OLE invariants from MS-XLS.

use super::super::semantic::{FtCf, FtPictFmla, FtPioGrbit, ObjSubrecord, OleObjectRecord};
use super::super::{
    FT_CBLS_DATA, FT_EDO_DATA, FT_GBO_DATA, FT_LBS_DATA, FT_RBO_DATA, FT_SBS, OBJ, invalid,
};
use crate::error::Result;

impl FtPioGrbit {
    pub(super) fn validate(self) -> Result<()> {
        if self.is_dde() && self.is_control() {
            return Err(invalid(
                OBJ,
                "FtPioGrbit DDE and control flags are mutually exclusive",
            ));
        }
        Ok(())
    }
}

impl FtCf {
    pub(crate) fn validate(self) -> Result<()> {
        if Self::new(self.format).is_none() {
            return Err(invalid(
                OBJ,
                format!(
                    "FtCf contains unsupported clipboard format 0x{:04X}",
                    self.format
                ),
            ));
        }
        Ok(())
    }
}

pub(crate) fn validate_picture_formula(formula: &FtPictFmla, flags: FtPioGrbit) -> Result<()> {
    // The object formula is an ObjFmla followed by optional embedInfo and
    // padding.  For an ordinary embedded object MS-XLS requires cce=5 and a
    // single five-byte PtgTbl expression.  The outer cbFmla is represented by
    // `formula.len()` in the semantic model and MUST be even.
    if formula.formula.len() < 14 || formula.formula.len() % 2 != 0 {
        return Err(invalid(
            OBJ,
            "FtPictFmla ObjectParsedFormula length must be even and include PtgTbl/embedInfo",
        ));
    }
    let cce = u16::from_le_bytes([formula.formula[0], formula.formula[1]]);
    if cce != 5 || cce & 0x8000 != 0 {
        return Err(invalid(
            OBJ,
            "embedded FtPictFmla requires reserved-bit-free cce=5",
        ));
    }
    if formula.formula[6] != 0x02 {
        return Err(invalid(
            OBJ,
            "embedded FtPictFmla requires a leading PtgTbl",
        ));
    }
    if formula.formula[11] != 0x03 || formula.formula[13] != 0 {
        return Err(invalid(
            OBJ,
            "FtPictFmlaEmbedInfo has an invalid reserved header",
        ));
    }
    let class_bytes = usize::from(formula.formula[12]);
    if class_bytes != 0 {
        let class_end = 14usize
            .checked_add(class_bytes)
            .ok_or_else(|| invalid(OBJ, "FtPictFmla class length overflow"))?;
        if class_end > formula.formula.len() {
            return Err(invalid(OBJ, "FtPictFmla class name is truncated"));
        }
        let high_byte = formula.formula[14];
        if high_byte & !1 != 0 {
            return Err(invalid(
                OBJ,
                "FtPictFmla class string has invalid reserved bits",
            ));
        }
        // XLUnicodeStringNoCch is a one-byte fHighByte followed by rgb.
        // When fHighByte is set, rgb contains two bytes per character; the
        // cbClass byte count must therefore leave an even number of bytes
        // after the flag.  Do not accept a one-byte RGB tail as a truncated
        // UTF-16 class name merely because it fits inside the formula.
        if high_byte & 1 != 0 && class_bytes.saturating_sub(1) % 2 != 0 {
            return Err(invalid(
                OBJ,
                "FtPictFmla class string has an odd UTF-16 byte count",
            ));
        }
    }
    if formula.storage_position.is_none() {
        return Err(invalid(OBJ, "FtPictFmla with PtgTbl requires lPosInCtlStm"));
    }
    if flags.uses_control_stream() != formula.control_buffer_size.is_some() {
        return Err(invalid(
            OBJ,
            "FtPictFmla control-stream fields do not match fPrstm",
        ));
    }
    Ok(())
}

fn validate_subrecord_owner(value: &ObjSubrecord, object_type: u16) -> Result<()> {
    let (allowed, name) = match value {
        ObjSubrecord::PictureFormat(_)
        | ObjSubrecord::PictureFlags(_)
        | ObjSubrecord::PictureFormula(_) => (object_type == 0x0008, "picture fields"),
        ObjSubrecord::CheckBoxData(_)
        | ObjSubrecord::Unknown {
            kind: FT_CBLS_DATA, ..
        } => (matches!(object_type, 0x000B | 0x000C), "FtCblsData"),
        ObjSubrecord::RadioButtonData(_)
        | ObjSubrecord::Unknown {
            kind: FT_RBO_DATA, ..
        } => (object_type == 0x000C, "FtRboData"),
        ObjSubrecord::EditBoxData(_)
        | ObjSubrecord::Unknown {
            kind: FT_EDO_DATA, ..
        } => (object_type == 0x000D, "FtEdoData"),
        ObjSubrecord::GroupBoxData(_)
        | ObjSubrecord::Unknown {
            kind: FT_GBO_DATA, ..
        } => (object_type == 0x0013, "FtGboData"),
        ObjSubrecord::ScrollBarData(_) | ObjSubrecord::Unknown { kind: FT_SBS, .. } => (
            matches!(object_type, 0x0010 | 0x0011 | 0x0012 | 0x0014),
            "FtSbs",
        ),
        ObjSubrecord::ListBoxData(_)
        | ObjSubrecord::Unknown {
            kind: FT_LBS_DATA, ..
        } => (matches!(object_type, 0x0012 | 0x0014), "FtLbsData"),
        _ => return Ok(()),
    };
    if allowed {
        Ok(())
    } else {
        Err(invalid(
            OBJ,
            format!("{name} is not valid for Obj cmo.ot 0x{object_type:04X}"),
        ))
    }
}

impl OleObjectRecord {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn validate(&self) -> Result<()> {
        if self.subrecords.len() > 1_024 {
            return Err(invalid(OBJ, "too many Obj subrecords"));
        }
        if self
            .subrecords
            .iter()
            .any(|value| matches!(value, ObjSubrecord::ClipboardFormat(_)))
        {
            return Err(invalid(
                OBJ,
                "OLE Obj contains a malformed FtCf pictFormat subrecord",
            ));
        }
        let common = self
            .subrecords
            .iter()
            .filter_map(|value| match value {
                ObjSubrecord::Common(value) => Some(value),
                _ => None,
            })
            .collect::<Vec<_>>();
        if common.len() != 1
            || common[0].object_type != 8
            || common[0].object_id == 0
            || !matches!(self.subrecords.first(), Some(ObjSubrecord::Common(_)))
        {
            return Err(invalid(
                OBJ,
                "OLE Obj requires a leading FtCmo type 8 with nonzero ID",
            ));
        }
        for value in &self.subrecords {
            validate_subrecord_owner(value, common[0].object_type)?;
        }
        // MS-XLS 2.4.181 places the picture fields in a fixed order.  Unknown
        // subrecords are retained between known fields, but they cannot make
        // a later mandatory field appear before an earlier one.  Checking the
        // order here also prevents a malformed extra FtCf from being hidden
        // behind an otherwise valid pictFormat.
        let mut picture_stage = 0u8;
        let mut end_count = 0usize;
        for (index, value) in self.subrecords.iter().enumerate() {
            match value {
                ObjSubrecord::Common(_) if index != 0 => {
                    return Err(invalid(OBJ, "OLE Obj contains a second FtCmo"));
                },
                ObjSubrecord::PictureFormat(_) => {
                    if picture_stage > 1 {
                        return Err(invalid(
                            OBJ,
                            "OLE Obj pictFormat is out of MS-XLS field order",
                        ));
                    }
                    picture_stage = 1;
                },
                ObjSubrecord::PictureFlags(_) => {
                    if picture_stage > 2 {
                        return Err(invalid(
                            OBJ,
                            "OLE Obj pictFlags is out of MS-XLS field order",
                        ));
                    }
                    picture_stage = 2;
                },
                ObjSubrecord::PictureFormula(_) => {
                    if picture_stage > 3 {
                        return Err(invalid(
                            OBJ,
                            "OLE Obj pictFmla is out of MS-XLS field order",
                        ));
                    }
                    picture_stage = 3;
                },
                ObjSubrecord::End => {
                    end_count = end_count.saturating_add(1);
                    if index + 1 != self.subrecords.len() {
                        return Err(invalid(OBJ, "OLE Obj FtEnd is not the final subrecord"));
                    }
                },
                _ => {},
            }
        }
        if end_count != 1 {
            return Err(invalid(OBJ, "OLE Obj requires exactly one FtEnd"));
        }
        if picture_stage != 3 {
            return Err(invalid(
                OBJ,
                "OLE Obj mandatory picture fields are incomplete",
            ));
        }
        let pio = self
            .subrecords
            .iter()
            .filter_map(|value| match value {
                ObjSubrecord::PictureFlags(value) => Some(*value),
                _ => None,
            })
            .collect::<Vec<_>>();
        if pio.len() != 1 {
            return Err(invalid(OBJ, "OLE Obj requires one FtPioGrbit"));
        }
        pio[0].validate()?;
        if pio[0].is_control() || pio[0].uses_control_stream() {
            return Err(invalid(
                OBJ,
                "OLE Obj data must be in an embedding or link storage",
            ));
        }
        let picture_formats = self
            .subrecords
            .iter()
            .filter_map(|value| match value {
                ObjSubrecord::PictureFormat(value) => Some(*value),
                _ => None,
            })
            .collect::<Vec<_>>();
        if picture_formats.len() != 1 {
            return Err(invalid(OBJ, "OLE Obj requires exactly one FtCf pictFormat"));
        }
        picture_formats[0].validate()?;
        let formulas = self
            .subrecords
            .iter()
            .filter_map(|value| match value {
                ObjSubrecord::PictureFormula(value) => Some(value),
                _ => None,
            })
            .collect::<Vec<_>>();
        if formulas.len() > 1 {
            return Err(invalid(OBJ, "duplicate FtPictFmla"));
        }
        if common[0].object_type == 8 {
            if formulas.len() != 1 {
                return Err(invalid(OBJ, "OLE Obj requires one FtPictFmla"));
            }
            if !pio[0].is_dde() && !pio[0].camera_picture() {
                validate_picture_formula(formulas[0], pio[0])?;
            }
        }
        if !matches!(self.subrecords.last(), Some(ObjSubrecord::End)) {
            return Err(invalid(OBJ, "OLE Obj must end with FtEnd"));
        }
        Ok(())
    }
}
