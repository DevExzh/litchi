//! Obj/OLE invariants from MS-XLS.

use super::super::semantic::{FtPioGrbit, ObjSubrecord, OleObjectRecord};
use super::super::{OBJ, invalid};
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

impl OleObjectRecord {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn validate(&self) -> Result<()> {
        if self.subrecords.len() > 1_024 {
            return Err(invalid(OBJ, "too many Obj subrecords"));
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
        if self
            .subrecords
            .iter()
            .filter(|value| matches!(value, ObjSubrecord::PictureFormula(_)))
            .count()
            > 1
        {
            return Err(invalid(OBJ, "duplicate FtPictFmla"));
        }
        if !matches!(self.subrecords.last(), Some(ObjSubrecord::End)) {
            return Err(invalid(OBJ, "OLE Obj must end with FtEnd"));
        }
        Ok(())
    }
}
