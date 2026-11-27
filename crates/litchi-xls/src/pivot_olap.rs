//! Layered BIFF8 PivotTable OLAP extension owner.
//!
//! Semantic models are separate from binary codecs, validation, and complete
//! `SXVIEWEX` sequence packaging while this file remains the stable public
//! facade used by the XLS crate.

mod codec;
mod model;
mod package;
mod semantic;
mod validation;

#[cfg(test)]
mod tests;

/// Record type of the `SXPIEx` record (MS-XLS 2.4.299).
pub(crate) const SXPI_EX_RECORD_TYPE: u16 = 0x080C;
/// Record type of the `SXViewEx` record (MS-XLS 2.4.314).
pub(crate) const SX_VIEW_EX_RECORD_TYPE: u16 = 0x080E;
/// Record type of the `SXVDTEx` record (MS-XLS 2.4.311).
pub(crate) const SXVDT_EX_RECORD_TYPE: u16 = 0x080F;
/// Record type of the `SXTH` record (MS-XLS 2.4.308).
pub(crate) const SXTH_RECORD_TYPE: u16 = 0x00DB;

/// `FrtHeaderOld.rt` mandated inside `SXViewEx` (MS-XLS 2.4.314). This is
/// not the `SXViewEx` record type; see the module documentation.
pub(crate) const SX_VIEW_EX_FRT_RT: u16 = 0x080C;
/// `FrtHeaderOld.rt` mandated inside `SXPIEx` (MS-XLS 2.4.299). This is not
/// the `SXPIEx` record type; see the module documentation.
pub(crate) const SXPI_EX_FRT_RT: u16 = 0x080E;
/// `FrtHeaderOld.rt` mandated inside `SXVDTEx` (MS-XLS 2.4.311).
pub(crate) const SXVDT_EX_FRT_RT: u16 = SXVDT_EX_RECORD_TYPE;
/// `FrtHeaderOld.rt` mandated inside `SXTH` (MS-XLS 2.4.308). This is not
/// the `SXTH` record type; see the module documentation.
pub(crate) const SXTH_FRT_RT: u16 = 0x080D;

/// Size in bytes of an `FrtHeaderOld` (MS-XLS 2.5.136).
pub(crate) const FRT_HEADER_OLD_LEN: usize = 4;
/// Maximum character count of the `XLUnicodeString` fields in these records
/// (MS-XLS 2.4.299, 2.4.308).
pub(crate) const MAX_OLAP_STRING_CHARS: usize = 255;
/// Maximum byte count of `SXViewEx.rgbFuture` (MS-XLS 2.4.314).
pub(crate) const MAX_FUTURE_BYTES: usize = 1_024;

pub use model::{
    HiddenMemberSet, PivotFieldOlapExt, PivotHierarchy, PivotHierarchyAxis, PivotItemOlapFlags,
    PivotPageItemOlapExt, PivotViewOlapHeader,
};
pub use package::OlapSequence;
