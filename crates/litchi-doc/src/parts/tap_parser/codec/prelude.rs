//! Shared imports for TAP binary codecs.

pub(super) use super::super::TapParser;
pub(super) use super::super::model::WidthUsage;
pub(super) use super::primitives::{binary_to_doc_result, read_byte};
pub(super) use crate::package::{Error as PackageError, Result};
pub(super) use crate::parts::tap::{
    BorderStyle, BorderType, CellBorderTypes, CellMergeStatus, CellProperties, CellShading,
    CellSpacing, CellSpacingSource, ShadingPattern, TableHorizontalPosition, TableJustification,
    TableProperties, TableStyleBorder, TableStyleShading, TableVerticalPosition, TableWidth,
    TextDirection, VerticalAlignment, VerticalMergeStatus, WidthType,
};
pub(super) use crate::sprm::Sprm;
pub(super) use litchi_core::binary::{BinaryResult, read_i16_le, read_u16_le};
