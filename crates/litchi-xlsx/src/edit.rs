//! Transactional workbook editing entry points.
//!
//! The canonical edit implementation remains next to the immutable workbook
//! snapshot; this module exposes that owner as the standalone edit layer.

pub use crate::workbook::edit::{
    ActiveTab, Change, Column, Commit, Conflict, ConflictSet, Defaults, Edit, FormControl,
    JoinError, JoinFailure, NewSheet, PackageChange, Patch, PictureSelector, Row, Sheet, State,
    SvgInput, Tab,
};
