//! Checked extended BIFF range-sort configuration.
//!
//! These values model the wider `Rw12`/`Col12` domains used by `SortData`;
//! they intentionally do not reuse the smaller BIFF8 cell-grid coordinates.

pub use crate::sort_data::{
    Axis, CONTINUE_FRT12_RECORD_TYPE, Col, Config, Dxf, Icon, IconSet, Key, Method, On, Parent,
    Range, Row, SORT_DATA_RECORD_TYPE,
};
