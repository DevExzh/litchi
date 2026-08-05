//! Private bridge for the chart authoring owner.
//!
//! The public pivot-chart vocabulary lives under [`crate::pivot::chart`].
//! This private module remains only because the legacy chart builder is kept
//! outside the pivot owner for now and must not create a second public API.

pub(crate) use crate::pivot::chart::{
    DEFAULT_FORMAT_ID as DEFAULT_PIVOT_CHART_FORMAT_ID,
    default_options_extension_xml as default_pivot_options_extension_xml,
};
