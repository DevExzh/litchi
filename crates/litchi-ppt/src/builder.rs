//! Canonical PPT authoring facade.
//!
//! [`Builder`] retains the complete native authoring surface while its
//! publication boundary is [`Builder::build`], which returns the immutable
//! [`crate::Package`] owner. Low-level codec integrations can use
//! [`crate::raw::encode`] explicitly.
//!
//! Drafts reached from the builder live here without a prefix (record 0351
//! §2.1 rule 7): [`Chart`](crate::builder::Chart) is a standalone Graph chart draft that
//! [`Builder::add_chart`](crate::builder::Builder::add_chart) places on a slide.

pub use crate::writer::Builder;
pub use crate::writer::chart::{Chart, ChartKind, ChartSeries};

/// Bounded drafts of standalone Graph chart packages.
pub mod graph;
