//! Canonical PPT authoring facade.
//!
//! [`Builder`] retains the complete native authoring surface while its
//! publication boundary is [`Builder::build`], which returns the immutable
//! [`crate::Package`] owner. Low-level codec integrations can use
//! [`crate::raw::encode`] explicitly.

pub use crate::writer::Builder;
