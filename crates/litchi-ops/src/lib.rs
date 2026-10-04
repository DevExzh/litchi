//! Unified operations layer for Litchi.
//!
//! This crate is the home of the format-agnostic operations surface agreed in
//! `docs/UNIFIED_OPS_API_DESIGN.md` and proposed for acceptance as ADR 0033:
//! one way to open, read, query, and edit any document Litchi understands,
//! dispatched statically to the existing per-format owners and their
//! snapshot/Edit/Commit/Patch transactions. It is a dispatch-and-coercion
//! layer, not a second document model.
//!
//! The `litchi` facade will depend on this crate and re-export its entry
//! points, so users continue to see a single `litchi::open()`.
//!
//! # Status
//!
//! The crate is being built step by step (record 0003 §4). Its public API so
//! far is the unified [`error`] module and the open-time vocabulary:
//! [`OpenOptions`], [`SourceMode`], [`Family`], and the [`Locked`] and
//! [`Signed`] shells. The module layout below fixes where each remaining part
//! of the design will live.
//!
//! # Layout
//!
//! - `file`: the flat `File` enum, including `Locked` and `Signed` states.
//! - `options`: `OpenOptions` and `SourceMode`.
//! - `error`: the unified typed error and its closed code list.
//! - `path`: re-export point for the `litchi_core::selector` grammar.
//! - `node`: the generic node view, projected on demand.
//! - `registry`: the static property registry and coercions.
//! - `batch`: the atomic and best-effort batch executor.
//! - `capabilities`: capability reporting.
//! - `handler`: the internal, statically dispatched per-format handlers.

#![forbid(unsafe_code)]

mod batch;
mod capabilities;
pub mod error;
mod file;
mod handler;
mod node;
mod options;
mod path;
mod registry;

pub use error::{Code, Error, Result};
pub use file::{EncryptionKind, Family, Locked, SignatureOrigin, Signed};
#[cfg(feature = "doc")]
pub use options::DocOptions;
pub use options::{OpenOptions, SourceMode};
