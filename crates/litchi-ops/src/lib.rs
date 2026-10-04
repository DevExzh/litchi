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
//! The crate is a scaffold. It exposes no public API yet: the module layout
//! below only fixes where each part of the design will live. Format features
//! are declared but do not yet pull in any leaf crate.
//!
//! # Layout
//!
//! - `file`: the flat `File` enum, including `Locked` and `Signed` states.
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
mod error;
mod file;
mod handler;
mod node;
mod path;
mod registry;
