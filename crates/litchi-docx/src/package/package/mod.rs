//! Semantic DOCX package facade.
//!
//! The public Package API is assembled from focused relationship, document,
//! data-store, settings, merge, and part-publication layers. The sibling
//! model and codec modules own state and archive I/O respectively.

mod access;
mod alternatives;
mod data_stores;
mod document;
mod merge;
mod parts;
mod settings;
mod validation;
