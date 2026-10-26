//! Shared drawing-style and media engines of the iWork owners (record
//! 0375a, ADR 0024 2026-10-26 amendment).
//!
//! The Pages, Numbers and Keynote packages consume this crate privately; no
//! facade re-exports it (ADR 0028's rule for focused low-level crates). It
//! joins the archive-free values of `litchi-iwa-common` with the bounded
//! wire codecs of `litchi-iwa-protos` and the IWA objects of
//! `litchi-iwa-core`:
//!
//! - [`drawing`]: the stored drawing values (colors, strokes, fills) in both
//!   directions;
//! - [`media`]: the embedded-media inventory of a package;
//! - `graph` (next): the object graph edits of style variations —
//!   allocating, registering and releasing the style objects a property
//!   write needs — over a workspace of parsed IWA components.
//!
//! Every entry point that reads untrusted bytes takes explicit limits.

#![forbid(unsafe_code)]

pub mod drawing;
pub mod media;
