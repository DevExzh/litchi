//! Shared drawing-style and media engines of the iWork owners (record
//! 0375a, ADR 0024 2026-10-26 amendment).
//!
//! The Pages, Numbers and Keynote packages consume this crate privately; no
//! facade re-exports it (ADR 0028's rule for focused low-level crates). It
//! joins the archive-free values of `litchi-iwa-common` with the bounded
//! wire codecs of `litchi-iwa-protos` and the IWA objects of
//! `litchi-iwa-core`:
//!
//! - [`chart_font`]: the chart-font graph engine (paragraph-style font
//!   variations of a chart's text);
//! - [`chart_style`]: the chart-level and axis-level style properties and
//!   their in-place rewrites;
//! - [`chart_text`]: the shared chart title and axis-title wire engine
//!   (caption authoring remains a deferred host-only follow-up);
//! - [`drawing`]: the stored drawing values (colors, strokes, fills) in both
//!   directions;
//! - [`media`]: the embedded-media inventory of a package;
//! - [`graph`]: a copy-on-write workspace over the parsed IWA components of
//!   a package and its package metadata registry, for the object-graph
//!   edits of style variations;
//! - [`series`]: the style-variation graph engine of per-series chart
//!   overrides (allocating, registering and releasing the override objects
//!   a property write needs) and the per-series properties;
//! - [`number_format`]: the dual chart number-format codec (axis and
//!   value-label formatters);
//! - [`reference_line`]: the reference-line graph engine of a chart's
//!   primary value axis;
//! - [`shape`]: the shape-style variation graph engine (copy-on-write
//!   `TSWP.ShapeStyleArchive` variations) and the fill and stroke;
//! - [`shape_path`]: the preset path sources a shape draws, in both
//!   directions.
//!
//! Every entry point that reads untrusted bytes takes explicit limits.

#![forbid(unsafe_code)]

pub mod chart_font;
pub mod chart_style;
pub mod chart_text;
pub mod drawing;
pub mod graph;
pub mod legend_style;
pub mod line_end;
pub mod media;
pub mod number_format;
pub mod reference_line;
pub mod series;
pub mod shape;
pub mod shape_path;
