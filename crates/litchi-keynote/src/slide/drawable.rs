//! Selector-first edits of the drawables a slide owns: read the inventory,
//! set geometry, remove, reorder (z-order) and create text boxes, preset
//! shapes and images, as one exact-source transaction (record 0298).
//!
//! [`Package::slide_drawable_list`](crate::Package::slide_drawable_list)
//! reads a slide's drawables in source order, the order
//! [`DrawableSelector`](crate::DrawableSelector) addresses. Start an edit
//! with [`Package::edit_slide_drawables`](crate::Package::edit_slide_drawables),
//! stage any number of operations and [`Edit::commit`] them as one verified
//! package; [`Package::apply_slide_drawables`](crate::Package::apply_slide_drawables)
//! replays a [`Patch`] or its inverse.
//!
//! ```no_run
//! use litchi_keynote::slide::drawable::{Geometry, Preset};
//! use litchi_keynote::{DrawableSelector, Package, Position, SlideSelector};
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::open("deck.key")?;
//! let slide = SlideSelector::index(0);
//! let mut edit = package.edit_slide_drawables();
//! edit.add_text_box(slide, Geometry::new(80.0, 80.0, 400.0, 60.0), "Agenda")?
//!     .add_shape(slide, Preset::Rectangle, Geometry::new(80.0, 200.0, 200.0, 100.0), "")?
//!     .move_to(slide, DrawableSelector::index(0), Position::new(0))?;
//! let (edited, patch) = edit.commit()?.into_parts();
//! let restored = edited.apply_slide_drawables(&patch.inverse())?;
//! # let _ = restored;
//! # Ok(())
//! # }
//! ```
//!
//! Supported: geometry of text boxes, shapes, images and placeholders;
//! plain text of text boxes and shapes ([`Edit::set_text`], a minimal-span
//! rewrite of the drawable's own storage);
//! removal of text boxes, shapes, lines and images with their private
//! objects (registry UUIDs and data owners are dropped; data items stay);
//! z-order moves; new text boxes, rectangle / rounded-rectangle / ellipse
//! shapes and PNG / JPEG / GIF images styled from the theme as the legacy
//! host does; new tables cloned from the deck's first slide table
//! ([`Edit::add_table`], record 0329); table cell values; character runs. Refused with
//! [`Refusal::Unsupported`]: removing tables, charts, groups, movies, audio
//! and placeholders, removing drawables referenced by builds or other
//! objects, geometry of tables, charts, groups, lines and movies, and
//! character-style writes ([`Unsupported::TextStyling`]).

pub use crate::package::slide_drawable::{
    Commit, Created, Diagnostics, Drawable, Edit, Error, Geometry, ImageFormat, Kind, LimitKind,
    NewTable, Patch, Preset, Refusal, Style, TableCellValue, TextRun, TextSpan, Unsupported,
};
/// Hyperlink, lock and accessibility properties ([`Edit::set_properties`]).
pub use crate::slide::media::MediaProperties;
/// The axis of an Arrange Flip ([`Edit::flip`]).
pub use litchi_iwa_common::shape::geometry::FlipAxis;
/// A straight line's segment ([`Edit::set_line_segment`]).
pub use litchi_iwa_common::shape::line::Segment;
/// Validated preset parameters ([`Preset::RegularPolygon`], [`Preset::Star`]),
/// the shared iWork shape vocabulary.
pub use litchi_iwa_common::shape::path::{InnerRadiusRatio, PolygonSides, StarPoints};
