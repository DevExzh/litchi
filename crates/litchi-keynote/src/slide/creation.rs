//! Selector-first creation of Keynote slides: duplicate an existing slide or
//! add a fresh slide from a theme layout, as one exact-source transaction.
//!
//! Start an edit with [`Package::edit_slide_creation`](crate::Package::edit_slide_creation),
//! stage any number of [`Edit::duplicate_slide`] and [`Edit::add_slide`]
//! operations, and [`Edit::commit`] them as one verified package. The
//! theme's layouts are listed by
//! [`Package::slide_layouts`](crate::Package::slide_layouts).
//!
//! ```no_run
//! use litchi_keynote::{Package, Position};
//! use litchi_keynote::slide::creation::LayoutSelector;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::open("deck.key")?;
//! let mut edit = package.edit_slide_creation();
//! edit.add_slide(LayoutSelector::Name("Title & Bullets"), Position::new(1))?;
//! let (package, patch) = edit.commit()?.into_parts();
//! let restored = package.apply_slide_creation(&patch.inverse())?;
//! # let _ = restored;
//! # Ok(())
//! # }
//! ```
//!
//! Supported content: slides whose objects live in their own slide
//! component and reference only that component, the document, the document
//! stylesheet and their layout; layouts with placeholders, text boxes,
//! images, live-video and media-placeholder movies. Refused with
//! [`Refusal::UnsupportedContent`]: slides with tables, charts or other
//! content stored in other components or held by native messages without a
//! reference schema, layouts with build or notes state, and decks with no
//! speaker-notes graph to reuse for a fresh slide.

pub use crate::package::slide_creation::{
    Commit, Diagnostics, Edit, Error, Layout, LayoutSelector, LimitKind, Patch, Refusal,
    Unsupported,
};
