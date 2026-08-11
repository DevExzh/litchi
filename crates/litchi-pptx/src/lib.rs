//! Typed `PowerPoint` Open XML documents.
//!
//! The crate is being extracted one semantic capability at a time. The
//! backgrounds module owns package-independent slide-background values and
//! its XML fill codec. The transition module owns slide-transition values and
//! its bounded codec. The
//! laser module owns inert laser-trace values and their bounded codec. The
//! font module owns embedded-font values and atomic package CRUD. The tag
//! module owns inert programmable tag lists and package CRUD. The notes module
//! owns bounded speaker-notes graphs, text encoding, and transactional package
//! mutation.
//! [`table::style`] owns typed table-style catalogs and their package graph.

#![forbid(unsafe_code)]
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the public facade and OOXML models are grouped by presentation feature and schema order"
)]
#![allow(
    clippy::format_push_string,
    reason = "the XML codecs intentionally append infallible formatted fragments to compact String buffers"
)]
#![allow(
    clippy::match_wildcard_for_single_variants,
    clippy::wildcard_enum_match_arm,
    reason = "streaming XML decoders must ignore future quick-xml event variants and unsupported schema content"
)]
#![allow(
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    reason = "codec scopes reuse short XML token and relationship names after earlier values are no longer live"
)]
#![allow(
    clippy::many_single_char_names,
    clippy::module_name_repetitions,
    clippy::similar_names,
    clippy::struct_excessive_bools,
    clippy::struct_field_names,
    reason = "public names preserve established PowerPoint and OOXML vocabulary"
)]
#![allow(
    clippy::needless_pass_by_value,
    reason = "public mutation APIs consume owned values so callers cannot accidentally reuse stale document state"
)]
#![allow(
    clippy::option_option,
    clippy::ptr_arg,
    clippy::ref_option,
    reason = "patch APIs use nested and referenced options to distinguish omission, removal, and replacement without changing their established signatures"
)]
#![allow(
    clippy::unnecessary_wraps,
    reason = "codec and validation helpers retain a uniform Result contract so callers can compose validation transactionally"
)]
#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    reason = "package resource limits bound XML offsets and schema collection sizes below every destination range before these conversions"
)]
#![allow(
    clippy::match_same_arms,
    reason = "separate schema-token arms document distinct OOXML spellings even when their current semantic action is identical"
)]
#![cfg_attr(
    test,
    allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "unit-test fixture construction and assertions panic by design"
    )
)]

pub mod actions;
pub mod animations;
pub mod backgrounds;
pub mod change_tracking;
pub mod chart;
pub mod comments;
pub mod custom;
mod error;
pub mod font;
pub mod format;
pub mod hyperlinks;
pub mod laser;
pub mod media_parts;
pub mod model3d;
pub mod modern_comments;

pub mod master_layout;
pub(crate) mod namespace;
pub mod notes;
pub mod opened;
pub mod package;
pub mod parts;
pub mod presentation;
pub mod presentation_properties;
pub mod shape;
pub mod slide;
pub mod table;
pub mod tag;
pub mod time;
pub mod transition;
pub mod view_properties;
/// Inert Office Add-in and persisted task-pane metadata.
///
/// This module exposes the bounded MS-OWEXML value model used by
/// [`Package::task_panes`]. Add-ins are retained as document data only:
/// external references, bindings, and snapshots are never contacted or
/// executed.
pub mod web;
pub mod writer;

pub(crate) mod resources;

pub use actions::{Jump, Kind, Setting, Target, Trigger};
pub use animations::*;
pub use backgrounds::{GradientStop, GradientType, PatternType, PictureStyle, SlideBackground};
pub use chart::{Chart, Info as ChartInfo, Series as ChartSeries, Type as ChartType};
pub use comments::{
    Author, Comment, Comments, Conformance, List, add_presentation_comment,
    add_presentation_comment_author, find_presentation_comment, find_presentation_comment_author,
    load_presentation_comments, parse_comment_authors, parse_slide_comments,
    remove_presentation_comment, remove_presentation_comment_author,
    reorder_presentation_comment_authors, reorder_presentation_comments,
    replace_presentation_comment, replace_presentation_comment_author, store_presentation_comments,
    update_presentation_comment, update_presentation_comment_author, write_comment_authors,
    write_slide_comments,
};
pub use error::{Error, Result, ShapeTransferRefusal};
pub use format::{ImageFormat, TextFormat};
pub use hyperlinks::Hyperlink;
#[cfg(feature = "encryption")]
pub use litchi_crypto::ooxml as encryption;
#[cfg(feature = "encryption")]
pub use litchi_crypto::ooxml::{
    Limits as EncryptionLimits, Mode as EncryptionMode, Password as EncryptionPassword,
};
#[cfg(feature = "encryption")]
pub use litchi_ooxml_common::package_encryption::{
    PackageEncryption, PolicyError as EncryptionPolicyError,
};
/// Resource policy for package ingestion through [`Package`].
pub use litchi_opc::ReadLimits;
pub use master_layout::{
    AuthoredSlideLayout, AuthoredSlideMaster, MIN_MASTER_OR_LAYOUT_ID, PlaceholderKind,
    PlaceholderSpec, SlideLayoutKind,
};
pub use media_parts::{
    Bookmark, Data, ExtensionList, Fade, Picture, Poster, Resource, Transform, Trim,
    load_slide_media, parse_slide_media, store_slide_media, write_slide_media_pictures,
};
pub use modern_comments::*;
pub use opened::{MAX_SHAPE_TEXT_REPLACEMENTS, ShapeTextReplacement};
pub use package::Package;
pub use presentation::{
    Presentation, SourceBackedPresentation, SourceBackedPresentationEditor,
    SourceBackedSlideCommit, SourceBackedSlideEdit, SourceBackedSlidePatch,
    SourceBackedSlideSnapshot, SourceSlide,
};
pub use presentation_properties::{
    BrowserSupport, Color, ColorKind, Extension, HtmlPublish, HtmlTarget, OpaqueExtension, Print,
    PrintColorMode, PrintOutput, Properties, Show, ShowExtension, ShowMode, SlideSelection, Web,
    WebColor, WebScreenSize, load_from_package as load_presentation_properties,
};
pub use slide::{Key, Slide, SlideLayout, SlideMaster};
pub use view_properties::{
    CommonSlideView, CommonView, GridSpacing, Guide, GuideOrientation, NormalView, OutlineSlide,
    OutlineView, Point, Ratio, RestoredPane, SimpleView, SlideLikeView, SorterView, SplitterState,
    ViewKind, ViewProperties, load_from_package as load_view_properties,
};
pub use writer::{MutablePresentation, MutableShape, MutableSlide};
