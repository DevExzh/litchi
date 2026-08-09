//! Shared frame-edit contract for flat and packaged ODI transactions.
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The public contract precedes its compact implementation macro."
)]

use crate::source::Source;
use litchi_core::Result;

/// The common lossless frame mutation surface implemented by both artifact kinds.
///
/// Generic callers can stage the same semantic edits on [`crate::Edit`] and
/// [`crate::FlatImageTransaction`]. Publication remains artifact-specific so a
/// package can also stage resource and metadata changes.
pub trait FrameEditor {
    /// Replaces one optional frame name.
    fn set_frame_name(&mut self, frame: usize, value: Option<String>) -> Result<()>;
    /// Replaces one linked URI or inline byte payload without changing representation.
    fn set_source(&mut self, frame: usize, value: Source) -> Result<()>;
    /// Replaces one optional graphic style reference.
    fn set_style_name(&mut self, frame: usize, value: Option<String>) -> Result<()>;
    /// Replaces one optional text style reference.
    fn set_text_style_name(&mut self, frame: usize, value: Option<String>) -> Result<()>;
    /// Replaces one optional drawing layer.
    fn set_layer(&mut self, frame: usize, value: Option<String>) -> Result<()>;
    /// Replaces one optional non-negative z-index.
    fn set_z_index(&mut self, frame: usize, value: Option<u32>) -> Result<()>;
    /// Replaces one optional lexical transform.
    fn set_transform(&mut self, frame: usize, value: Option<String>) -> Result<()>;
    /// Replaces one optional anchoring mode.
    fn set_anchor_type(&mut self, frame: usize, value: Option<String>) -> Result<()>;
    /// Replaces lexical position and size values.
    fn set_geometry(
        &mut self,
        frame: usize,
        x: Option<String>,
        y: Option<String>,
        width: Option<String>,
        height: Option<String>,
    ) -> Result<()>;
    /// Replaces lexical relative size values.
    fn set_relative_size(
        &mut self,
        frame: usize,
        width: Option<String>,
        height: Option<String>,
    ) -> Result<()>;
}

macro_rules! impl_frame_editor {
    ($type:ty) => {
        impl FrameEditor for $type {
            fn set_frame_name(&mut self, frame: usize, value: Option<String>) -> Result<()> {
                <$type>::set_frame_name(self, frame, value)
            }
            fn set_source(&mut self, frame: usize, value: Source) -> Result<()> {
                <$type>::set_source(self, frame, value)
            }
            fn set_style_name(&mut self, frame: usize, value: Option<String>) -> Result<()> {
                <$type>::set_style_name(self, frame, value)
            }
            fn set_text_style_name(&mut self, frame: usize, value: Option<String>) -> Result<()> {
                <$type>::set_text_style_name(self, frame, value)
            }
            fn set_layer(&mut self, frame: usize, value: Option<String>) -> Result<()> {
                <$type>::set_layer(self, frame, value)
            }
            fn set_z_index(&mut self, frame: usize, value: Option<u32>) -> Result<()> {
                <$type>::set_z_index(self, frame, value)
            }
            fn set_transform(&mut self, frame: usize, value: Option<String>) -> Result<()> {
                <$type>::set_transform(self, frame, value)
            }
            fn set_anchor_type(&mut self, frame: usize, value: Option<String>) -> Result<()> {
                <$type>::set_anchor_type(self, frame, value)
            }
            fn set_geometry(
                &mut self,
                frame: usize,
                x: Option<String>,
                y: Option<String>,
                width: Option<String>,
                height: Option<String>,
            ) -> Result<()> {
                <$type>::set_geometry(self, frame, x, y, width, height)
            }
            fn set_relative_size(
                &mut self,
                frame: usize,
                width: Option<String>,
                height: Option<String>,
            ) -> Result<()> {
                <$type>::set_relative_size(self, frame, width, height)
            }
        }
    };
}

impl_frame_editor!(crate::FlatImageTransaction);
impl_frame_editor!(crate::Edit<'_>);
