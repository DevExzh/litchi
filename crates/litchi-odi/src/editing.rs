//! Shared frame-edit contract for flat and packaged ODI transactions.
use crate::map::ImageMap;
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
    /// Replaces one optional typed client-side image map.
    fn set_image_map(&mut self, frame: usize, value: Option<ImageMap>) -> Result<()>;
    /// Replaces the lexical horizontal position.
    fn set_x(&mut self, frame: usize, value: Option<String>) -> Result<()>;
    /// Replaces the lexical vertical position.
    fn set_y(&mut self, frame: usize, value: Option<String>) -> Result<()>;
    /// Replaces the lexical width.
    fn set_width(&mut self, frame: usize, value: Option<String>) -> Result<()>;
    /// Replaces the lexical height.
    fn set_height(&mut self, frame: usize, value: Option<String>) -> Result<()>;
    /// Replaces the lexical relative width.
    fn set_relative_width(&mut self, frame: usize, value: Option<String>) -> Result<()>;
    /// Replaces the lexical relative height.
    fn set_relative_height(&mut self, frame: usize, value: Option<String>) -> Result<()>;
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
            fn set_image_map(&mut self, frame: usize, value: Option<ImageMap>) -> Result<()> {
                <$type>::set_image_map(self, frame, value)
            }
            fn set_x(&mut self, frame: usize, value: Option<String>) -> Result<()> {
                <$type>::set_x(self, frame, value)
            }
            fn set_y(&mut self, frame: usize, value: Option<String>) -> Result<()> {
                <$type>::set_y(self, frame, value)
            }
            fn set_width(&mut self, frame: usize, value: Option<String>) -> Result<()> {
                <$type>::set_width(self, frame, value)
            }
            fn set_height(&mut self, frame: usize, value: Option<String>) -> Result<()> {
                <$type>::set_height(self, frame, value)
            }
            fn set_relative_width(&mut self, frame: usize, value: Option<String>) -> Result<()> {
                <$type>::set_relative_width(self, frame, value)
            }
            fn set_relative_height(&mut self, frame: usize, value: Option<String>) -> Result<()> {
                <$type>::set_relative_height(self, frame, value)
            }
        }
    };
}

impl_frame_editor!(crate::FlatImageTransaction);
impl_frame_editor!(crate::Edit<'_>);

/// Common metadata mutation contract for flat and packaged ODI transactions.
pub trait MetadataEditor {
    /// Replaces the document title.
    fn set_title(&mut self, value: Option<String>) -> Result<()>;
    /// Replaces the document author.
    fn set_author(&mut self, value: Option<String>) -> Result<()>;
    /// Replaces the document subject.
    fn set_subject(&mut self, value: Option<String>) -> Result<()>;
    /// Replaces the document description.
    fn set_description(&mut self, value: Option<String>) -> Result<()>;
    /// Replaces the comma-separated keyword value.
    fn set_keywords(&mut self, value: Option<String>) -> Result<()>;
}

macro_rules! impl_metadata_editor {
    ($type:ty) => {
        impl MetadataEditor for $type {
            fn set_title(&mut self, value: Option<String>) -> Result<()> {
                <$type>::set_title(self, value)
            }
            fn set_author(&mut self, value: Option<String>) -> Result<()> {
                <$type>::set_author(self, value)
            }
            fn set_subject(&mut self, value: Option<String>) -> Result<()> {
                <$type>::set_subject(self, value)
            }
            fn set_description(&mut self, value: Option<String>) -> Result<()> {
                <$type>::set_description(self, value)
            }
            fn set_keywords(&mut self, value: Option<String>) -> Result<()> {
                <$type>::set_keywords(self, value)
            }
        }
    };
}

impl_metadata_editor!(crate::FlatImageTransaction);
impl_metadata_editor!(crate::Edit<'_>);
