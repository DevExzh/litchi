//! Forward-only, bounded authoring of small semantic PresentationML decks.
//!
//! This module deliberately does not hydrate [`super::MutablePresentation`].
//! The caller declares the slide count up front and emits one slide at a time;
//! each completed slide is released after its OPC member is finalized.  The
//! writer is fresh-authoring only and intentionally models plain text boxes.
//! The semantic authoring window is bounded; the ZIP transport still retains
//! central-directory and member-name metadata that grows with the part count.
//! The output budget's preflight check covers only mandatory ZIP structure and
//! member-name metadata; payload, descriptor, and compressor bytes remain
//! runtime work that can report typed incomplete output.

use std::io::{self, Write};

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::packuri::PackURI;
use litchi_opc::phys_pkg::{PartWriter, PhysPkgWriter};

use crate::resources;
use crate::{Error, Result};

const XML_DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";
const DRAWINGML_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PRESENTATIONML_NAMESPACE: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const OFFICE_RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const PRESENTATION_PART: &str = "/ppt/presentation.xml";
const MASTER_PART: &str = "/ppt/slideMasters/slideMaster1.xml";
const THEME_PART: &str = "/ppt/theme/theme1.xml";
const NOTES_THEME_PART: &str = "/ppt/theme/theme2.xml";
const NOTES_MASTER_PART: &str = "/ppt/notesMasters/notesMaster1.xml";
const VIEW_PROPERTIES_PART: &str = "/ppt/viewProps.xml";
const PRESENTATION_PROPERTIES_PART: &str = "/ppt/presProps.xml";
const TABLE_STYLES_PART: &str = "/ppt/tableStyles.xml";
const CORE_PROPERTIES_PART: &str = "/docProps/core.xml";
const EXTENDED_PROPERTIES_PART: &str = "/docProps/app.xml";

const FIRST_SLIDE_ID: u32 = 256;
const FIRST_SLIDE_RELATIONSHIP_ID: u32 = 4;
// `soapberry-zip`'s default streaming transport is ZIP32 with 65,534
// members.  This tranche emits 37 fixed members and two members per slide.
const MAX_PHYSICAL_ZIP_ENTRIES: usize = 65_534;
const FIXED_PHYSICAL_ENTRY_COUNT: usize = 15 + resources::SLIDE_LAYOUTS.len() * 2;
const MAX_STREAMING_SLIDES: usize = (MAX_PHYSICAL_ZIP_ENTRIES - FIXED_PHYSICAL_ENTRY_COUNT) / 2;
const PHYSICAL_MAX_ENTRY_BYTES: u64 = 512 * 1024 * 1024;
const PHYSICAL_MAX_TOTAL_BYTES: u64 = 2 * 1024 * 1024 * 1024;
const PHYSICAL_MAX_OUTPUT_BYTES: u64 = 512 * 1024 * 1024;
const ZIP_LOCAL_HEADER_BYTES: u64 = 30;
const ZIP_CENTRAL_DIRECTORY_HEADER_BYTES: u64 = 46;
const ZIP_END_OF_CENTRAL_DIRECTORY_BYTES: u64 = 22;
const STANDARD_SLIDE_WIDTH: i64 = 9_144_000;
const STANDARD_SLIDE_HEIGHT: i64 = 6_858_000;
const WIDESCREEN_SLIDE_HEIGHT: i64 = 5_143_500;
const NOTES_WIDTH: i64 = 6_858_000;
const NOTES_HEIGHT: i64 = 9_144_000;

const SLIDE_PREFIX_START: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Slide "#;
const SLIDE_PREFIX_MIDDLE: &str = "\"><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>";
const SLIDE_SUFFIX: &str =
    r#"</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"#;

const SHAPE_START: &str = "<p:sp>";
const SHAPE_PREFIX: &str = "<p:nvSpPr><p:cNvPr id=\"";
const SHAPE_AFTER_ID: &str =
    "\" name=\"TextBox\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"";
const SHAPE_AFTER_X: &str = "\" y=\"";
const SHAPE_AFTER_Y: &str = "\"/><a:ext cx=\"";
const SHAPE_AFTER_WIDTH: &str = "\" cy=\"";
const SHAPE_AFTER_HEIGHT: &str = "\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"/><a:t xml:space=\"preserve\">";
const SHAPE_SUFFIX: &str = r#"</a:t></a:r></a:p></p:txBody></p:sp>"#;

/// A checked set of finite limits for a streaming presentation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StreamingPresentationLimits {
    /// Maximum number of slides declared for one deck. The physical ZIP32
    /// transport imposes a stricter hard cap of 32,748 slides for this
    /// topology, even when this policy field is larger.
    pub max_slides: usize,
    /// Maximum number of caller-supplied text boxes on one slide.
    pub max_text_boxes_per_slide: usize,
    /// Maximum UTF-8 byte length of one title or text box.
    pub max_text_bytes_per_box: usize,
    /// Maximum UTF-8 bytes across all titles and text boxes.
    pub max_total_text_bytes: usize,
    /// Maximum uncompressed XML bytes in one slide member.
    pub max_slide_xml_bytes: usize,
    /// Maximum bytes accepted by the physical output sink, including ZIP
    /// headers, compressed payloads, central-directory records, and EOCD.
    /// This must fit the underlying ZIP transport's 512 MiB ceiling and the
    /// fixed topology's structural metadata lower bound. A positive budget
    /// above that lower bound is a runtime ceiling, not a proof that all
    /// static resources and authored XML will fit; such output can fail with
    /// typed incomplete progress after streaming has begun.
    pub max_output_bytes: u64,
}

impl Default for StreamingPresentationLimits {
    fn default() -> Self {
        Self {
            max_slides: MAX_STREAMING_SLIDES,
            max_text_boxes_per_slide: 100_000,
            max_text_bytes_per_box: 16 * 1024 * 1024,
            max_total_text_bytes: 256 * 1024 * 1024,
            max_slide_xml_bytes: 64 * 1024 * 1024,
            max_output_bytes: 512 * 1024 * 1024,
        }
    }
}

/// Deterministic slide dimensions accepted by the bounded writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StreamingPresentationOptions {
    slide_width: i64,
    slide_height: i64,
}

impl Default for StreamingPresentationOptions {
    fn default() -> Self {
        Self::standard()
    }
}

impl StreamingPresentationOptions {
    /// Return standard 4:3 dimensions.
    #[must_use]
    pub const fn standard() -> Self {
        Self {
            slide_width: STANDARD_SLIDE_WIDTH,
            slide_height: STANDARD_SLIDE_HEIGHT,
        }
    }

    /// Return standard 16:9 dimensions.
    #[must_use]
    pub const fn widescreen() -> Self {
        Self {
            slide_width: STANDARD_SLIDE_WIDTH,
            slide_height: WIDESCREEN_SLIDE_HEIGHT,
        }
    }

    /// Construct a supported dimension pair in EMUs.
    ///
    /// Only the built-in 4:3 and 16:9 dimensions are accepted so the
    /// generated `p:sldSz@type` remains deterministic.
    ///
    /// # Errors
    ///
    /// Returns an error for unsupported, zero, or negative dimensions.
    pub fn new(slide_width: i64, slide_height: i64) -> Result<Self> {
        let supported = slide_width == STANDARD_SLIDE_WIDTH
            && (slide_height == STANDARD_SLIDE_HEIGHT || slide_height == WIDESCREEN_SLIDE_HEIGHT);
        if !supported {
            return Err(Error::Invalid(
                "streaming PPTX supports only standard 4:3 or 16:9 dimensions".into(),
            ));
        }
        Ok(Self {
            slide_width,
            slide_height,
        })
    }

    /// Return the slide width in EMUs.
    #[must_use]
    pub const fn slide_width(self) -> i64 {
        self.slide_width
    }

    /// Return the slide height in EMUs.
    #[must_use]
    pub const fn slide_height(self) -> i64 {
        self.slide_height
    }

    fn size_type(self) -> &'static str {
        if self.slide_height == STANDARD_SLIDE_HEIGHT {
            "screen4x3"
        } else {
            "screen16x9"
        }
    }
}

/// One plain text box accepted by [`StreamingSlideWriter`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextBoxSpec<'text> {
    /// Unformatted UTF-8 text.
    pub text: &'text str,
    /// X position in EMUs.
    pub x: i64,
    /// Y position in EMUs.
    pub y: i64,
    /// Positive width in EMUs.
    pub width: i64,
    /// Positive height in EMUs.
    pub height: i64,
}

impl<'text> TextBoxSpec<'text> {
    /// Construct one unformatted text-box specification.
    #[must_use]
    pub const fn new(text: &'text str, x: i64, y: i64, width: i64, height: i64) -> Self {
        Self {
            text,
            x,
            y,
            width,
            height,
        }
    }
}

/// A forward-only PPTX package writer over a caller-owned sequential sink.
pub struct StreamingPresentationWriter<W: Write> {
    archive: PhysPkgWriter<BudgetedSink<W>>,
    slide_count: usize,
    next_slide: usize,
    options: StreamingPresentationOptions,
    limits: StreamingPresentationLimits,
    total_text_bytes: usize,
}

/// A single active slide in a [`StreamingPresentationWriter`].
pub struct StreamingSlideWriter<W: Write> {
    part: PartWriter<BudgetedSink<W>>,
    slide_count: usize,
    slide_index: usize,
    options: StreamingPresentationOptions,
    limits: StreamingPresentationLimits,
    total_text_bytes: usize,
    text_box_count: usize,
    slide_xml_bytes: usize,
    poisoned: bool,
}

impl<W: Write> StreamingPresentationWriter<W> {
    /// Create a 4:3 streaming deck with default finite limits.
    ///
    /// The slide count must be at least one. Package metadata and the
    /// presentation manifest are emitted before this method returns; invalid
    /// arguments are rejected before the sink is touched.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid limits, dimensions, or a sink failure.
    pub fn new(writer: W, slide_count: usize) -> Result<Self> {
        Self::with_options(
            writer,
            slide_count,
            StreamingPresentationOptions::default(),
            StreamingPresentationLimits::default(),
        )
    }

    /// Create a streaming deck with explicit dimensions and finite limits.
    ///
    /// This is a fresh-authoring writer. It does not preserve or mutate an
    /// existing package and does not expose raw OPC part names to callers.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid limits, dimensions, or a sink failure.
    pub fn with_options(
        writer: W,
        slide_count: usize,
        options: StreamingPresentationOptions,
        limits: StreamingPresentationLimits,
    ) -> Result<Self> {
        validate_configuration(slide_count, options, limits)?;
        let sink = BudgetedSink::new(writer, limits.max_output_bytes);
        let mut archive = PhysPkgWriter::with_writer(sink);
        archive = write_content_types(archive, slide_count)?;
        archive = write_package_relationships(archive)?;
        archive = write_static_parts(archive)?;
        archive = write_presentation(archive, slide_count, options)?;
        archive = write_presentation_relationships(archive, slide_count)?;
        Ok(Self {
            archive,
            slide_count,
            next_slide: 0,
            options,
            limits,
            total_text_bytes: 0,
        })
    }

    /// Start the next slide in declaration order.
    ///
    /// `title`, when present, is emitted as the deterministic title shape
    /// with ID 2. It is validated before the slide ZIP local header is
    /// published. Calling this after all declared slides is refused.
    ///
    /// # Errors
    ///
    /// Returns an error for an extra slide, an invalid title, a limit
    /// violation, or a sink failure. Once output has started, sink failures
    /// are reported as [`litchi_opc::OpcError::IncompleteOutput`].
    pub fn start_slide(self, title: Option<&str>) -> Result<StreamingSlideWriter<W>> {
        if self.next_slide >= self.slide_count {
            return Err(Error::Invalid(
                "streaming PPTX received more slides than declared".into(),
            ));
        }
        if let Some(title) = title {
            validate_text(title, self.limits.max_text_bytes_per_box)?;
            checked_add_text(
                self.total_text_bytes,
                title.len(),
                self.limits.max_total_text_bytes,
            )?;
        }
        let slide_id = FIRST_SLIDE_ID
            .checked_add(u32::try_from(self.next_slide).map_err(|_| Error::Limit {
                resource: "streaming PPTX slide ID",
                limit: u32::MAX as usize,
            })?)
            .ok_or(Error::Limit {
                resource: "streaming PPTX slide ID",
                limit: u32::MAX as usize,
            })?;
        let prefix_bytes = slide_prefix_len(slide_id);
        let title_bytes = title.map_or(0, |value| {
            shape_xml_len(2, value, 914_400, 457_200, 7_315_200, 914_400)
        });
        let minimum = prefix_bytes
            .checked_add(title_bytes)
            .and_then(|bytes| bytes.checked_add(SLIDE_SUFFIX.len()))
            .ok_or_else(|| Error::Invalid("streaming PPTX slide XML length overflow".into()))?;
        if minimum > self.limits.max_slide_xml_bytes {
            return Err(Error::Limit {
                resource: "streaming PPTX slide XML bytes",
                limit: self.limits.max_slide_xml_bytes,
            });
        }
        let uri = PackURI::new(format!("/ppt/slides/slide{}.xml", self.next_slide + 1))
            .map_err(Error::Uri)?;
        let mut part = self.archive.start_part(&uri).map_err(Error::Opc)?;
        write_part_bytes(&mut part, SLIDE_PREFIX_START.as_bytes())?;
        write_unsigned_part(&mut part, u64::from(slide_id))?;
        write_part_bytes(&mut part, SLIDE_PREFIX_MIDDLE.as_bytes())?;
        let mut slide = StreamingSlideWriter {
            part,
            slide_count: self.slide_count,
            slide_index: self.next_slide,
            options: self.options,
            limits: self.limits,
            total_text_bytes: self.total_text_bytes,
            text_box_count: 0,
            slide_xml_bytes: prefix_bytes,
            poisoned: false,
        };
        if let Some(title) = title {
            slide.write_text_shape(2, title, 914_400, 457_200, 7_315_200, 914_400)?;
            slide.total_text_bytes = checked_add_text(
                slide.total_text_bytes,
                title.len(),
                slide.limits.max_total_text_bytes,
            )?;
        }
        Ok(slide)
    }

    /// Number of slides that have been finalized.
    #[must_use]
    pub const fn completed_slides(&self) -> usize {
        self.next_slide
    }

    /// Return the finite policy retained by this writer.
    #[must_use]
    pub const fn limits(&self) -> StreamingPresentationLimits {
        self.limits
    }

    /// Total UTF-8 text bytes accepted by finalized slides.
    #[must_use]
    pub const fn total_text_bytes(&self) -> usize {
        self.total_text_bytes
    }

    /// Number of bytes accepted by the output sink so far.
    #[must_use]
    pub fn output_bytes(&self) -> u64 {
        self.archive.output_bytes()
    }

    /// Finish the archive after all declared slides have been finalized.
    ///
    /// # Errors
    ///
    /// Returns an error when slides are missing or ZIP finalization fails.
    pub fn finish(self) -> Result<W> {
        if self.next_slide != self.slide_count {
            return Err(Error::Invalid(
                "streaming PPTX finished before all declared slides were emitted".into(),
            ));
        }
        let sink = self.archive.finish_into_inner().map_err(Error::Opc)?;
        Ok(sink.into_inner())
    }
}

impl<W: Write> StreamingSlideWriter<W> {
    /// Emit one plain text box and release no deck-sized state.
    ///
    /// All text and geometry validation occurs before the first byte of this
    /// box is sent to the sink. A later sink failure leaves the physical
    /// stream incomplete and is reported with OPC progress.
    ///
    /// # Errors
    ///
    /// Returns an error for hostile XML text, invalid geometry, or a finite
    /// limit violation.
    pub fn write_text_box(&mut self, spec: TextBoxSpec<'_>) -> Result<()> {
        self.ensure_usable()?;
        validate_text(spec.text, self.limits.max_text_bytes_per_box)?;
        validate_geometry(spec, self.options)?;
        if self.text_box_count >= self.limits.max_text_boxes_per_slide {
            return Err(Error::Limit {
                resource: "streaming PPTX text boxes per slide",
                limit: self.limits.max_text_boxes_per_slide,
            });
        }
        let total_text = checked_add_text(
            self.total_text_bytes,
            spec.text.len(),
            self.limits.max_total_text_bytes,
        )?;
        let shape_id = 3u32
            .checked_add(
                u32::try_from(self.text_box_count).map_err(|_| Error::Limit {
                    resource: "streaming PPTX shape ID",
                    limit: u32::MAX as usize,
                })?,
            )
            .ok_or(Error::Limit {
                resource: "streaming PPTX shape ID",
                limit: u32::MAX as usize,
            })?;
        let shape_bytes =
            shape_xml_len(shape_id, spec.text, spec.x, spec.y, spec.width, spec.height);
        let next_xml = self
            .slide_xml_bytes
            .checked_add(shape_bytes)
            .ok_or_else(|| Error::Invalid("streaming PPTX slide XML length overflow".into()))?;
        let next_xml_with_suffix = next_xml
            .checked_add(SLIDE_SUFFIX.len())
            .ok_or_else(|| Error::Invalid("streaming PPTX slide XML length overflow".into()))?;
        if next_xml_with_suffix > self.limits.max_slide_xml_bytes {
            return Err(Error::Limit {
                resource: "streaming PPTX slide XML bytes",
                limit: self.limits.max_slide_xml_bytes,
            });
        }
        self.write_text_shape(shape_id, spec.text, spec.x, spec.y, spec.width, spec.height)?;
        self.text_box_count = self.text_box_count.saturating_add(1);
        self.total_text_bytes = total_text;
        Ok(())
    }

    /// Number of text boxes emitted on this slide.
    #[must_use]
    pub const fn text_box_count(&self) -> usize {
        self.text_box_count
    }

    /// Number of uncompressed XML bytes committed to this slide so far.
    #[must_use]
    pub const fn slide_xml_bytes(&self) -> usize {
        self.slide_xml_bytes
    }

    /// Return the finite policy retained by this slide.
    #[must_use]
    pub const fn limits(&self) -> StreamingPresentationLimits {
        self.limits
    }

    /// Number of bytes accepted by the output sink so far.
    #[must_use]
    pub fn output_bytes(&self) -> u64 {
        self.part.output_bytes()
    }

    /// Finalize this slide and recover its parent writer.
    ///
    /// # Errors
    ///
    /// Returns an error if the slide or its relationship part cannot be
    /// finalized.
    pub fn finish(mut self) -> Result<StreamingPresentationWriter<W>> {
        self.ensure_usable()?;
        let next_xml = self
            .slide_xml_bytes
            .checked_add(SLIDE_SUFFIX.len())
            .ok_or_else(|| Error::Invalid("streaming PPTX slide XML length overflow".into()))?;
        if next_xml > self.limits.max_slide_xml_bytes {
            return Err(Error::Limit {
                resource: "streaming PPTX slide XML bytes",
                limit: self.limits.max_slide_xml_bytes,
            });
        }
        self.write_fragment(SLIDE_SUFFIX.as_bytes())?;
        let archive = self.part.finish().map_err(Error::Opc)?;
        let uri = PackURI::new(format!(
            "/ppt/slides/_rels/slide{}.xml.rels",
            self.slide_index + 1
        ))
        .map_err(Error::Uri)?;
        let mut rels = archive.start_part(&uri).map_err(Error::Opc)?;
        write_relationships_start(&mut rels)?;
        write_relationship(
            &mut rels,
            "rId1",
            rt::SLIDE_LAYOUT,
            "../slideLayouts/slideLayout1.xml",
        )?;
        write_relationships_end(&mut rels)?;
        let archive = rels.finish().map_err(Error::Opc)?;
        Ok(StreamingPresentationWriter {
            archive,
            slide_count: self.slide_count,
            next_slide: self.slide_index + 1,
            options: self.options,
            limits: self.limits,
            total_text_bytes: self.total_text_bytes,
        })
    }

    fn write_text_shape(
        &mut self,
        shape_id: u32,
        text: &str,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> Result<()> {
        self.ensure_usable()?;
        let escaped = escaped_len(text)?;
        let shape_bytes = shape_xml_len(shape_id, text, x, y, width, height);
        let next_xml = self
            .slide_xml_bytes
            .checked_add(shape_bytes)
            .ok_or_else(|| Error::Invalid("streaming PPTX slide XML length overflow".into()))?;
        let next_xml_with_suffix = next_xml
            .checked_add(SLIDE_SUFFIX.len())
            .ok_or_else(|| Error::Invalid("streaming PPTX slide XML length overflow".into()))?;
        if next_xml_with_suffix > self.limits.max_slide_xml_bytes {
            return Err(Error::Limit {
                resource: "streaming PPTX slide XML bytes",
                limit: self.limits.max_slide_xml_bytes,
            });
        }
        self.write_fragment(SHAPE_START.as_bytes())?;
        self.write_fragment(SHAPE_PREFIX.as_bytes())?;
        self.write_unsigned(u64::from(shape_id))?;
        self.write_fragment(SHAPE_AFTER_ID.as_bytes())?;
        self.write_unsigned(u64::try_from(x).unwrap_or_default())?;
        self.write_fragment(SHAPE_AFTER_X.as_bytes())?;
        self.write_unsigned(u64::try_from(y).unwrap_or_default())?;
        self.write_fragment(SHAPE_AFTER_Y.as_bytes())?;
        self.write_unsigned(u64::try_from(width).unwrap_or_default())?;
        self.write_fragment(SHAPE_AFTER_WIDTH.as_bytes())?;
        self.write_unsigned(u64::try_from(height).unwrap_or_default())?;
        self.write_fragment(SHAPE_AFTER_HEIGHT.as_bytes())?;
        self.write_escaped(text, escaped)?;
        self.write_fragment(SHAPE_SUFFIX.as_bytes())?;
        self.slide_xml_bytes = next_xml;
        Ok(())
    }

    fn ensure_usable(&self) -> Result<()> {
        if self.poisoned {
            return Err(Error::Invalid(
                "streaming PPTX slide writer is poisoned after sink failure".into(),
            ));
        }
        Ok(())
    }

    fn write_fragment(&mut self, bytes: &[u8]) -> Result<()> {
        self.ensure_usable()?;
        match write_part_bytes(&mut self.part, bytes) {
            Ok(()) => Ok(()),
            Err(error) => {
                self.poisoned = true;
                Err(error)
            },
        }
    }

    fn write_unsigned(&mut self, value: u64) -> Result<()> {
        let mut buffer = [0u8; 20];
        let bytes = decimal_bytes(value, &mut buffer);
        self.write_fragment(bytes)
    }

    fn write_escaped(&mut self, text: &str, expected_len: usize) -> Result<()> {
        self.ensure_usable()?;
        let mut start = 0;
        for (index, character) in text.char_indices() {
            let replacement = match character {
                '&' => Some("&amp;"),
                '<' => Some("&lt;"),
                '>' => Some("&gt;"),
                '"' => Some("&quot;"),
                '\'' => Some("&apos;"),
                _ => None,
            };
            if let Some(replacement) = replacement {
                if start < index {
                    self.write_fragment(&text.as_bytes()[start..index])?;
                }
                self.write_fragment(replacement.as_bytes())?;
                start = index + character.len_utf8();
            }
        }
        if start < text.len() {
            self.write_fragment(&text.as_bytes()[start..])?;
        }
        debug_assert_eq!(
            expected_len,
            escaped_len(text).unwrap_or(usize::MAX),
            "escaped text preflight must match direct writer"
        );
        Ok(())
    }
}

struct BudgetedSink<W: Write> {
    inner: W,
    max_output_bytes: u64,
    written: u64,
}

impl<W: Write> BudgetedSink<W> {
    fn new(inner: W, max_output_bytes: u64) -> Self {
        Self {
            inner,
            max_output_bytes,
            written: 0,
        }
    }

    fn into_inner(self) -> W {
        self.inner
    }
}

impl<W: Write> Write for BudgetedSink<W> {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        if buffer.is_empty() {
            return Ok(0);
        }
        let remaining = self.max_output_bytes.saturating_sub(self.written);
        if remaining == 0 {
            return Err(io::Error::new(
                io::ErrorKind::WriteZero,
                "streaming PPTX output limit exceeded",
            ));
        }
        let allowed = buffer
            .len()
            .min(usize::try_from(remaining).unwrap_or(usize::MAX));
        let written = self.inner.write(&buffer[..allowed])?;
        self.written = self.written.saturating_add(written as u64);
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

fn validate_configuration(
    slide_count: usize,
    options: StreamingPresentationOptions,
    limits: StreamingPresentationLimits,
) -> Result<()> {
    if slide_count == 0 {
        return Err(Error::Invalid(
            "streaming PPTX requires at least one slide".into(),
        ));
    }
    if slide_count > MAX_STREAMING_SLIDES {
        return Err(Error::Limit {
            resource: "streaming PPTX physical ZIP slides",
            limit: MAX_STREAMING_SLIDES,
        });
    }
    if slide_count > limits.max_slides {
        return Err(Error::Limit {
            resource: "streaming PPTX slides",
            limit: limits.max_slides,
        });
    }
    if limits.max_text_boxes_per_slide == 0
        || limits.max_text_bytes_per_box == 0
        || limits.max_total_text_bytes == 0
        || limits.max_slide_xml_bytes == 0
        || limits.max_output_bytes == 0
    {
        return Err(Error::Invalid(
            "streaming PPTX limits must all be positive".into(),
        ));
    }
    if limits.max_text_bytes_per_box as u64 > PHYSICAL_MAX_ENTRY_BYTES {
        return Err(Error::Limit {
            resource: "streaming PPTX text bytes",
            limit: usize::try_from(PHYSICAL_MAX_ENTRY_BYTES).unwrap_or(usize::MAX),
        });
    }
    if limits.max_slide_xml_bytes as u64 > PHYSICAL_MAX_ENTRY_BYTES {
        return Err(Error::Limit {
            resource: "streaming PPTX slide XML bytes",
            limit: usize::try_from(PHYSICAL_MAX_ENTRY_BYTES).unwrap_or(usize::MAX),
        });
    }
    if limits.max_total_text_bytes as u64 > PHYSICAL_MAX_TOTAL_BYTES {
        return Err(Error::Limit {
            resource: "streaming PPTX total text bytes",
            limit: usize::try_from(PHYSICAL_MAX_TOTAL_BYTES).unwrap_or(usize::MAX),
        });
    }
    if limits.max_output_bytes > PHYSICAL_MAX_OUTPUT_BYTES {
        return Err(Error::Limit {
            resource: "streaming PPTX output bytes",
            limit: usize::try_from(PHYSICAL_MAX_OUTPUT_BYTES).unwrap_or(usize::MAX),
        });
    }
    StreamingPresentationOptions::new(options.slide_width, options.slide_height)?;
    let maximum_slide_id = FIRST_SLIDE_ID
        .checked_add(u32::try_from(slide_count - 1).map_err(|_| Error::Limit {
            resource: "streaming PPTX slide ID",
            limit: u32::MAX as usize,
        })?)
        .ok_or(Error::Limit {
            resource: "streaming PPTX slide ID",
            limit: u32::MAX as usize,
        })?;
    let minimum_slide_xml = slide_prefix_len(maximum_slide_id)
        .checked_add(SLIDE_SUFFIX.len())
        .ok_or_else(|| Error::Invalid("streaming PPTX slide XML length overflow".into()))?;
    if minimum_slide_xml > limits.max_slide_xml_bytes {
        return Err(Error::Limit {
            resource: "streaming PPTX slide XML bytes",
            limit: limits.max_slide_xml_bytes,
        });
    }
    let structural_metadata_lower_bound = structural_metadata_lower_bound_bytes(slide_count)?;
    if limits.max_output_bytes < structural_metadata_lower_bound {
        return Err(Error::Limit {
            resource: "streaming PPTX structural ZIP metadata bytes",
            limit: usize::try_from(structural_metadata_lower_bound).unwrap_or(usize::MAX),
        });
    }
    Ok(())
}

fn validate_text(text: &str, maximum: usize) -> Result<()> {
    if text.len() > maximum {
        return Err(Error::Limit {
            resource: "streaming PPTX text bytes",
            limit: maximum,
        });
    }
    if !text.chars().all(is_xml_char) {
        return Err(Error::Invalid(
            "streaming PPTX text contains an invalid XML character".into(),
        ));
    }
    Ok(())
}

fn is_xml_char(character: char) -> bool {
    matches!(
        character as u32,
        0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    )
}

fn validate_geometry(spec: TextBoxSpec<'_>, options: StreamingPresentationOptions) -> Result<()> {
    if spec.x < 0
        || spec.y < 0
        || spec.width <= 0
        || spec.height <= 0
        || spec.x > options.slide_width
        || spec.y > options.slide_height
        || spec.width > options.slide_width - spec.x
        || spec.height > options.slide_height - spec.y
    {
        return Err(Error::Invalid(
            "streaming PPTX text-box geometry is outside the slide".into(),
        ));
    }
    Ok(())
}

fn checked_add_text(current: usize, additional: usize, maximum: usize) -> Result<usize> {
    let total = current
        .checked_add(additional)
        .ok_or_else(|| Error::Invalid("streaming PPTX text-byte count overflow".into()))?;
    if total > maximum {
        return Err(Error::Limit {
            resource: "streaming PPTX total text bytes",
            limit: maximum,
        });
    }
    Ok(total)
}

fn escaped_len(text: &str) -> Result<usize> {
    text.chars().try_fold(0usize, |length, character| {
        let encoded_length = match character {
            '&' => "&amp;".len(),
            '<' => "&lt;".len(),
            '>' => "&gt;".len(),
            '"' => "&quot;".len(),
            '\'' => "&apos;".len(),
            _ => character.len_utf8(),
        };
        length
            .checked_add(encoded_length)
            .ok_or_else(|| Error::Invalid("streaming PPTX escaped text length overflow".into()))
    })
}

fn slide_prefix_len(slide_id: u32) -> usize {
    SLIDE_PREFIX_START
        .len()
        .saturating_add(decimal_len(u64::from(slide_id)))
        .saturating_add(SLIDE_PREFIX_MIDDLE.len())
}

fn shape_xml_len(id: u32, text: &str, x: i64, y: i64, width: i64, height: i64) -> usize {
    SHAPE_START
        .len()
        .saturating_add(SHAPE_PREFIX.len())
        .saturating_add(decimal_len(u64::from(id)))
        .saturating_add(SHAPE_AFTER_ID.len())
        .saturating_add(decimal_len(u64::try_from(x).unwrap_or_default()))
        .saturating_add(SHAPE_AFTER_X.len())
        .saturating_add(decimal_len(u64::try_from(y).unwrap_or_default()))
        .saturating_add(SHAPE_AFTER_Y.len())
        .saturating_add(decimal_len(u64::try_from(width).unwrap_or_default()))
        .saturating_add(SHAPE_AFTER_WIDTH.len())
        .saturating_add(decimal_len(u64::try_from(height).unwrap_or_default()))
        .saturating_add(SHAPE_AFTER_HEIGHT.len())
        .saturating_add(escaped_len(text).unwrap_or(usize::MAX))
        .saturating_add(SHAPE_SUFFIX.len())
}

fn decimal_len(mut value: u64) -> usize {
    let mut length = 1;
    while value >= 10 {
        value /= 10;
        length += 1;
    }
    length
}

/// Return the mandatory ZIP framing and member-name metadata footprint.
///
/// This is intentionally only a structural lower bound: it omits compressed
/// and uncompressed payload bytes, data descriptors, extra fields, and
/// compressor finalization. Budgets above it remain runtime ceilings enforced
/// by [`BudgetedSink`] and can still produce typed incomplete output.
fn structural_metadata_lower_bound_bytes(slide_count: usize) -> Result<u64> {
    let mut total = ZIP_END_OF_CENTRAL_DIRECTORY_BYTES;
    for name in [
        "[Content_Types].xml",
        "_rels/.rels",
        "docProps/core.xml",
        "docProps/app.xml",
        "ppt/presProps.xml",
        "ppt/viewProps.xml",
        "ppt/tableStyles.xml",
        "ppt/slideMasters/slideMaster1.xml",
        "ppt/slideMasters/_rels/slideMaster1.xml.rels",
        "ppt/theme/theme1.xml",
        "ppt/theme/theme2.xml",
        "ppt/notesMasters/notesMaster1.xml",
        "ppt/notesMasters/_rels/notesMaster1.xml.rels",
        "ppt/presentation.xml",
        "ppt/_rels/presentation.xml.rels",
    ] {
        add_minimum_entry(&mut total, name.len())?;
    }
    for index in 1..=resources::SLIDE_LAYOUTS.len() {
        add_minimum_entry(
            &mut total,
            indexed_name_len("ppt/slideLayouts/slideLayout", index, ".xml")?,
        )?;
        add_minimum_entry(
            &mut total,
            indexed_name_len("ppt/slideLayouts/_rels/slideLayout", index, ".xml.rels")?,
        )?;
    }
    for index in 1..=slide_count {
        add_minimum_entry(
            &mut total,
            indexed_name_len("ppt/slides/slide", index, ".xml")?,
        )?;
        add_minimum_entry(
            &mut total,
            indexed_name_len("ppt/slides/_rels/slide", index, ".xml.rels")?,
        )?;
    }
    Ok(total)
}

fn indexed_name_len(prefix: &str, index: usize, suffix: &str) -> Result<usize> {
    prefix
        .len()
        .checked_add(decimal_len(u64::try_from(index).map_err(|_| {
            Error::Invalid("streaming PPTX physical member-name length overflow".into())
        })?))
        .and_then(|length| length.checked_add(suffix.len()))
        .ok_or_else(|| Error::Invalid("streaming PPTX physical member-name length overflow".into()))
}

fn add_minimum_entry(total: &mut u64, name_length: usize) -> Result<()> {
    let name_length = u64::try_from(name_length).map_err(|_| {
        Error::Invalid("streaming PPTX physical member-name length overflow".into())
    })?;
    let name_bytes = name_length
        .checked_mul(2)
        .ok_or_else(|| Error::Invalid("streaming PPTX physical output length overflow".into()))?;
    let entry_bytes = ZIP_LOCAL_HEADER_BYTES
        .checked_add(ZIP_CENTRAL_DIRECTORY_HEADER_BYTES)
        .and_then(|bytes| bytes.checked_add(name_bytes))
        .ok_or_else(|| Error::Invalid("streaming PPTX physical output length overflow".into()))?;
    *total = total
        .checked_add(entry_bytes)
        .ok_or_else(|| Error::Invalid("streaming PPTX physical output length overflow".into()))?;
    Ok(())
}

fn write_part_bytes<W: Write>(part: &mut PartWriter<BudgetedSink<W>>, bytes: &[u8]) -> Result<()> {
    match part.write_all(bytes) {
        Ok(()) => Ok(()),
        Err(error) => {
            let written = part.output_bytes();
            Err(Error::Opc(incomplete_io(written, error)))
        },
    }
}

fn write_unsigned_part<W: Write>(part: &mut PartWriter<BudgetedSink<W>>, value: u64) -> Result<()> {
    let mut buffer = [0u8; 20];
    let bytes = decimal_bytes(value, &mut buffer);
    write_part_bytes(part, bytes)
}

fn decimal_bytes(mut value: u64, buffer: &mut [u8; 20]) -> &[u8] {
    let mut index = buffer.len();
    loop {
        index -= 1;
        buffer[index] = b'0' + u8::try_from(value % 10).unwrap_or(0);
        value /= 10;
        if value == 0 {
            return &buffer[index..];
        }
    }
}

fn incomplete_io(written: u64, error: io::Error) -> litchi_opc::OpcError {
    let source = litchi_opc::OpcError::IoError(error);
    if written == 0 {
        source
    } else {
        litchi_opc::OpcError::IncompleteOutput {
            written,
            source: Box::new(source),
        }
    }
}

fn part_uri(path: &str) -> Result<PackURI> {
    PackURI::new(path).map_err(Error::Uri)
}

fn write_blob<W: Write>(
    archive: PhysPkgWriter<BudgetedSink<W>>,
    path: &str,
    bytes: &[u8],
) -> Result<PhysPkgWriter<BudgetedSink<W>>> {
    let uri = part_uri(path)?;
    let mut part = archive.start_part(&uri).map_err(Error::Opc)?;
    write_part_bytes(&mut part, bytes)?;
    part.finish().map_err(Error::Opc)
}

fn write_content_types<W: Write>(
    mut archive: PhysPkgWriter<BudgetedSink<W>>,
    slide_count: usize,
) -> Result<PhysPkgWriter<BudgetedSink<W>>> {
    let uri = part_uri("/[Content_Types].xml")?;
    let mut part = archive.start_part(&uri).map_err(Error::Opc)?;
    write_part_bytes(&mut part, XML_DECLARATION.as_bytes())?;
    write_part_bytes(
        &mut part,
        b"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>",
    )?;
    let overrides = [
        (EXTENDED_PROPERTIES_PART, ct::OFC_EXTENDED_PROPERTIES),
        (CORE_PROPERTIES_PART, ct::OPC_CORE_PROPERTIES),
        (NOTES_MASTER_PART, ct::PML_NOTES_MASTER),
        (PRESENTATION_PROPERTIES_PART, ct::PML_PRES_PROPS),
        (PRESENTATION_PART, ct::PML_PRESENTATION_MAIN),
    ];
    for (path, content_type) in overrides {
        write_override(&mut part, path, content_type)?;
    }
    for index in 1..=resources::SLIDE_LAYOUTS.len() {
        let path = format!("/ppt/slideLayouts/slideLayout{index}.xml");
        write_override(&mut part, &path, ct::PML_SLIDE_LAYOUT)?;
    }
    write_override(&mut part, MASTER_PART, ct::PML_SLIDE_MASTER)?;
    for index in 1..=slide_count {
        let path = format!("/ppt/slides/slide{index}.xml");
        write_override(&mut part, &path, ct::PML_SLIDE)?;
    }
    write_override(&mut part, TABLE_STYLES_PART, ct::PML_TABLE_STYLES)?;
    write_override(&mut part, THEME_PART, ct::OFC_THEME)?;
    write_override(&mut part, NOTES_THEME_PART, ct::OFC_THEME)?;
    write_override(&mut part, VIEW_PROPERTIES_PART, ct::PML_VIEW_PROPS)?;
    write_part_bytes(&mut part, b"</Types>")?;
    archive = part.finish().map_err(Error::Opc)?;
    Ok(archive)
}

fn write_override<W: Write>(
    part: &mut PartWriter<BudgetedSink<W>>,
    path: &str,
    content_type: &str,
) -> Result<()> {
    write_part_bytes(
        part,
        format!("<Override PartName=\"{path}\" ContentType=\"{content_type}\"/>").as_bytes(),
    )
}

fn write_package_relationships<W: Write>(
    archive: PhysPkgWriter<BudgetedSink<W>>,
) -> Result<PhysPkgWriter<BudgetedSink<W>>> {
    let uri = part_uri("/_rels/.rels")?;
    let mut part = archive.start_part(&uri).map_err(Error::Opc)?;
    write_relationships_start(&mut part)?;
    write_relationship(
        &mut part,
        "rId1",
        rt::OFFICE_DOCUMENT,
        "ppt/presentation.xml",
    )?;
    write_relationship(&mut part, "rId2", rt::CORE_PROPERTIES, "docProps/core.xml")?;
    write_relationship(
        &mut part,
        "rId3",
        rt::EXTENDED_PROPERTIES,
        "docProps/app.xml",
    )?;
    write_relationships_end(&mut part)?;
    part.finish().map_err(Error::Opc)
}

fn write_static_parts<W: Write>(
    mut archive: PhysPkgWriter<BudgetedSink<W>>,
) -> Result<PhysPkgWriter<BudgetedSink<W>>> {
    archive = write_blob(
        archive,
        CORE_PROPERTIES_PART,
        resources::CORE_PROPERTIES.as_bytes(),
    )?;
    archive = write_blob(
        archive,
        EXTENDED_PROPERTIES_PART,
        resources::EXTENDED_PROPERTIES.as_bytes(),
    )?;
    archive = write_blob(
        archive,
        PRESENTATION_PROPERTIES_PART,
        resources::PRESENTATION_PROPERTIES.as_bytes(),
    )?;
    archive = write_blob(
        archive,
        VIEW_PROPERTIES_PART,
        resources::VIEW_PROPERTIES.as_bytes(),
    )?;
    archive = write_blob(
        archive,
        TABLE_STYLES_PART,
        crate::table::style::default_xml().as_bytes(),
    )?;
    archive = write_blob(archive, MASTER_PART, resources::SLIDE_MASTER.as_bytes())?;
    archive = write_master_relationships(archive)?;
    for (index, xml) in resources::SLIDE_LAYOUTS.iter().enumerate() {
        let path = format!("/ppt/slideLayouts/slideLayout{}.xml", index + 1);
        archive = write_blob(archive, &path, xml.as_bytes())?;
        archive = write_layout_relationships(archive, index + 1)?;
    }
    archive = write_blob(archive, THEME_PART, resources::THEME.as_bytes())?;
    archive = write_blob(archive, NOTES_THEME_PART, resources::THEME.as_bytes())?;
    archive = write_blob(
        archive,
        NOTES_MASTER_PART,
        crate::notes::master_xml().as_bytes(),
    )?;
    write_notes_master_relationships(archive)
}

fn write_presentation<W: Write>(
    archive: PhysPkgWriter<BudgetedSink<W>>,
    slide_count: usize,
    options: StreamingPresentationOptions,
) -> Result<PhysPkgWriter<BudgetedSink<W>>> {
    let uri = part_uri(PRESENTATION_PART)?;
    let mut part = archive.start_part(&uri).map_err(Error::Opc)?;
    write_part_bytes(
        &mut part,
        format!(
            "{XML_DECLARATION}<p:presentation xmlns:a=\"{DRAWINGML_NAMESPACE}\" xmlns:r=\"{OFFICE_RELATIONSHIPS_NAMESPACE}\" xmlns:p=\"{PRESENTATIONML_NAMESPACE}\"><p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id=\"rIdNotesMaster\"/></p:notesMasterIdLst><p:sldIdLst>"
        )
        .as_bytes(),
    )?;
    for index in 0..slide_count {
        write_part_bytes(&mut part, b"<p:sldId id=\"")?;
        write_unsigned_part(
            &mut part,
            u64::from(FIRST_SLIDE_ID)
                .checked_add(index as u64)
                .ok_or_else(|| Error::Invalid("streaming PPTX slide ID overflow".into()))?,
        )?;
        write_part_bytes(&mut part, b"\" r:id=\"rId")?;
        write_unsigned_part(
            &mut part,
            u64::from(FIRST_SLIDE_RELATIONSHIP_ID)
                .checked_add(index as u64)
                .ok_or_else(|| Error::Invalid("streaming PPTX relationship ID overflow".into()))?,
        )?;
        write_part_bytes(&mut part, b"\"/>")?;
    }
    write_part_bytes(
        &mut part,
        format!(
            "</p:sldIdLst><p:sldSz cx=\"{}\" cy=\"{}\" type=\"{}\"/><p:notesSz cx=\"{NOTES_WIDTH}\" cy=\"{NOTES_HEIGHT}\"/><p:defaultTextStyle/></p:presentation>",
            options.slide_width,
            options.slide_height,
            options.size_type()
        )
        .as_bytes(),
    )?;
    part.finish().map_err(Error::Opc)
}

fn write_presentation_relationships<W: Write>(
    archive: PhysPkgWriter<BudgetedSink<W>>,
    slide_count: usize,
) -> Result<PhysPkgWriter<BudgetedSink<W>>> {
    let uri = part_uri("/ppt/_rels/presentation.xml.rels")?;
    let mut part = archive.start_part(&uri).map_err(Error::Opc)?;
    write_relationships_start(&mut part)?;
    write_relationship(
        &mut part,
        "rId1",
        rt::SLIDE_MASTER,
        "slideMasters/slideMaster1.xml",
    )?;
    write_relationship(&mut part, "rId2", rt::VIEW_PROPS, "viewProps.xml")?;
    write_relationship(&mut part, "rId3", rt::PRES_PROPS, "presProps.xml")?;
    for index in 0..slide_count {
        let relationship_id = format!("rId{}", FIRST_SLIDE_RELATIONSHIP_ID as usize + index);
        let target = format!("slides/slide{}.xml", index + 1);
        write_relationship(&mut part, &relationship_id, rt::SLIDE, &target)?;
    }
    write_relationship(
        &mut part,
        "rIdNotesMaster",
        rt::NOTES_MASTER,
        "notesMasters/notesMaster1.xml",
    )?;
    write_relationship(
        &mut part,
        "rIdTableStyles",
        rt::TABLE_STYLES,
        "tableStyles.xml",
    )?;
    write_relationships_end(&mut part)?;
    part.finish().map_err(Error::Opc)
}

fn write_master_relationships<W: Write>(
    archive: PhysPkgWriter<BudgetedSink<W>>,
) -> Result<PhysPkgWriter<BudgetedSink<W>>> {
    let uri = part_uri("/ppt/slideMasters/_rels/slideMaster1.xml.rels")?;
    let mut part = archive.start_part(&uri).map_err(Error::Opc)?;
    write_relationships_start(&mut part)?;
    for index in 1..=resources::SLIDE_LAYOUTS.len() {
        let target = format!("../slideLayouts/slideLayout{index}.xml");
        let relationship_id = format!("rId{index}");
        write_relationship(&mut part, &relationship_id, rt::SLIDE_LAYOUT, &target)?;
    }
    write_relationship(&mut part, "rId12", rt::THEME, "../theme/theme1.xml")?;
    write_relationships_end(&mut part)?;
    part.finish().map_err(Error::Opc)
}

fn write_layout_relationships<W: Write>(
    archive: PhysPkgWriter<BudgetedSink<W>>,
    index: usize,
) -> Result<PhysPkgWriter<BudgetedSink<W>>> {
    let uri = part_uri(&format!(
        "/ppt/slideLayouts/_rels/slideLayout{index}.xml.rels"
    ))?;
    let mut part = archive.start_part(&uri).map_err(Error::Opc)?;
    write_relationships_start(&mut part)?;
    write_relationship(
        &mut part,
        "rId1",
        rt::SLIDE_MASTER,
        "../slideMasters/slideMaster1.xml",
    )?;
    write_relationships_end(&mut part)?;
    part.finish().map_err(Error::Opc)
}

fn write_notes_master_relationships<W: Write>(
    archive: PhysPkgWriter<BudgetedSink<W>>,
) -> Result<PhysPkgWriter<BudgetedSink<W>>> {
    let uri = part_uri("/ppt/notesMasters/_rels/notesMaster1.xml.rels")?;
    let mut part = archive.start_part(&uri).map_err(Error::Opc)?;
    write_relationships_start(&mut part)?;
    write_relationship(&mut part, "rId1", rt::THEME, "../theme/theme2.xml")?;
    write_relationships_end(&mut part)?;
    part.finish().map_err(Error::Opc)
}

fn write_relationships_start<W: Write>(part: &mut PartWriter<BudgetedSink<W>>) -> Result<()> {
    write_part_bytes(
        part,
        format!("{XML_DECLARATION}<Relationships xmlns=\"{RELATIONSHIPS_NAMESPACE}\">").as_bytes(),
    )
}

fn write_relationships_end<W: Write>(part: &mut PartWriter<BudgetedSink<W>>) -> Result<()> {
    write_part_bytes(part, b"</Relationships>")
}

fn write_relationship<W: Write>(
    part: &mut PartWriter<BudgetedSink<W>>,
    id: &str,
    relationship_type: &str,
    target: &str,
) -> Result<()> {
    write_part_bytes(
        part,
        format!("<Relationship Id=\"{id}\" Type=\"{relationship_type}\" Target=\"{target}\"/>")
            .as_bytes(),
    )
}
