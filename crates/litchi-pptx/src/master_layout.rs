//! Semantic master/layout authoring vocabulary.
//!
//! Package graph mutation remains intentionally narrow in this slice: the
//! new-package writer publishes one validated default master and layout, while
//! these values provide the canonical vocabulary for the next graph mutation
//! step.

use litchi_opc::PackURI;

use crate::{Error, Result};

/// Minimum legal `p:sldMasterId` and `p:sldLayoutId` value.
pub const MIN_MASTER_OR_LAYOUT_ID: u32 = 2_147_483_648;

/// PresentationML slide-layout type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideLayoutKind {
    /// Title layout.
    Title,
    /// Text layout.
    Text,
    /// Two-column text layout.
    TwoColumnText,
    /// Table layout.
    Table,
    /// Text and chart layout.
    TextAndChart,
    /// Chart and text layout.
    ChartAndText,
    /// Diagram layout.
    Diagram,
    /// Chart layout.
    Chart,
    /// Title-only layout.
    TitleOnly,
    /// Blank layout.
    Blank,
    /// Object layout.
    Object,
    /// Custom layout token.
    Custom,
}

impl SlideLayoutKind {
    /// Return the PresentationML lexical token.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Title => "title",
            Self::Text => "tx",
            Self::TwoColumnText => "twoColTx",
            Self::Table => "tbl",
            Self::TextAndChart => "txAndChart",
            Self::ChartAndText => "chartAndTx",
            Self::Diagram => "dgm",
            Self::Chart => "chart",
            Self::TitleOnly => "titleOnly",
            Self::Blank => "blank",
            Self::Object => "obj",
            Self::Custom => "cust",
        }
    }
}

/// PresentationML placeholder type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PlaceholderKind {
    /// Title placeholder.
    Title,
    /// Body placeholder.
    Body,
    /// Centered title placeholder.
    CenteredTitle,
    /// Subtitle placeholder.
    Subtitle,
    /// Date/time placeholder.
    DateTime,
    /// Slide-number placeholder.
    SlideNumber,
    /// Footer placeholder.
    Footer,
    /// Header placeholder.
    Header,
    /// Object placeholder.
    Object,
    /// Chart placeholder.
    Chart,
    /// Table placeholder.
    Table,
    /// Picture placeholder.
    Picture,
}

impl PlaceholderKind {
    /// Return the PresentationML lexical token.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Title => "title",
            Self::Body => "body",
            Self::CenteredTitle => "ctrTitle",
            Self::Subtitle => "subTitle",
            Self::DateTime => "dt",
            Self::SlideNumber => "sldNum",
            Self::Footer => "ftr",
            Self::Header => "hdr",
            Self::Object => "obj",
            Self::Chart => "chart",
            Self::Table => "tbl",
            Self::Picture => "pic",
        }
    }
}

/// One bounded placeholder authored into a layout.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PlaceholderSpec {
    /// Placeholder kind.
    pub kind: PlaceholderKind,
    /// Optional placeholder index.
    pub index: Option<u32>,
    /// Optional producer-visible shape name.
    pub name: Option<String>,
    /// Optional inert default text.
    pub text: Option<String>,
}

impl PlaceholderSpec {
    /// Create a placeholder specification with no optional values.
    pub fn new(kind: PlaceholderKind) -> Self {
        Self {
            kind,
            index: None,
            name: None,
            text: None,
        }
    }

    /// Set the placeholder index.
    pub fn with_index(mut self, index: u32) -> Self {
        self.index = Some(index);
        self
    }

    /// Set the placeholder name.
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Set inert placeholder text.
    pub fn with_text(mut self, text: impl Into<String>) -> Self {
        self.text = Some(text.into());
        self
    }
}

/// Identity returned after a master graph mutation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuthoredSlideMaster {
    /// `p:sldMasterId@id` value.
    pub master_id: u32,
    /// Master part name.
    pub part_name: PackURI,
    /// Presentation relationship ID.
    pub relationship_id: String,
}

/// Identity returned after a layout graph mutation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuthoredSlideLayout {
    /// `p:sldLayoutId@id` value.
    pub layout_id: u32,
    /// Layout part name.
    pub part_name: PackURI,
    /// Master relationship ID.
    pub relationship_id: String,
}

/// Validate the bounded master/layout graph exposed by the semantic reader.
pub fn validate_master_layout_graph(package: &litchi_opc::OpcPackage) -> Result<()> {
    let presentation = crate::presentation::Presentation::new(
        crate::parts::PresentationPart::from_package(package)?,
        package,
    );
    for master in presentation.slide_masters()? {
        for layout in master.layouts()? {
            let _ = layout.master()?;
        }
    }
    Ok(())
}

/// Explicitly report that graph mutation is outside the current bounded slice.
pub fn add_slide_master(_package: &mut litchi_opc::OpcPackage) -> Result<AuthoredSlideMaster> {
    Err(Error::UnsafeEdit {
        operation: "add_slide_master",
        reason: "standalone master graph authoring is staged after the default new-package graph",
    })
}

/// Explicitly report that graph mutation is outside the current bounded slice.
pub fn add_slide_layout(
    _package: &mut litchi_opc::OpcPackage,
    _master_part_name: &PackURI,
    _kind: SlideLayoutKind,
    _name: &str,
    _placeholders: &[PlaceholderSpec],
) -> Result<AuthoredSlideLayout> {
    Err(Error::UnsafeEdit {
        operation: "add_slide_layout",
        reason: "standalone master graph authoring is staged after the default new-package graph",
    })
}

/// Explicitly report that graph mutation is outside the current bounded slice.
pub fn remove_slide_layout(
    _package: &mut litchi_opc::OpcPackage,
    _layout_part_name: &PackURI,
) -> Result<bool> {
    Err(Error::UnsafeEdit {
        operation: "remove_slide_layout",
        reason: "standalone master graph authoring is staged after the default new-package graph",
    })
}

/// Explicitly report that placeholder-shape mutation is outside this slice.
pub fn store_placeholder_shape(
    _package: &mut litchi_opc::OpcPackage,
    _layout_part_name: &PackURI,
    _spec: &PlaceholderSpec,
) -> Result<()> {
    Err(Error::UnsafeEdit {
        operation: "store_placeholder_shape",
        reason: "standalone placeholder authoring is staged after the default new-package graph",
    })
}
