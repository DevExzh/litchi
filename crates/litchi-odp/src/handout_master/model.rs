//! Semantic values for one ODF `style:handout-master`.

use super::codec;
use crate::model::page_layout::{Collection, Layout};
use litchi_core::Result;
use litchi_odf_common::style::master::{Child, ChildKind};

/// A typed handout master and its losslessly retained direct drawing children.
///
/// ODF handout masters do not have a `style:name` and the package schema
/// permits at most one of them.  `page_layout_name` is the required physical
/// page-layout reference.  `presentation_page_layout_name`, when present,
/// supplies the one presentation-layout layer used for handout placeholders.
/// It is resolved at most once; ODF does not define a recursive handout-parent
/// chain.
#[derive(Clone, Debug)]
pub struct Master {
    /// Required `style:page-layout-name` reference.
    pub page_layout_name: String,
    /// Optional `presentation:presentation-page-layout-name` reference.
    pub presentation_page_layout_name: Option<String>,
    /// Optional `draw:style-name` drawing-page style reference.
    pub drawing_style_name: Option<String>,
    /// Optional `presentation:use-header-name` declaration reference.
    pub header_name: Option<String>,
    /// Optional `presentation:use-footer-name` declaration reference.
    pub footer_name: Option<String>,
    /// Optional `presentation:use-date-time-name` declaration reference.
    pub date_time_name: Option<String>,
    /// Direct drawing children in source order.
    pub children: Vec<Child>,
    /// Exact source fragment, when this value originated in XML.
    pub(crate) source: String,
}

/// One-hop resolved handout-master view.
///
/// The handout master remains the authoritative local layer.  The optional
/// presentation layout is an inherited placeholder layer and is copied only
/// when a caller explicitly asks for resolution.
#[derive(Clone, Debug, PartialEq)]
pub struct Resolved {
    pub master: Master,
    pub presentation_layout: Option<Layout>,
}

impl PartialEq for Master {
    fn eq(&self, other: &Self) -> bool {
        self.page_layout_name == other.page_layout_name
            && self.presentation_page_layout_name == other.presentation_page_layout_name
            && self.drawing_style_name == other.drawing_style_name
            && self.header_name == other.header_name
            && self.footer_name == other.footer_name
            && self.date_time_name == other.date_time_name
            && self.children == other.children
    }
}

impl Eq for Master {}

impl Master {
    /// Create an empty schema-valid handout master.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(page_layout_name: impl Into<String>) -> Result<Self> {
        let value = Self {
            page_layout_name: page_layout_name.into(),
            presentation_page_layout_name: None,
            drawing_style_name: None,
            header_name: None,
            footer_name: None,
            date_time_name: None,
            children: Vec::new(),
            source: String::new(),
        };
        value.validate()?;
        Ok(value)
    }

    /// Parse one exact `style:handout-master` fragment.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn from_xml_fragment(xml: &str) -> Result<Self> {
        codec::parse_fragment(xml)
    }

    /// Return the exact source fragment or the next validated serialization.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn xml(&self) -> Result<String> {
        codec::write(self)
    }

    /// Serialize this master as one validated XML element.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn to_xml_fragment(&self) -> Result<String> {
        codec::write(self)
    }

    /// Validate semantic fields and every direct drawing child.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> Result<()> {
        super::validation::validate(self)
    }

    /// Append one direct drawing child after validating its complete XML.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn push_child(&mut self, child: Child) -> Result<()> {
        if child.kind != ChildKind::Shape {
            return Err(litchi_core::Error::InvalidFormat(
                "handout-master children must be drawing shapes".to_string(),
            ));
        }
        codec::validate_shape_fragment(&child.xml)?;
        self.children.push(child);
        Ok(())
    }

    /// Replace the direct drawing-child list atomically.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn set_children(&mut self, children: Vec<Child>) -> Result<()> {
        let candidate = Self {
            children,
            ..self.clone()
        };
        candidate.validate()?;
        self.children = candidate.children;
        Ok(())
    }

    /// Resolve the optional presentation page-layout layer exactly once.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn resolve(&self, layouts: &Collection) -> Result<Resolved> {
        self.validate()?;
        let presentation_layout = self
            .presentation_page_layout_name
            .as_deref()
            .map(|name| {
                layouts.get(name).cloned().ok_or_else(|| {
                    litchi_core::Error::InvalidFormat(format!(
                        "handout presentation page layout '{name}' does not exist"
                    ))
                })
            })
            .transpose()?;
        Ok(Resolved {
            master: self.clone(),
            presentation_layout,
        })
    }
}
