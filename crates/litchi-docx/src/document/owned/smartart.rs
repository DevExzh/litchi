//! Inline SmartArt insertion staged with a main-document edit (record
//! 0295, record 0284 follow-up 0284a).
//!
//! A new SmartArt graphic is an authored [`Spec`]
//! (`litchi_drawingml::diagram::authoring`, the PPTX vocabulary): its data
//! model, Litchi's layout definition, quick style and colors, and its
//! cached drawing, related from the main document as Word relates them
//! (`diagramData`, `diagramLayout`, `diagramQuickStyle`, `diagramColors`
//! and Microsoft's `diagramDrawing`), and a run holding the inline drawing
//! whose `dgm:relIds` names the first four — all one graph change and one
//! document operation, published by one commit.

use std::sync::Arc;

use litchi_core::Position;
use litchi_drawingml::diagram::authoring::{Spec, colors_xml, quick_style_xml};
use litchi_drawingml::diagram::{
    DIAGRAM_COLORS_REL, DIAGRAM_DATA_REL, DIAGRAM_LAYOUT_REL, DIAGRAM_QUICK_STYLE_REL,
    MS_DIAGRAM_DRAWING_REL, STRICT_DIAGRAM_COLORS_REL, STRICT_DIAGRAM_DATA_REL,
    STRICT_DIAGRAM_LAYOUT_REL, STRICT_DIAGRAM_QUICK_STYLE_REL,
};
use litchi_opc::PackURI;
use litchi_opc::constants::content_type as ct;

use super::super::transaction::{GraphChange, SmartArtMarkup};
use super::super::{ElementPath, Error, Refusal, Result};
use super::Edit;
use crate::Package;
use crate::document::{TransferPart, TransferRelationship};

/// The default extent of a new SmartArt graphic: 16 cm × 9 cm.
const DEFAULT_WIDTH: i64 = 5_760_000;
const DEFAULT_HEIGHT: i64 = 3_240_000;
/// The largest accepted extent: 22 inches, Word's page-size limit.
const MAX_EXTENT: i64 = 22 * 914_400;

fn invalid() -> Error {
    Error::Refused {
        position: 0,
        reason: Refusal::InvalidPropertyValue,
    }
}

/// A SmartArt graphic to insert: its authored spec, an extent in EMUs, and
/// its frame name and description (alternative text).
///
/// ```
/// use litchi_docx::document::NewSmartArt;
/// use litchi_drawingml::diagram::authoring::{Layout, Spec};
///
/// let spec = Spec::parse_outline(Layout::Process, "Plan\nBuild\nShip")?;
/// let smartart = NewSmartArt::new(spec).with_size(Some(5_400_000), None).with_name("Flow");
/// assert_eq!(smartart.size(), (5_400_000, 3_240_000));
/// # Ok::<(), litchi_drawingml::Error>(())
/// ```
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NewSmartArt {
    spec: Spec,
    width: Option<i64>,
    height: Option<i64>,
    name: Option<String>,
    description: String,
}

impl NewSmartArt {
    /// A graphic of `spec`, 16 cm × 9 cm.
    #[must_use]
    pub const fn new(spec: Spec) -> Self {
        Self {
            spec,
            width: None,
            height: None,
            name: None,
            description: String::new(),
        }
    }

    /// The extent in EMUs; `None` keeps the default of that side.
    #[must_use]
    pub const fn with_size(mut self, width: Option<i64>, height: Option<i64>) -> Self {
        self.width = width;
        self.height = height;
        self
    }

    /// The frame name (`wp:docPr/@name`); by default `SmartArt N`.
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// The alternative text (`wp:docPr/@descr`).
    #[must_use]
    pub fn with_description(mut self, description: impl Into<String>) -> Self {
        self.description = description.into();
        self
    }

    /// The extent in EMUs, defaults applied.
    #[must_use]
    pub fn size(&self) -> (i64, i64) {
        (
            self.width.unwrap_or(DEFAULT_WIDTH),
            self.height.unwrap_or(DEFAULT_HEIGHT),
        )
    }
}

impl Edit<Package> {
    /// Insert `smartart` as an inline drawing in a new run of the paragraph
    /// at `paragraph`, before its direct run `at` (at its end for `None`),
    /// and return the data-model relationship identifier (`r:dm`), which
    /// identifies the graphic in [`Package::smart_arts`].
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::Package;
    /// use litchi_docx::document::{ElementPath, ElementStep, NewSmartArt};
    /// use litchi_drawingml::diagram::authoring::{Layout, Spec};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Steps: ");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// let spec = Spec::parse_outline(Layout::Process, "Plan\nBuild\nShip")?;
    /// let id = edit.insert_smartart_at(&paragraph, None, &NewSmartArt::new(spec))?;
    /// let mut next = edit.commit()?.into_package();
    /// let reopened = Package::from_bytes(&next.to_bytes()?)?;
    /// let diagrams = reopened.smart_arts()?;
    /// assert_eq!(diagrams.len(), 1);
    /// assert_eq!(diagrams[0].data_relationship_id, id);
    /// assert_eq!(diagrams[0].text(), "Plan\nBuild\nShip");
    /// assert!(diagrams[0].drawing_part_name.is_some());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidPropertyValue`] for an extent outside 1 EMU
    /// to 22 inches, a name or description holding NUL, or a spec the
    /// authoring model refuses; a not-found refusal for a missing paragraph.
    pub fn insert_smartart_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        smartart: &NewSmartArt,
    ) -> Result<String> {
        let (width, height) = smartart.size();
        if !(1..=MAX_EXTENT).contains(&width) || !(1..=MAX_EXTENT).contains(&height) {
            return Err(invalid());
        }
        let description = &smartart.description;
        if smartart
            .name
            .as_deref()
            .is_some_and(|name| name.contains('\0'))
            || description.contains('\0')
        {
            return Err(invalid());
        }
        smartart.spec.validate().map_err(|_| invalid())?;
        let id = self
            .inner
            .max_drawing_id()?
            .map_or(Some(1), |maximum| maximum.checked_add(1))
            .ok_or_else(invalid)?;
        let strict = self.inner.is_strict();
        let base = self.main_folder()?;
        let used = self.part_names()?;
        let number = (1_usize..)
            .find(|number| {
                ["data", "layout", "quickStyle", "colors", "drawing"]
                    .iter()
                    .all(|stem| !used.contains(&format!("{base}diagrams/{stem}{number}.xml")))
            })
            .ok_or_else(invalid)?;
        let ids = self.next_relationship_ids(5)?;
        let relationship_types = if strict {
            [
                STRICT_DIAGRAM_DATA_REL,
                STRICT_DIAGRAM_LAYOUT_REL,
                STRICT_DIAGRAM_QUICK_STYLE_REL,
                STRICT_DIAGRAM_COLORS_REL,
                MS_DIAGRAM_DRAWING_REL,
            ]
        } else {
            [
                DIAGRAM_DATA_REL,
                DIAGRAM_LAYOUT_REL,
                DIAGRAM_QUICK_STYLE_REL,
                DIAGRAM_COLORS_REL,
                MS_DIAGRAM_DRAWING_REL,
            ]
        };
        let content_types = [
            ct::DML_DIAGRAM_DATA,
            ct::DML_DIAGRAM_LAYOUT,
            ct::DML_DIAGRAM_STYLE,
            ct::DML_DIAGRAM_COLORS,
            ct::DML_DIAGRAM_DRAWING,
        ];
        let blobs = [
            smartart
                .spec
                .data_xml(Some(&ids[4]))
                .map_err(|_| invalid())?
                .into_bytes(),
            smartart.spec.layout_xml().into_bytes(),
            quick_style_xml().into_bytes(),
            colors_xml().into_bytes(),
            smartart
                .spec
                .drawing_xml(width, height)
                .map_err(|_| invalid())?
                .into_bytes(),
        ];
        let stems = ["data", "layout", "quickStyle", "colors", "drawing"];
        let mut parts = Vec::with_capacity(5);
        let mut relationships = Vec::with_capacity(5);
        for (index, blob) in blobs.into_iter().enumerate() {
            let target = format!("diagrams/{}{number}.xml", stems[index]);
            let name = format!("{base}{target}");
            PackURI::new(name.clone()).map_err(crate::Error::Invalid)?;
            parts.push(TransferPart {
                name,
                content_type: content_types[index].to_owned(),
                blob: Arc::new(blob),
                relationships: Arc::new([]),
            });
            relationships.push(TransferRelationship {
                id: ids[index].clone(),
                relationship_type: relationship_types[index].to_owned(),
                target,
                external: false,
            });
        }
        let name = smartart
            .name
            .clone()
            .unwrap_or_else(|| format!("SmartArt {id}"));
        // The drawing first: a refused target stages nothing.
        self.inner.insert_smartart_markup_at(
            paragraph,
            at,
            [&ids[0], &ids[1], &ids[2], &ids[3]],
            &SmartArtMarkup {
                width,
                height,
                id,
                name: &name,
                description,
            },
        )?;
        self.inner.stage_graph_change(
            Arc::new(GraphChange {
                relationships: relationships.into(),
                parts: parts.into(),
                replaced: Arc::new([]),
            }),
            true,
        )?;
        Ok(ids[0].clone())
    }
}
