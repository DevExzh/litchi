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
use litchi_drawingml::diagram::edit as diagram_edit;
use litchi_drawingml::diagram::{
    DIAGRAM_COLORS_REL, DIAGRAM_DATA_REL, DIAGRAM_LAYOUT_REL, DIAGRAM_QUICK_STYLE_REL, Limits,
    MS_DIAGRAM_DRAWING_REL, STRICT_DIAGRAM_COLORS_REL, STRICT_DIAGRAM_DATA_REL,
    STRICT_DIAGRAM_LAYOUT_REL, STRICT_DIAGRAM_QUICK_STYLE_REL,
};
use litchi_opc::PackURI;
use litchi_opc::constants::content_type as ct;

use super::super::transaction::{GraphChange, ReplacedPart, SmartArtMarkup};
use super::super::{ElementPath, Error, Refusal, Result};
use super::Edit;
use crate::Package;
use crate::document::{TransferPart, TransferRelationship};

/// The default extent of a new SmartArt graphic: 16 cm × 9 cm.
const DEFAULT_WIDTH: i64 = 5_760_000;
const DEFAULT_HEIGHT: i64 = 3_240_000;
/// The largest accepted extent: 22 inches, Word's page-size limit.
const MAX_EXTENT: i64 = 22 * 914_400;

const fn not_found() -> Error {
    Error::Refused {
        position: 0,
        reason: Refusal::BlockNotFound,
    }
}

fn invalid() -> Error {
    Error::Refused {
        position: 0,
        reason: Refusal::InvalidPropertyValue,
    }
}

fn drawing_error(error: litchi_drawingml::Error) -> Error {
    match error {
        litchi_drawingml::Error::Invalid(_) | litchi_drawingml::Error::Xml(_) => invalid(),
        error => Error::Document(crate::Error::Drawing(error)),
    }
}

/// A SmartArt graphic to insert: its authored spec, an extent in EMUs, and
/// its frame name and description (alternative text).
///
/// ```
/// use litchi_docx::document::NewSmartArt;
/// use litchi_drawingml::diagram::authoring::{Layout, Spec};
///
/// let spec = Spec::parse_outline(
///     Layout::Process,
///     "Plan\nBuild\nShip",
///     &litchi_drawingml::diagram::Limits::default(),
/// )?;
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

/// The parts of one SmartArt graphic (record 0335), as
/// [`Edit::smartart_parts`] reads them and [`Edit::set_smartart`] writes
/// them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SmartArtParts {
    /// The data model (`dgm:dataModel`).
    pub data: Vec<u8>,
    /// The layout definition (`dgm:layoutDef`).
    pub layout: Vec<u8>,
    /// The quick style (`dgm:styleDef`).
    pub quick_style: Vec<u8>,
    /// The colors (`dgm:colorsDef`).
    pub colors: Vec<u8>,
    /// The cached drawing (`dsp:drawing`), if the graphic has one.
    pub drawing: Option<Vec<u8>>,
}

/// The parts of one graphic, by role, and the main relationships naming
/// them.
struct Located {
    data: (String, String),
    layout: (String, String),
    quick_style: (String, String),
    colors: (String, String),
    /// The drawing part and, when the main part relates it, that
    /// relationship.
    drawing: Option<(Option<String>, String)>,
}

impl Edit<Package> {
    /// The parts of the graphic whose data model the main document relates
    /// through `data_relationship_id` (`r:dm`), as the base package holds
    /// them; `None` when no graphic has that relationship.
    ///
    /// # Errors
    ///
    /// Returns a package-graph error for unreadable diagram parts, and
    /// [`Refusal::InvalidPropertyValue`] for a graphic without a layout,
    /// quick style or colors part.
    pub fn smartart_parts(&self, data_relationship_id: &str) -> Result<Option<SmartArtParts>> {
        let Some(located) = self.locate_smartart(data_relationship_id)? else {
            return Ok(None);
        };
        let read = |name: &str| -> Result<Vec<u8>> {
            Ok(self
                .part_payload(name)?
                .ok_or_else(|| crate::Error::PartNotFound(name.to_owned()))?
                .as_ref()
                .clone())
        };
        Ok(Some(SmartArtParts {
            data: read(&located.data.1)?,
            layout: read(&located.layout.1)?,
            quick_style: read(&located.quick_style.1)?,
            colors: read(&located.colors.1)?,
            drawing: match &located.drawing {
                Some((_, name)) => Some(read(name)?),
                None => None,
            },
        }))
    }

    /// Replace the parts of the graphic whose data model the main document
    /// relates through `data_relationship_id`.
    ///
    /// The data model must read as one. A `drawing` replaces the cached
    /// drawing, or creates one (`diagrams/drawingN.xml` and its main
    /// relationship) when the graphic had none; `None` removes the cached
    /// drawing and the main relationship naming it, so consumers lay the
    /// graphic out again from the data model. Either way the data model's
    /// drawing reference (`dsp:dataModelExt/@relId`) is rewritten to name
    /// exactly the drawing that remains. Returns whether anything changed;
    /// identical parts are an exact no-op.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::Package;
    /// use litchi_docx::document::{ElementPath, ElementStep, NewSmartArt};
    /// use litchi_drawingml::diagram::{Limits, authoring::{Layout, Spec}};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Steps: ");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// let spec = Spec::parse_outline(
    ///     Layout::Process,
    ///     "Plan\nBuild",
    ///     &Limits::default(),
    /// )?;
    /// let id = edit.insert_smartart_at(&paragraph, None, &NewSmartArt::new(spec))?;
    /// let package = edit.commit()?.into_package();
    /// let mut edit = package.edit()?;
    /// let mut parts = edit.smartart_parts(&id)?.expect("the graphic");
    /// let spec = Spec::parse_outline(
    ///     Layout::Process,
    ///     "Plan\nBuild\nShip",
    ///     &Limits::default(),
    /// )?;
    /// let drawing = parts.drawing.take().expect("a cached drawing");
    /// let limits = Limits::default();
    /// parts.data = spec.data_xml(Some("rId0"), &limits)?.into_bytes();
    /// parts.drawing = Some(spec.drawing_xml(5_760_000, 3_240_000, &limits)?.into_bytes());
    /// assert_ne!(parts.drawing.as_deref(), Some(drawing.as_slice()));
    /// assert!(edit.set_smartart(&id, parts)?);
    /// let package = edit.commit()?.into_package();
    /// assert_eq!(package.smart_arts()?[0].text(), "Plan\nBuild\nShip");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown graphic,
    /// [`Refusal::InvalidPropertyValue`] for a data model that does not
    /// read, and a package-graph error for unreadable parts.
    pub fn set_smartart(
        &mut self,
        data_relationship_id: &str,
        parts: SmartArtParts,
    ) -> Result<bool> {
        let located = self
            .locate_smartart(data_relationship_id)?
            .ok_or_else(not_found)?;
        diagram_edit::Outline::read(&parts.data, &Limits::default()).map_err(drawing_error)?;
        let SmartArtParts {
            data,
            layout,
            quick_style,
            colors,
            drawing,
        } = parts;
        let mut replaced = Vec::new();
        let mut replace = |name: &str, bytes: Vec<u8>, this: &Self| -> Result<()> {
            let before = this
                .part_payload(name)?
                .ok_or_else(|| crate::Error::PartNotFound(name.to_owned()))?;
            if before.as_slice() != bytes.as_slice() {
                replaced.push(ReplacedPart {
                    name: name.to_owned(),
                    before,
                    after: Arc::new(bytes),
                });
            }
            Ok(())
        };
        // Settle the drawing first: its relationship goes into the data.
        let mut removal = None;
        let mut addition = None;
        let relationship = match (drawing, &located.drawing) {
            (Some(bytes), Some((relationship, name))) => {
                replace(name, bytes, self)?;
                relationship.clone()
            },
            (Some(bytes), None) => {
                let base = self.main_folder()?;
                let used = self.part_names()?;
                let target = (1_usize..)
                    .map(|number| format!("diagrams/drawing{number}.xml"))
                    .find(|target| !used.contains(&format!("{base}{target}")))
                    .ok_or_else(invalid)?;
                let name = format!("{base}{target}");
                PackURI::new(name.clone()).map_err(crate::Error::Invalid)?;
                let id = self.next_relationship_id()?;
                addition = Some(GraphChange {
                    relationships: Arc::new([TransferRelationship {
                        id: id.clone(),
                        relationship_type: MS_DIAGRAM_DRAWING_REL.to_owned(),
                        target,
                        external: false,
                    }]),
                    parts: Arc::new([TransferPart {
                        name,
                        content_type: ct::DML_DIAGRAM_DRAWING.to_owned(),
                        blob: Arc::new(bytes),
                        relationships: Arc::new([]),
                    }]),
                    replaced: Arc::new([]),
                });
                Some(id)
            },
            // A drawing the main part relates goes with its relationship;
            // one the data part relates stays, unreferenced by the data.
            (None, Some((Some(relationship), name))) => {
                removal =
                    Some(self.removal(vec![relationship.clone()], std::slice::from_ref(name))?);
                None
            },
            (None, Some((None, _))) => None,
            (None, None) => None,
        };
        let data =
            diagram_edit::set_drawing_reference(&data, relationship.as_deref(), &Limits::default())
                .map_err(drawing_error)?;
        replace(&located.data.1, data, self)?;
        replace(&located.layout.1, layout, self)?;
        replace(&located.quick_style.1, quick_style, self)?;
        replace(&located.colors.1, colors, self)?;
        if replaced.is_empty() && removal.is_none() && addition.is_none() {
            return Ok(false);
        }
        if !replaced.is_empty() {
            self.inner_mut()?.stage_graph_change(
                Arc::new(GraphChange {
                    relationships: Arc::new([]),
                    parts: Arc::new([]),
                    replaced: replaced.into(),
                }),
                true,
            )?;
        }
        if let Some(addition) = addition {
            self.inner_mut()?
                .stage_graph_change(Arc::new(addition), true)?;
        }
        if let Some(removal) = removal {
            self.inner_mut()?
                .stage_graph_change(Arc::new(removal), false)?;
        }
        Ok(true)
    }

    /// Remove the graphic whose data model the main document relates
    /// through `data_relationship_id`: its drawing in the main story (with
    /// its run when the run holds nothing else), its data, layout, quick
    /// style, colors and cached-drawing parts, and the main relationships
    /// naming them. A part another main relationship still names stays.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a graphic the main story does
    /// not draw, and a package-graph error for unreadable parts.
    pub fn remove_smartart(&mut self, data_relationship_id: &str) -> Result<()> {
        let located = self
            .locate_smartart(data_relationship_id)?
            .ok_or_else(not_found)?;
        if self
            .inner_mut()?
            .remove_smartart_drawings(data_relationship_id)?
            == 0
        {
            return Err(not_found());
        }
        let mut relationships = vec![
            located.data.0.clone(),
            located.layout.0.clone(),
            located.quick_style.0.clone(),
            located.colors.0.clone(),
        ];
        let mut names = vec![
            located.data.1.clone(),
            located.layout.1.clone(),
            located.quick_style.1.clone(),
            located.colors.1.clone(),
        ];
        if let Some((relationship, name)) = &located.drawing {
            relationships.extend(relationship.iter().cloned());
            names.push(name.clone());
        }
        relationships.dedup();
        let removal = self.removal(relationships, &names)?;
        self.inner_mut()?
            .stage_graph_change(Arc::new(removal), false)?;
        Ok(())
    }

    /// The parts and main relationships of one graphic.
    fn locate_smartart(&self, data_relationship_id: &str) -> Result<Option<Located>> {
        let diagrams = self.package.smart_arts()?;
        let Some(diagram) = diagrams
            .into_iter()
            .find(|diagram| diagram.data_relationship_id == data_relationship_id)
        else {
            return Ok(None);
        };
        let pair = |id: Option<String>, name: Option<String>| -> Result<(String, String)> {
            match (id, name) {
                (Some(id), Some(name)) => Ok((id, name)),
                _ => Err(invalid()),
            }
        };
        let drawing = match diagram.drawing_part_name {
            Some(name) => {
                let main = self
                    .package
                    .opc_package()
                    .main_document_part()
                    .map_err(crate::Error::from)?;
                let relationship = main
                    .rels()
                    .iter()
                    .filter(|relationship| {
                        relationship.reltype() == MS_DIAGRAM_DRAWING_REL
                            && !relationship.is_external()
                    })
                    .find(|relationship| {
                        relationship
                            .target_partname()
                            .is_ok_and(|target| target.as_str() == name)
                    })
                    .map(|relationship| relationship.r_id().to_owned());
                Some((relationship, name))
            },
            None => None,
        };
        Ok(Some(Located {
            data: (diagram.data_relationship_id, diagram.data_part_name),
            layout: pair(diagram.layout_relationship_id, diagram.layout_part_name)?,
            quick_style: pair(
                diagram.quick_style_relationship_id,
                diagram.quick_style_part_name,
            )?,
            colors: pair(diagram.colors_relationship_id, diagram.colors_part_name)?,
            drawing,
        }))
    }

    /// The graph change that, staged backward, removes the main
    /// relationships `relationships` and the parts `names` (a part another
    /// main relationship still names is kept), as the base package holds
    /// them.
    fn removal(&self, relationships: Vec<String>, names: &[String]) -> Result<GraphChange> {
        let opc = self.package.opc_package();
        let main = opc.main_document_part().map_err(crate::Error::from)?;
        let mut removed = Vec::with_capacity(relationships.len());
        for id in &relationships {
            let relationship = main.rels().get(id).ok_or_else(invalid)?;
            removed.push(TransferRelationship {
                id: id.clone(),
                relationship_type: relationship.reltype().to_owned(),
                target: relationship.target_ref().to_owned(),
                external: relationship.is_external(),
            });
        }
        let kept: std::collections::BTreeSet<String> = main
            .rels()
            .iter()
            .filter(|relationship| {
                !relationship.is_external()
                    && !relationships.iter().any(|id| id == relationship.r_id())
            })
            .filter_map(|relationship| {
                relationship
                    .target_partname()
                    .ok()
                    .map(|name| name.as_str().to_owned())
            })
            .collect();
        let mut parts = Vec::with_capacity(names.len());
        for name in names {
            if kept.contains(name) || parts.iter().any(|part: &TransferPart| &part.name == name) {
                continue;
            }
            let uri = PackURI::new(name.clone()).map_err(crate::Error::Invalid)?;
            let part = opc.get_part(&uri).map_err(crate::Error::from)?;
            parts.push(TransferPart {
                name: name.clone(),
                content_type: part.content_type().to_owned(),
                blob: part.blob_arc(),
                relationships: part
                    .rels()
                    .iter()
                    .map(|relationship| TransferRelationship {
                        id: relationship.r_id().to_owned(),
                        relationship_type: relationship.reltype().to_owned(),
                        target: relationship.target_ref().to_owned(),
                        external: relationship.is_external(),
                    })
                    .collect(),
            });
        }
        Ok(GraphChange {
            relationships: removed.into(),
            parts: parts.into(),
            replaced: Arc::new([]),
        })
    }

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
    /// let spec = Spec::parse_outline(
    ///     Layout::Process,
    ///     "Plan\nBuild\nShip",
    ///     &litchi_drawingml::diagram::Limits::default(),
    /// )?;
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
            .inner_mut()?
            .max_drawing_id()?
            .map_or(Some(1), |maximum| maximum.checked_add(1))
            .ok_or_else(invalid)?;
        let strict = self.is_strict()?;
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
        let limits = Limits::default();
        let blobs = [
            smartart
                .spec
                .data_xml(Some(&ids[4]), &limits)
                .map_err(|_| invalid())?
                .into_bytes(),
            smartart
                .spec
                .layout_xml(&limits)
                .map_err(|_| invalid())?
                .into_bytes(),
            quick_style_xml(&limits)
                .map_err(|_| invalid())?
                .into_bytes(),
            colors_xml(&limits).map_err(|_| invalid())?.into_bytes(),
            smartart
                .spec
                .drawing_xml(width, height, &limits)
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
        self.inner_mut()?.insert_smartart_markup_at(
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
        self.inner_mut()?.stage_graph_change(
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
