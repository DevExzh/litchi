//! Placeholder authoring and slide-layout retargeting for the unified API
//! (record 0346).
//!
//! A new placeholder follows OfficeCLI's `AddPlaceholder` binding rules: a
//! title adopts the layout's title-family type (a `title` on a "Title
//! Slide" layout is its `ctrTitle`), so it inherits the layout slot; other
//! types take the index of the layout slot of their type, else the
//! smallest index the slide does not use. A placeholder bound to a layout
//! slot keeps an empty `p:spPr`, inheriting its position, as `PowerPoint`
//! writes one; an unbound one gets OfficeCLI's default rectangle for its
//! type, so it renders.

use litchi_core::Position;
use litchi_opc::constants::relationship_type as rt;

use super::Edit;
use crate::master_layout::PlaceholderKind;
use crate::opened::{TextAssignment, TextProperties, TextTarget};
use crate::{Error, Result};

const fn is_title(kind: PlaceholderKind) -> bool {
    matches!(
        kind,
        PlaceholderKind::Title | PlaceholderKind::CenteredTitle
    )
}

/// The placeholder kind a `p:ph` names (an absent `type` is `obj`).
fn kind_of(placeholder: crate::shape::Placeholder<'_>) -> Option<PlaceholderKind> {
    placeholder
        .kind()
        .map_or(Some(PlaceholderKind::Object), PlaceholderKind::from_token)
}

/// OfficeCLI's slot rectangle for an unbound placeholder of `kind` (EMU,
/// 16:9 slide).
const fn fallback_bounds(kind: PlaceholderKind) -> (i64, i64, i64, i64) {
    match kind {
        PlaceholderKind::Title | PlaceholderKind::CenteredTitle => {
            (838_200, 365_125, 10_515_600, 1_325_563)
        },
        PlaceholderKind::Subtitle => (1_371_600, 3_886_200, 6_400_800, 1_752_600),
        PlaceholderKind::DateTime => (838_200, 6_356_350, 2_895_600, 365_125),
        PlaceholderKind::Footer => (3_884_613, 6_356_350, 4_351_338, 365_125),
        PlaceholderKind::SlideNumber => (8_506_463, 6_356_350, 2_847_338, 365_125),
        _ => (838_200, 1_825_625, 10_515_600, 4_351_338),
    }
}

impl Edit {
    /// Append a placeholder of `kind` to a slide, holding `text` (a line
    /// feed starts a paragraph). The type and index bind it to the slide
    /// layout's slot as the module documentation describes. Returns the
    /// new shape's positional scene index.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] for a second title placeholder on the
    /// slide, an error for an unknown slide, invalid text, malformed slide
    /// or layout XML, or an exhausted shape-ID space. A failure leaves the
    /// edit unchanged.
    pub fn add_placeholder<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        kind: PlaceholderKind,
        text: &str,
    ) -> Result<Position> {
        let selected = self.resolve_slide(slide.into())?;
        let before = self.working.clone();
        let result = self.append_placeholder(&selected, kind, None, text);
        if result.is_err() {
            self.working = before;
        }
        result
    }

    /// Give the slide's placeholder of `kind` (any title-family placeholder
    /// for a title; the one at `index` when given) the text `text`, adding
    /// one as [`Self::add_placeholder`] does, with `index`, when the slide
    /// has none: OfficeCLI's slide `title` and `text` shorthand (record
    /// 0346). Returns the placeholder's positional scene index.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::add_placeholder`] and of the text
    /// writer. A failure leaves the edit unchanged.
    pub fn put_placeholder_text<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        kind: PlaceholderKind,
        index: Option<u32>,
        text: &str,
    ) -> Result<Position> {
        let selected = self.resolve_slide(slide.into())?;
        let scene_xml = self.working.get_part(&selected.part_name)?.blob().to_vec();
        let scene = crate::shape::Scene::read(&scene_xml, self.semantic_limits().shape())?;
        let existing = scene.iter().position(|shape| {
            shape.placeholder().is_some_and(|placeholder| {
                let found = kind_of(placeholder);
                if is_title(kind) {
                    found.is_some_and(is_title)
                } else if let Some(index) = index {
                    placeholder.index() == index && !found.is_some_and(is_title)
                } else {
                    found == Some(kind)
                }
            })
        });
        drop(scene);
        let before = self.working.clone();
        let result = match existing {
            Some(position) => self
                .set_placeholder_body(&selected, Position::new(position), text)
                .map(|()| Position::new(position)),
            None => self.append_placeholder(&selected, kind, index, text),
        };
        if result.is_err() {
            self.working = before;
        }
        result
    }

    /// Point a slide at the layout named `layout_name` exactly (OfficeCLI's
    /// slide `layout` set, record 0346): the slide's layout relationship is
    /// retargeted and its shapes are kept as they are. Returns whether
    /// anything changed.
    ///
    /// # Errors
    ///
    /// Returns [`Error::SlideLayoutNameNotFound`] or
    /// [`Error::AmbiguousSlideLayoutName`] for the name, and an error for a
    /// slide without exactly one internal layout relationship or a signed
    /// package.
    pub fn set_slide_layout<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        layout_name: &str,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let presentation_name = self.working.main_document_part()?.partname().clone();
        let layout = {
            let main = self.working.get_part(&presentation_name)?;
            let presentation = crate::Presentation::new_with_semantic_limits(
                crate::parts::PresentationPart::from_part(main)?,
                &self.working,
                self.semantic_limits_ref(),
            );
            let mut layout = None;
            let mut matches = 0;
            for candidate in presentation.slide_layouts()? {
                if candidate.name()? == layout_name {
                    matches += 1;
                    layout = Some(candidate.part().part().partname().clone());
                }
            }
            if matches > 1 {
                return Err(Error::AmbiguousSlideLayoutName {
                    name: layout_name.to_owned(),
                    matches,
                });
            }
            layout.ok_or_else(|| Error::SlideLayoutNameNotFound(layout_name.to_owned()))?
        };
        let part = self.working.get_part(&selected.part_name)?;
        let edge = crate::opened::remove_plan::removal_layout_edge(part.rels())?;
        if edge.target_query().is_some() || edge.target_fragment().is_some() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "set_slide_layout",
                "the slide's layout target contains a query or fragment",
            ));
        }
        if edge.target_partname()? == layout {
            return Ok(false);
        }
        let relationship_id = edge.r_id().to_owned();
        let target = layout.relative_ref(selected.part_name.base_uri());
        self.working
            .get_part_mut(&selected.part_name)?
            .rels_mut()
            .retarget(&relationship_id, target)?;
        Ok(true)
    }

    /// Append a placeholder shape and give it `text`.
    fn append_placeholder(
        &mut self,
        slide: &crate::opened::model::Slide,
        kind: PlaceholderKind,
        index: Option<u32>,
        text: &str,
    ) -> Result<Position> {
        let owner = self.working.get_part(&slide.part_name)?;
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(owner.blob(), shape_limits)?;
        let used: Vec<(Option<PlaceholderKind>, u32)> = scene
            .roots()
            .filter_map(|shape| shape.placeholder())
            .map(|placeholder| (kind_of(placeholder), placeholder.index()))
            .collect();
        if is_title(kind) && used.iter().any(|(found, _)| found.is_some_and(is_title)) {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "add_placeholder",
                "the slide already has a title placeholder",
            ));
        }
        let shape_id = super::next_shape_id(&scene, "placeholder")?;
        let position = scene.len();
        let conformance = crate::media_parts::document_conformance_with_limits(
            owner.blob(),
            self.semantic_limits()
                .media_parts()
                .constrained_by_source(self.working.read_limits()),
        )?;
        drop(scene);
        let slots = self.layout_slots(slide)?;
        let (kind, index, bound) = if is_title(kind) {
            match slots.iter().find(|(found, _)| found.is_some_and(is_title)) {
                Some((Some(found), _)) => (*found, None, true),
                _ => (kind, None, false),
            }
        } else if let Some(index) = index {
            (
                kind,
                Some(index),
                slots.iter().any(|(_, slot)| *slot == index),
            )
        } else if kind == PlaceholderKind::Subtitle {
            (
                kind,
                None,
                slots.iter().any(|(found, _)| *found == Some(kind)),
            )
        } else if let Some((_, slot)) = slots
            .iter()
            .find(|(found, slot)| *found == Some(kind) && *slot != 0)
        {
            (kind, Some(*slot), true)
        } else {
            let next = (1..=u32::MAX)
                .find(|candidate| !used.iter().any(|(_, index)| index == candidate))
                .ok_or_else(|| Error::Invalid("placeholder indexes are exhausted".into()))?;
            (kind, Some(next), false)
        };
        let (pml, dml, _) = super::conformance_namespaces(conformance);
        let mut fragment = format!(
            "<p:sp xmlns:p=\"{pml}\" xmlns:a=\"{dml}\"><p:nvSpPr><p:cNvPr id=\"{shape_id}\" name=\"{}\"/><p:cNvSpPr/><p:nvPr><p:ph type=\"{}\"",
            crate::opened::new_slide::placeholder_name(kind, false, shape_id),
            kind.as_str(),
        );
        if let Some(index) = index {
            fragment.push_str(&format!(" idx=\"{index}\""));
        }
        fragment.push_str("/></p:nvPr></p:nvSpPr>");
        if bound {
            fragment.push_str("<p:spPr/>");
        } else {
            let (x, y, width, height) = fallback_bounds(kind);
            fragment.push_str(&format!(
                "<p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></a:xfrm></p:spPr>"
            ));
        }
        fragment.push_str("<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>");
        self.append_checked_shape(slide, shape_id, fragment.as_bytes())?;
        let position = Position::new(position);
        if !text.is_empty() {
            self.set_placeholder_body(slide, position, text)?;
        }
        Ok(position)
    }

    /// Replace the text of the shape at `position` through the text writer.
    fn set_placeholder_body(
        &mut self,
        slide: &crate::opened::model::Slide,
        position: Position,
        text: &str,
    ) -> Result<()> {
        let edits = [TextAssignment::new(
            TextTarget::Shape { shape: position },
            TextProperties::default().with_text(text),
        )];
        let slide_position = self
            .slides
            .iter()
            .position(|candidate| candidate.part_name == slide.part_name)
            .ok_or_else(|| Error::Invalid("the slide left the edit".into()))?;
        self.set_text_properties(slide_position, &edits)?;
        Ok(())
    }

    /// The placeholders of the slide's layout: their kind and index.
    fn layout_slots(
        &self,
        slide: &crate::opened::model::Slide,
    ) -> Result<Vec<(Option<PlaceholderKind>, u32)>> {
        let part = self.working.get_part(&slide.part_name)?;
        let Some(edge) = part.rels().iter().find(|relationship| {
            !relationship.is_external()
                && crate::parts::is_relationship_type(
                    relationship.reltype(),
                    rt::SLIDE_LAYOUT,
                    "slideLayout",
                )
        }) else {
            return Ok(Vec::new());
        };
        let layout = self.working.get_part(&edge.target_partname()?)?;
        let scene = crate::shape::Scene::read(layout.blob(), self.semantic_limits().shape())?;
        Ok(scene
            .roots()
            .filter_map(|shape| shape.placeholder())
            .map(|placeholder| (kind_of(placeholder), placeholder.index()))
            .collect())
    }
}
