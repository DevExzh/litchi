//! Office Math equation shapes: add, replace and remove.
//!
//! PowerPoint 2010 and later write a shape whose text holds an equation as
//! one `mc:AlternateContent` block ([MS-ODRAWXML] §2.3.3.41 for `a14:m`,
//! §3.5 for the example; ECMA-376 Part 3 for markup compatibility): the
//! `mc:Choice Requires="a14"` branch holds the shape with the equation's
//! OMML in `p:txBody/a:p/a14:m`, and the `mc:Fallback` branch the same shape
//! with locked text, filled with a picture of the equation, for readers
//! without Office Math. [`Edit::add_equation`] writes that form for a new
//! text box, [`Edit::set_equation`] replaces the OMML and the picture of an
//! existing one, and [`Edit::remove_equation`] removes the whole block.
//! Every other byte of the slide stays as it was.
//!
//! Only a shape whose text is exactly one equation is rewritten: the
//! fallback picture shows the whole text body, which the picture renderer
//! cannot draw when it mixes equations with other text.

use std::ops::Range;

use litchi_drawingml::math::{
    A14, Bindings, FALLBACK_FILL, FALLBACK_LOCKS, Form, HostRun, MCE, Prepared, carry_declarations,
    cnvpr_id, fallback_picture, fitted, horizontal, linear_run, max_raw_shape_id, render,
    require_declared_prefixes, root_bindings, tag_name, vertical,
};
use litchi_equation::Justification;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::{BlobPart, PackURI};
use quick_xml::Reader;
use quick_xml::events::Event;

use super::picture::attribute_value;
use super::{
    Edit, conformance_image_relationship, conformance_namespaces, ensure_shape_id, free_part_name,
    invalid, max_shape_reference_id, remove_unreferenced_dependencies, shape_relationship_ids,
};
use crate::shape::math::sites::{self, Block, DEFAULT_INSETS, Site};
use crate::shape::math::{EquationSource, EquationWrite};
use crate::{Error, Result};

/// The text size of a new text box when the caller names none, in points:
/// PowerPoint's default for text boxes.
const DEFAULT_FONT_SIZE_PT: f64 = 18.0;

fn refused(operation: &'static str, reason: &'static str) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        operation,
        reason,
    )
}

/// The OMML of `source` for a slide ([`litchi_drawingml::math::prepare`]),
/// refused when the caller's text exceeds the patch byte bound.
fn prepare(
    source: EquationSource<'_>,
    form: Form,
    run: &HostRun,
    dml: &str,
    bound: &Bindings,
    max_bytes: usize,
) -> Result<Prepared> {
    let input = source.text();
    if input.len() > max_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::PatchBytes,
            observed: input.len(),
            maximum: max_bytes,
        });
    }
    Ok(litchi_drawingml::math::prepare(
        source, form, run, dml, bound,
    )?)
}

/// What a placeholder inherits; see [`Edit::placeholder_inheritance`].
#[derive(Debug, Default)]
struct Inherited {
    extent: Option<(i64, i64)>,
    text_size_pt: Option<f64>,
}

/// The first-level text size, in points, of a master's `p:txStyles` style
/// (`titleStyle` or `bodyStyle`): its `a:lvl1pPr/a:defRPr@sz`.
fn master_text_size(xml: &[u8], style: &[u8]) -> Result<Option<f64>> {
    let mut reader = Reader::from_reader(xml);
    // How far into `p:txStyles`, the style and its first level we are.
    let mut path: Vec<Vec<u8>> = Vec::new();
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let local = element.local_name().as_ref().to_vec();
                let inside = path.len() >= 3
                    && path[path.len() - 3] == b"txStyles"
                    && path[path.len() - 2] == style
                    && path[path.len() - 1] == b"lvl1pPr";
                if inside && local == b"defRPr" {
                    return Ok(element
                        .checked_attributes()
                        .flatten()
                        .find(|attribute| attribute.key.as_ref() == b"sz")
                        .and_then(|attribute| {
                            std::str::from_utf8(&attribute.value)
                                .ok()?
                                .trim()
                                .parse::<u32>()
                                .ok()
                        })
                        .filter(|size| (100..=400_000).contains(size))
                        .map(|size| f64::from(size) / 100.0));
                }
                if matches!(event, Event::Start(_)) {
                    path.push(local);
                }
            },
            Event::End(_) => {
                path.pop();
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
    }
}

/// The `cx` and `cy` of an `a:ext` start tag.
fn extent(tag: &[u8]) -> Option<(i64, i64)> {
    let value = |name: &[u8]| -> Option<i64> {
        let (span, _, _) = attribute_value(tag, name)?;
        std::str::from_utf8(&tag[span]).ok()?.trim().parse().ok()
    };
    Some((value(b"cx")?, value(b"cy")?)).filter(|(width, height)| *width > 0 && *height > 0)
}

/// Replacements of byte ranges of one slide, applied in one pass.
#[derive(Default)]
pub(super) struct Splice {
    edits: Vec<(Range<usize>, String)>,
}

impl Splice {
    pub(super) fn replace(&mut self, range: Range<usize>, with: impl Into<String>) {
        self.edits.push((range, with.into()));
    }

    pub(super) fn apply(mut self, xml: &[u8], max_bytes: usize) -> Result<Vec<u8>> {
        self.edits.sort_by_key(|(range, _)| range.start);
        let mut out = Vec::new();
        let mut at = 0usize;
        let size = self.edits.iter().fold(xml.len(), |size, (range, with)| {
            size.saturating_sub(range.len()).saturating_add(with.len())
        });
        if size > max_bytes {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
                observed: size,
                maximum: max_bytes,
            });
        }
        out.try_reserve_exact(size)
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation equation XML",
                source,
            })?;
        for (range, with) in &self.edits {
            if range.start < at || range.end > xml.len() {
                return Err(invalid("equation edits overlap"));
            }
            out.extend_from_slice(&xml[at..range.start]);
            out.extend_from_slice(with.as_bytes());
            at = range.end;
        }
        out.extend_from_slice(&xml[at..]);
        Ok(out)
    }
}

/// Replace the `cx` and `cy` values of an `a:ext` start tag at `tag`.
pub(super) fn resize(
    splice: &mut Splice,
    xml: &[u8],
    tag: &Range<usize>,
    size: (i64, i64),
) -> Result<()> {
    let bytes = &xml[tag.clone()];
    for (name, value) in [(b"cx".as_slice(), size.0), (b"cy", size.1)] {
        let (span, _, _) = attribute_value(bytes, name)
            .ok_or_else(|| invalid("the equation frame has no extent"))?;
        splice.replace(
            tag.start + span.start..tag.start + span.end,
            value.to_string(),
        );
    }
    Ok(())
}

/// The equation at `index` and the block holding it, refused when the
/// equation is not in the text of a shape-level markup-compatibility block
/// with one Choice and no markup compatibility of its own.
fn located(
    sites: &[Site],
    blocks: &[Block],
    index: usize,
    operation: &'static str,
) -> Result<(Site, usize)> {
    let site = sites.get(index).ok_or(Error::IndexOutOfBounds {
        index,
        len: sites.len(),
    })?;
    let block = site
        .block
        .filter(|block| blocks.get(*block).is_some_and(|block| block.shape_level))
        .ok_or_else(|| {
            refused(
                operation,
                "only an equation in the text of a shape-level markup-compatibility block is edited",
            )
        })?;
    editable_block(&blocks[block], operation)?;
    Ok((site.clone(), block))
}

fn editable_block(block: &Block, operation: &'static str) -> Result<()> {
    if block.nested || block.choices != 1 || block.choice_shape.is_none() {
        return Err(refused(
            operation,
            "the equation's shape holds markup-compatibility content of its own",
        ));
    }
    Ok(())
}

/// Whether the block's shape holds this one equation and no other text.
const fn equation_only(block: &Block) -> bool {
    block.equations == 1 && !block.other_content
}

/// The edit that gives block `index` of `xml` a text Fallback: a copy of
/// its Choice shape in which every equation is a run of its linear form,
/// with no picture fill. Office ignores the Fallback; other readers show
/// the text.
fn text_fallback(xml: &[u8], index: usize) -> Result<(Range<usize>, String)> {
    let (sites, blocks) = sites::scan(xml)?;
    let block = blocks
        .get(index)
        .ok_or_else(|| invalid("the equation block disappeared"))?;
    let shape = block
        .choice_shape
        .clone()
        .ok_or_else(|| invalid("the equation block has no Choice shape"))?;
    let mut copy = Splice::default();
    for site in sites
        .iter()
        .filter(|site| shape.start <= site.element.start && site.element.end <= shape.end)
    {
        let omml = std::str::from_utf8(&xml[site.content.clone()])
            .map_err(|error| Error::Xml(error.to_string()))?;
        copy.replace(
            site.element.start - shape.start..site.element.end - shape.start,
            linear_run(omml),
        );
    }
    let copy = String::from_utf8(copy.apply(&xml[shape], usize::MAX)?)
        .map_err(|error| Error::Xml(error.to_string()))?;
    let block_name = tag_name(&xml[block.range.clone()])?;
    match (&block.fallback, &block.fallback_tag) {
        (Some((range, empty)), Some(tag)) => {
            let name = tag_name(&xml[tag.clone()])?;
            let open = std::str::from_utf8(&xml[tag.clone()])
                .map_err(|error| Error::Xml(error.to_string()))?;
            let open = if *empty {
                let trimmed = open.trim_end_matches('>').trim_end().trim_end_matches('/');
                format!("{trimmed}>")
            } else {
                open.to_owned()
            };
            Ok((range.clone(), format!("{open}{copy}</{name}>")))
        },
        _ => {
            // No Fallback yet: one goes before the block's end tag, under
            // the block's own prefix.
            let prefix = block_name
                .split_once(':')
                .map_or(String::new(), |(prefix, _)| format!("{prefix}:"));
            let close = format!("</{block_name}>");
            let at = block
                .range
                .end
                .checked_sub(close.len())
                .filter(|at| xml[*at..block.range.end] == *close.as_bytes())
                .ok_or_else(|| invalid("the equation block is malformed"))?;
            Ok((
                at..at,
                format!("<{prefix}Fallback>{copy}</{prefix}Fallback>"),
            ))
        },
    }
}

impl Edit {
    /// What a slide placeholder shape (`shape_id`) inherits from its layout
    /// and master (ECMA-376 Part 1, §19.3.1.36: matched by index on the
    /// layout and by type on the master): its frame size, when it has none
    /// of its own, and the master's first-level text size for its kind
    /// (§19.3.1.51 `txStyles`). Both `None` for a shape that is no
    /// placeholder.
    fn placeholder_inheritance(
        &self,
        slide: &PackURI,
        xml: &[u8],
        shape_id: u32,
    ) -> Result<Inherited> {
        let limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(xml, limits)?;
        let Some(placeholder) = scene
            .iter()
            .find(|shape| shape.id() == Some(shape_id))
            .and_then(crate::shape::Shape::placeholder)
        else {
            return Ok(Inherited::default());
        };
        // An omitted type is `obj`, whose master counterpart is the body.
        let kind = placeholder.kind().unwrap_or("obj");
        let master_kind = match kind {
            "obj" | "subTitle" => "body",
            "ctrTitle" => "title",
            other => other,
        };
        let related = |owner: &PackURI, kind: &str| -> Result<Option<PackURI>> {
            let part = self.working.get_part(owner)?;
            part.rels()
                .iter()
                .filter(|relationship| !relationship.is_external())
                .find(|relationship| relationship.reltype().ends_with(kind))
                .map(litchi_opc::Relationship::target_partname)
                .transpose()
                .map_err(Error::Opc)
        };
        let mut inherited = Inherited::default();
        let mut owner = slide.clone();
        for (relationship, by_index) in [("/slideLayout", true), ("/slideMaster", false)] {
            let Some(next) = related(&owner, relationship)? else {
                break;
            };
            owner = next;
            let blob = self.working.get_part(&owner)?.blob();
            if inherited.extent.is_none() {
                let scene = crate::shape::Scene::read(blob, limits)?;
                let found = scene.iter().find_map(|shape| {
                    let candidate = shape.placeholder()?;
                    let candidate_kind = candidate.kind().unwrap_or("obj");
                    let same = if by_index && placeholder.index() != 0 {
                        candidate.index() == placeholder.index()
                    } else if by_index {
                        candidate_kind == kind
                    } else {
                        candidate_kind == master_kind
                    };
                    same.then(|| shape.bounds())
                });
                if let Some(Some(bounds)) = found {
                    inherited.extent = Some((bounds.width(), bounds.height()))
                        .filter(|(width, height)| *width > 0 && *height > 0);
                }
            }
            if !by_index {
                let style: &[u8] = if master_kind == "title" {
                    b"titleStyle"
                } else {
                    b"bodyStyle"
                };
                inherited.text_size_pt = master_text_size(blob, style)?;
            }
        }
        Ok(inherited)
    }

    /// The conformance-specific DrawingML namespace of a slide.
    fn drawing_namespace(&self, xml: &[u8]) -> Result<crate::media_parts::Conformance> {
        crate::media_parts::document_conformance_with_limits(
            xml,
            self.semantic_limits()
                .media_parts()
                .constrained_by_source(self.working.read_limits()),
        )
    }

    /// Stage `next` as the XML of `slide` in `candidate` (which may hold a
    /// new picture part), after checking that every shape but the edited
    /// one (`shape`, which a reader without Office Math may only now see)
    /// keeps its place and ID, that a `removed` shape is gone, and that
    /// `equations` remain; the relationships `before` referenced and `next`
    /// no longer does are removed, with the parts nothing else uses.
    #[allow(
        clippy::too_many_arguments,
        reason = "one staging step; the arguments are its inputs and its checks"
    )]
    fn publish_equation_edit(
        &mut self,
        slide: &PackURI,
        mut candidate: super::Working,
        before: &[u8],
        next: Vec<u8>,
        equations: usize,
        shape: u32,
        removed: bool,
    ) -> Result<()> {
        let shape_limits = self.semantic_limits().shape();
        require_declared_prefixes(&next)?;
        if removed {
            crate::opened::xml::validate_shape_removal_dependencies(
                &next,
                &std::collections::BTreeSet::from([shape]),
            )?;
        }
        let others = |xml: &[u8]| -> Result<Vec<Option<u32>>> {
            Ok(crate::shape::Scene::read(xml, shape_limits)?
                .iter()
                .map(crate::shape::Shape::id)
                .filter(|id| *id != Some(shape))
                .collect())
        };
        let gone = !crate::shape::Scene::read(&next, shape_limits)?
            .iter()
            .any(|staged| staged.id() == Some(shape));
        if others(before)? != others(&next)?
            || (removed && !gone)
            || sites::scan(&next)?.0.len() != equations
        {
            return Err(invalid("the equation edit did not round-trip"));
        }
        let owner = candidate.get_part(slide)?;
        let referenced = shape_relationship_ids(before, owner.rels())?;
        let retained = shape_relationship_ids(&next, owner.rels())?;
        let mut roots = Vec::new();
        {
            let owner = candidate.get_part_mut(slide)?;
            owner.set_blob(next);
            for id in referenced.difference(&retained) {
                if let Some(relationship) = owner.rels_mut().remove(id)
                    && !relationship.is_external()
                {
                    roots.push(relationship.target_partname()?);
                }
            }
        }
        remove_unreferenced_dependencies(&mut candidate, roots)?;
        self.working = candidate;
        self.note_scene_read(slide);
        Ok(())
    }

    /// Add an equation as a new text box at `origin` (`x`, `y` in EMU), in
    /// the form PowerPoint writes: a markup-compatibility block whose Choice
    /// holds the OMML in a display paragraph and whose Fallback shows a PNG
    /// picture of it, with the picture part and its relationship.
    ///
    /// The text box sizes itself to the equation at `font_size_pt` (18 when
    /// `None`); a given size is also written on the equation's runs.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Equation`] for LaTeX or OMML that does not convert
    /// or render, an error for malformed OMML, a missing slide, malformed
    /// slide XML or exceeded bounds. A failure leaves the edit unchanged.
    pub fn add_equation<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        source: EquationSource<'_>,
        origin: (i64, i64),
        font_size_pt: Option<f64>,
    ) -> Result<EquationWrite> {
        let selected = self.resolve_slide(slide.into())?;
        let max_patch_bytes = self.source_ref()?.limits.max_patch_bytes();
        let shape_limits = self.semantic_limits().shape();
        let owner = self.working.get_part(&selected.part_name)?;
        let xml = owner.blob();
        let conformance = self.drawing_namespace(xml)?;
        let (pml, dml, rel) = conformance_namespaces(conformance);
        let size = match font_size_pt {
            Some(points) => {
                // ST_TextFontSize: 1 to 4000 points in hundredths.
                if !(points.is_finite() && (1.0..=4000.0).contains(&points)) {
                    return Err(invalid("the equation text size must be 1 to 4000 points"));
                }
                #[allow(
                    clippy::cast_possible_truncation,
                    clippy::cast_sign_loss,
                    reason = "the size is checked to lie in 1..=4000 points"
                )]
                Some((points * 100.0).round() as u32)
            },
            None => None,
        };
        let run = HostRun::native(size);
        let prepared = prepare(
            source,
            Form::Display(Some(Justification::CenterGroup)),
            &run,
            dml,
            &root_bindings(xml)?,
            max_patch_bytes,
        )?;
        let picture = render(&prepared.omml, font_size_pt.unwrap_or(DEFAULT_FONT_SIZE_PT))?;
        let (width, height) = fitted(&picture, DEFAULT_INSETS)?;
        let png = fallback_picture(&picture, (width, height), DEFAULT_INSETS, (0.5, 0.0))?;
        let (x, y) = origin;
        if x.checked_add(width).is_none() || y.checked_add(height).is_none() {
            return Err(invalid(
                "the equation frame lies outside the coordinate space",
            ));
        }
        let shape_id = max_raw_shape_id(xml)?
            .max(max_shape_reference_id(xml)?.unwrap_or(1))
            .checked_add(1)
            .ok_or_else(|| invalid("opened-presentation equation shape ID overflow"))?;
        let before = sites::scan(xml)?.0.len();
        let part_name = free_part_name(&self.working, "/ppt/media/image", ".png")?;
        let mut candidate = self.working.clone();
        candidate.try_add_part(Box::new(BlobPart::new(
            part_name.clone(),
            "image/png".to_owned(),
            png,
        )))?;
        let target = part_name.relative_ref(selected.part_name.base_uri());
        let relationship = candidate
            .get_part_mut(&selected.part_name)?
            .relate_to(&target, conformance_image_relationship(conformance));
        let name = format!("TextBox {}", shape_id - 1);
        let end_size = size.map_or_else(String::new, |size| format!(r#" sz="{size}""#));
        let geometry = format!(
            r#"<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#
        );
        let fragment = format!(
            concat!(
                r#"<mc:AlternateContent xmlns:mc="{mce}" xmlns:a14="{a14}" xmlns:p="{pml}" xmlns:a="{dml}" xmlns:r="{rel}">"#,
                r#"<mc:Choice Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#,
                r#"<p:spPr>{geometry}<a:noFill/></p:spPr>"#,
                r#"<p:txBody><a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a14:m>{omml}</a14:m><a:endParaRPr lang="en-US"{end_size} dirty="0"/></a:p></p:txBody></p:sp></mc:Choice>"#,
                r#"<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr txBox="1">{locks}</p:cNvSpPr><p:nvPr/></p:nvSpPr>"#,
                r#"<p:spPr>{geometry}<a:blipFill rotWithShape="1"><a:blip r:embed="{relationship}"/>{fill}</a:blipFill></p:spPr>"#,
                "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"><a:noFill/></a:rPr><a:t>\u{a0}</a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>",
            ),
            mce = MCE,
            a14 = A14,
            pml = pml,
            dml = dml,
            rel = rel,
            id = shape_id,
            name = name,
            geometry = geometry,
            omml = prepared.omml,
            end_size = end_size,
            locks = FALLBACK_LOCKS,
            relationship = relationship,
            fill = FALLBACK_FILL,
        );
        let next = crate::opened::xml::append_shape_with_limit(
            candidate.get_part(&selected.part_name)?.blob(),
            fragment.as_bytes(),
            shape_limits.output_bytes(),
        )?;
        ensure_shape_id(&next, shape_id, "equation", shape_limits)?;
        let (after, _) = sites::scan(&next)?;
        if after.len() != before + 1 {
            return Err(invalid("the equation shape did not round-trip"));
        }
        candidate.get_part_mut(&selected.part_name)?.set_blob(next);
        self.working = candidate;
        self.note_scene_read(&selected.part_name);
        Ok(EquationWrite {
            shape_id,
            index: before,
            losses: prepared.losses,
        })
    }

    /// Add an equation at the end of the text of the shape `shape_id` (its
    /// `cNvPr` ID): inline, after the last paragraph's runs. A shape that
    /// holds no equation yet is wrapped in the markup-compatibility block
    /// PowerPoint writes; its Fallback (here, and whenever a shape's text
    /// mixes equations with other text) is a copy of the shape in which each
    /// equation is a run of its linear form (UnicodeMath), for readers
    /// without Office Math.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing shape or one without a text paragraph,
    /// a refusal for a shape with markup-compatibility content of its own,
    /// [`Error::Equation`] for input that does not convert, and an error for
    /// malformed OMML or slide XML. A failure leaves the edit unchanged.
    pub fn add_equation_to_shape<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape_id: u32,
        source: EquationSource<'_>,
    ) -> Result<EquationWrite> {
        const OPERATION: &str = "add_equation_to_shape";
        let selected = self.resolve_slide(slide.into())?;
        let max_patch_bytes = self.source_ref()?.limits.max_patch_bytes();
        let shape_limits = self.semantic_limits().shape();
        let xml = self.working.get_part(&selected.part_name)?.blob().to_vec();
        let (_, dml, _) = conformance_namespaces(self.drawing_namespace(&xml)?);
        let (sites, blocks) = sites::scan(&xml)?;
        let mut held = None;
        for (index, block) in blocks.iter().enumerate() {
            if let Some(shape) = &block.choice_shape
                && cnvpr_id(&xml[shape.clone()])? == Some(shape_id)
            {
                editable_block(block, OPERATION)?;
                held = Some((index, shape.clone()));
                break;
            }
        }
        let (shape, wrapped) = match held {
            Some((_, ref shape)) => (shape.clone(), false),
            None => (
                sites::plain_shape(&xml, shape_id)?
                    .ok_or_else(|| refused(OPERATION, "the target is not a text shape"))?,
                true,
            ),
        };
        let end = sites::text_end(&xml[shape.clone()])?
            .ok_or_else(|| refused(OPERATION, "the shape has no text paragraph"))?;
        let prepared = prepare(
            source,
            Form::Inline,
            // The equation takes the text size of the paragraph it ends.
            &HostRun::native(end.size),
            dml,
            &root_bindings(&xml)?,
            max_patch_bytes,
        )?;
        let math = format!(r#"<a14:m xmlns:a14="{A14}">{}</a14:m>"#, prepared.omml);
        // An empty paragraph (`<a:p/>`) opens to take the equation.
        let (at, math) = match end.place {
            sites::TextPlace::Before(at) => (at..at, math),
            sites::TextPlace::Empty(tag) => {
                let name = tag_name(&xml[shape.start + tag.start..shape.start + tag.end])?;
                let open =
                    std::str::from_utf8(&xml[shape.start + tag.start..shape.start + tag.end])
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .trim_end_matches('>')
                        .trim_end()
                        .trim_end_matches('/')
                        .to_owned();
                (tag, format!("{open}>{math}</{name}>"))
            },
        };
        let mut splice = Splice::default();
        if wrapped {
            let mut inner = Splice::default();
            inner.replace(at, math);
            let inner = String::from_utf8(inner.apply(&xml[shape.clone()], usize::MAX)?)
                .map_err(|error| Error::Xml(error.to_string()))?;
            splice.replace(
                shape.clone(),
                format!(
                    r#"<mc:AlternateContent xmlns:mc="{MCE}" xmlns:a14="{A14}"><mc:Choice Requires="a14">{inner}</mc:Choice></mc:AlternateContent>"#
                ),
            );
        } else {
            splice.replace(shape.start + at.start..shape.start + at.end, math);
        }
        let staged = splice.apply(&xml, shape_limits.output_bytes())?;
        let (_, staged_blocks) = sites::scan(&staged)?;
        let block = if wrapped {
            staged_blocks
                .iter()
                .position(|block| block.range.start == shape.start)
        } else {
            held.as_ref().map(|(index, _)| *index)
        }
        .ok_or_else(|| invalid("the equation block disappeared"))?;
        let (range, fallback) = text_fallback(&staged, block)?;
        let mut splice = Splice::default();
        splice.replace(range, fallback);
        let next = splice.apply(&staged, shape_limits.output_bytes())?;
        let (after, after_blocks) = sites::scan(&next)?;
        let choice = after_blocks
            .get(block)
            .and_then(|block| block.choice_shape.clone())
            .ok_or_else(|| invalid("the equation block disappeared"))?;
        let index = after
            .iter()
            .rposition(|site| choice.start <= site.element.start && site.element.end <= choice.end)
            .ok_or_else(|| invalid("the added equation did not round-trip"))?;
        let candidate = self.working.clone();
        self.publish_equation_edit(
            &selected.part_name,
            candidate,
            &xml,
            next,
            sites.len() + 1,
            shape_id,
            false,
        )?;
        Ok(EquationWrite {
            shape_id,
            index,
            losses: prepared.losses,
        })
    }

    /// Replace the content of the slide's equation at `equation` (its
    /// position among [`Slide::equations`](crate::slide::Slide::equations)).
    /// The Choice branch gets the new OMML. When the shape holds only this
    /// equation and its Fallback shows a picture, the picture is redrawn (a
    /// shape that resizes to its text, `a:spAutoFit`, takes the new size);
    /// otherwise the Fallback becomes a copy of the shape with every
    /// equation as a run of its linear form. The shape keeps its ID,
    /// position and other properties; pictures and relationships nothing
    /// else uses are removed. LaTeX keeps the equation's form (inline or
    /// display, with its justification) and its first run's host properties
    /// (text size, fonts). Returns `None` when the OMML is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing equation, a
    /// refusal for an equation outside a shape-level markup-compatibility
    /// block, [`Error::Equation`] for input that does not convert or render,
    /// and an error for malformed OMML or slide XML. A failure leaves the
    /// edit unchanged.
    pub fn set_equation<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        equation: usize,
        source: EquationSource<'_>,
    ) -> Result<Option<EquationWrite>> {
        let selected = self.resolve_slide(slide.into())?;
        let max_patch_bytes = self.source_ref()?.limits.max_patch_bytes();
        let shape_limits = self.semantic_limits().shape();
        let xml = self.working.get_part(&selected.part_name)?.blob().to_vec();
        let (sites, blocks) = sites::scan(&xml)?;
        let (site, index) = located(&sites, &blocks, equation, "set_equation")?;
        let block = blocks[index].clone();
        let conformance = self.drawing_namespace(&xml)?;
        let (_, dml, _) = conformance_namespaces(conformance);
        let old = std::str::from_utf8(&xml[site.content.clone()])
            .map_err(|error| Error::Xml(error.to_string()))?;
        let run = HostRun::of(old)?.unwrap_or_else(|| HostRun::native(None));
        let prepared = prepare(
            source,
            Form::of(old)?,
            &run,
            dml,
            &root_bindings(&xml)?,
            max_patch_bytes,
        )?;
        if prepared.omml == old.trim() {
            return Ok(None);
        }
        let shape_id = cnvpr_id(&xml[block.range.clone()])?
            .ok_or_else(|| invalid("the equation shape has no ID"))?;
        let mut splice = Splice::default();
        let candidate = if equation_only(&block) && block.blip.is_some() {
            let inherited = self.placeholder_inheritance(&selected.part_name, &xml, shape_id)?;
            let picture = render(
                &prepared.omml,
                run.font_size_pt()
                    .or(inherited.text_size_pt)
                    .unwrap_or(DEFAULT_FONT_SIZE_PT),
            )?;
            let blip = block
                .blip
                .clone()
                .ok_or_else(|| invalid("the equation fallback has no picture"))?;
            let (embed, _, _) = attribute_value(&xml[blip.clone()], b"embed")
                .ok_or_else(|| invalid("the equation picture is not embedded"))?;
            let embed = blip.start + embed.start..blip.start + embed.end;
            // The frame the picture covers: resized to the new picture when
            // the shape fits its text, otherwise its own or its
            // placeholder's; a placeholder frame that cannot be found is
            // taken as the picture's.
            let frame = match (&block.choice_extent, &block.fallback_extent) {
                (Some(choice), Some(fallback)) if block.autofit => {
                    let size = fitted(&picture, block.insets)?;
                    resize(&mut splice, &xml, choice, size)?;
                    resize(&mut splice, &xml, fallback, size)?;
                    Some(size)
                },
                (Some(choice), _) => extent(&xml[choice.clone()]),
                (None, _) => inherited.extent,
            };
            let frame = match frame {
                Some(frame) => frame,
                None => fitted(&picture, block.insets)?,
            };
            if let Some(stretch) = block.stretch.clone() {
                splice.replace(stretch, FALLBACK_FILL);
            }
            let place = (
                horizontal(&prepared.omml, site.align.as_deref())?,
                vertical(block.anchor.as_deref()),
            );
            let png = fallback_picture(&picture, frame, block.insets, place)?;
            let part_name = free_part_name(&self.working, "/ppt/media/image", ".png")?;
            let mut candidate = self.working.clone();
            candidate.try_add_part(Box::new(BlobPart::new(
                part_name.clone(),
                "image/png".to_owned(),
                png,
            )))?;
            let target = part_name.relative_ref(selected.part_name.base_uri());
            let new_id = candidate
                .get_part_mut(&selected.part_name)?
                .relate_to(&target, conformance_image_relationship(conformance));
            splice.replace(embed, new_id);
            splice.replace(site.content.clone(), prepared.omml);
            candidate
        } else {
            splice.replace(site.content.clone(), prepared.omml);
            self.working.clone()
        };
        let mut next = splice.apply(&xml, shape_limits.output_bytes())?;
        if !(equation_only(&block) && block.blip.is_some()) {
            let (range, fallback) = text_fallback(&next, index)?;
            let mut splice = Splice::default();
            splice.replace(range, fallback);
            next = splice.apply(&next, shape_limits.output_bytes())?;
        }
        self.publish_equation_edit(
            &selected.part_name,
            candidate,
            &xml,
            next,
            sites.len(),
            shape_id,
            false,
        )?;
        Ok(Some(EquationWrite {
            shape_id,
            index: equation,
            losses: prepared.losses,
        }))
    }

    /// Remove the slide's equation at `equation` (its position among
    /// [`Slide::equations`](crate::slide::Slide::equations)). A shape whose
    /// text is only this equation leaves with it: the whole
    /// markup-compatibility block, with or without a picture Fallback. In a
    /// shape whose text holds more, only the equation leaves: the Fallback
    /// becomes a copy of the shape with every remaining equation as a run of
    /// its linear form, and a shape left without equations is unwrapped to
    /// its plain Choice shape. Pictures and relationships nothing else uses
    /// are removed. Returns the ID of the shape that held the equation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing equation, a
    /// refusal for an equation outside a shape-level markup-compatibility
    /// block or a removed shape other markup refers to (connectors,
    /// animations), and an error for malformed slide XML. A failure leaves
    /// the edit unchanged.
    pub fn remove_equation<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        equation: usize,
    ) -> Result<u32> {
        let selected = self.resolve_slide(slide.into())?;
        let shape_limits = self.semantic_limits().shape();
        let xml = self.working.get_part(&selected.part_name)?.blob().to_vec();
        let (sites, blocks) = sites::scan(&xml)?;
        let (site, index) = located(&sites, &blocks, equation, "remove_equation")?;
        let block = &blocks[index];
        let shape_id = cnvpr_id(&xml[block.range.clone()])?
            .ok_or_else(|| invalid("the equation shape has no ID"))?;
        let mut splice = Splice::default();
        if equation_only(block) {
            splice.replace(block.range.clone(), String::new());
            let next = splice.apply(&xml, shape_limits.output_bytes())?;
            let candidate = self.working.clone();
            self.publish_equation_edit(
                &selected.part_name,
                candidate,
                &xml,
                next,
                sites.len() - 1,
                shape_id,
                true,
            )?;
            return Ok(shape_id);
        }
        splice.replace(site.element.clone(), String::new());
        let staged = splice.apply(&xml, shape_limits.output_bytes())?;
        let (_, staged_blocks) = sites::scan(&staged)?;
        let staged_block = staged_blocks
            .get(index)
            .ok_or_else(|| invalid("the equation block disappeared"))?;
        let choice_shape = staged_block
            .choice_shape
            .clone()
            .ok_or_else(|| invalid("the equation block has no Choice shape"))?;
        let mut splice = Splice::default();
        if staged_block.equations == 0 {
            // No equation is left: the shape stands without the block, with
            // the block's declarations it relies on.
            let shape = carry_declarations(
                &staged[staged_block.range.start..choice_shape.start],
                &staged[choice_shape],
            )?;
            splice.replace(staged_block.range.clone(), shape);
        } else {
            let (range, fallback) = text_fallback(&staged, index)?;
            splice.replace(range, fallback);
        }
        let next = splice.apply(&staged, shape_limits.output_bytes())?;
        let candidate = self.working.clone();
        self.publish_equation_edit(
            &selected.part_name,
            candidate,
            &xml,
            next,
            sites.len() - 1,
            shape_id,
            false,
        )?;
        Ok(shape_id)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn master_text_sizes_are_read_per_style() {
        let master = br#"<p:sldMaster xmlns:p="p" xmlns:a="a"><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr marL="0"><a:defRPr sz="2800"><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr><a:lvl2pPr><a:defRPr sz="2400"/></a:lvl2pPr></p:bodyStyle></p:txStyles></p:sldMaster>"#;
        assert_eq!(master_text_size(master, b"titleStyle").unwrap(), Some(44.0));
        assert_eq!(master_text_size(master, b"bodyStyle").unwrap(), Some(28.0));
        assert_eq!(master_text_size(master, b"otherStyle").unwrap(), None);
    }

    #[test]
    fn frames_and_extents() {
        assert_eq!(extent(br#"<a:ext cx="10" cy="20"/>"#), Some((10, 20)));
        assert_eq!(extent(br#"<a:ext cx="0" cy="20"/>"#), None);
        assert_eq!(extent(br#"<a:ext cy="20"/>"#), None);
        assert_eq!(
            max_raw_shape_id(br#"<r><p:cNvPr id="7"/><x><a:cNvPr id="12"/></x></r>"#).unwrap(),
            12
        );
    }
}
