//! Shape and text-run click actions (record 0321).
//!
//! [`Edit::set_shape_hyperlink`] writes a shape's click action on its
//! `p:cNvPr` and on every text run, as OfficeCLI does so readers that look
//! at either place agree; [`Edit::set_run_hyperlink`] writes one run's.
//! `None` removes them. An external target reuses the slide's hyperlink
//! relationship to the same URL, a slide target its relationship to that
//! slide; relationships a removed click leaves unused are kept, as the
//! source-preserving editors keep them. Every byte outside the clicks the
//! verb names is kept.

use std::ops::Range;

use litchi_core::Position;
use litchi_opc::constants::relationship_type as rt;

use super::{Edit, invalid};
use crate::Result;
use crate::hyperlinks::link::{Element, REL, STRICT_REL, ShapeLinks};
use crate::hyperlinks::{Link, Target};

const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const STRICT_SLIDE_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/slide";

/// Which clicks of a shape a verb writes.
#[derive(Clone, Copy)]
enum Scope {
    Shape,
    Run { paragraph: usize, run: usize },
}

fn escape_attribute(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\t' => out.push_str("&#9;"),
            '\n' => out.push_str("&#10;"),
            '\r' => out.push_str("&#13;"),
            _ => out.push(ch),
        }
    }
    out
}

fn valid_text(value: &str) -> bool {
    value.len() <= 1 << 20
        && value.chars().all(|ch| {
            matches!(ch, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')
        })
}

/// Close an empty element's start tag around `content`.
fn expand(tag: &[u8], element: &Element, content: &[u8]) -> Result<Vec<u8>> {
    let open = tag
        .strip_suffix(b"/>")
        .ok_or_else(|| invalid("an empty element does not end with />"))?;
    let open = open.trim_ascii_end();
    let mut out = Vec::with_capacity(tag.len() + content.len() + 16);
    out.extend_from_slice(open);
    out.push(b'>');
    out.extend_from_slice(content);
    out.extend_from_slice(b"</");
    if !element.prefix.is_empty() {
        out.extend_from_slice(&element.prefix);
        out.push(b':');
    }
    out.extend_from_slice(&element.local);
    out.push(b'>');
    Ok(out)
}

impl Edit {
    /// Set (`Some`) or remove (`None`) the click action of a shape (`p:sp`)
    /// at a depth-first position of a slide: on its `p:cNvPr` and on every
    /// text run of its text body. A picture (`p:pic`) or a group
    /// (`p:grpSp`) carries it on its `p:cNvPr` only (record 0325). Returns
    /// whether the slide changed; an equal action is an exact no-op.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_pptx::hyperlinks::{Link, Target};
    /// let source = litchi_pptx::Package::new()?;
    /// let mut edit = source.edit()?;
    /// let slide = edit.insert_slide(Position::new(0), "Blank")?;
    /// edit.add_text_box(slide.name(), "Docs", (0, 0, 914_400, 914_400))?;
    /// let link = Link { target: Target::Url("https://example.com/".into()), tooltip: None };
    /// assert!(edit.set_shape_hyperlink(0_usize, 0_usize, Some(&link))?);
    /// assert!(!edit.set_shape_hyperlink(0_usize, 0_usize, Some(&link))?);
    /// let _package = edit.commit()?.into_package();
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide or shape, a shape that is not a
    /// `p:sp`, `p:pic` or `p:grpSp`, a slide target past the last slide, text XML cannot carry,
    /// or markup-compatibility content.
    pub fn set_shape_hyperlink<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        link: Option<&Link>,
    ) -> Result<bool> {
        self.write_clicks(slide.into(), shape.into(), Scope::Shape, link)
    }

    /// Set (`Some`) or remove (`None`) the click action of one text run: the
    /// run at a zero-based position among the run-like children (`a:r`,
    /// `a:br`, `a:fld`) of a zero-based paragraph of a shape's text body.
    /// Returns whether the slide changed; an equal action is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide, shape, paragraph or run, a run
    /// that is not a text run, a slide target past the last slide, text XML
    /// cannot carry, or markup-compatibility content.
    pub fn set_run_hyperlink<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        paragraph: Position,
        run: Position,
        link: Option<&Link>,
    ) -> Result<bool> {
        self.write_clicks(
            slide.into(),
            shape.into(),
            Scope::Run {
                paragraph: paragraph.get(),
                run: run.get(),
            },
            link,
        )
    }

    fn write_clicks(
        &mut self,
        slide: crate::slide::Selector<'_>,
        key: crate::shape::Key<'_>,
        scope: Scope,
        link: Option<&Link>,
    ) -> Result<bool> {
        self.ensure_ready()?;
        if let Some(link) = link {
            let text_ok = match &link.target {
                Target::Url(url) => !url.is_empty() && valid_text(url),
                Target::Action(action) => !action.is_empty() && valid_text(action),
                Target::Slide(position) => position.get() < self.slides.len(),
            };
            if !text_ok || link.tooltip.as_deref().is_some_and(|tip| !valid_text(tip)) {
                return Err(invalid("the hyperlink target or tooltip is not valid"));
            }
        }
        let selected = self.resolve_slide(slide)?;
        let owner = self.working.get_part(&selected.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let xml = owner.blob().to_vec();
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(&xml, shape_limits)?;
        if scene.is_rewritten() {
            return Err(crate::Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "set_hyperlink",
                "hyperlink edits do not select markup-compatibility branches",
            ));
        }
        let shape = scene.shape(key)?;
        // A picture's and a group's click lives on its `p:cNvPr` alone: a
        // picture has no text body, and a group's runs belong to its members.
        match (shape, scope) {
            (crate::shape::Shape::Auto(_), _)
            | (crate::shape::Shape::Picture(_) | crate::shape::Shape::Group(_), Scope::Shape) => {},
            _ => {
                return Err(invalid(
                    "hyperlinks are written on p:sp shapes, and on p:pic and p:grpSp non-visual properties",
                ));
            },
        }
        let span = shape.common().span()?.range(xml.len())?;
        let fragment = shape.common().self_contained_xml()?;
        let links = ShapeLinks::scan(&fragment)?;
        let delta = fragment
            .len()
            .checked_sub(span.len())
            .ok_or_else(|| invalid("the self-contained shape is shorter than its span"))?;
        let map = |position: usize| -> Result<usize> {
            if position < links.root_open_end {
                return Err(invalid(
                    "a hyperlink edit falls inside the shape's root tag",
                ));
            }
            Ok(span.start + position - delta)
        };
        let shape_id = shape.id();

        // The elements the verb writes and the runs that lack properties.
        let mut targets: Vec<&Element> = Vec::new();
        let mut bare: Vec<(usize, &[u8])> = Vec::new();
        match scope {
            Scope::Shape => {
                if let Some(non_visual) = &links.non_visual {
                    targets.push(non_visual);
                }
                for run in links.paragraphs.iter().flatten().filter(|run| run.text) {
                    match &run.properties {
                        Some(properties) => targets.push(properties),
                        None => bare.push((run.open_end, &run.prefix)),
                    }
                }
            },
            Scope::Run { paragraph, run } => {
                let slot = links
                    .paragraphs
                    .get(paragraph)
                    .and_then(|runs| runs.get(run))
                    .ok_or_else(|| invalid("the run position is past the last run"))?;
                if !slot.text {
                    return Err(invalid("hyperlinks are written on text runs"));
                }
                match &slot.properties {
                    Some(properties) => targets.push(properties),
                    None => bare.push((slot.open_end, &slot.prefix)),
                }
            },
        }
        // An exact no-op: every slot already carries an equal action (or,
        // for a removal, none).
        let unchanged = match link {
            None => targets.iter().all(|element| element.clicks.is_empty()),
            Some(link) => {
                bare.is_empty()
                    && targets.iter().all(|element| {
                        element.clicks.len() == 1
                            && element.click.as_ref().is_some_and(|click| {
                                self.resolve_click(&selected.part_name, click)
                                    .ok()
                                    .flatten()
                                    .is_some_and(|current| current == *link)
                            })
                    })
            },
        };
        if unchanged {
            return Ok(false);
        }

        // The new click, with its relationship staged on a candidate.
        let mut candidate = self.working.clone();
        let click = match link {
            None => None,
            Some(link) => {
                let part = candidate.get_part_mut(&selected.part_name)?;
                let (id, action) = match &link.target {
                    Target::Url(url) => {
                        let reltype = if links.strict {
                            rt::STRICT_HYPERLINK
                        } else {
                            rt::HYPERLINK
                        };
                        (part.relate_to_ext(url, reltype), None)
                    },
                    Target::Slide(position) => {
                        let target = &self.slides[position.get()].part_name;
                        let reference = target.relative_ref(selected.part_name.base_uri());
                        let reltype = if links.strict {
                            STRICT_SLIDE_REL
                        } else {
                            rt::SLIDE
                        };
                        (
                            part.relate_to(&reference, reltype),
                            Some("ppaction://hlinksldjump"),
                        )
                    },
                    Target::Action(action) => (String::new(), Some(action.as_str())),
                };
                let (prefix, declare) = match links.drawing_prefix.as_deref() {
                    Some(prefix) if !prefix.is_empty() => (
                        String::from_utf8(prefix.to_vec())
                            .map_err(|_| invalid("the DrawingML prefix is not UTF-8"))?,
                        false,
                    ),
                    _ => ("a".to_owned(), true),
                };
                let mut out = format!("<{prefix}:hlinkClick");
                if declare {
                    out.push_str(&format!(
                        " xmlns:{prefix}=\"{}\"",
                        if links.strict { STRICT_DML } else { DML }
                    ));
                }
                out.push_str(&format!(
                    " xmlns:r=\"{}\" r:id=\"{}\"",
                    if links.strict { STRICT_REL } else { REL },
                    escape_attribute(&id)
                ));
                if let Some(action) = action {
                    out.push_str(&format!(" action=\"{}\"", escape_attribute(action)));
                }
                if let Some(tooltip) = link.tooltip.as_deref().filter(|tip| !tip.is_empty()) {
                    out.push_str(&format!(" tooltip=\"{}\"", escape_attribute(tooltip)));
                }
                out.push_str("/>");
                Some(out.into_bytes())
            },
        };

        let mut edits: Vec<(Range<usize>, Vec<u8>)> = Vec::new();
        for element in targets {
            let tag = &fragment[element.tag.clone()];
            match (&click, element.empty) {
                (Some(click), true) => {
                    edits.push((element.tag.clone(), expand(tag, element, click)?));
                },
                (Some(click), false) => match element.clicks.split_first() {
                    // Replace the first click in place, drop any others.
                    Some((first, rest)) => {
                        edits.push((first.clone(), click.clone()));
                        for span in rest {
                            edits.push((span.clone(), Vec::new()));
                        }
                    },
                    None => {
                        let at = element
                            .insertion()
                            .ok_or_else(|| invalid("an element has no end tag"))?;
                        edits.push((at..at, click.clone()));
                    },
                },
                (None, _) => {
                    for span in &element.clicks {
                        edits.push((span.clone(), Vec::new()));
                    }
                },
            }
        }
        if let Some(click) = &click {
            for (at, prefix) in bare {
                let name = if prefix.is_empty() {
                    "rPr".to_owned()
                } else {
                    format!("{}:rPr", String::from_utf8_lossy(prefix))
                };
                let mut content = format!("<{name}>").into_bytes();
                content.extend_from_slice(click);
                content.extend_from_slice(format!("</{name}>").as_bytes());
                edits.push((at..at, content));
            }
        }
        if edits.is_empty() {
            return Ok(false);
        }
        let mut mapped = Vec::with_capacity(edits.len());
        for (range, bytes) in edits {
            mapped.push((map(range.start)?..map(range.end)?, bytes));
        }
        mapped.sort_by_key(|(range, _)| (range.start, range.end));
        if mapped
            .windows(2)
            .any(|pair| pair[0].0.end > pair[1].0.start)
        {
            return Err(invalid("hyperlink edits overlap"));
        }
        let mut out = Vec::with_capacity(xml.len() + 256);
        let mut cursor = 0;
        for (range, bytes) in mapped {
            out.extend_from_slice(&xml[cursor..range.start]);
            out.extend_from_slice(&bytes);
            cursor = range.end;
        }
        out.extend_from_slice(&xml[cursor..]);
        // Read the candidate back: the same shapes, the click in place.
        let staged = crate::shape::Scene::read(&out, shape_limits)?;
        let staged_shape = staged.shape(key)?;
        if staged_shape.id() != shape_id {
            return Err(invalid("the hyperlink edit moved the shape"));
        }
        let read = ShapeLinks::scan(&staged_shape.common().self_contained_xml()?)?;
        let present = match scope {
            Scope::Shape => read.shape_click.is_some(),
            Scope::Run { paragraph, run } => read
                .paragraphs
                .get(paragraph)
                .and_then(|runs| runs.get(run))
                .and_then(|slot| slot.properties.as_ref())
                .is_some_and(|properties| !properties.clicks.is_empty()),
        };
        if present != click.is_some() {
            return Err(invalid("the hyperlink edit did not round-trip"));
        }
        candidate.get_part_mut(&selected.part_name)?.set_blob(out);
        self.working = candidate;
        Ok(true)
    }
}

impl Edit {
    /// Resolve a click action of the slide part `slide` against the staged
    /// relationships and slide order.
    fn resolve_click(
        &self,
        slide: &litchi_opc::PackURI,
        click: &crate::hyperlinks::Click,
    ) -> Result<Option<Link>> {
        let rels = self.working.get_part(slide)?.rels();
        Link::resolve(
            click,
            |id| {
                let Some(relationship) = rels.get(id) else {
                    return Ok(None);
                };
                Ok(Some(if relationship.is_external() {
                    (true, relationship.target_ref().to_owned())
                } else {
                    (false, relationship.target_partname()?.to_string())
                }))
            },
            |part| {
                Ok(self
                    .slides
                    .iter()
                    .position(|slide| slide.part_name.as_str() == part)
                    .map(Position::new))
            },
        )
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use litchi_core::Position;

    use crate::hyperlinks::{Click, Link, Target};

    fn deck() -> crate::Package {
        let source = crate::Package::new().unwrap();
        let mut edit = source.edit().unwrap();
        for at in 0..2 {
            let slide = edit.insert_slide(Position::new(at), "Blank").unwrap();
            edit.add_text_box(slide.name(), "Docs", (0, 0, 914_400, 914_400))
                .unwrap();
        }
        edit.commit().unwrap().into_package()
    }

    fn shape_link(package: &crate::Package) -> Option<Link> {
        let presentation = package.presentation().unwrap();
        let slide = presentation.slide(0).unwrap().unwrap();
        let scene = slide.scene().unwrap();
        let handle = scene.shape(0_usize).unwrap().unwrap();
        let click = Click::of_shape(handle.view()).unwrap()?;
        slide.hyperlink(&click).unwrap()
    }

    #[test]
    fn shape_and_run_links_round_trip() {
        let source = deck();
        let url = Link {
            target: Target::Url("https://example.com/a?b&c".into()),
            tooltip: Some("Open \"docs\"".into()),
        };
        let mut edit = source.edit().unwrap();
        assert!(
            edit.set_shape_hyperlink(0_usize, 0_usize, Some(&url))
                .unwrap()
        );
        let package = edit.commit().unwrap().into_package();
        assert_eq!(shape_link(&package), Some(url.clone()));
        // Equal again: an exact no-op, even after a reopen.
        let reopened = crate::Package::from_bytes(&package.to_bytes().unwrap()).unwrap();
        let mut edit = reopened.edit().unwrap();
        assert!(
            !edit
                .set_shape_hyperlink(0_usize, 0_usize, Some(&url))
                .unwrap()
        );
        // A slide jump replaces it; a removal drops it.
        let jump = Link {
            target: Target::Slide(Position::new(1)),
            tooltip: None,
        };
        assert!(
            edit.set_shape_hyperlink(0_usize, 0_usize, Some(&jump))
                .unwrap()
        );
        let package = edit.commit().unwrap().into_package();
        assert_eq!(shape_link(&package), Some(jump));
        let mut edit = package.edit().unwrap();
        assert!(edit.set_shape_hyperlink(0_usize, 0_usize, None).unwrap());
        assert!(!edit.set_shape_hyperlink(0_usize, 0_usize, None).unwrap());
        let package = edit.commit().unwrap().into_package();
        assert_eq!(shape_link(&package), None);
        // A run action.
        let next = Link {
            target: Target::Action("ppaction://hlinkshowjump?jump=nextslide".into()),
            tooltip: None,
        };
        let mut edit = package.edit().unwrap();
        assert!(
            edit.set_run_hyperlink(
                0_usize,
                0_usize,
                Position::new(0),
                Position::new(0),
                Some(&next)
            )
            .unwrap()
        );
        assert!(
            edit.set_run_hyperlink(
                0_usize,
                0_usize,
                Position::new(0),
                Position::new(9),
                Some(&next)
            )
            .is_err()
        );
        let package = edit.commit().unwrap().into_package();
        let presentation = package.presentation().unwrap();
        let slide = presentation.slide(0).unwrap().unwrap();
        let scene = slide.scene().unwrap();
        let handle = scene.shape(0_usize).unwrap().unwrap();
        let body = handle.view().text_body().unwrap().unwrap();
        let run = &body.paragraphs[0].runs[0];
        let click = Click::from_text(run.properties.hyperlink.as_ref().unwrap());
        assert_eq!(slide.hyperlink(&click).unwrap(), Some(next));
    }
}
