//! Per-slide reads shared by the eager and source-backed PPTX owners.
//!
//! The adapter deliberately keeps this layer small.  It asks the leaf for a
//! detached scene or related-part snapshot and lets the node layer retain the
//! result.  No source slide or borrowed eager slide escapes this module.

use litchi_core::{Metadata, Position};
use litchi_pptx::master_layout::SlideLayoutKind;
use litchi_pptx::presentation::{modern_comments, related};
use litchi_pptx::shape::Shape;
use litchi_pptx::shape::owned::Scene;

use super::{Pptx, PptxOwner};
use crate::Result;
use crate::error::Error;
use crate::pptx;

/// Master and layout reads kept per handle (record 0280): the layout counts
/// and the strict scenes of masters and layouts. A query from the root
/// visits every master and layout, and the source-backed leaf re-reads the
/// presentation part and the related part for each of them, which cost
/// about a third of a full query on a 500-slide deck. A handle's content
/// never changes (a changed source fails the handle's freshness fence, and a
/// write returns a new handle with its own cache), so the values stay valid;
/// a failed read is not kept.
#[derive(Default)]
pub(super) struct Related {
    layout_counts: std::sync::Mutex<Vec<(usize, usize)>>,
    masters: std::sync::Mutex<Vec<(usize, Scene)>>,
    layouts: std::sync::Mutex<Vec<((usize, usize), Scene)>>,
}

/// The value cached under `key` in `cache`, else `read()`'s, kept when it
/// is present.
fn cached<K: PartialEq + Copy, V: Clone>(
    cache: &std::sync::Mutex<Vec<(K, V)>>,
    key: K,
    read: impl FnOnce() -> Result<Option<V>>,
) -> Result<Option<V>> {
    if let Ok(entries) = cache.lock()
        && let Some((_, value)) = entries.iter().find(|(at, _)| *at == key)
    {
        return Ok(Some(value.clone()));
    }
    let value = read()?;
    if let (Some(value), Ok(mut entries)) = (&value, cache.lock())
        && !entries.iter().any(|(at, _)| *at == key)
    {
        entries.push((key, value.clone()));
    }
    Ok(value)
}

/// The values a slide node exposes from the presentation and its layout.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct SlideFacts {
    /// `p:cSld@name`, when the slide has one.
    pub(crate) name: Option<String>,
    /// `p:sld@show="0"`.
    pub(crate) hidden: bool,
    /// The layout's `p:cSld@name`, when it has one.
    pub(crate) layout: Option<String>,
    /// The layout's `p:sldLayout@type` token, when it has one.
    pub(crate) layout_type: Option<String>,
}

impl Pptx {
    /// Bracket a source-backed read with freshness and execution fences.
    /// When a source changes while the read is in flight, that typed source
    /// error wins over the read's own result.
    pub(super) fn finish<T>(&self, read: impl FnOnce() -> Result<T>) -> Result<T> {
        self.check_owner()?;
        let result = read();
        self.check_owner()?;
        result
    }
}

impl Pptx {
    /// Acquire a slide from the effective native revision, including Parts
    /// restored by inverse replay after reopening.
    fn revised_slide(&self, position: usize) -> Result<Option<pptx::source::Slide>> {
        let Some(revision) = &self.0.revision else {
            return Ok(None);
        };
        let presentation = revision.presentation()?;
        if position >= presentation.slide_count()? {
            return Ok(None);
        }
        Ok(Some(presentation.slide(position)?))
    }

    /// The number of slides in the validated presentation catalog.
    pub(super) fn slides_in_catalog(&self) -> Result<usize> {
        self.slide_count()
    }

    /// The detached shape scene of the selected slide, or `None` past the
    /// catalog end. Source parser and scene reservations remain governed by
    /// the native source/execution limits; the facade snapshot ceiling is
    /// reserved for the retained native patch wire proof.
    pub(super) fn slide_scene(&self, position: usize) -> Result<Option<Scene>> {
        self.finish(|| {
            if let Some(revision) = &self.0.revision {
                return self.finish(|| {
                    let presentation = revision.presentation()?;
                    if position >= presentation.slide_count()? {
                        return Ok(None);
                    }
                    strict_scene(presentation.slide(position)?.scene()?).map(Some)
                });
            }
            match &self.0.owner {
                PptxOwner::Source(presentation) => {
                    self.check_owner()?;
                    let Some(slide) = presentation.slide(position) else {
                        return Ok(None);
                    };
                    strict_scene(slide.scene()?).map(Some)
                },
                PptxOwner::Eager(package) => {
                    let Some(part) = super::eager_slide_part(self, package, position)? else {
                        return Ok(None);
                    };
                    let payload = part.payload.clone();
                    // A cached scene was read from this very payload, so the
                    // part is read as a slide only on a miss (record 0326).
                    if let Some(scene) = payload.as_ref().and_then(|key| self.0.scenes.get(key)) {
                        return Ok(Some(scene));
                    }
                    let scene = strict_scene(part.slide()?.scene()?)?;
                    if let Some(payload) = payload {
                        self.0.scenes.insert(payload, scene.clone());
                    }
                    Ok(Some(scene))
                },
            }
        })
    }

    /// All slide text for one selected slide, or `None` past the catalog end.
    pub(super) fn slide_text(&self, position: usize) -> Result<Option<String>> {
        Ok(self.slide(position)?.map(|slide| slide.into_parts().0))
    }

    /// The number of slide masters in the owner catalog.
    pub(super) fn masters_in_catalog(&self) -> Result<usize> {
        self.finish(|| {
            if let Some(&count) = self.0.masters.get() {
                return Ok(count);
            }
            let count = match &self.0.owner {
                PptxOwner::Source(presentation) => {
                    self.check_owner()?;
                    presentation.masters()?.len()
                },
                PptxOwner::Eager(package) => package.presentation()?.slide_masters()?.len(),
            };
            Ok(*self.0.masters.get_or_init(|| count))
        })
    }

    /// The number of layouts owned by one slide master.
    pub(super) fn layouts_in_catalog(&self, master: usize) -> Result<usize> {
        self.finish(|| {
            cached(&self.0.related.layout_counts, master, || {
                self.read_layouts_in_catalog(master).map(Some)
            })
            .map(Option::unwrap_or_default)
        })
    }

    fn read_layouts_in_catalog(&self, master: usize) -> Result<usize> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                Ok(presentation
                    .layouts(Position::new(master))?
                    .map_or(0, |layouts| layouts.len()))
            },
            PptxOwner::Eager(package) => {
                let presentation = package.presentation()?;
                let Some(master) = presentation.slide_masters()?.into_iter().nth(master) else {
                    return Ok(0);
                };
                Ok(master.layouts()?.len())
            },
        })
    }

    /// The detached shape scene of one slide master.
    pub(super) fn master_scene(&self, position: usize) -> Result<Option<Scene>> {
        self.finish(|| {
            cached(&self.0.related.masters, position, || {
                self.read_master_scene(position)
            })
        })
    }

    fn read_master_scene(&self, position: usize) -> Result<Option<Scene>> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                let Some(master) = presentation.master(Position::new(position))? else {
                    return Ok(None);
                };
                strict_scene(master.scene().clone()).map(Some)
            },
            PptxOwner::Eager(package) => {
                let presentation = package.presentation()?;
                let Some(master) = presentation.slide_masters()?.into_iter().nth(position) else {
                    return Ok(None);
                };
                strict_scene(master.scene()?).map(Some)
            },
        })
    }

    /// The name and preservation switch of one slide master.
    pub(super) fn master_facts(&self, position: usize) -> Result<Option<super::node::MasterFacts>> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                let Some(master) = presentation.master(Position::new(position))? else {
                    return Ok(None);
                };
                Ok(Some(super::node::MasterFacts {
                    name: master.name().map(str::to_owned),
                    preserved: master.is_preserved(),
                }))
            },
            PptxOwner::Eager(package) => {
                let presentation = package.presentation()?;
                let Some(master) = presentation.slide_masters()?.into_iter().nth(position) else {
                    return Ok(None);
                };
                Ok(Some(super::node::MasterFacts {
                    name: master.part().authored_name()?,
                    preserved: master.is_preserved()?,
                }))
            },
        })
    }

    /// The detached shape scene of one layout owned by one master.
    pub(super) fn layout_scene(
        &self,
        master_position: usize,
        layout_position: usize,
    ) -> Result<Option<Scene>> {
        self.finish(|| {
            cached(
                &self.0.related.layouts,
                (master_position, layout_position),
                || self.read_layout_scene(master_position, layout_position),
            )
        })
    }

    fn read_layout_scene(
        &self,
        master_position: usize,
        layout_position: usize,
    ) -> Result<Option<Scene>> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                let Some(layout) = presentation.layout(
                    Position::new(master_position),
                    Position::new(layout_position),
                )?
                else {
                    return Ok(None);
                };
                strict_scene(layout.scene().clone()).map(Some)
            },
            PptxOwner::Eager(package) => {
                let presentation = package.presentation()?;
                let Some(master) = presentation
                    .slide_masters()?
                    .into_iter()
                    .nth(master_position)
                else {
                    return Ok(None);
                };
                let Some(layout) = master.layouts()?.into_iter().nth(layout_position) else {
                    return Ok(None);
                };
                strict_scene(layout.scene()?).map(Some)
            },
        })
    }

    /// The authored name and typed kind of one layout.
    pub(super) fn layout_facts(
        &self,
        master_position: usize,
        layout_position: usize,
    ) -> Result<Option<super::node::LayoutFacts>> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                let Some(layout) = presentation.layout(
                    Position::new(master_position),
                    Position::new(layout_position),
                )?
                else {
                    return Ok(None);
                };
                Ok(Some(super::node::LayoutFacts {
                    name: layout.name().map(str::to_owned),
                    kind: layout.kind(),
                }))
            },
            PptxOwner::Eager(package) => {
                let presentation = package.presentation()?;
                let Some(master) = presentation
                    .slide_masters()?
                    .into_iter()
                    .nth(master_position)
                else {
                    return Ok(None);
                };
                let Some(layout) = master.layouts()?.into_iter().nth(layout_position) else {
                    return Ok(None);
                };
                let kind = layout
                    .kind()?
                    .map(|token| {
                        SlideLayoutKind::from_token(&token).ok_or_else(|| {
                            Error::from(pptx::Error::Invalid(
                                "slide layout has an unsupported type token".into(),
                            ))
                        })
                    })
                    .transpose()?;
                Ok(Some(super::node::LayoutFacts {
                    name: layout.part().authored_name()?,
                    kind,
                }))
            },
        })
    }

    /// The detached semantic facts needed by a slide node.
    pub(super) fn slide_facts(&self, position: usize) -> Result<Option<SlideFacts>> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                if self.0.revision.is_some() {
                    let Some(slide) = self.revised_slide(position)? else {
                        return Ok(None);
                    };
                    let name = slide.authored_name()?;
                    let hidden = slide.is_hidden()?;
                    let (layout, layout_type) = match slide.layout()? {
                        Some(layout) => (
                            layout.name().map(str::to_owned),
                            layout.kind().map(|kind| kind.as_str().to_owned()),
                        ),
                        None => (None, None),
                    };
                    return Ok(Some(SlideFacts {
                        name,
                        hidden,
                        layout,
                        layout_type,
                    }));
                }
                let Some(slide) = presentation.slide(position) else {
                    return Ok(None);
                };
                let name = slide.authored_name()?;
                let hidden = slide.is_hidden()?;
                let (layout, layout_type) = match slide.layout()? {
                    Some(layout) => (
                        layout.name().map(str::to_owned),
                        layout.kind().map(|kind| kind.as_str().to_owned()),
                    ),
                    None => (None, None),
                };
                Ok(Some(SlideFacts {
                    name,
                    hidden,
                    layout,
                    layout_type,
                }))
            },
            PptxOwner::Eager(package) => {
                let Some(slide) = super::eager_slide(self, package, position)? else {
                    return Ok(None);
                };
                let name = slide.part().authored_name()?;
                let hidden = slide.is_hidden()?;
                let (layout, layout_type) = match slide.layout()? {
                    Some(layout) => (layout.part().authored_name()?, layout.kind()?),
                    None => (None, None),
                };
                Ok(Some(SlideFacts {
                    name,
                    hidden,
                    layout,
                    layout_type,
                }))
            },
        })
    }

    /// The speaker-notes text of the selected slide, when the notes part has
    /// a primary text body.
    pub(super) fn slide_notes(&self, position: usize) -> Result<Option<String>> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                if self.0.revision.is_some() {
                    let Some(slide) = self.revised_slide(position)? else {
                        return Ok(None);
                    };
                    return Ok(slide
                        .notes()?
                        .and_then(|notes| notes.text().map(str::to_owned)));
                }
                let Some(slide) = presentation.slide(position) else {
                    return Ok(None);
                };
                Ok(slide
                    .notes()?
                    .and_then(|notes| notes.text().map(str::to_owned)))
            },
            PptxOwner::Eager(package) => {
                let Some(slide) = super::eager_slide(self, package, position)? else {
                    return Ok(None);
                };
                Ok(slide
                    .notes()?
                    .and_then(|notes| notes.text().map(str::to_owned)))
            },
        })
    }

    /// The notes page of the selected slide as the navigator sees it
    /// (record 0348): its shape scene, and whether it has notes text, which
    /// [`Self::slide_notes`] requires before `/slide[N]/notes` resolves —
    /// read once.
    pub(super) fn slide_notes_view(&self, position: usize) -> Result<Option<(bool, Scene)>> {
        let view = |notes: Option<related::Notes>| {
            notes.map(|notes| (notes.text().is_some(), notes.scene().clone()))
        };
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                if self.0.revision.is_some() {
                    let Some(slide) = self.revised_slide(position)? else {
                        return Ok(None);
                    };
                    return Ok(view(slide.notes()?));
                }
                let Some(slide) = presentation.slide(position) else {
                    return Ok(None);
                };
                Ok(view(slide.notes()?))
            },
            PptxOwner::Eager(package) => {
                let Some(slide) = super::eager_slide(self, package, position)? else {
                    return Ok(None);
                };
                Ok(view(slide.notes()?))
            },
        })
    }

    /// The shape scene of the selected slide's notes page, when it has one
    /// (record 0285: find reads its shapes' text bodies).
    pub(super) fn slide_notes_scene(&self, position: usize) -> Result<Option<Scene>> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                if self.0.revision.is_some() {
                    let Some(slide) = self.revised_slide(position)? else {
                        return Ok(None);
                    };
                    return Ok(slide.notes()?.map(|notes| notes.scene().clone()));
                }
                let Some(slide) = presentation.slide(position) else {
                    return Ok(None);
                };
                Ok(slide.notes()?.map(|notes| notes.scene().clone()))
            },
            PptxOwner::Eager(package) => {
                let Some(slide) = super::eager_slide(self, package, position)? else {
                    return Ok(None);
                };
                Ok(slide.notes()?.map(|notes| notes.scene().clone()))
            },
        })
    }

    /// The chart-space XML a slide relates through `relationship_id` (a
    /// chart frame's payload), or `None` when the slide has no such chart
    /// (record 0270). Only that chart part is read.
    pub(super) fn slide_chart_xml(
        &self,
        position: usize,
        relationship_id: &str,
    ) -> Result<Option<Vec<u8>>> {
        self.finish(|| {
            if self.0.revision.is_some() {
                return self
                    .revised_slide(position)?
                    .map(|slide| slide.chart_xml(relationship_id))
                    .transpose()
                    .map(Option::flatten)
                    .map_err(Error::from);
            }
            match &self.0.owner {
                PptxOwner::Source(presentation) => {
                    self.check_owner()?;
                    let Some(slide) = presentation.slide(position) else {
                        return Ok(None);
                    };
                    Ok(slide.chart_xml(relationship_id)?)
                },
                PptxOwner::Eager(package) => {
                    let Some(slide) = super::eager_slide(self, package, position)? else {
                        return Ok(None);
                    };
                    Ok(slide.chart_xml(relationship_id)?.map(<[u8]>::to_vec))
                },
            }
        })
    }

    /// The data-model XML a slide relates through `relationship_id` (a
    /// SmartArt frame's `r:dm`), or `None` when the slide has no such
    /// diagram (record 0284). Only that data part is read.
    pub(super) fn slide_diagram_xml(
        &self,
        position: usize,
        relationship_id: &str,
    ) -> Result<Option<Vec<u8>>> {
        self.finish(|| {
            if self.0.revision.is_some() {
                return self
                    .revised_slide(position)?
                    .map(|slide| slide.diagram_xml(relationship_id))
                    .transpose()
                    .map(Option::flatten)
                    .map_err(Error::from);
            }
            match &self.0.owner {
                PptxOwner::Source(presentation) => {
                    self.check_owner()?;
                    let Some(slide) = presentation.slide(position) else {
                        return Ok(None);
                    };
                    Ok(slide.diagram_xml(relationship_id)?)
                },
                PptxOwner::Eager(package) => {
                    let Some(slide) = super::eager_slide(self, package, position)? else {
                        return Ok(None);
                    };
                    Ok(slide.diagram_xml(relationship_id)?.map(<[u8]>::to_vec))
                },
            }
        })
    }

    /// The detached legacy-comments snapshot of the selected slide.
    pub(super) fn slide_comments(&self, position: usize) -> Result<Option<related::Comments>> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                if self.0.revision.is_some() {
                    return self
                        .revised_slide(position)?
                        .map(|slide| slide.comments())
                        .transpose()
                        .map(Option::flatten)
                        .map_err(Error::from);
                }
                let Some(slide) = presentation.slide(position) else {
                    return Ok(None);
                };
                Ok(slide.comments()?)
            },
            PptxOwner::Eager(package) => {
                let Some(slide) = super::eager_slide(self, package, position)? else {
                    return Ok(None);
                };
                Ok(slide.comment_snapshot()?)
            },
        })
    }

    /// The detached modern-comments snapshot of the selected slide.
    pub(super) fn slide_modern_comments(
        &self,
        position: usize,
    ) -> Result<Option<modern_comments::List>> {
        self.finish(|| match &self.0.owner {
            PptxOwner::Source(presentation) => {
                self.check_owner()?;
                if self.0.revision.is_some() {
                    return self
                        .revised_slide(position)?
                        .map(|slide| slide.modern_comments())
                        .transpose()
                        .map(Option::flatten)
                        .map_err(Error::from);
                }
                let Some(slide) = presentation.slide(position) else {
                    return Ok(None);
                };
                Ok(slide.modern_comments()?)
            },
            PptxOwner::Eager(package) => {
                let Some(slide) = super::eager_slide(self, package, position)? else {
                    return Ok(None);
                };
                Ok(slide.modern_comments()?)
            },
        })
    }

    /// Borrow the cached core-properties metadata after enforcing the source
    /// freshness fence.
    pub(super) fn cached_metadata(&self) -> Result<Option<&Metadata>> {
        self.finish(|| {
            self.check_owner()?;
            if self.0.metadata.get().is_none() {
                let PptxOwner::Source(presentation) = &self.0.owner else {
                    // The eager owner fills this cell at open.
                    return Ok(None);
                };
                let metadata = presentation
                    .properties_with_tolerance(&mut pptx::ToleranceReport::default())?
                    .unwrap_or_default();
                let metadata = crate::format::metadata::with_source_app_props(
                    metadata,
                    presentation.source_backed_package(),
                );
                let metadata = Some(metadata).filter(Metadata::has_data);
                let _ = self.0.metadata.set(metadata);
            }
            Ok(self.0.metadata.get().and_then(Option::as_ref))
        })
    }
}

fn strict_scene(scene: Scene) -> Result<Scene> {
    if !scene.tolerance_report().is_clean() {
        Err(Error::from(pptx::Error::Invalid(
            "slide scene requires strict shape XML".into(),
        )))
    } else if scene.views().any(|shape| {
        matches!(
            shape,
            Shape::Frame(_) | Shape::Content(_) | Shape::Unknown(_)
        )
    }) {
        Err(Error::new(
            crate::error::Code::UnsupportedCapability,
            "slide scene contains a shape kind without a unified projection",
        )
        .with_format(litchi_core::FileFormat::Pptx))
    } else {
        Ok(scene)
    }
}
