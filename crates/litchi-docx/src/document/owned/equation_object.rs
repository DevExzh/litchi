#![cfg_attr(
    not(feature = "equation"),
    allow(
        dead_code,
        reason = "equation objects are written with the equation feature"
    )
)]
//! OLE equation objects (MathType and Equation Editor 3.0) staged with a
//! main-document edit.
//!
//! A new object is a MathType one (`Equation.DSMT4`, MTEF 5) unless raw
//! MTEF 3 is given, which makes an Equation Editor 3.0 one (`Equation.3`).
//! It becomes two new parts related from the main part, the embedded
//! compound file (`embeddings/oleObjectN.bin`, an `oleObject` relationship;
//! ECMA-376 Part 1 §15.2.13) and a PNG preview (`media/imageN.png`, an
//! `image` relationship), and a new run holding the `w:object` markup Word
//! writes for an equation ([MS-OI29500] 2.1.1641): the VML picture shape
//! showing the preview and the `o:OLEObject` naming the compound file. Word
//! and PowerPoint write EMF or WMF previews; the preview is an ordinary
//! image part to them, which the equation server redraws when the object is
//! activated, so a PNG serves the same purpose.
//!
//! The preview is drawn at the text size of the run the object stands in
//! (resolved through the run's and the paragraph's styles and the document
//! defaults, as Word resolves it), shown at its size in points, and the run
//! is lowered by the depth of the equation below its baseline
//! (`w:position`), as Word lowers an inline equation.
//!
//! Objects of a header, a footer or a note are edited the same way; their
//! parts are related from the story's own part.
//!
//! A replacement keeps the object's kind (an MTEF 3 object stays MTEF 3),
//! writes the new `Equation Native` stream into the object's compound file
//! (dropping the cached presentations that would show the old equation),
//! draws a new preview and resizes the markup. A compound file or PNG
//! preview only this object uses is replaced in place; otherwise, and for an
//! EMF or WMF preview, a new part is staged and the old one goes with its
//! relationship once nothing names it. A removal drops the object's run (or
//! its `w:object`, in a run that holds other content), and the embedding and
//! preview parts with their relationships once nothing else names them.

use std::collections::BTreeSet;
use std::sync::Arc;

use litchi_core::Position;
use litchi_opc::PackURI;

use super::super::raw::Tree;
use super::super::transaction::{FontHint, GraphChange, PartRelationship, object_numbers};
use super::super::{ElementPath, Error, Refusal, Result};
use super::{Edit, StoryTarget};
use crate::Package;
use crate::document::{TransferPart, TransferRelationship};
use crate::object::EquationSpan;

const OLE_OBJECT_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
const IMAGE_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const OLE_OBJECT_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.oleObject";
/// The first VML shape number Word gives an object (`_x0000_i1025`).
const FIRST_SHAPE: u32 = 1025;
/// The object number a new object follows when the document has none.
const FIRST_OBJECT: u64 = 1_000_000_000;

fn refused(reason: Refusal) -> Error {
    Error::Refused {
        position: 0,
        reason,
    }
}

/// Where the content of a new or replaced equation object comes from.
#[cfg(feature = "equation")]
#[derive(Debug, Clone, Copy)]
#[non_exhaustive]
pub enum EquationSource<'a> {
    /// LaTeX math, converted to MTEF.
    Latex(&'a str),
    /// An OMML equation (`m:oMath`, `m:oMathPara` or a fragment of them),
    /// converted to MTEF.
    Omml(&'a str),
    /// Raw MTEF: a complete `Equation Native` stream, or bare MTEF to which
    /// the 28-byte OLE header is added.
    Mtef(&'a [u8]),
}

/// The content of a new or replaced equation object: its `Equation Native`
/// stream, its kind, its preview picture and what the conversion lost.
///
/// ```
/// use litchi_docx::document::{EquationContent, EquationSource};
/// use litchi_equation::Limits;
///
/// let content =
///     EquationContent::new(EquationSource::Latex(r"\frac{a}{b}"), None, 12.0, &Limits::DEFAULT)?;
/// assert_eq!(content.prog_id(), "Equation.DSMT4");
/// assert!(content.losses().is_empty());
/// assert!(content.preview_png().starts_with(b"\x89PNG"));
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[cfg(feature = "equation")]
#[derive(Debug, Clone)]
pub struct EquationContent {
    native: Vec<u8>,
    target: litchi_equation::MtefTarget,
    preview: Vec<u8>,
    size_pt: (f64, f64),
    lowering: u32,
    losses: Vec<String>,
}

/// Pixels per point of a preview, relative to its shown size.
#[cfg(feature = "equation")]
const PREVIEW_DENSITY: f64 = 2.0;
/// The smallest scale a preview is drawn at to fit the pixel ceiling.
#[cfg(feature = "equation")]
const MIN_PREVIEW_SCALE: f64 = 1.0 / 64.0;
/// The `EQNOLEFILEHDR` length.
#[cfg(feature = "equation")]
const OLE_HEADER_LEN: usize = 28;

#[cfg(feature = "equation")]
impl EquationContent {
    /// Convert `source` to an `Equation Native` stream of `kind` (a new
    /// MathType object for `None`; an existing object's kind on a
    /// replacement) and draw its preview for text of `font_size_pt` points
    /// (see [`Edit::equation_object_font_size_pt`]).
    ///
    /// Raw MTEF keeps its own kind (MTEF 3: `Equation.3`, MTEF 5:
    /// `Equation.DSMT4`); with `kind`, its version must be the kind's.
    ///
    /// # Errors
    ///
    /// Returns the conversion error for input that does not parse or that
    /// exceeds `limits`, [`litchi_equation::Error::Object`] for raw MTEF of
    /// another version than `kind`'s, and a rendering error.
    pub fn new(
        source: EquationSource<'_>,
        kind: Option<&litchi_equation::MtefTarget>,
        font_size_pt: f64,
        limits: &litchi_equation::Limits,
    ) -> litchi_equation::Result<Self> {
        let (native, target, losses) = equation_native(source, kind, limits)?;
        let options =
            litchi_equation::PngOptions::new(font_size_pt)?.with_scale(PREVIEW_DENSITY)?;
        let png = render_fitting(&native, options, limits)?;
        let size_pt = png.size_pt();
        #[expect(
            clippy::cast_possible_truncation,
            clippy::cast_sign_loss,
            reason = "the value is clamped to the w:position range Word writes first"
        )]
        let lowering = (png.depth_pt() * 2.0).round().clamp(0.0, 3168.0) as u32;
        Ok(Self {
            native,
            target,
            preview: png.into_bytes(),
            size_pt,
            lowering,
            losses,
        })
    }

    /// The `Equation Native` stream.
    #[must_use]
    pub fn equation_native(&self) -> &[u8] {
        &self.native
    }

    /// The ProgID of the object kind: `Equation.DSMT4` or `Equation.3`.
    #[must_use]
    pub fn prog_id(&self) -> &'static str {
        self.target.prog_id()
    }

    /// The preview picture, a PNG.
    #[must_use]
    pub fn preview_png(&self) -> &[u8] {
        &self.preview
    }

    /// The shown size of the preview, in points.
    #[must_use]
    pub const fn size_pt(&self) -> (f64, f64) {
        self.size_pt
    }

    /// One message per piece of the source MTEF cannot hold; empty when
    /// nothing was lost (always empty for raw MTEF).
    #[must_use]
    pub fn losses(&self) -> &[String] {
        &self.losses
    }
}

/// The preview of `native` at `options`, drawn at a smaller scale when the
/// picture would exceed the pixel ceiling.
#[cfg(feature = "equation")]
fn render_fitting(
    native: &[u8],
    mut options: litchi_equation::PngOptions,
    limits: &litchi_equation::Limits,
) -> litchi_equation::Result<litchi_equation::Png> {
    loop {
        match litchi_equation::mtef_to_png(native, &options, limits) {
            Err(litchi_equation::Error::Render(_)) if options.scale() > MIN_PREVIEW_SCALE => {
                options = options.with_scale((options.scale() / 2.0).max(MIN_PREVIEW_SCALE))?;
            },
            result => return result,
        }
    }
}

/// The MTEF version an object of `target`'s kind holds: 3 or 5.
#[cfg(feature = "equation")]
fn version_of(target: &litchi_equation::MtefTarget) -> u8 {
    if target.prog_id() == "Equation.3" {
        3
    } else {
        5
    }
}

/// The `Equation Native` stream of `source`, the kind it is written as, and
/// the conversion losses.
#[cfg(feature = "equation")]
fn equation_native(
    source: EquationSource<'_>,
    kind: Option<&litchi_equation::MtefTarget>,
    limits: &litchi_equation::Limits,
) -> litchi_equation::Result<(Vec<u8>, litchi_equation::MtefTarget, Vec<String>)> {
    use litchi_equation::MtefTarget;
    let expected = kind.map(version_of);
    let target = kind.cloned().unwrap_or(MtefTarget::MathType);
    match source {
        EquationSource::Latex(latex) => {
            let native = litchi_equation::latex_to_equation_native(latex, &target, limits)?;
            let losses = native.losses().to_vec();
            Ok((native.into_bytes(), target, losses))
        },
        EquationSource::Omml(omml) => {
            let native = litchi_equation::omml_to_equation_native(omml, &target, limits)?;
            let losses = native.losses().to_vec();
            Ok((native.into_bytes(), target, losses))
        },
        EquationSource::Mtef(bytes) => {
            if bytes.len() as u64 > litchi_equation::MAX_EQUATION_NATIVE_BYTES {
                return Err(litchi_equation::Error::Object(
                    "the MTEF exceeds the Equation Native size limit".to_owned(),
                ));
            }
            let read = litchi_equation::read_mtef(bytes, limits)?;
            let info = read.info().cloned().ok_or_else(|| {
                litchi_equation::Error::Object("the MTEF has no header".to_owned())
            })?;
            let own = MtefTarget::SameAs(info.clone());
            if expected.is_some_and(|expected| expected != version_of(&own)) {
                return Err(litchi_equation::Error::Object(
                    "the MTEF version differs from the equation object's".to_owned(),
                ));
            }
            let native = if info.ole_header.is_some() {
                bytes.to_vec()
            } else {
                with_ole_header(bytes)?
            };
            Ok((native, own, Vec::new()))
        },
    }
}

/// Bare MTEF behind an `EQNOLEFILEHDR` (`cbHdr` 28, version 0x00020000, the
/// MathType EF clipboard format, the MTEF length and four reserved zero
/// fields), as the equation servers write it.
#[cfg(feature = "equation")]
fn with_ole_header(mtef: &[u8]) -> litchi_equation::Result<Vec<u8>> {
    let len = u32::try_from(mtef.len())
        .map_err(|_| litchi_equation::Error::Object("the MTEF exceeds 4 GiB".to_owned()))?;
    let mut out = Vec::with_capacity(OLE_HEADER_LEN + mtef.len());
    out.extend_from_slice(&28u16.to_le_bytes());
    out.extend_from_slice(&0x0002_0000u32.to_le_bytes());
    out.extend_from_slice(&0xC1C6u16.to_le_bytes());
    out.extend_from_slice(&len.to_le_bytes());
    out.extend_from_slice(&[0; 16]);
    out.extend_from_slice(mtef);
    Ok(out)
}

/// The story whose part an equation-object edit changes: the main
/// document, or a header, footer or note story with the part that holds
/// it, whose relationships its objects use.
#[derive(Clone)]
enum Owner<'a> {
    Main,
    Story {
        target: StoryTarget<'a>,
        part: String,
    },
}

/// The text size Word gives a run that no style or property sizes, in
/// half-points (ECMA-376 Part 1, §17.3.2.38).
const DEFAULT_SIZE_HALF_POINTS: u32 = 20;
/// The deepest `w:basedOn` chain a size is looked up through.
const MAX_STYLE_DEPTH: usize = 32;

impl Edit<Package> {
    /// The owner of `story`'s objects.
    fn owner<'a>(&self, story: Option<StoryTarget<'a>>) -> Result<Owner<'a>> {
        Ok(match story {
            None => Owner::Main,
            Some(target @ StoryTarget::Part(name)) => {
                if self.part_payload(name)?.is_none() {
                    return Err(refused(Refusal::BlockNotFound));
                }
                Owner::Story {
                    target,
                    part: name.to_owned(),
                }
            },
            Some(target @ StoryTarget::Note(kind, _)) => Owner::Story {
                target,
                part: self.notes_part(kind)?.0,
            },
        })
    }

    /// Run `edit` on the owner's transaction: the main document's, or a
    /// story's, whose edited content replaces the story's.
    fn on_owner<T>(
        &mut self,
        owner: &Owner<'_>,
        edit: impl FnOnce(&mut Tree) -> Result<T>,
    ) -> Result<T> {
        match owner {
            Owner::Main => edit(self.inner_mut()?),
            Owner::Story { target, .. } => self.edit_story(*target, edit),
        }
    }

    /// The owner part's XML as this edit leaves it.
    fn owner_xml(&mut self, owner: &Owner<'_>) -> Result<Vec<u8>> {
        match owner {
            Owner::Main => Ok(self.inner_mut()?.main_xml().to_vec()),
            Owner::Story { part, .. } => Ok(self
                .part_payload(part)?
                .ok_or_else(|| crate::Error::PartNotFound(part.clone()))?
                .to_vec()),
        }
    }

    /// The owner part's relationships as this edit leaves them: (id, type,
    /// target part name for an internal one).
    fn owner_relationships(
        &self,
        owner: &Owner<'_>,
    ) -> Result<Vec<(String, String, Option<String>)>> {
        let package = self.package.opc_package();
        let main = package.main_document_part().map_err(crate::Error::from)?;
        let (uri, existing) = match owner {
            Owner::Main => (main.partname().clone(), Some(main.rels())),
            Owner::Story { part, .. } => {
                let uri = PackURI::new(part.clone()).map_err(crate::Error::Invalid)?;
                let rels = package.get_part(&uri).ok().map(|found| found.rels());
                (uri, rels)
            },
        };
        let mut found = Vec::new();
        if let Some(relationships) = existing {
            for relationship in relationships.iter() {
                let target = (!relationship.is_external())
                    .then(|| relationship.target_partname().ok())
                    .flatten()
                    .map(|name| name.as_str().to_owned());
                found.push((
                    relationship.r_id().to_owned(),
                    relationship.reltype().to_owned(),
                    target,
                ));
            }
        }
        let base = uri.base_uri().to_owned();
        let main_owner = matches!(owner, Owner::Main);
        let staged = self.inner_ref(|inner| {
            let mut staged: Vec<(String, String, Option<String>, bool)> = Vec::new();
            for (change, forward) in inner.staged_graph_changes() {
                let owned = change.part_relationships.iter().filter_map(|owned| {
                    (owned.part == uri.as_str()).then_some(&owned.relationship)
                });
                let relationships: Vec<&TransferRelationship> = if main_owner {
                    change.relationships.iter().collect()
                } else {
                    owned.collect()
                };
                for relationship in relationships {
                    let target = (!relationship.external)
                        .then(|| PackURI::from_rel_ref(&base, &relationship.target).ok())
                        .flatten()
                        .map(|name| name.as_str().to_owned());
                    staged.push((
                        relationship.id.clone(),
                        relationship.relationship_type.clone(),
                        target,
                        forward,
                    ));
                }
            }
            staged
        })?;
        for (id, kind, target, forward) in staged {
            found.retain(|(existing, _, _)| *existing != id);
            if forward {
                found.push((id, kind, target));
            }
        }
        Ok(found)
    }

    /// The internal part the owner relates through `id`.
    fn owner_target(&self, owner: &Owner<'_>, id: &str) -> Result<Option<String>> {
        Ok(self
            .owner_relationships(owner)?
            .into_iter()
            .find(|(existing, _, _)| existing == id)
            .and_then(|(_, _, target)| target))
    }

    /// Stage a new part related from the owner part and return its
    /// relationship identifier.
    fn add_owner_part(
        &mut self,
        owner: &Owner<'_>,
        (relationship_type, directory, stem, extension): (&str, &str, &str, &str),
        content_type: &str,
        payload: Vec<u8>,
    ) -> Result<String> {
        let part = match owner {
            Owner::Main => {
                return Ok(self
                    .add_related_part(
                        relationship_type,
                        directory,
                        stem,
                        extension,
                        content_type,
                        payload,
                    )?
                    .1);
            },
            Owner::Story { part, .. } => part.clone(),
        };
        let folder = part[..=part.rfind('/').unwrap_or(0)].to_owned();
        let used = self.part_names()?;
        let mut index = 1usize;
        let (name, target) = loop {
            let target = format!("{directory}{stem}{index}{extension}");
            let name = format!("{folder}{target}");
            if !used.contains(&name) {
                break (name, target);
            }
            index += 1;
        };
        PackURI::new(name.clone()).map_err(crate::Error::Invalid)?;
        let ids: BTreeSet<String> = self
            .owner_relationships(owner)?
            .into_iter()
            .map(|(id, _, _)| id)
            .collect();
        let mut number = ids.len() + 1;
        let id = loop {
            let candidate = format!("rId{number}");
            if !ids.contains(&candidate) {
                break candidate;
            }
            number += 1;
        };
        self.inner_mut()?.stage_graph_change(
            Arc::new(GraphChange {
                relationships: Arc::new([]),
                parts: Arc::new([TransferPart {
                    name,
                    content_type: content_type.to_owned(),
                    blob: Arc::new(payload),
                    relationships: Arc::new([]),
                }]),
                replaced: Arc::new([]),
                part_relationships: Arc::new([PartRelationship {
                    part,
                    relationship: TransferRelationship {
                        id: id.clone(),
                        relationship_type: relationship_type.to_owned(),
                        target,
                        external: false,
                    },
                }]),
            }),
            true,
        )?;
        Ok(id)
    }

    /// The owner relationships an attribute of the owner part names.
    fn owner_referenced(&mut self, owner: &Owner<'_>) -> Result<BTreeSet<String>> {
        if matches!(owner, Owner::Main) {
            return self.referenced_main_relationships();
        }
        let ids: BTreeSet<String> = self
            .owner_relationships(owner)?
            .into_iter()
            .map(|(id, _, _)| id)
            .collect();
        let xml = self.owner_xml(owner)?;
        let mut found = BTreeSet::new();
        let mut reader = quick_xml::Reader::from_reader(xml.as_slice());
        loop {
            match reader
                .read_event()
                .map_err(|_| crate::Error::Xml("invalid Word XML".into()))?
            {
                quick_xml::events::Event::Start(element)
                | quick_xml::events::Event::Empty(element) => {
                    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
                    for attribute in element.unchecked_attributes().flatten() {
                        if let Ok(value) = std::str::from_utf8(attribute.value.as_ref())
                            && ids.contains(value)
                        {
                            found.insert(value.to_owned());
                        }
                    }
                },
                quick_xml::events::Event::Eof => return Ok(found),
                _ => {},
            }
        }
    }

    /// Stage the removal of the owner relationships in `before` the owner
    /// part no longer names, with the parts only they reach.
    fn drop_released(&mut self, owner: &Owner<'_>, before: &BTreeSet<String>) -> Result<()> {
        let part = match owner {
            Owner::Main => return self.drop_released_relationships(before),
            Owner::Story { part, .. } => part.clone(),
        };
        let after = self.owner_referenced(owner)?;
        let released: BTreeSet<String> = before.difference(&after).cloned().collect();
        if released.is_empty() {
            return Ok(());
        }
        let uri = PackURI::new(part.clone()).map_err(crate::Error::Invalid)?;
        let package_relationships: Vec<&litchi_opc::rel::Relationship> =
            match self.package.opc_package().get_part(&uri) {
                Ok(source) => source.rels().iter().collect(),
                Err(_) => Vec::new(),
            };
        let mut relationships = Vec::new();
        let mut roots = Vec::new();
        for relationship in package_relationships {
            if !released.contains(relationship.r_id()) {
                continue;
            }
            relationships.push(PartRelationship {
                part: part.clone(),
                relationship: TransferRelationship {
                    id: relationship.r_id().to_owned(),
                    relationship_type: relationship.reltype().to_owned(),
                    target: relationship.target_ref().to_owned(),
                    external: relationship.is_external(),
                },
            });
            if !relationship.is_external() {
                roots.push(relationship.target_partname().map_err(crate::Error::from)?);
            }
        }
        let dropped: BTreeSet<(String, String)> = relationships
            .iter()
            .map(|owned| (owned.part.clone(), owned.relationship.id.clone()))
            .collect();
        let mut parts = if relationships.is_empty() {
            Vec::new()
        } else {
            self.unreachable_parts_without(&roots, &BTreeSet::new(), &dropped, false)?
        };
        // A relationship this edit added and no longer needs goes with the
        // part it added, so nothing it staged survives.
        let added: Vec<TransferRelationship> = self
            .staged_relationships(Some(&part))?
            .into_iter()
            .filter(|relationship| released.contains(&relationship.id))
            .collect();
        let pairs: Vec<(Option<String>, String)> = added
            .iter()
            .map(|relationship| (Some(part.clone()), relationship.id.clone()))
            .chain(
                dropped
                    .iter()
                    .map(|(owner, id)| (Some(owner.clone()), id.clone())),
            )
            .collect();
        parts.extend(self.released_staged_parts(Some(&part), &added, &pairs)?);
        relationships.extend(added.into_iter().map(|relationship| PartRelationship {
            part: part.clone(),
            relationship,
        }));
        if relationships.is_empty() {
            return Ok(());
        }
        self.inner_mut()?.stage_graph_change(
            Arc::new(GraphChange {
                relationships: Arc::new([]),
                parts: parts.into(),
                replaced: Arc::new([]),
                part_relationships: relationships.into(),
            }),
            false,
        )?;
        Ok(())
    }

    /// Whether the part `name` is reachable other than through the owner
    /// relationship `id` once: through another relationship of any part, or
    /// through `id` named more than once in the owner part.
    fn shared_beyond(&mut self, owner: &Owner<'_>, name: &str, id: &str) -> Result<bool> {
        let owner_name = match owner {
            Owner::Main => self
                .package
                .opc_package()
                .main_document_part()
                .map_err(crate::Error::from)?
                .partname()
                .as_str()
                .to_owned(),
            Owner::Story { part, .. } => part.clone(),
        };
        for part in self.package.opc_package().iter_parts() {
            for relationship in part.rels().iter() {
                if relationship.is_external()
                    || (part.partname().as_str() == owner_name && relationship.r_id() == id)
                {
                    continue;
                }
                if relationship
                    .target_partname()
                    .is_ok_and(|target| target.as_str() == name)
                {
                    return Ok(true);
                }
            }
        }
        let quoted = format!("\"{id}\"");
        let named = self
            .owner_xml(owner)?
            .windows(quoted.len())
            .filter(|window| *window == quoted.as_bytes())
            .count();
        Ok(named > 1)
    }

    /// The next free VML shape and OLE object numbers of the document:
    /// every WordprocessingML part is read, so a number a header, a footer
    /// or a note uses is not reused.
    fn next_object_numbers(&mut self) -> Result<(u32, u64)> {
        let (mut shape, mut object) = object_numbers(self.inner_mut()?.main_xml());
        let main_name = self
            .package
            .opc_package()
            .main_document_part()
            .map_err(crate::Error::from)?
            .partname()
            .as_str()
            .to_owned();
        let names: Vec<String> = self
            .package
            .opc_package()
            .iter_parts()
            .filter(|part| part.content_type().contains("wordprocessingml"))
            .map(|part| part.partname().as_str().to_owned())
            .filter(|name| *name != main_name)
            .collect();
        for name in names {
            if let Some(payload) = self.part_payload(&name)? {
                let (part_shape, part_object) = object_numbers(&payload);
                shape = shape.max(part_shape);
                object = object.max(part_object);
            }
        }
        let shape = shape
            .map_or(Some(FIRST_SHAPE), |maximum| {
                maximum.max(FIRST_SHAPE - 1).checked_add(1)
            })
            .ok_or_else(|| refused(Refusal::InvalidPropertyValue))?;
        let object = object
            .map_or(Some(FIRST_OBJECT + 1), |maximum| {
                maximum.max(FIRST_OBJECT).checked_add(1)
            })
            .ok_or_else(|| refused(Refusal::InvalidPropertyValue))?;
        Ok((shape, object))
    }

    /// The text size `hint` resolves to, in points: the run's own size, then
    /// its character style's, then its paragraph style's (the default
    /// paragraph style's when it names none), each through `w:basedOn`, then
    /// the document defaults, then Word's 10 points.
    fn resolved_size_pt(&self, hint: &FontHint) -> Result<f64> {
        let half_points = |size: u32| f64::from(size) / 2.0;
        if let Some(size) = hint.size.filter(|size| *size > 0) {
            return Ok(half_points(size));
        }
        let Some(catalog) = self.style_catalog()? else {
            return Ok(half_points(DEFAULT_SIZE_HALF_POINTS));
        };
        let chain_size = |id: &str| -> Result<Option<u32>> {
            let mut current = catalog.get(id);
            for _ in 0..MAX_STYLE_DEPTH {
                let Some(entry) = current else {
                    return Ok(None);
                };
                let run = catalog.run_properties(entry).map_err(Error::Document)?;
                if let Some(size) = run.size.and_then(|size| u32::try_from(size.0).ok()) {
                    return Ok(Some(size));
                }
                current = entry.based_on().and_then(|parent| catalog.get(parent));
            }
            Ok(None)
        };
        if let Some(style) = &hint.run_style
            && let Some(size) = chain_size(style)?
        {
            return Ok(half_points(size));
        }
        let paragraph_style = hint.paragraph_style.clone().or_else(|| {
            catalog
                .entries()
                .iter()
                .find(|entry| {
                    entry.is_default() && entry.kind() == Some(crate::styles::Type::Paragraph)
                })
                .map(|entry| entry.id().to_owned())
        });
        if let Some(style) = paragraph_style
            && let Some(size) = chain_size(&style)?
        {
            return Ok(half_points(size));
        }
        Ok(half_points(
            default_size(catalog.xml()).unwrap_or(DEFAULT_SIZE_HALF_POINTS),
        ))
    }

    /// The text size, in points, a new equation object of the paragraph at
    /// `paragraph` in `story` (the main document for `None`) inserted before
    /// its direct run `at` is drawn at: that run's size (for `None`, the
    /// paragraph's last run's), resolved through the styles as Word resolves
    /// it. Pass it to `EquationContent::new` (`equation` feature).
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a path that names no
    /// paragraph, or a malformed-part error.
    pub fn equation_object_font_size_pt(
        &mut self,
        story: Option<StoryTarget<'_>>,
        paragraph: &ElementPath,
        at: Option<Position>,
    ) -> Result<f64> {
        let owner = self.owner(story)?;
        let hint = self.on_owner(&owner, |tree| tree.insertion_font_hint(paragraph, at))?;
        self.resolved_size_pt(&hint)
    }

    /// The text size, in points, of the run holding the equation object at
    /// `index` among the paragraph's, resolved as
    /// [`Self::equation_object_font_size_pt`] resolves it: the size a
    /// replacement's preview is drawn at.
    ///
    /// # Errors
    ///
    /// As [`Self::equation_object_payload`].
    pub fn existing_equation_object_font_size_pt(
        &mut self,
        story: Option<StoryTarget<'_>>,
        paragraph: &ElementPath,
        index: Position,
    ) -> Result<f64> {
        let owner = self.owner(story)?;
        let hint = self.on_owner(&owner, |tree| tree.object_font_hint(paragraph, index))?;
        self.resolved_size_pt(&hint)
    }

    /// The equation objects of the paragraph at `paragraph` of `story` (the
    /// main document for `None`), as this edit leaves them.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a path that names no
    /// paragraph or a missing story.
    pub fn equation_objects_at(
        &mut self,
        story: Option<StoryTarget<'_>>,
        paragraph: &ElementPath,
    ) -> Result<Vec<crate::object::EquationObject>> {
        let owner = self.owner(story)?;
        self.on_owner(&owner, |inner| {
            let mut objects = Vec::new();
            let mut index = 0;
            loop {
                match inner.equation_object_span(paragraph, Position::new(index)) {
                    Ok(span) => objects.push(span.object()),
                    Err(Error::Refused {
                        reason: Refusal::EquationNotFound,
                        ..
                    }) => return Ok(objects),
                    Err(Error::Refused {
                        reason: Refusal::ComplexContent,
                        ..
                    }) => {},
                    Err(error) => return Err(error),
                }
                index += 1;
            }
        })
    }

    /// The object at `index` and the payload of its compound file.
    fn object_payload(
        &mut self,
        owner: &Owner<'_>,
        paragraph: &ElementPath,
        index: Position,
    ) -> Result<(EquationSpan, String, Arc<Vec<u8>>)> {
        let span = self.on_owner(owner, |tree| tree.equation_object_span(paragraph, index))?;
        if span
            .object_type
            .as_deref()
            .is_some_and(|kind| kind.eq_ignore_ascii_case("Link"))
        {
            return Err(refused(Refusal::ComplexContent));
        }
        let id = span
            .object_relationship_id
            .as_deref()
            .ok_or_else(|| refused(Refusal::ComplexContent))?;
        let name = self
            .owner_target(owner, id)?
            .ok_or_else(|| refused(Refusal::ComplexContent))?;
        let payload = self
            .part_payload(&name)?
            .ok_or_else(|| refused(Refusal::ComplexContent))?;
        Ok((span, name, payload))
    }

    /// The payload of the compound file of the equation object at `index`
    /// among the paragraph's, in `story` (the main document for `None`).
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::EquationNotFound`] for a missing object,
    /// [`Refusal::ComplexContent`] for a linked object, one without an
    /// embedded part, or one in alternate content.
    pub fn equation_object_payload(
        &mut self,
        story: Option<StoryTarget<'_>>,
        paragraph: &ElementPath,
        index: Position,
    ) -> Result<Arc<Vec<u8>>> {
        let owner = self.owner(story)?;
        Ok(self.object_payload(&owner, paragraph, index)?.2)
    }

    /// Insert a new equation object holding `content` into the paragraph at
    /// `paragraph` of `story` (the main document for `None`; a header, a
    /// footer or a note relates the new parts from its own part), before its
    /// direct run `at` (after its last child for `None`), and return its
    /// position among the paragraph's equation objects (the order of
    /// [`crate::Paragraph::equation_objects`]).
    ///
    /// Requires the `equation` feature.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::Package;
    /// use litchi_docx::document::{ElementPath, ElementStep, EquationContent, EquationSource};
    /// use litchi_equation::Limits;
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Area: ");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// let size = edit.equation_object_font_size_pt(None, &paragraph, None)?;
    /// let content =
    ///     EquationContent::new(EquationSource::Latex(r"\pi r^2"), None, size, &Limits::DEFAULT)?;
    /// let position = edit.insert_equation_object_at(None, &paragraph, None, &content)?;
    /// assert_eq!(position, Position::new(0));
    /// let mut next = edit.commit()?.into_package();
    /// let reopened = Package::from_bytes(&next.to_bytes()?)?;
    /// let objects = reopened.document()?.paragraphs()?[0].equation_objects()?;
    /// assert_eq!(objects[0].prog_id(), "Equation.DSMT4");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::TransitionalValue`] in a Strict document (which
    /// has no VML), [`Refusal::RunNotFound`] for a run position past the
    /// end, [`Refusal::FieldInterior`] inside a complex field,
    /// [`Refusal::BlockNotFound`] for a path that names no paragraph or a
    /// missing story, or a resource-limit error.
    #[cfg(feature = "equation")]
    pub fn insert_equation_object_at(
        &mut self,
        story: Option<StoryTarget<'_>>,
        paragraph: &ElementPath,
        at: Option<Position>,
        content: &EquationContent,
    ) -> Result<Position> {
        if self.is_strict()? {
            return Err(refused(Refusal::TransitionalValue));
        }
        // The paragraph first, so a missing one stages no part.
        self.equation_objects_at(story, paragraph)?;
        let owner = self.owner(story)?;
        let compound = litchi_equation::equation_object(&content.target, &content.native)
            .map_err(|_| refused(Refusal::InvalidPropertyValue))?;
        let (shape, object) = self.next_object_numbers()?;
        let object_relationship = self.add_owner_part(
            &owner,
            (OLE_OBJECT_RELATIONSHIP, "embeddings/", "oleObject", ".bin"),
            OLE_OBJECT_CONTENT_TYPE,
            compound,
        )?;
        let preview_relationship = self.add_owner_part(
            &owner,
            (IMAGE_RELATIONSHIP, "media/", "image", ".png"),
            "image/png",
            content.preview.clone(),
        )?;
        let markup = super::super::transaction::ObjectMarkup {
            prog_id: content.prog_id(),
            shape,
            object,
            object_relationship: &object_relationship,
            preview_relationship: &preview_relationship,
            size_pt: content.size_pt,
            lowering: content.lowering,
        };
        self.on_owner(&owner, |tree| {
            tree.insert_equation_object_markup_at(paragraph, at, &markup)
        })
    }

    /// The kind of the equation object at `index` among the paragraph's,
    /// the kind a replacement keeps: the MTEF header of its `Equation
    /// Native` stream, or, when that cannot be read, its ProgID's.
    ///
    /// Requires the `equation` feature.
    ///
    /// # Errors
    ///
    /// As [`Self::equation_object_payload`].
    #[cfg(feature = "equation")]
    pub fn equation_object_kind(
        &mut self,
        story: Option<StoryTarget<'_>>,
        paragraph: &ElementPath,
        index: Position,
        limits: &litchi_equation::Limits,
    ) -> Result<litchi_equation::MtefTarget> {
        let owner = self.owner(story)?;
        let (span, _, payload) = self.object_payload(&owner, paragraph, index)?;
        let by_prog_id = litchi_equation::MtefTarget::from_prog_id(&span.prog_id)
            .unwrap_or(litchi_equation::MtefTarget::MathType);
        Ok(litchi_equation::equation_object_native(&payload)
            .ok()
            .and_then(|native| litchi_equation::read_mtef(&native, limits).ok())
            .and_then(|read| read.info().cloned())
            .map_or(by_prog_id, litchi_equation::MtefTarget::SameAs))
    }

    /// Replace the equation of the object at `index` among the paragraph's
    /// with `content` (made for the object's kind, see
    /// [`Self::equation_object_kind`]), redraw its preview and resize its
    /// markup; return whether anything changed.
    ///
    /// Requires the `equation` feature.
    ///
    /// # Errors
    ///
    /// As [`Self::equation_object_payload`]. An embedding that is not a
    /// readable equation object is replaced by a new compound file of
    /// `content`'s kind.
    #[cfg(feature = "equation")]
    pub fn set_equation_object_at(
        &mut self,
        story: Option<StoryTarget<'_>>,
        paragraph: &ElementPath,
        index: Position,
        content: &EquationContent,
    ) -> Result<bool> {
        let owner = self.owner(story)?;
        let (span, object_name, payload) = self.object_payload(&owner, paragraph, index)?;
        // An embedding that is not a readable equation object becomes a
        // new compound file of the content's kind (the ProgID's, see
        // `equation_object_kind`); its markup keeps the ProgID.
        let compound =
            match litchi_equation::replace_equation_object_native(&payload, &content.native) {
                Ok(compound) => compound,
                Err(_) => litchi_equation::equation_object(&content.target, &content.native)
                    .map_err(|_| refused(Refusal::InvalidPropertyValue))?,
            };
        let before = self.owner_referenced(&owner)?;
        let object_id = span
            .object_relationship_id
            .clone()
            .ok_or_else(|| refused(Refusal::ComplexContent))?;
        let mut new_object = None;
        if self.shared_beyond(&owner, &object_name, &object_id)? {
            new_object = Some(self.add_owner_part(
                &owner,
                (OLE_OBJECT_RELATIONSHIP, "embeddings/", "oleObject", ".bin"),
                OLE_OBJECT_CONTENT_TYPE,
                compound,
            )?);
        } else {
            self.replace_part(&object_name, compound)?;
        }
        // The preview: replaced in place when it is a PNG only this object
        // shows.
        let mut new_preview = None;
        let preview = match span.preview_relationship_id.as_deref() {
            Some(id) => self
                .owner_target(&owner, id)?
                .map(|name| (id.to_owned(), name)),
            None => None,
        };
        let in_place = match &preview {
            Some((id, name)) => {
                let png = PackURI::new(name.clone())
                    .ok()
                    .and_then(|uri| self.package.opc_package().get_part(&uri).ok())
                    .is_some_and(|part| part.content_type() == "image/png");
                png && !self.shared_beyond(&owner, name, id)?
            },
            None => false,
        };
        match (&preview, in_place) {
            (Some((_, name)), true) => {
                self.replace_part(name, content.preview.clone())?;
            },
            // A shape without a `v:imagedata` names no preview to replace.
            _ if span.imagedata_tag.is_none() => {},
            _ => {
                new_preview = Some(self.add_owner_part(
                    &owner,
                    (IMAGE_RELATIONSHIP, "media/", "image", ".png"),
                    "image/png",
                    content.preview.clone(),
                )?);
            },
        }
        let rewrite = super::super::transaction::ObjectRewrite {
            size_pt: content.size_pt,
            lowering: content.lowering,
            preview_relationship: new_preview.as_deref(),
            object_relationship: new_object.as_deref(),
        };
        self.on_owner(&owner, |tree| {
            tree.rewrite_equation_object_at(paragraph, index, &rewrite)
        })?;
        self.drop_released(&owner, &before)?;
        Ok(true)
    }

    /// Remove the equation object at `index` among the paragraph's of
    /// `story` (the main document for `None`): its run when the run holds
    /// nothing else, its `w:object` otherwise. Its compound file and preview
    /// go with their relationships once nothing else names them.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::EquationNotFound`] for a missing object,
    /// [`Refusal::ComplexContent`] for one in alternate content, or
    /// [`Refusal::BlockNotFound`] for a path that names no paragraph or a
    /// missing story.
    pub fn remove_equation_object_at(
        &mut self,
        story: Option<StoryTarget<'_>>,
        paragraph: &ElementPath,
        index: Position,
    ) -> Result<()> {
        let owner = self.owner(story)?;
        let before = self.owner_referenced(&owner)?;
        self.on_owner(&owner, |tree| {
            tree.remove_equation_object_at(paragraph, index).map(|_| ())
        })?;
        self.drop_released(&owner, &before)
    }
}

/// The document-default run size of a styles part
/// (`w:docDefaults/w:rPrDefault/w:rPr/w:sz`), in half-points.
fn default_size(xml: &[u8]) -> Option<u32> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut stack: Vec<Vec<u8>> = Vec::new();
    loop {
        match reader.read_event().ok()? {
            quick_xml::events::Event::Start(element) => {
                stack.push(element.local_name().as_ref().to_vec());
            },
            quick_xml::events::Event::Empty(element)
                if element.local_name().as_ref() == b"sz"
                    && stack.iter().map(Vec::as_slice).eq([
                        b"styles".as_slice(),
                        b"docDefaults",
                        b"rPrDefault",
                        b"rPr",
                    ]) =>
            {
                use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
                return element
                    .unchecked_attributes()
                    .filter_map(std::result::Result::ok)
                    .find(|attribute| attribute.key.local_name().as_ref() == b"val")
                    .and_then(|attribute| {
                        std::str::from_utf8(&attribute.value)
                            .ok()?
                            .trim()
                            .parse()
                            .ok()
                    });
            },
            quick_xml::events::Event::End(_) => {
                stack.pop();
                if stack.is_empty() {
                    return None;
                }
            },
            quick_xml::events::Event::Eof => return None,
            _ => {},
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_document_default_size_is_read() {
        let xml = br#"<w:styles xmlns:w="w"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>"#;
        assert_eq!(default_size(xml), Some(22));
        assert_eq!(default_size(br#"<w:styles xmlns:w="w"/>"#), None);
    }
}
