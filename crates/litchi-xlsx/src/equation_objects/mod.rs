//! OLE equation objects of worksheets: MathType (`Equation.DSMT4` and the
//! other `Equation.DSMT*` ProgIDs) and Equation Editor 3.0 (`Equation.3`)
//! objects embedded as `xl/embeddings/oleObjectN.bin`.
//!
//! [`crate::Package::equation_objects`] and
//! [`crate::workbook::SourceWorksheet::equation_objects`] list a worksheet's
//! equation objects in `<oleObjects>` order, each an [`EquationObject`] with
//! its shape id (the stable address of the object), ProgID, kind, anchor,
//! embedding and preview part names, and its `Equation Native` stream, read
//! bounded per object: an embedding that cannot be read yields that object's
//! [`EquationObject::error`], never a failed listing. Other OLE objects are
//! not listed.
//!
//! [`crate::Package::edit_equation_objects`] starts an [`Edit`]:
//! [`Edit::add`] writes a new object as Excel writes an embedded object (see
//! below) from LaTeX, OMML or MTEF, with a PNG preview drawn from the
//! equation; [`Edit::set`] replaces an object's `Equation Native` stream
//! (keeping its kind), redraws its preview and resizes it; [`Edit::remove`]
//! deletes the object with every part, relationship and drawing shape only
//! it used. [`Edit::commit`] returns a [`Commit`] with the new package and
//! the reversible [`Patch`]; [`crate::Package::apply_equation_objects`]
//! replays a patch. Every byte not owned by the edited object is kept.
//!
//! # Markup
//!
//! A new object is the embedding part (content type
//! `application/vnd.openxmlformats-officedocument.oleObject`, an
//! `oleObject` relationship of the worksheet), a PNG preview part (an
//! `image` relationship of the worksheet and of the VML drawing), and three
//! pieces of markup sharing the shape id `N`:
//!
//! ```xml
//! <!-- worksheet <oleObjects> (ECMA-376 Part 1 §18.3.1.59-60; [MS-XLSX]) -->
//! <mc:AlternateContent><mc:Choice Requires="x14">
//!   <oleObject progId="Equation.DSMT4" shapeId="N" r:id="rIdObject">
//!     <objectPr defaultSize="0" r:id="rIdPreview">
//!       <anchor moveWithCells="1"><from>…</from><to>…</to></anchor>
//!     </objectPr></oleObject>
//! </mc:Choice><mc:Fallback>
//!   <oleObject progId="Equation.DSMT4" shapeId="N" r:id="rIdObject"/>
//! </mc:Fallback></mc:AlternateContent>
//! <!-- legacy VML drawing ([MS-OI29500]) -->
//! <v:shape id="_x0000_sN" type="#_x0000_t75" style="…">
//!   <v:imagedata o:relid="rIdPreview" o:title=""/>
//!   <x:ClientData ObjectType="Pict">…<x:CF>Pict</x:CF><x:AutoPict/></x:ClientData>
//! </v:shape>
//! <!-- DrawingML drawing ([MS-ODRAWXML]) -->
//! <mc:AlternateContent><mc:Choice Requires="a14"><xdr:twoCellAnchor editAs="oneCell">…
//!   <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="N" name="Object k" hidden="1">
//!     <a:extLst><a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}">
//!       <a14:compatExt spid="_x0000_sN"/></a:ext></a:extLst>…
//! </mc:Choice><mc:Fallback/></mc:AlternateContent>
//! ```
//!
//! Office writes an EMF or WMF preview; this module writes a PNG, an
//! ordinary image part that the VML `v:imagedata` and the `objectPr`
//! relationship name. The equation server redraws the preview when the
//! object is activated.

use std::borrow::Cow;
use std::collections::HashMap;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, OnceLock};

use litchi_equation::MtefTarget;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{OpcPackage, PackURI, Relationships, SourceBackedPackage};

use crate::error::Result;
use crate::ole_objects::{OleObjectAnchor, parse_ole_objects};

mod edit;
mod markup;
#[cfg(test)]
mod tests;

pub use edit::{Commit, Diagnostics, Edit, Error, Patch, Source, Written};

/// The most bytes one listing reads and decodes for the `Equation Native`
/// streams of its objects, whatever the package's own part budget allows.
pub const MAX_LISTING_NATIVE_BYTES: u64 = 256 * 1024 * 1024;

/// The bytes one listing may still read and decode: the package's total part
/// budget, at most [`MAX_LISTING_NATIVE_BYTES`].
#[derive(Debug)]
struct Budget {
    remaining: AtomicU64,
}

impl Budget {
    fn new(limit: u64) -> Self {
        Self {
            remaining: AtomicU64::new(limit.min(MAX_LISTING_NATIVE_BYTES)),
        }
    }

    fn charge(&self, bytes: u64) -> std::result::Result<(), String> {
        self.remaining
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |left| {
                left.checked_sub(bytes)
            })
            .map(drop)
            .map_err(|_| {
                "the equation objects of this worksheet exceed the listing's byte budget".to_owned()
            })
    }
}

/// Where an embedding part's bytes come from.
#[derive(Clone)]
pub(crate) enum Loader {
    /// An eager package's blob, shared.
    Eager(Arc<Vec<u8>>),
    /// A source-backed package, read on first use.
    Source(Box<SourceBackedPackage>),
}

/// One embedding part, shared by every object that names it: its
/// `Equation Native` stream is read and decoded once, on first use, against
/// the listing's byte budget.
struct Embedded {
    part: PackURI,
    loader: Loader,
    budget: Arc<Budget>,
    native: OnceLock<std::result::Result<Arc<[u8]>, String>>,
    /// The stream as LaTeX under the default limits; `None` when it cannot
    /// be read or does not convert.
    latex: OnceLock<Option<String>>,
}

impl Embedded {
    fn latex(&self) -> Option<&str> {
        self.latex
            .get_or_init(|| {
                let native = self.native().as_ref().ok()?;
                litchi_equation::mtef_to_latex(native, &litchi_equation::Limits::DEFAULT).ok()
            })
            .as_deref()
    }

    fn native(&self) -> &std::result::Result<Arc<[u8]>, String> {
        self.native
            .get_or_init(|| self.load().map_err(|error| error.to_string()))
    }

    fn load(&self) -> Result<Arc<[u8]>> {
        let maximum = litchi_equation::MAX_EQUATION_OBJECT_BYTES;
        let size = match &self.loader {
            Loader::Eager(bytes) => u64::try_from(bytes.len()).unwrap_or(u64::MAX),
            Loader::Source(package) => package.part(&self.part)?.declared_uncompressed_size()?,
        };
        if size > maximum {
            return Err(too_large(&self.part));
        }
        self.budget.charge(size).map_err(crate::error::invalid)?;
        let native = match &self.loader {
            Loader::Eager(bytes) => litchi_equation::equation_object_native(bytes)?,
            Loader::Source(package) => {
                let data = package.part(&self.part)?.data()?;
                litchi_equation::equation_object_native(data.as_bytes())?
            },
        };
        self.budget
            .charge(u64::try_from(native.len()).unwrap_or(u64::MAX))
            .map_err(crate::error::invalid)?;
        #[cfg(test)]
        tests::DECODED.with(|decoded| decoded.set(decoded.get() + native.len()));
        Ok(Arc::from(native))
    }
}

/// An object's `Equation Native` stream: its shared embedding, or why its
/// relationship does not resolve.
#[derive(Clone)]
enum Native {
    Embedded(Arc<Embedded>),
    Unresolved(String),
}

impl Native {
    fn get(&self) -> std::result::Result<&[u8], &str> {
        match self {
            Self::Embedded(embedded) => match embedded.native() {
                Ok(native) => Ok(native),
                Err(error) => Err(error),
            },
            Self::Unresolved(error) => Err(error),
        }
    }
}

/// One OLE equation object of a worksheet. Its `Equation Native` stream is
/// read on first use, once per embedding part however many objects name
/// it, within the listing's byte budget.
#[derive(Clone)]
pub struct EquationObject {
    shape_id: u32,
    prog_id: String,
    kind: MtefTarget,
    anchor: Option<OleObjectAnchor>,
    embedding: Option<String>,
    preview: Option<String>,
    native: Native,
}

impl std::fmt::Debug for EquationObject {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("EquationObject")
            .field("shape_id", &self.shape_id)
            .field("prog_id", &self.prog_id)
            .field("kind", &self.kind)
            .field("anchor", &self.anchor)
            .field("embedding", &self.embedding)
            .field("preview", &self.preview)
            .finish_non_exhaustive()
    }
}

impl EquationObject {
    /// The object's shape id (`shapeId`, the VML shape `_x0000_sN`): its
    /// stable address within the worksheet.
    #[must_use]
    pub const fn shape_id(&self) -> u32 {
        self.shape_id
    }

    /// The object's ProgID as stored (`Equation.DSMT4`, `Equation.3`).
    #[must_use]
    pub fn prog_id(&self) -> &str {
        &self.prog_id
    }

    /// The kind of equation object the ProgID names.
    #[must_use]
    pub const fn kind(&self) -> &MtefTarget {
        &self.kind
    }

    /// The object's two-cell anchor (`objectPr/anchor`); `None` for an
    /// object Excel 2007 wrote, which has only its VML shape.
    #[must_use]
    pub const fn anchor(&self) -> Option<&OleObjectAnchor> {
        self.anchor.as_ref()
    }

    /// The embedding part's name, when its relationship resolves.
    #[must_use]
    pub fn embedding(&self) -> Option<&str> {
        self.embedding.as_deref()
    }

    /// The preview image part's name (`objectPr` relationship), when the
    /// object has one.
    #[must_use]
    pub fn preview(&self) -> Option<&str> {
        self.preview.as_deref()
    }

    /// The complete `Equation Native` stream (the 28-byte `EQNOLEFILEHDR`
    /// and the MTEF data), read on first use; `None` when the embedding
    /// cannot be read or the listing's byte budget is spent.
    #[must_use]
    pub fn native(&self) -> Option<&[u8]> {
        self.native.get().ok()
    }

    /// Why the embedding cannot be read, when it cannot (reads it on first
    /// use).
    #[must_use]
    pub fn error(&self) -> Option<&str> {
        self.native.get().err()
    }

    /// The equation as LaTeX under the default limits, converted once per
    /// embedding part and kept: read on first use, within the listing's byte
    /// budget; `None` when the embedding cannot be read, the budget is spent
    /// or the MTEF does not convert.
    #[must_use]
    pub fn default_latex(&self) -> Option<&str> {
        match &self.native {
            Native::Embedded(embedded) => embedded.latex(),
            Native::Unresolved(_) => None,
        }
    }

    /// The equation as LaTeX.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::Invalid`] when the embedding cannot be read
    /// and [`crate::Error::Equation`] when its MTEF does not convert.
    pub fn latex(&self, limits: &litchi_equation::Limits) -> Result<String> {
        if *limits == litchi_equation::Limits::DEFAULT
            && let Native::Embedded(embedded) = &self.native
            && let Some(latex) = embedded.latex()
        {
            return Ok(latex.to_owned());
        }
        let native = self
            .native
            .get()
            .map_err(|error| crate::error::invalid(error.to_owned()))?;
        Ok(litchi_equation::mtef_to_latex(native, limits)?)
    }
}

/// Read access to the parts of an eager or a source-backed package.
pub(crate) trait Parts {
    fn relationships(&self, part: &PackURI) -> Result<&Relationships>;
    /// The part's bytes, refused when larger than `maximum`.
    fn bytes(&self, part: &PackURI, maximum: u64) -> Result<Cow<'_, [u8]>>;
    /// How an embedding part is read later.
    fn loader(&self, part: &PackURI) -> Result<Loader>;
    /// The package's total part budget.
    fn budget(&self) -> u64;
}

fn too_large(part: &PackURI) -> Error {
    crate::error::invalid(format!("part '{part}' exceeds the equation object limit"))
}

impl Parts for OpcPackage {
    fn relationships(&self, part: &PackURI) -> Result<&Relationships> {
        Ok(self.get_part(part)?.rels())
    }

    fn bytes(&self, part: &PackURI, maximum: u64) -> Result<Cow<'_, [u8]>> {
        let blob = self.get_part(part)?.blob();
        if u64::try_from(blob.len()).unwrap_or(u64::MAX) > maximum {
            return Err(too_large(part));
        }
        Ok(Cow::Borrowed(blob))
    }

    fn loader(&self, part: &PackURI) -> Result<Loader> {
        Ok(Loader::Eager(self.get_part(part)?.blob_arc()))
    }

    fn budget(&self) -> u64 {
        self.read_limits().max_total_part_bytes()
    }
}

impl Parts for SourceBackedPackage {
    fn relationships(&self, part: &PackURI) -> Result<&Relationships> {
        Ok(self.part(part)?.rels())
    }

    fn bytes(&self, part: &PackURI, maximum: u64) -> Result<Cow<'_, [u8]>> {
        let view = self.part(part)?;
        if view.declared_uncompressed_size()? > maximum {
            return Err(too_large(part));
        }
        Ok(Cow::Owned(view.data()?.as_bytes().to_vec()))
    }

    fn loader(&self, part: &PackURI) -> Result<Loader> {
        self.part(part)?;
        Ok(Loader::Source(Box::new(self.clone())))
    }

    fn budget(&self) -> u64 {
        self.read_limits().max_total_part_bytes()
    }
}

/// The equation objects of the worksheet part `worksheet`.
pub(crate) fn list(parts: &impl Parts, worksheet: &PackURI) -> Result<Vec<EquationObject>> {
    list_within(parts, worksheet, parts.budget())
}

/// [`list`] whose objects read at most `budget` bytes of embeddings and
/// `Equation Native` streams.
fn list_within(
    parts: &impl Parts,
    worksheet: &PackURI,
    budget: u64,
) -> Result<Vec<EquationObject>> {
    let xml = parts.bytes(worksheet, u64::MAX)?;
    if markup::find(&xml, b"oleObject", 0).is_none() {
        return Ok(Vec::new());
    }
    let Some(objects) = parse_ole_objects(&xml)? else {
        return Ok(Vec::new());
    };
    let relationships = parts.relationships(worksheet)?;
    let budget = Arc::new(Budget::new(budget));
    let mut shared: HashMap<PackURI, Arc<Embedded>> = HashMap::new();
    let mut out = Vec::new();
    for object in objects.objects {
        let Some(kind) = object
            .program_id
            .as_deref()
            .and_then(MtefTarget::from_prog_id)
        else {
            continue;
        };
        let embedding = embedding_part(relationships, &object.relationship_id);
        let native = match &embedding {
            Ok(part) => match shared.get(part) {
                Some(embedded) => Native::Embedded(Arc::clone(embedded)),
                None => match parts.loader(part) {
                    Ok(loader) => {
                        let embedded = Arc::new(Embedded {
                            part: part.clone(),
                            loader,
                            budget: Arc::clone(&budget),
                            native: OnceLock::new(),
                            latex: OnceLock::new(),
                        });
                        shared.insert(part.clone(), Arc::clone(&embedded));
                        Native::Embedded(embedded)
                    },
                    Err(error) => Native::Unresolved(error.to_string()),
                },
            },
            Err(error) => Native::Unresolved(error.clone()),
        };
        let preview = object.properties.as_ref().and_then(|properties| {
            let relationship = relationships.get(&properties.preview_relationship_id)?;
            (!relationship.is_external())
                .then(|| relationship.target_partname().ok())
                .flatten()
                .map(|part| part.to_string())
        });
        out.push(EquationObject {
            shape_id: object.shape_id,
            prog_id: object.program_id.unwrap_or_default(),
            kind,
            anchor: object.properties.map(|properties| properties.anchor),
            embedding: embedding.ok().map(|part| part.to_string()),
            preview,
            native,
        });
    }
    Ok(out)
}

/// The embedding part an object relationship targets, or why it does not
/// resolve to an embedded object.
fn embedding_part(relationships: &Relationships, id: &str) -> std::result::Result<PackURI, String> {
    let relationship = relationships
        .get(id)
        .ok_or_else(|| format!("the object relationship '{id}' is missing"))?;
    if relationship.is_external() {
        return Err("the object is linked, not embedded".to_owned());
    }
    if !matches!(
        relationship.reltype(),
        rt::OLE_OBJECT | rt::STRICT_OLE_OBJECT
    ) {
        return Err("the object relationship is not an embedded OLE object".to_owned());
    }
    relationship
        .target_partname()
        .map_err(|error| error.to_string())
}

impl crate::Package {
    /// The OLE equation objects of the worksheet at catalog position
    /// `sheet`, in `<oleObjects>` order.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet, a chartsheet, or a worksheet
    /// whose `<oleObjects>` markup is malformed.
    pub fn equation_objects(&self, sheet: usize) -> Result<Vec<EquationObject>> {
        let worksheet = edit::worksheet_part(self, sheet)?;
        list(self.opc(), &worksheet)
    }
}
