//! OLE equation objects of a worksheet drawing.
//!
//! Excel keeps a MathType (`Equation.DSMT4`) or Equation Editor 3.0
//! (`Equation.3`) equation as an embedded OLE object ([MS-XLS] 2.4.181 Obj,
//! 2.5.150 FtPictFmla):
//! a picture-frame shape in the worksheet's `MsoDrawing` (its `pib` names
//! the preview picture in the drawing group's blip store), the shape's
//! `Obj` record (`ftCmo.ot` 8, `ftCf`, `ftPioGrbit` and an `ftPictFmla`
//! naming the object's ProgID and its `MBDxxxxxxxx` storage), and that
//! storage in the compound file, whose `Equation Native` stream holds the
//! MTEF data.

use super::drawing::{Block, Object};
use super::read;
use super::stream::{Globals, Sheet};
use crate::Package;

/// The largest `Equation Native` stream an equation read or edit accepts.
pub const MAX_EQUATION_NATIVE_BYTES: u64 = 16 * 1024 * 1024;

/// The `Equation Native` stream of an OLE equation storage.
pub(super) const EQUATION_NATIVE: &str = "Equation Native";

/// A change to an embedded-object storage that a worksheet-feature edit
/// publishes together with its Workbook stream.
#[derive(Debug, Clone)]
pub(super) enum StagedStorage {
    /// A new equation storage of `kind` holding `native`.
    New {
        /// The object's kind.
        kind: EquationKind,
        /// Its `Equation Native` stream.
        native: Vec<u8>,
    },
    /// A new `Equation Native` stream for an existing storage.
    Native(Vec<u8>),
    /// The storage is removed with its whole subtree.
    Removed,
}

/// The kind of OLE equation object, from its ProgID.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum EquationKind {
    /// MathType (`Equation.DSMT4` and other `Equation.DSMT*` ProgIDs),
    /// holding MTEF 5.
    MathType,
    /// Microsoft Equation 3.0 (`Equation.3`), holding MTEF 3.
    Equation3,
    /// Another `Equation.*` class.
    Other,
}

impl EquationKind {
    /// The kind a ProgID names.
    #[must_use]
    pub fn from_prog_id(prog_id: &str) -> Self {
        if prog_id.eq_ignore_ascii_case("Equation.3") {
            Self::Equation3
        } else if prog_id
            .get(..13)
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case("Equation.DSMT"))
        {
            Self::MathType
        } else {
            Self::Other
        }
    }
}

/// One OLE equation of a worksheet, in drawing order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Equation {
    /// The `Obj` identifier.
    pub object_id: u16,
    /// The OfficeArt shape identifier.
    pub shape_id: u32,
    /// The cell anchor.
    pub anchor: Option<crate::SheetAnchor>,
    /// The object's ProgID, as its `ftPictFmla` names it.
    pub prog_id: String,
    /// The compound-file storage that holds the object (`MBDxxxxxxxx`).
    pub storage: String,
    /// The `Equation Native` stream (the 28-byte `EQNOLEFILEHDR` and the
    /// MTEF data); `None` when the storage has no such stream or it could
    /// not be read.
    pub native: Option<Vec<u8>>,
    /// Why the `Equation Native` stream could not be read (for example, it
    /// is larger than [`MAX_EQUATION_NATIVE_BYTES`]); `None` when it was
    /// read or is absent.
    pub native_error: Option<String>,
    /// The preview picture's MIME type and size, from its blip store
    /// entry.
    pub preview: Option<(&'static str, u32)>,
}

impl Equation {
    /// The object's kind, from its ProgID.
    #[must_use]
    pub fn kind(&self) -> EquationKind {
        EquationKind::from_prog_id(&self.prog_id)
    }

    /// The equation as LaTeX, read from its MTEF data.
    ///
    /// # Errors
    ///
    /// Returns [`litchi_equation::Error::Malformed`] when the object has no
    /// `Equation Native` stream or the stream is not valid MTEF, and the
    /// limit errors when it exceeds `limits`.
    #[cfg(feature = "equation")]
    pub fn latex(
        &self,
        limits: &litchi_equation::Limits,
    ) -> Result<String, litchi_equation::Error> {
        let native = self
            .native
            .as_deref()
            .ok_or_else(|| litchi_equation::Error::Malformed {
                format: litchi_equation::Format::Mtef,
                message: "the OLE object has no Equation Native stream".to_owned(),
            })?;
        litchi_equation::mtef_to_latex(native, limits)
    }
}

/// The positions, among a drawing's top-level objects, of its OLE
/// equations.
pub(super) fn positions(block: &Block) -> Vec<usize> {
    block
        .objects
        .iter()
        .enumerate()
        .filter(|(_, object)| is_equation(object))
        .map(|(position, _)| position)
        .collect()
}

/// Whether a drawing object is an OLE equation: a picture `Obj` whose
/// formula names an `Equation.*` class.
pub(super) fn is_equation(object: &Object) -> bool {
    !object.group
        && object.obj.is_some_and(|(ot, _)| ot == 8)
        && object
            .ole
            .as_ref()
            .is_some_and(super::drawing::OleLink::is_equation)
}

/// The equation an object projects, with its `Equation Native` stream.
pub(super) fn equation_of(
    object: &Object,
    store: &[(u8, u32)],
    native: Option<Vec<u8>>,
) -> Option<Equation> {
    let (_, object_id) = object.obj?;
    let ole = object.ole.as_ref()?;
    let preview = object
        .prop(0x0104)
        .and_then(|prop| store.get((prop.value as usize).checked_sub(1)?))
        .map(|(win, size)| (super::drawing::content_type(*win), *size));
    Some(Equation {
        object_id,
        shape_id: object.shape_id,
        anchor: object.anchor.as_ref().and_then(read::sheet_anchor),
        prog_id: ole.class.clone(),
        storage: ole.storage_name(),
        native,
        native_error: None,
        preview,
    })
}

/// The `Equation Native` stream of `storage`, or why it could not be read.
/// A changed source still fails: every other read of the package would.
pub(super) fn read_native(
    package: &Package,
    storage: &str,
) -> crate::package::Result<(Option<Vec<u8>>, Option<String>)> {
    match package.package_stream(&[storage, EQUATION_NATIVE], MAX_EQUATION_NATIVE_BYTES) {
        Ok(native) => Ok((native, None)),
        Err(
            error @ (crate::package::Error::Source(_) | crate::package::Error::CachedSource(_)),
        ) => Err(error),
        Err(error) => Ok((None, Some(error.to_string()))),
    }
}

/// The worksheet drawing of `sheet` and the blip store, for reading.
pub(super) fn drawing(
    stream: &[u8],
    sheet: usize,
) -> crate::Result<(Option<Block>, Vec<(u8, u32)>)> {
    let globals = Globals::parse(stream)?;
    let start = globals.worksheet(sheet)?;
    let layout = Sheet::parse(stream, start)?;
    let block = Block::parse(stream, &layout)?;
    let store = if block.is_some() {
        read::drawing_store(stream, &globals)?
    } else {
        Vec::new()
    };
    Ok((block, store))
}

impl Package {
    /// The OLE equations of the worksheet at catalog position `sheet`, in
    /// drawing order, each with its `Equation Native` stream.
    ///
    /// # Errors
    ///
    /// Returns an error when the position is not an ordinary worksheet or
    /// its drawing records are malformed. An `Equation Native` stream that
    /// cannot be read (one larger than [`MAX_EQUATION_NATIVE_BYTES`], for
    /// example) leaves that equation's `native` empty and says why in
    /// `native_error`.
    pub fn equations(&self, sheet: usize) -> crate::package::Result<Vec<Equation>> {
        let stream = self.workbook_stream()?;
        let (block, store) = drawing(&stream, sheet).map_err(crate::package::Error::Reader)?;
        let Some(block) = block else {
            return Ok(Vec::new());
        };
        let mut out = Vec::new();
        for position in positions(&block) {
            let object = &block.objects[position];
            let Some(ole) = object.ole.as_ref() else {
                continue;
            };
            let storage = ole.storage_name();
            let (native, error) = read_native(self, &storage)?;
            if let Some(mut equation) = equation_of(object, &store, native) {
                equation.native_error = error;
                out.push(equation);
            }
        }
        Ok(out)
    }
}
