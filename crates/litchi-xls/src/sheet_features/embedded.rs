//! Embedded OLE objects of the worksheet-feature edit.
//!
//! Excel keeps an embedded OLE object as three linked parts ([MS-XLS]
//! 2.4.181 Obj and the `OBJECTS` grammar, 2.5.150 FtPictFmla, 2.5.187
//! ObjFmla, 2.5.202 PictFmlaEmbedInfo):
//!
//! - a picture-frame shape (`msosptPictureFrame`, `fOleShape`) in the
//!   worksheet's `MsoDrawing`, with its cell anchor and a `pib` naming the
//!   preview picture in the drawing group's blip store;
//! - the shape's `Obj` record right after it (`ftCmo.ot` 8, `ftCf`
//!   unspecified, `ftPioGrbit` `fAutoPict`, and an `ftPictFmla` naming the
//!   ProgID and the storage position);
//! - the `MBDxxxxxxxx` storage of the compound file.
//!
//! Adding writes all three and removing takes all three away (the preview
//! picture loses one reference and goes when it is the store's last blip),
//! so the drawing shapes and the `Obj` records stay paired.
//!
//! A form control is added the same way: a host-control shape
//! (`msosptHostControl`) with its anchor, followed by the control's `Obj`.

use litchi_core::edit::RefusalKind;
use litchi_odraw::builder::PropertyBuilder;

use super::super::super::drawing::ole_link;
use super::super::super::equation::StagedStorage;
use super::super::super::stream::{Globals, OBJ, Sheet};
use super::refused;
use super::{Anchor, Edit, NewObject, Result, image_kind, obj_record};

/// An embedded OLE object to add.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedObjectSpec {
    /// The object's storage: a complete compound file whose root becomes
    /// the `MBDxxxxxxxx` storage. It is kept as opaque data.
    pub compound_file: Vec<u8>,
    /// The ProgID the `Obj` names (for example `Word.Document.8`); empty
    /// when the object names none. ASCII, at most 255 bytes.
    pub prog_id: String,
    /// The preview picture (PNG or JPEG) shown in place of the object.
    pub preview: Vec<u8>,
    /// The cell anchor of the object's frame.
    pub anchor: Anchor,
}

/// The `Obj` of an embedded OLE object: `ftCmo` (`ot` 8, locked, printed,
/// auto-fill, auto-line), `ftCf` unspecified, `ftPioGrbit` `fAutoPict`, and
/// `ftPictFmla` with a `PtgTbl` formula, the ProgID (when there is one) and
/// the storage position, as Excel writes them.
pub(in crate::sheet_features) fn embedded_obj(
    object_id: u16,
    prog_id: &str,
    storage: u32,
) -> Result<Vec<u8>> {
    let class = prog_id.as_bytes();
    let count = u8::try_from(class.len())
        .ok()
        .filter(|_| prog_id.is_ascii())
        .ok_or(refused(
            RefusalKind::InvalidValue,
            "the ProgID is not ASCII of at most 255 bytes",
        ))?;
    // ObjectParsedFormula: cce = 5, four unused bytes, PtgTbl; then
    // PictFmlaEmbedInfo: ttb 3, cbClass, reserved and, when cbClass is not
    // zero, an 8-bit XLUnicodeStringNoCch; padded to an even length.
    let mut formula = vec![0x05, 0x00, 0, 0, 0, 0, 0x02, 0, 0, 0, 0, 0x03, count, 0x00];
    if !class.is_empty() {
        formula.push(0x00);
        formula.extend_from_slice(class);
    }
    if formula.len() % 2 == 1 {
        formula.push(0);
    }
    let formula_length = u16::try_from(formula.len())
        .map_err(|_| refused(RefusalKind::InvalidValue, "the ProgID is too long"))?;
    let mut extra = vec![
        0x07, 0x00, 0x02, 0x00, 0xff, 0xff, 0x08, 0x00, 0x02, 0x00, 0x01, 0x00,
    ];
    extra.extend_from_slice(&0x0009u16.to_le_bytes());
    extra.extend_from_slice(&(formula_length + 6).to_le_bytes());
    extra.extend_from_slice(&formula_length.to_le_bytes());
    extra.extend(formula);
    extra.extend_from_slice(&storage.to_le_bytes());
    obj_record(8, object_id, &extra)
}

/// The shape container of an embedded OLE object: a picture frame with
/// `fOleShape`, the properties Excel writes for one, its anchor and its
/// client data.
fn ole_shape(shape_id: u32, pib: u32, anchor: [u8; 18]) -> Result<Vec<u8>> {
    let mut properties = PropertyBuilder::new();
    let pib = i32::try_from(pib).unwrap_or(i32::MAX);
    #[allow(
        clippy::cast_possible_wrap,
        reason = "OfficeArt property values are raw 32-bit fields"
    )]
    for (id, value) in [
        (0x00bfu16, 0x0008_0008u32),
        (0x0104, 0),
        (0x010b, 1),
        (0x0181, 0x0800_0041),
        (0x0183, 0x0800_0041),
        (0x01bf, 0x0010_0010),
        (0x01c0, 0x0800_0040),
        (0x01ff, 0x0008_0008),
        (0x033f, 0x0010_0010),
    ] {
        if id == 0x0104 {
            properties.add_blip_id(litchi_odraw::prop::Id::from(id), pib);
        } else {
            properties.add_simple(litchi_odraw::prop::Id::from(id), value as i32);
        }
    }
    super::shape_container_with(
        75,
        shape_id,
        litchi_odraw::shape::Flags::OLE_SHAPE,
        (false, false),
        &properties,
        anchor,
        false,
    )
}

/// The shape container of a form control without a label, as Excel writes
/// one: a host-control shape (`msosptHostControl`, [MS-ODRAW] 2.4.24) with
/// the properties Excel gives it, its anchor and its client data.
fn control_shape(shape_id: u32, anchor: [u8; 18]) -> Result<Vec<u8>> {
    let mut properties = PropertyBuilder::new();
    #[allow(
        clippy::cast_possible_wrap,
        reason = "OfficeArt property values are raw 32-bit fields"
    )]
    for (id, value) in [
        (0x007fu16, 0x0104_0104u32),
        (0x00bf, 0x0008_0008),
        (0x01c0, 0x0800_0040),
        (0x01ff, 0x0008_0000),
    ] {
        properties.add_simple(litchi_odraw::prop::Id::from(id), value as i32);
    }
    super::shape_container_with(
        201,
        shape_id,
        litchi_odraw::shape::Flags::empty(),
        (false, false),
        &properties,
        anchor,
        false,
    )
}

impl Edit {
    /// Append a form control: its host-control shape at `anchor`, then the
    /// `Obj` that `obj` renders for the identifier it is given. Returns that
    /// identifier.
    pub(crate) fn add_form_control_object(
        &mut self,
        sheet: usize,
        anchor: Anchor,
        obj: impl FnOnce(u16) -> crate::Result<Vec<u8>>,
    ) -> Result<u16> {
        self.ensure_ready()?;
        self.step()?;
        anchor.check()?;
        let (object_id, _) = self.append_object(sheet, |shape_id, object_id| {
            Ok(NewObject {
                sp: control_shape(shape_id, anchor.bytes())?,
                obj: obj(object_id)?,
                textbox: None,
            })
        })?;
        Ok(object_id)
    }

    pub(in crate::sheet_features) fn stage(&mut self, storage: String, change: StagedStorage) {
        let staged = self.storages.iter().position(|(name, _)| *name == storage);
        match (staged, change) {
            (Some(index), StagedStorage::Removed) => {
                if matches!(
                    self.storages[index].1,
                    StagedStorage::New { .. } | StagedStorage::File(_)
                ) {
                    self.storages.remove(index);
                } else {
                    self.storages[index].1 = StagedStorage::Removed;
                }
            },
            (Some(index), StagedStorage::Native(native)) => match &mut self.storages[index].1 {
                StagedStorage::New { native: staged, .. } => *staged = native,
                other => *other = StagedStorage::Native(native),
            },
            (Some(index), change) => self.storages[index].1 = change,
            (None, change) => self.storages.push((storage, change)),
        }
    }

    /// Every `MBD` storage position an `Obj` of the workbook names.
    fn referenced_storages(&self) -> Result<Vec<u32>> {
        let globals = Globals::parse(&self.stream)?;
        let mut out = Vec::new();
        for bound in globals.bounds.iter().filter(|bound| bound.kind == 0) {
            let layout = Sheet::parse(&self.stream, bound.position)?;
            for (index, record) in layout.records.iter().enumerate() {
                if record.kind == OBJ {
                    if let Some(link) = ole_link(record.payload(&self.stream), index) {
                        out.push(link.storage);
                    }
                }
            }
        }
        Ok(out)
    }

    /// A storage position no `Obj`, staged change or compound-file entry
    /// uses.
    fn free_storage(&self) -> Result<u32> {
        let referenced = self.referenced_storages()?;
        let mut candidate = referenced.iter().max().map_or(1, |max| max.wrapping_add(1));
        for _ in 0..4_096 {
            let name = format!("MBD{candidate:08X}");
            let taken = candidate == 0
                || referenced.contains(&candidate)
                || self.storages.iter().any(|(staged, _)| *staged == name)
                || self.base.package_entry_exists(&[name.as_str()])?;
            if !taken {
                return Ok(candidate);
            }
            candidate = candidate.wrapping_add(1);
        }
        Err(refused(
            RefusalKind::UnsupportedOperation,
            "no free embedded-object storage name was found",
        ))
    }

    /// Append an embedded object: its picture-frame shape at `anchor` with
    /// the preview, its `Obj` naming `prog_id` and a fresh storage, and the
    /// storage `change` makes. Returns the `Obj` identifier.
    pub(in crate::sheet_features) fn add_embedded(
        &mut self,
        sheet: usize,
        prog_id: &str,
        preview: &[u8],
        anchor: [u8; 18],
        change: StagedStorage,
    ) -> Result<u16> {
        let kind = image_kind(preview).ok_or(refused(
            RefusalKind::InvalidValue,
            "the preview is not PNG or JPEG",
        ))?;
        // Check the Obj before the blip store and the drawing change.
        embedded_obj(1, prog_id, 1)?;
        let storage = self.free_storage()?;
        // Refuse a drawing the object cannot join before the blip store
        // grows.
        let (_, _, block, group) = self.drawing(sheet)?;
        if block.as_ref().is_some_and(|block| !block.canonical) {
            return Err(refused(
                RefusalKind::UnsupportedOperation,
                "the worksheet drawing's record layout is not rewritten",
            ));
        }
        Self::place(block.as_ref(), group.as_ref())?;
        let pib = self.add_blip(kind, preview)?;
        let (object_id, _) = self.append_object(sheet, |shape_id, object_id| {
            Ok(NewObject {
                sp: ole_shape(shape_id, pib, anchor)?,
                obj: embedded_obj(object_id, prog_id, storage)?,
                textbox: None,
            })
        })?;
        self.stage(format!("MBD{storage:08X}"), change);
        Ok(object_id)
    }

    /// Remove the embedded object at drawing `position`: its shape, its
    /// `Obj`, its preview's blip reference and its storage (another `Obj`
    /// naming the same storage keeps it).
    pub(in crate::sheet_features) fn remove_embedded(
        &mut self,
        sheet: usize,
        position: usize,
    ) -> Result<()> {
        let object = self.remove_object(sheet, position)?;
        if let Some(prop) = object.prop(0x0104) {
            self.release_blip(prop.value)?;
        }
        self.drop_empty_group()?;
        if let Some(ole) = &object.ole {
            // A storage the compound file does not have is not removed.
            let name = ole.storage_name();
            let staged = self.storages.iter().any(|(staged, _)| *staged == name);
            if !self.referenced_storages()?.contains(&ole.storage)
                && (staged || self.base.package_entry_exists(&[name.as_str()])?)
            {
                self.stage(name, StagedStorage::Removed);
            }
        }
        Ok(())
    }

    /// Add an embedded OLE object: a picture frame at `spec.anchor` showing
    /// `spec.preview`, its `Obj`, and a new `MBDxxxxxxxx` storage holding
    /// `spec.compound_file`. Returns the new `Obj` identifier.
    ///
    /// # Errors
    ///
    /// Returns [`Refused`](super::Error::Refused) for a signed or encrypted
    /// payload, a ProgID that is not short ASCII, a preview that is not PNG
    /// or JPEG, or an anchor outside the grid, and
    /// [`Records`](super::Error::Records) for a payload that is not a
    /// compound file or exceeds the default OLE-object limits.
    pub fn add_embedded_object(&mut self, sheet: usize, spec: &EmbeddedObjectSpec) -> Result<u16> {
        self.ensure_ready()?;
        self.step()?;
        spec.anchor.check()?;
        crate::ole_object::validate_compound_file_for_publication(
            &spec.compound_file,
            crate::ole_object::Limits::default(),
        )?;
        self.add_embedded(
            sheet,
            &spec.prog_id,
            &spec.preview,
            spec.anchor.bytes(),
            StagedStorage::File(spec.compound_file.clone()),
        )
    }

    /// Remove the embedded OLE object whose `Obj` identifier is
    /// `object_id`, with its shape, its preview's blip reference and its
    /// storage.
    ///
    /// # Errors
    ///
    /// Returns [`Refused`](super::Error::Refused) when the worksheet drawing
    /// has no storage-backed embedded object with that identifier (a linked
    /// object, a control or another drawing is not removed here) or its
    /// record layout is not rewritten.
    pub fn remove_embedded_object(&mut self, sheet: usize, object_id: u16) -> Result<()> {
        self.ensure_ready()?;
        self.step()?;
        let (_, _, block, _) = self.drawing(sheet)?;
        let position = block
            .and_then(|block| {
                block.objects.iter().position(|object| {
                    !object.group && object.obj == Some((8, object_id)) && object.ole.is_some()
                })
            })
            .ok_or(refused(
                RefusalKind::NotFound,
                "the embedded object does not exist",
            ))?;
        self.remove_embedded(sheet, position)
    }
}
