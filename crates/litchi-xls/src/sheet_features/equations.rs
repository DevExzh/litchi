//! OLE equation operations of the worksheet-feature edit.
//!
//! A new equation is written as Excel writes an embedded OLE object: a
//! picture-frame shape (`msosptPictureFrame`, `fOleShape`) whose `pib`
//! names the preview picture, followed by its `Obj` (`ftCmo.ot` 8, `ftCf`
//! unspecified, `ftPioGrbit` `fAutoPict`, and an `ftPictFmla` naming the
//! ProgID and the `MBDxxxxxxxx` storage), and the storage itself with the
//! `\x01CompObj`, `\x01Ole` and `Equation Native` streams ([MS-XLS]
//! 2.4.181 Obj, 2.5.150 FtPictFmla, 2.5.187 ObjFmla, 2.5.202
//! PictFmlaEmbedInfo, 2.5.198.92 PtgTbl; [MS-OLEDS] 2.3.7 and 2.3.8).

use litchi_core::edit::RefusalKind;

use super::super::super::drawing::Block;
use super::super::super::equation::{
    self as model, EQUATION_NATIVE, Equation, EquationKind, MAX_EQUATION_NATIVE_BYTES,
    StagedStorage,
};
use super::super::super::stream::{Globals, Sheet, u16_at, u32_at};
#[cfg(feature = "equation")]
use super::Error;
use super::refused;
use super::{Edit, Property, Result, encode_fopt, fopt_of, image_kind, io};

/// New content of an OLE equation: its `Equation Native` stream, its
/// preview picture and the size it is shown at.
#[derive(Debug, Clone, PartialEq)]
pub struct EquationContent {
    /// The complete `Equation Native` stream: the 28-byte `EQNOLEFILEHDR`
    /// and the MTEF data.
    pub native: Vec<u8>,
    /// The preview picture (PNG or JPEG) shown in place of the equation.
    pub preview: Vec<u8>,
    /// The shown width and height, in points.
    pub size_pt: (f64, f64),
}

/// An OLE equation to add.
#[derive(Debug, Clone, PartialEq)]
pub struct EquationSpec {
    /// The object's kind: [`EquationKind::MathType`] (MTEF 5) or
    /// [`EquationKind::Equation3`] (MTEF 3).
    pub kind: EquationKind,
    /// The content.
    pub content: EquationContent,
    /// The zero-based row of the top-left cell.
    pub row: u16,
    /// The zero-based column of the top-left cell.
    pub column: u16,
}

/// The ProgID, CLSID (on-disk byte order), `CompObj` user type and
/// clipboard format name of an equation kind, and its MTEF version.
pub(in crate::sheet_features) struct Identity {
    pub(in crate::sheet_features) prog_id: &'static str,
    pub(in crate::sheet_features) clsid: [u8; 16],
    pub(in crate::sheet_features) user_type: &'static str,
    pub(in crate::sheet_features) clipboard: &'static str,
    pub(in crate::sheet_features) mtef_version: u8,
}

pub(in crate::sheet_features) fn identity(kind: EquationKind) -> Option<Identity> {
    match kind {
        EquationKind::MathType => Some(Identity {
            prog_id: "Equation.DSMT4",
            clsid: [
                0x03, 0xCE, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x46,
            ],
            user_type: "MathType 6.0 Equation",
            clipboard: "MathType EF",
            mtef_version: 5,
        }),
        EquationKind::Equation3 => Some(Identity {
            prog_id: "Equation.3",
            clsid: [
                0x02, 0xCE, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x46,
            ],
            user_type: "Microsoft Equation 3.0",
            clipboard: "DS Equation",
            mtef_version: 3,
        }),
        EquationKind::Other => None,
    }
}

/// Check an `Equation Native` stream's `EQNOLEFILEHDR` and that its MTEF
/// version is `version`.
fn check_native(native: &[u8], version: Option<u8>) -> Result<()> {
    let header_ok = native.len() > 28
        && native.len() as u64 <= MAX_EQUATION_NATIVE_BYTES
        && u16_at(native, 0) == 28
        && u32_at(native, 8) as usize == native.len() - 28;
    if !header_ok {
        return Err(refused(
            RefusalKind::InvalidValue,
            "the Equation Native stream has no valid EQNOLEFILEHDR",
        ));
    }
    if version.is_some_and(|version| native[28] != version) {
        return Err(refused(
            RefusalKind::InvalidValue,
            "the MTEF version does not match the equation object's kind",
        ));
    }
    Ok(())
}

fn check_size(size_pt: (f64, f64)) -> Result<()> {
    let valid = |value: f64| value.is_finite() && value > 0.0 && value <= 20_000.0;
    if valid(size_pt.0) && valid(size_pt.1) {
        Ok(())
    } else {
        Err(refused(
            RefusalKind::InvalidValue,
            "the equation size is not a positive number of points",
        ))
    }
}

/// The `CompObj` stream ([MS-OLEDS] 2.3.8) of an equation object, as
/// MathType and Equation Editor write it.
fn comp_obj(identity: &Identity) -> Vec<u8> {
    let mut out = vec![
        0x01, 0x00, 0xFE, 0xFF, 0x03, 0x0A, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF,
    ];
    out.extend_from_slice(&identity.clsid);
    for text in [identity.user_type, identity.clipboard, identity.prog_id] {
        let length = u32::try_from(text.len() + 1).unwrap_or(0);
        out.extend_from_slice(&length.to_le_bytes());
        out.extend_from_slice(text.as_bytes());
        out.push(0);
    }
    // UnicodeMarker, then empty Unicode user type, clipboard format and
    // reserved strings.
    out.extend_from_slice(&0x71B2_39F4u32.to_le_bytes());
    out.extend_from_slice(&[0; 12]);
    out
}

/// The `\x01Ole` stream ([MS-OLEDS] 2.3.7) of an embedded object.
const OLE_STREAM: [u8; 20] = [
    0x01, 0x00, 0x00, 0x02, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00,
];

/// The compound file of a new equation storage.
pub(in crate::sheet_features) fn storage_file(
    kind: EquationKind,
    native: &[u8],
) -> crate::Result<Vec<u8>> {
    let identity = identity(kind)
        .ok_or_else(|| crate::Error::InvalidData("an equation kind has no OLE identity".into()))?;
    let mut writer = litchi_cfb::OleWriter::new();
    writer.set_root_clsid(identity.clsid);
    writer.create_stream(&["\u{1}CompObj"], &comp_obj(&identity))?;
    writer.create_stream(&["\u{1}Ole"], &OLE_STREAM)?;
    writer.create_stream(&[EQUATION_NATIVE], native)?;
    let mut output = std::io::Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

/// The sizes, in points, of a worksheet's columns and rows: the default
/// size and the explicit sizes `ColInfo` and `Row` records give.
struct Grid {
    default_column: f64,
    columns: Vec<(u16, u16, f64)>,
    default_row: f64,
    rows: std::collections::HashMap<u16, f64>,
}

impl Grid {
    fn parse(stream_bytes: &[u8], layout: &Sheet) -> Self {
        // A column of n characters is n * 7 + 5 pixels, rounded up to a
        // multiple of 8, at the default font's 7-pixel digit width.
        let mut default_column = 64.0 * 0.75;
        let mut default_row = 15.0;
        let mut columns = Vec::new();
        let mut rows = std::collections::HashMap::new();
        for record in &layout.records {
            let payload = record.payload(stream_bytes);
            match record.kind {
                0x0055 if payload.len() >= 2 => {
                    let pixels = (u32::from(u16_at(payload, 0)) * 7 + 5).div_ceil(8) * 8;
                    default_column = f64::from(pixels) * 0.75;
                },
                0x0225 if payload.len() >= 4 => {
                    default_row = f64::from(u16_at(payload, 2) & 0x7fff) / 20.0;
                },
                0x007d if payload.len() >= 10 => {
                    let hidden = u16_at(payload, 8) & 1 != 0;
                    let pixels = (u32::from(u16_at(payload, 4)) * 7 + 128) / 256;
                    let width = if hidden {
                        0.0
                    } else {
                        f64::from(pixels) * 0.75
                    };
                    columns.push((u16_at(payload, 0), u16_at(payload, 2), width));
                },
                0x0208 if payload.len() >= 16 => {
                    let hidden = u16_at(payload, 12) & 0x20 != 0;
                    let height = if hidden {
                        0.0
                    } else {
                        f64::from(u16_at(payload, 6) & 0x7fff) / 20.0
                    };
                    rows.insert(u16_at(payload, 0), height);
                },
                _ => {},
            }
        }
        Self {
            default_column,
            columns,
            default_row,
            rows,
        }
    }

    fn column(&self, column: u16) -> f64 {
        self.columns
            .iter()
            .find(|(first, last, _)| (*first..=*last).contains(&column))
            .map_or(self.default_column, |(_, _, width)| *width)
    }

    fn row(&self, row: u16) -> f64 {
        self.rows.get(&row).copied().unwrap_or(self.default_row)
    }

    /// The cell and offset (in `scale` parts of that cell) `length` points
    /// after `offset` parts into `start`.
    fn reach(
        start: u16,
        offset: u16,
        length: f64,
        scale: f64,
        last: u16,
        size: impl Fn(u16) -> f64,
    ) -> (u16, u16) {
        let mut index = start;
        let mut remaining = length + size(start) * f64::from(offset) / scale;
        loop {
            let extent = size(index);
            if remaining < extent || index == last {
                let fraction = if extent > 0.0 {
                    (remaining / extent).clamp(0.0, 1.0)
                } else {
                    0.0
                };
                #[allow(
                    clippy::cast_possible_truncation,
                    clippy::cast_sign_loss,
                    reason = "the fraction is clamped to [0, 1] and scaled to at most 1024"
                )]
                let parts = (fraction * scale).round() as u16;
                return (index, parts.min(scale as u16));
            }
            remaining -= extent;
            index += 1;
        }
    }

    /// An `OfficeArtClientAnchorSheet` (move without resizing, as Excel
    /// anchors an OLE object) from a top-left cell and offsets to a size in
    /// points.
    fn anchor(&self, column: (u16, u16), row: (u16, u16), size_pt: (f64, f64)) -> [u8; 18] {
        let (last_column, dx) = Self::reach(column.0, column.1, size_pt.0, 1024.0, 255, |index| {
            self.column(index)
        });
        let (last_row, dy) = Self::reach(row.0, row.1, size_pt.1, 256.0, u16::MAX, |index| {
            self.row(index)
        });
        let mut out = [0u8; 18];
        for (index, value) in [
            2u16,
            column.0,
            column.1,
            row.0,
            row.1,
            last_column,
            dx,
            last_row,
            dy,
        ]
        .into_iter()
        .enumerate()
        {
            out[index * 2..index * 2 + 2].copy_from_slice(&value.to_le_bytes());
        }
        out
    }
}

impl Edit {
    /// The drawing positions of the OLE equations of a worksheet.
    fn equation_positions(&self, sheet: usize) -> Result<(Block, Vec<usize>)> {
        let (_, _, block, _) = self.drawing(sheet)?;
        let block = block.ok_or(refused(
            RefusalKind::NotFound,
            "the worksheet has no equation",
        ))?;
        let positions = model::positions(&block);
        Ok((block, positions))
    }

    /// The staged `Equation Native` stream of a storage, else the
    /// source's, or why the source's could not be read.
    fn staged_native(&self, storage: &str) -> Result<(Option<Vec<u8>>, Option<String>)> {
        match self.storages.iter().find(|(name, _)| name == storage) {
            Some((_, StagedStorage::New { native, .. } | StagedStorage::Native(native))) => {
                Ok((Some(native.clone()), None))
            },
            Some((_, StagedStorage::File(_) | StagedStorage::Removed)) => Ok((None, None)),
            None => Ok(model::read_native(&self.base, storage)?),
        }
    }

    /// The staged OLE equations of a worksheet, in drawing order.
    ///
    /// # Errors
    ///
    /// Returns an error when the position is not a worksheet or its records
    /// are malformed.
    pub fn equations(&mut self, sheet: usize) -> Result<Vec<Equation>> {
        self.ensure_ready()?;
        let (block, store) = model::drawing(&self.stream, sheet)?;
        let Some(block) = block else {
            return Ok(Vec::new());
        };
        let mut out = Vec::new();
        for position in model::positions(&block) {
            let object = &block.objects[position];
            let Some(ole) = object.ole.as_ref() else {
                continue;
            };
            let (native, error) = self.staged_native(&ole.storage_name())?;
            if let Some(mut equation) = model::equation_of(object, &store, native) {
                equation.native_error = error;
                out.push(equation);
            }
        }
        Ok(out)
    }

    /// Add an OLE equation anchored at a cell, sized to `size_pt`; returns
    /// its position among the worksheet's equations.
    ///
    /// # Errors
    ///
    /// Returns [`Refused`](super::Error::Refused) for a kind other than MathType or
    /// Equation Editor, an `Equation Native` stream that is malformed or of
    /// another MTEF version, a preview that is not PNG or JPEG, a size that
    /// is not positive, or a cell outside the grid.
    pub fn add_equation(&mut self, sheet: usize, spec: &EquationSpec) -> Result<usize> {
        self.ensure_ready()?;
        self.step()?;
        let identity = identity(spec.kind).ok_or(refused(
            RefusalKind::InvalidValue,
            "a new equation is a MathType or Equation Editor object",
        ))?;
        check_native(&spec.content.native, Some(identity.mtef_version))?;
        check_size(spec.content.size_pt)?;
        if spec.column > 255 {
            return Err(refused(
                RefusalKind::InvalidValue,
                "the anchor is outside the BIFF8 grid",
            ));
        }
        if image_kind(&spec.content.preview).is_none() {
            return Err(refused(
                RefusalKind::InvalidValue,
                "the preview is not PNG or JPEG",
            ));
        }
        let count = self.equations(sheet)?.len();
        let (_, layout) = self.sheet(sheet)?;
        let anchor = Grid::parse(&self.stream, &layout).anchor(
            (spec.column, 0),
            (spec.row, 0),
            spec.content.size_pt,
        );
        self.add_embedded(
            sheet,
            identity.prog_id,
            &spec.content.preview,
            anchor,
            StagedStorage::New {
                kind: spec.kind,
                native: spec.content.native.clone(),
            },
        )?;
        Ok(count)
    }

    /// Replace an OLE equation's content: its `Equation Native` stream,
    /// its preview picture and its size (the top-left corner stays). The
    /// object keeps its kind, so the stream must carry the MTEF version of
    /// that kind.
    ///
    /// # Errors
    ///
    /// Returns [`Refused`](super::Error::Refused) when there is no such equation, the
    /// object has no `Equation Native` stream, or the content is invalid.
    pub fn set_equation(
        &mut self,
        sheet: usize,
        index: usize,
        content: &EquationContent,
    ) -> Result<()> {
        self.ensure_ready()?;
        self.step()?;
        check_size(content.size_pt)?;
        let kind = image_kind(&content.preview).ok_or(refused(
            RefusalKind::InvalidValue,
            "the preview is not PNG or JPEG",
        ))?;
        let (block, positions) = self.equation_positions(sheet)?;
        let position = *positions.get(index).ok_or(refused(
            RefusalKind::NotFound,
            "the equation does not exist",
        ))?;
        let object = block.objects[position].clone();
        let ole = object.ole.clone().ok_or(refused(
            RefusalKind::NotFound,
            "the equation does not exist",
        ))?;
        let storage = ole.storage_name();
        let current = self.staged_native(&storage)?.0.ok_or(refused(
            RefusalKind::UnsupportedContent,
            "the OLE equation has no Equation Native stream",
        ))?;
        let version = identity(EquationKind::from_prog_id(&ole.class)).map_or_else(
            || current.get(28).copied(),
            |identity| Some(identity.mtef_version),
        );
        check_native(&content.native, version)?;
        // The preview: replace the blip, and point the shape at it.
        let old = object.prop(0x0104).map(|prop| prop.value);
        let pib = self.replace_blip(old, kind, &content.preview)?;
        if old != Some(pib) {
            let (_, _, block, _) = self.drawing(sheet)?;
            let block = block.ok_or(refused(
                RefusalKind::NotFound,
                "the worksheet has no drawing",
            ))?;
            let (start, end, mut props) = fopt_of(&block.ds, &block.objects[position])?;
            props.retain(|prop| prop.0 != 0x0104);
            props.push((0x0104, Property::Blip(pib)));
            props.sort_by_key(|prop| prop.0);
            let fopt = encode_fopt(&props)?;
            self.rewrite_object(sheet, position, start, end, &fopt)?;
        }
        // The anchor keeps its top-left corner and takes the new size.
        let (_, _, block, _) = self.drawing(sheet)?;
        let block = block.ok_or(refused(
            RefusalKind::NotFound,
            "the worksheet has no drawing",
        ))?;
        let object = &block.objects[position];
        if let (Some(bytes), Some(at)) = (object.anchor, object.anchor_at) {
            let (_, layout) = self.sheet(sheet)?;
            let mut anchor = Grid::parse(&self.stream, &layout).anchor(
                (u16_at(&bytes, 2), u16_at(&bytes, 4)),
                (u16_at(&bytes, 6), u16_at(&bytes, 8)),
                content.size_pt,
            );
            anchor[..2].copy_from_slice(&bytes[..2]);
            self.patch_drawing(sheet, at, &anchor)?;
        }
        self.stage(storage, StagedStorage::Native(content.native.clone()));
        Ok(())
    }

    /// Remove an OLE equation: its shape, its `Obj`, its preview picture's
    /// blip reference and its storage.
    ///
    /// # Errors
    ///
    /// Returns [`Refused`](super::Error::Refused) when there is no such equation.
    pub fn remove_equation(&mut self, sheet: usize, index: usize) -> Result<()> {
        self.ensure_ready()?;
        self.step()?;
        let (_, positions) = self.equation_positions(sheet)?;
        let position = *positions.get(index).ok_or(refused(
            RefusalKind::NotFound,
            "the equation does not exist",
        ))?;
        self.remove_embedded(sheet, position)
    }

    /// Point a picture at new image bytes: the blip `old` is replaced in
    /// place when nothing else uses it, else released; returns the blip's
    /// one-based index.
    fn replace_blip(
        &mut self,
        old: Option<u32>,
        kind: litchi_odraw::image::Kind,
        data: &[u8],
    ) -> Result<u32> {
        let globals = Globals::parse(&self.stream)?;
        let group = Self::group(&self.stream, &globals)?;
        let entries = match &group {
            Some(group) => group.store()?,
            None => None,
        };
        let sole = old.and_then(|pib| {
            let (store_at, entries) = entries.as_ref()?;
            let index = (pib as usize).checked_sub(1)?;
            let entry = entries.get(index)?;
            let refs = u32_at(&group.as_ref()?.bytes, entry.at + 8 + 24);
            (entry.fbse && entry.win != 0 && refs == 1).then_some((
                *store_at,
                entry.at,
                entry.length,
            ))
        });
        let (Some((store_at, entry_at, length)), Some(group), Some(pib)) = (sole, group, old)
        else {
            if let Some(pib) = old {
                self.release_blip(pib)?;
            }
            return self.add_blip(kind, data);
        };
        let entry = Self::blip_entry(kind, data)?;
        let range = (
            globals.records[group.first].start,
            globals.records[group.end - 1].end,
        );
        let mut bytes = group.bytes.clone();
        // `check` keeps the entry inside the store and the store inside the
        // container.
        let removed = 8 + length;
        let delta = i64::try_from(entry.len()).unwrap_or(i64::MAX)
            - i64::try_from(removed).unwrap_or(i64::MAX);
        let grow = |value: u32| {
            i64::from(value)
                .checked_add(delta)
                .and_then(|value| u32::try_from(value).ok())
                .ok_or(refused(
                    RefusalKind::InvalidValue,
                    "a drawing-group count or length is at its limit",
                ))
        };
        let store_length = grow(u32_at(&bytes, store_at + 4))?;
        let container = grow(u32_at(&bytes, 4))?;
        bytes.splice(entry_at..entry_at + removed, entry.iter().copied());
        bytes[store_at + 4..store_at + 8].copy_from_slice(&store_length.to_le_bytes());
        bytes[4..8].copy_from_slice(&container.to_le_bytes());
        self.put_group(bytes, Some(range), group.continuation)?;
        Ok(pib)
    }

    /// The `OfficeArtFBSE` (with its embedded blip) of an image, referenced
    /// once.
    fn blip_entry(kind: litchi_odraw::image::Kind, data: &[u8]) -> Result<Vec<u8>> {
        let mut store = litchi_odraw::image::builder::StoreBuilder::new();
        let blip = litchi_odraw::image::builder::BlipBuilder::bitmap(kind, data)
            .map_err(|_| refused(RefusalKind::InvalidValue, "the image cannot be stored"))?;
        store
            .add_embedded(blip)
            .map_err(|_| refused(RefusalKind::InvalidValue, "the image cannot be stored"))?;
        let mut written = Vec::new();
        store
            .write_store(&mut written, litchi_odraw::image::Limits::default())
            .map_err(io)?;
        Ok(written
            .get(8..)
            .ok_or(refused(
                RefusalKind::InvalidValue,
                "the image cannot be stored",
            ))?
            .to_vec())
    }
}

/// Where an equation's content comes from.
#[cfg(feature = "equation")]
#[derive(Debug, Clone, Copy)]
pub enum EquationSource<'a> {
    /// LaTeX math.
    Latex(&'a str),
    /// An OMML equation (`m:oMath` or `m:oMathPara`).
    Omml(&'a str),
    /// MTEF data: a complete `Equation Native` stream, or bare MTEF that
    /// is given the `EQNOLEFILEHDR`.
    Mtef(&'a [u8]),
}

/// The `Equation Native` stream, kind and conversion losses of a source.
#[cfg(feature = "equation")]
fn native_of(
    source: EquationSource<'_>,
    target: &litchi_equation::MtefTarget,
    limits: &litchi_equation::Limits,
) -> Result<(Vec<u8>, u8, Vec<String>)> {
    let native = match source {
        EquationSource::Latex(latex) => {
            litchi_equation::latex_to_equation_native(latex, target, limits)
                .map_err(Error::Equation)?
        },
        EquationSource::Omml(omml) => {
            litchi_equation::omml_to_equation_native(omml, target, limits)
                .map_err(Error::Equation)?
        },
        EquationSource::Mtef(bytes) => {
            let read = litchi_equation::read_mtef(bytes, limits).map_err(Error::Equation)?;
            let info = read
                .info()
                .ok_or(refused(RefusalKind::InvalidValue, "the MTEF data is empty"))?;
            let native = if info.ole_header.is_some() {
                bytes.to_vec()
            } else {
                with_ole_header(bytes)?
            };
            return Ok((native, info.version, Vec::new()));
        },
    };
    let version = native.info().version;
    let losses = native.losses().to_vec();
    Ok((native.into_bytes(), version, losses))
}

/// Bare MTEF data behind an `EQNOLEFILEHDR` (MTEF specification, "OLE
/// equation objects").
#[cfg(feature = "equation")]
fn with_ole_header(mtef: &[u8]) -> Result<Vec<u8>> {
    let length = u32::try_from(mtef.len())
        .map_err(|_| refused(RefusalKind::InvalidValue, "the MTEF data is too large"))?;
    let mut out = Vec::with_capacity(28 + mtef.len());
    out.extend_from_slice(&28u16.to_le_bytes());
    out.extend_from_slice(&0x0002_0000u32.to_le_bytes());
    out.extend_from_slice(&0xC295u16.to_le_bytes());
    out.extend_from_slice(&length.to_le_bytes());
    out.extend_from_slice(&[0; 16]);
    out.extend_from_slice(mtef);
    Ok(out)
}

/// The preview of an `Equation Native` stream, drawn at twice the
/// resolution it is shown at, and its size in points.
#[cfg(feature = "equation")]
fn preview_of(native: &[u8], limits: &litchi_equation::Limits) -> Result<(Vec<u8>, (f64, f64))> {
    let options = litchi_equation::PngOptions::DEFAULT
        .with_scale(2.0)
        .map_err(Error::Equation)?;
    let png = litchi_equation::mtef_to_png(native, &options, limits).map_err(Error::Equation)?;
    let size = png.size_pt();
    Ok((png.into_bytes(), size))
}

#[cfg(feature = "equation")]
impl Edit {
    /// Add an equation from LaTeX, OMML or MTEF, anchored at a cell: a
    /// MathType object (MTEF data keeps its own kind) with a rendered
    /// preview. Returns its position among the worksheet's equations and
    /// what the conversion could not carry over.
    ///
    /// # Errors
    ///
    /// Returns [`Equation`](super::Error::Equation) error when the source does not convert or
    /// render, and the errors of [`Edit::add_equation`].
    pub fn add_equation_from(
        &mut self,
        sheet: usize,
        source: EquationSource<'_>,
        row: u16,
        column: u16,
        limits: &litchi_equation::Limits,
    ) -> Result<(usize, Vec<String>)> {
        let (native, version, losses) =
            native_of(source, &litchi_equation::MtefTarget::MathType, limits)?;
        let (preview, size_pt) = preview_of(&native, limits)?;
        let kind = if version < 5 {
            EquationKind::Equation3
        } else {
            EquationKind::MathType
        };
        let index = self.add_equation(
            sheet,
            &EquationSpec {
                kind,
                content: EquationContent {
                    native,
                    preview,
                    size_pt,
                },
                row,
                column,
            },
        )?;
        Ok((index, losses))
    }

    /// Replace an equation's content from LaTeX, OMML or MTEF, keeping the
    /// object's kind (an Equation Editor 3.0 object stays MTEF 3); the
    /// preview is redrawn and the anchor resized. Returns what the
    /// conversion could not carry over.
    ///
    /// # Errors
    ///
    /// Returns [`Equation`](super::Error::Equation) error when the source does not convert or
    /// render, and the errors of [`Edit::set_equation`].
    pub fn set_equation_from(
        &mut self,
        sheet: usize,
        index: usize,
        source: EquationSource<'_>,
        limits: &litchi_equation::Limits,
    ) -> Result<Vec<String>> {
        let equations = self.equations(sheet)?;
        let equation = equations.get(index).ok_or(refused(
            RefusalKind::NotFound,
            "the equation does not exist",
        ))?;
        let current = equation.native.as_deref().ok_or(refused(
            RefusalKind::UnsupportedContent,
            "the OLE equation has no Equation Native stream",
        ))?;
        let info = litchi_equation::read_mtef(current, limits)
            .map_err(Error::Equation)?
            .info()
            .cloned()
            .ok_or(refused(
                RefusalKind::UnsupportedContent,
                "the OLE equation's MTEF data is empty",
            ))?;
        let target = litchi_equation::MtefTarget::SameAs(info);
        let (native, _, losses) = native_of(source, &target, limits)?;
        let (preview, size_pt) = preview_of(&native, limits)?;
        self.set_equation(
            sheet,
            index,
            &EquationContent {
                native,
                preview,
                size_pt,
            },
        )?;
        Ok(losses)
    }
}

#[cfg(test)]
mod tests {
    use super::super::super::super::drawing::ole_link;
    use super::super::embedded::embedded_obj;
    use super::*;

    #[test]
    fn equation_obj_matches_the_excel_layout() {
        let obj = embedded_obj(3, "Equation.DSMT4", 0x00DD_3649).unwrap();
        // Obj header, then ftCmo.
        assert_eq!(&obj[..4], &[0x5d, 0x00, 0x4e, 0x00]);
        assert_eq!(&obj[4..10], &[0x15, 0x00, 0x12, 0x00, 0x08, 0x00]);
        let payload = &obj[4..];
        let link = ole_link(payload, 0).unwrap();
        assert_eq!(link.class, "Equation.DSMT4");
        assert_eq!(link.storage_name(), "MBD00DD3649");
        assert!(link.is_equation());
        // ftCf unspecified, ftPioGrbit fAutoPict.
        assert_eq!(
            &payload[22..34],
            &[7, 0, 2, 0, 0xff, 0xff, 8, 0, 2, 0, 1, 0]
        );
        // The typed OLE-object model reads it back as a valid MBD object.
        let record = crate::ole_object::OleObjectRecord::parse(payload, None).expect("typed Obj");
        assert_eq!(record.storage_name().as_deref(), Some("MBD00DD3649"));
        record.validate().unwrap();
    }

    #[test]
    fn comp_obj_matches_mathtype() {
        let bytes = comp_obj(&identity(EquationKind::MathType).unwrap());
        assert_eq!(&bytes[12..16], &[0x03, 0xCE, 0x02, 0x00]);
        let text = String::from_utf8_lossy(&bytes);
        assert!(text.contains("MathType EF\0"));
        assert!(text.contains("Equation.DSMT4\0"));
        assert!(bytes.ends_with(&[0xF4, 0x39, 0xB2, 0x71, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]));
    }

    #[test]
    fn grid_reaches_the_cell_holding_the_far_corner() {
        let grid = Grid {
            default_column: 48.0,
            columns: vec![(2, 2, 96.0)],
            default_row: 15.0,
            rows: std::collections::HashMap::new(),
        };
        // 120 pt from the start of column 1: 48 in column 1, 72 of column 2's 96.
        let anchor = grid.anchor((1, 0), (0, 0), (120.0, 15.0));
        assert_eq!(u16_at(&anchor, 10), 2);
        assert_eq!(u16_at(&anchor, 12), 768);
        assert_eq!(u16_at(&anchor, 14), 1);
        assert_eq!(u16_at(&anchor, 16), 0);
    }
}
