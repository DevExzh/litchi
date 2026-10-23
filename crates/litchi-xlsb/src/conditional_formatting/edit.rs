//! Immutable conditional-formatting edits of an XLSB [`Package`]
//! (unified-ops record 0322).
//!
//! A worksheet stores each classic conditional format as one
//! `BrtBeginConditionalFormatting` collection (its ranges and rules).
//! [`Edit`] lists, adds, replaces and removes those collections one at a
//! time; every other record keeps its bytes. Office 2013 formats
//! (`BrtBeginConditionalFormatting14`, stored with the future records) are
//! not listed here and keep their bytes. A rule's formatting is a
//! differential format (`BrtDXF`) of the styles part:
//! [`Edit::add_differential`] appends one and returns its index.
//!
//! Formula text of an authored rule (`Rule::formula_texts` with no binary
//! formulas) is compiled with the workbook formula compiler; relative
//! references are stored relative to the top-left cell of the first range,
//! as the reader renders them.
//!
//! ```
//! use litchi_xlsb::Package;
//! use litchi_xlsb::conditional_formatting::{Formatting, Rule, RuleType};
//! use litchi_xlsb::conditional_formatting::edit::DifferentialStyle;
//!
//! # fn main() -> Result<(), litchi_xlsb::package::PackageError> {
//! let package = Package::new()?;
//! let mut edit = package.edit_conditional_formats()?;
//! let style = DifferentialStyle { fill: Some(0xFFC7CE), ..DifferentialStyle::default() };
//! let dxf = edit.add_differential(&style)?;
//! let mut rule = Rule::new(RuleType::CellIs, 1);
//! rule.operator = Some(5);
//! rule.formula_texts = vec!["5".to_owned()];
//! rule.dxf_id = Some(dxf);
//! let mut format = Formatting::new(vec!["A1:A9".to_owned()]);
//! format.add_rule(rule);
//! assert_eq!(edit.add(0, format)?, 0);
//! let after = edit.commit()?.into_package();
//! let edit = after.edit_conditional_formats()?;
//! assert_eq!(edit.formats(0)?[0].ranges, ["A1:A9"]);
//! assert_eq!(edit.differential_formats()?[dxf as usize], style);
//! # Ok(())
//! # }
//! ```

use litchi_opc::PackURI;
use litchi_opc::constants::relationship_type;

use super::{Formatting, RecordKind, write_conditional_formattings};
use crate::package::Package;
use crate::package::error::{Error as PackageError, Result};
use crate::part_edit::{Rank, SheetLayout, Staged, family, invalid, snapshot_options, splice};
use crate::raw::{Encoder, Records, kind};

family!("conditional-format");
snapshot_options!("conditional-format");

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Classic conditional formats of one worksheet ([`MAX_FORMATS`]).
    Formats,
    /// Differential formats of the styles part ([`MAX_DIFFERENTIALS`]).
    Differentials,
}

/// The most classic conditional formats one worksheet holds through this
/// edit.
pub const MAX_FORMATS: usize = 65_534;
/// The most differential formats the styles part holds through this edit.
pub const MAX_DIFFERENTIALS: usize = 65_534;

/// The underline of a differential font.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Underline {
    /// No underline.
    None,
    /// A single underline.
    Single,
    /// A double underline.
    Double,
    /// A single accounting underline.
    SingleAccounting,
    /// A double accounting underline.
    DoubleAccounting,
}

impl Underline {
    const fn code(self) -> u16 {
        match self {
            Self::None => 0,
            Self::Single => 1,
            Self::Double => 2,
            Self::SingleAccounting => 0x21,
            Self::DoubleAccounting => 0x22,
        }
    }

    const fn from_code(code: u16) -> Option<Self> {
        Some(match code {
            0 => Self::None,
            1 => Self::Single,
            2 => Self::Double,
            0x21 => Self::SingleAccounting,
            0x22 => Self::DoubleAccounting,
            _ => return None,
        })
    }
}

/// The formatting of a differential format (`BrtDXF`) that conditional
/// formats apply: a solid fill and font attributes. Colors are `0xRRGGBB`.
/// Properties this model does not name (borders, number formats, theme
/// colors) read as absent and are never authored.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct DifferentialStyle {
    /// The solid fill color.
    pub fill: Option<u32>,
    /// The font color.
    pub font_color: Option<u32>,
    /// Bold.
    pub bold: Option<bool>,
    /// Italic.
    pub italic: Option<bool>,
    /// Strikethrough.
    pub strike: Option<bool>,
    /// The underline.
    pub underline: Option<Underline>,
    /// The font size in points.
    pub size: Option<f64>,
    /// The font name.
    pub name: Option<String>,
}

const FILL_PATTERN: u16 = 0x0000;
const FOREGROUND: u16 = 0x0001;
const BACKGROUND: u16 = 0x0002;
const FONT_COLOR: u16 = 0x0005;
const FONT_NAME: u16 = 0x0018;
const FONT_BOLD: u16 = 0x0019;
const FONT_UNDERLINE: u16 = 0x001A;
const FONT_ITALIC: u16 = 0x001C;
const FONT_STRIKE: u16 = 0x001D;
const FONT_HEIGHT: u16 = 0x0024;

impl DifferentialStyle {
    /// Decode a `BrtDXF` payload ([MS-XLSB] sections 2.4.356, 2.5.159 and
    /// 2.5.164) under an explicit raw profile.
    ///
    /// # Errors
    ///
    /// Returns an error for a truncated property array, a payload above the
    /// profile's payload ceiling or a font name above its UTF-16-unit ceiling.
    pub fn parse(payload: &[u8], limits: crate::raw::Limits) -> Result<Self> {
        let limits = admit_payload(payload.len(), limits)?;
        let short = || invalid("a differential format is truncated");
        let u16_at = |offset: usize| -> Result<u16> {
            payload
                .get(offset..offset + 2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
                .ok_or_else(short)
        };
        let count = u16_at(4)?;
        let mut style = Self::default();
        let mut offset = 6usize;
        let mut foreground = None;
        let mut background = None;
        for _ in 0..count {
            let kind = u16_at(offset)?;
            let size = usize::from(u16_at(offset + 2)?);
            if size < 4 {
                return Err(short());
            }
            let data = payload.get(offset + 4..offset + size).ok_or_else(short)?;
            let color = || {
                (data.len() == 8 && data[0] >> 1 == 0x02)
                    .then(|| u32::from_be_bytes([0, data[4], data[5], data[6]]))
            };
            let flag = || data.first().map(|value| *value != 0);
            match kind {
                FOREGROUND => foreground = color(),
                BACKGROUND => background = color(),
                FONT_COLOR => style.font_color = color(),
                FONT_BOLD => {
                    style.bold = data
                        .get(..2)
                        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]) >= 0x02BC);
                },
                FONT_ITALIC => style.italic = flag(),
                FONT_STRIKE => style.strike = flag(),
                FONT_UNDERLINE => {
                    style.underline = data.get(..2).and_then(|bytes| {
                        Underline::from_code(u16::from_le_bytes([bytes[0], bytes[1]]))
                    });
                },
                FONT_HEIGHT => {
                    style.size = data.get(..4).map(|bytes| {
                        f64::from(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
                            / 20.0
                    });
                },
                FONT_NAME => {
                    if let Some(count) = data.get(..2) {
                        admit_units(
                            usize::from(u16::from_le_bytes([count[0], count[1]])),
                            limits,
                            offset + 4,
                        )?;
                    }
                    style.name = data.get(..2).and_then(|count| {
                        let units = usize::from(u16::from_le_bytes([count[0], count[1]]));
                        let text = data.get(2..2 + units * 2)?;
                        let units: Vec<u16> = text
                            .chunks_exact(2)
                            .map(|unit| u16::from_le_bytes([unit[0], unit[1]]))
                            .collect();
                        String::from_utf16(&units).ok()
                    });
                },
                _ => {},
            }
            offset += size;
        }
        // A conditional format's solid fill is its background color
        // (`<bgColor>` in SpreadsheetML), the foreground when only that is
        // set.
        style.fill = background.or(foreground);
        Ok(style)
    }

    /// Encode as a `BrtDXF` payload under an explicit raw output profile.
    ///
    /// # Errors
    ///
    /// Returns an error for a color above `0xFFFFFF`, a size outside 1–409.55
    /// points, a font name longer than 32 characters, or output above the
    /// profile's payload or UTF-16-unit ceilings.
    pub fn encode(&self, limits: crate::raw::Limits) -> Result<Vec<u8>> {
        let limits = limits.validate()?;
        let mut props: Vec<(u16, Vec<u8>)> = Vec::new();
        let color = |rgb: u32| -> Result<Vec<u8>> {
            if rgb > 0x00FF_FFFF {
                return Err(invalid("a differential color is not 0xRRGGBB"));
            }
            let [_, red, green, blue] = rgb.to_be_bytes();
            Ok(vec![0x05, 0, 0, 0, red, green, blue, 0xFF])
        };
        if let Some(fill) = self.fill {
            props.push((FILL_PATTERN, vec![0x01]));
            props.push((BACKGROUND, color(fill)?));
        }
        if let Some(rgb) = self.font_color {
            props.push((FONT_COLOR, color(rgb)?));
        }
        if let Some(name) = &self.name {
            let units: Vec<u16> = name.encode_utf16().collect();
            if units.is_empty() || units.len() > 32 {
                return Err(invalid("a differential font name has 1 to 32 characters"));
            }
            admit_units(units.len(), limits, 0)?;
            let mut data = u16::try_from(units.len())
                .map_err(|_overflow| invalid("font name"))?
                .to_le_bytes()
                .to_vec();
            data.extend(units.iter().flat_map(|unit| unit.to_le_bytes()));
            props.push((FONT_NAME, data));
        }
        if let Some(bold) = self.bold {
            let weight: u16 = if bold { 0x02BC } else { 0x0190 };
            props.push((FONT_BOLD, weight.to_le_bytes().to_vec()));
        }
        if let Some(underline) = self.underline {
            props.push((FONT_UNDERLINE, underline.code().to_le_bytes().to_vec()));
        }
        if let Some(italic) = self.italic {
            props.push((FONT_ITALIC, vec![u8::from(italic)]));
        }
        if let Some(strike) = self.strike {
            props.push((FONT_STRIKE, vec![u8::from(strike)]));
        }
        if let Some(size) = self.size {
            let twips = (size * 20.0).round();
            if !(20.0..=8191.0).contains(&twips) {
                return Err(invalid(
                    "a differential font size is outside 1–409.55 points",
                ));
            }
            #[allow(
                clippy::cast_possible_truncation,
                clippy::cast_sign_loss,
                reason = "the range check above bounds the value"
            )]
            let twips = twips as u32;
            props.push((FONT_HEIGHT, twips.to_le_bytes().to_vec()));
        }
        let count =
            u16::try_from(props.len()).map_err(|_overflow| invalid("too many properties"))?;
        let mut payload = Vec::with_capacity(6 + props.len() * 12);
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&count.to_le_bytes());
        for (kind, data) in props {
            let size =
                u16::try_from(data.len() + 4).map_err(|_overflow| invalid("property size"))?;
            payload.extend_from_slice(&kind.to_le_bytes());
            payload.extend_from_slice(&size.to_le_bytes());
            payload.extend_from_slice(&data);
        }
        admit_payload(payload.len(), limits)?;
        Ok(payload)
    }
}

/// Validate `limits` and admit a `BrtDXF` payload of `length` bytes.
fn admit_payload(length: usize, limits: crate::raw::Limits) -> Result<crate::raw::Limits> {
    Ok(limits.admit(length)?)
}

/// Refuse a font name above the profile's UTF-16-unit ceiling.
fn admit_units(units: usize, limits: crate::raw::Limits, offset: usize) -> Result<()> {
    Ok(limits.admit_units(units, offset)?)
}

/// A staged conditional-format edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    options: EditOptions,
}

impl Package {
    /// Start a conditional-format edit.
    ///
    /// # Errors
    ///
    /// Returns the workbook materialization errors of this package.
    pub fn edit_conditional_formats(&self) -> Result<Edit> {
        self.edit_conditional_formats_with_options(EditOptions::default())
    }

    /// Start a conditional-format edit with an explicit retained-payload ceiling.
    pub fn edit_conditional_formats_with_options(&self, options: EditOptions) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
            options,
        })
    }

    /// Apply a conditional-format patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_conditional_formats(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

/// The styles part of a workbook package.
pub(crate) fn styles_part(package: &litchi_opc::OpcPackage) -> Result<Option<PackURI>> {
    let workbook = package.main_document_part()?;
    for relationship in workbook.rels().iter() {
        if !relationship.is_external()
            && matches!(
                relationship.reltype(),
                relationship_type::STYLES | relationship_type::STRICT_STYLES
            )
        {
            return Ok(Some(relationship.target_partname()?));
        }
    }
    Ok(None)
}

/// The differential formats of a styles stream.
pub(crate) fn read_differentials(
    styles: &[u8],
    limits: crate::raw::Limits,
) -> Result<Vec<DifferentialStyle>> {
    crate::cell_values::resources::collection(
        styles,
        kind::BEGIN_DXFS,
        kind::END_DXFS,
        kind::DXF,
        limits,
    )?
    .iter()
    .map(|payload| DifferentialStyle::parse(payload, limits))
    .collect()
}

/// The classic conditional formats of a worksheet, in stream order.
pub(crate) fn classic(formats: Vec<Formatting>) -> Vec<Formatting> {
    formats
        .into_iter()
        .filter(|format| format.record_kind == RecordKind::Classic)
        .collect()
}

impl Edit {
    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// The staged classic conditional formats of the worksheet at
    /// complete-catalog position `sheet`, in stream order.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab or a malformed
    /// stream.
    pub fn formats(&self, sheet: usize) -> Result<Vec<Formatting>> {
        let worksheet = self.staged.worksheet(sheet)?;
        Ok(classic(
            self.staged.feature_pass(&worksheet, sheet)?.formats,
        ))
    }

    /// The staged differential formats of the styles part.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed styles part.
    pub fn differential_formats(&self) -> Result<Vec<DifferentialStyle>> {
        match styles_part(self.staged.opc())? {
            Some(styles) => {
                read_differentials(self.staged.blob(&styles)?, self.staged.raw_limits())
            },
            None => Ok(Vec::new()),
        }
    }

    /// Append a differential format to the styles part (creating its
    /// collection at its stream position when absent); returns its index.
    ///
    /// # Errors
    ///
    /// Returns an error for a style [`DifferentialStyle::encode`] refuses, a
    /// workbook without a styles part, more than [`MAX_DIFFERENTIALS`]
    /// formats, or a changed edit of a signed workbook.
    pub fn add_differential(&mut self, style: &DifferentialStyle) -> Result<u32> {
        let payload = style.encode(self.staged.raw_limits())?;
        let styles = styles_part(self.staged.opc())?
            .ok_or_else(|| invalid("the workbook has no styles part"))?;
        let mut bytes = self.staged.blob(&styles)?.to_vec();
        let existing = crate::cell_values::resources::collection(
            &bytes,
            kind::BEGIN_DXFS,
            kind::END_DXFS,
            kind::DXF,
            self.staged.raw_limits(),
        )?;
        if existing.len() >= MAX_DIFFERENTIALS {
            return Err(invalid(
                "the styles part holds the most differential formats it can",
            ));
        }
        let limits = self.staged.raw_limits();
        if !Records::new(&bytes, limits)
            .any(|record| record.is_ok_and(|record| record.kind() == kind::BEGIN_DXFS))
        {
            bytes = insert_dxfs(&bytes, limits)?;
        }
        let bytes = crate::cell_values::resources::append_collection_item(
            &bytes,
            kind::BEGIN_DXFS,
            kind::END_DXFS,
            kind::DXF,
            &payload,
            limits,
        )?;
        self.staged.replace_blob(&styles, bytes)?;
        u32::try_from(existing.len()).map_err(|_overflow| invalid("too many differential formats"))
    }

    /// The priorities of every rule on the worksheet except the classic
    /// format at `skip`.
    fn priorities(&self, sheet: usize, skip: Option<usize>) -> Result<Vec<i64>> {
        let worksheet = self.staged.worksheet(sheet)?;
        let mut priorities = Vec::new();
        let classic = self
            .staged
            .feature_pass(&worksheet, sheet)?
            .formats
            .into_iter()
            .scan(0usize, |classic_index, format| {
                let position = (format.record_kind == RecordKind::Classic).then(|| {
                    let position = *classic_index;
                    *classic_index += 1;
                    position
                });
                Some((position, format))
            });
        for (position, format) in classic {
            if position.is_some() && position == skip {
                continue;
            }
            for rule in &format.rules {
                priorities.push(
                    rule.extension14
                        .map_or(i64::from(rule.priority), |metadata| {
                            i64::from(metadata.priority)
                        }),
                );
            }
        }
        Ok(priorities)
    }

    fn encode(&self, sheet: usize, skip: Option<usize>, format: &Formatting) -> Result<Vec<u8>> {
        if format.record_kind != RecordKind::Classic {
            return Err(invalid(
                "only classic conditional formats are authored; Office 2013 formats are kept as they are",
            ));
        }
        let taken = self.priorities(sheet, skip)?;
        if format
            .rules
            .iter()
            .any(|rule| taken.contains(&i64::from(rule.priority)))
        {
            return Err(invalid("a rule priority is already used on the worksheet"));
        }
        let mut format = format.clone();
        let base = first_cell(&format)?;
        for rule in &mut format.rules {
            if rule.formulas.is_empty() && !rule.formula_texts.is_empty() {
                for text in &rule.formula_texts {
                    let compiled = crate::package::formula::text::Compiler::compile_shared(
                        text, base.0, base.1,
                    )?;
                    rule.formulas.push(compiled.rgce);
                    rule.formula_extras.push(compiled.rgcb);
                }
            }
        }
        let mut out = Vec::new();
        write_conditional_formattings(
            &mut Encoder::new(&mut out, self.staged.raw_limits()),
            std::slice::from_ref(&format),
        )
        .map_err(map_authoring_error)?;
        Ok(out)
    }

    fn splice(&mut self, sheet: usize, index: Option<usize>, encoded: &[u8]) -> Result<bool> {
        let worksheet = self.staged.worksheet(sheet)?;
        let bytes = self.staged.blob(&worksheet)?;
        let layout = SheetLayout::scan_with_limits(bytes, self.staged.raw_limits())?;
        let units: Vec<_> = layout.units_of(kind::BEGIN_COND_FORMATTING).collect();
        let span = match index {
            Some(index) => units
                .get(index)
                .map(|unit| unit.span.clone())
                .ok_or_else(|| invalid("the conditional format is missing"))?,
            None => {
                let at = units.last().map_or_else(
                    || layout.insertion(Rank::ConditionalFormatting),
                    |unit| unit.span.end,
                );
                at..at
            },
        };
        let updated = splice(bytes, span, encoded);
        self.staged.replace_worksheet(&worksheet, updated)
    }

    /// Append a classic conditional format after the worksheet's last one;
    /// returns its position.
    ///
    /// # Errors
    ///
    /// Returns an error for a format the records cannot encode (an Office
    /// 2013 record kind, an invalid range, rule or formula), a rule priority
    /// already used on the worksheet, more than [`MAX_FORMATS`] formats, or a
    /// changed edit of a signed workbook.
    pub fn add(&mut self, sheet: usize, format: Formatting) -> Result<usize> {
        let count = self.formats(sheet)?.len();
        if count >= MAX_FORMATS {
            return Err(invalid(
                "the worksheet holds the most conditional formats it can",
            ));
        }
        let encoded = self.encode(sheet, None, &format)?;
        self.splice(sheet, None, &encoded)?;
        Ok(count)
    }

    /// Replace the classic conditional format at `index`. Returns whether
    /// anything changed.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::add`], and an error for a missing position.
    pub fn set(&mut self, sheet: usize, index: usize, format: Formatting) -> Result<bool> {
        let encoded = self.encode(sheet, Some(index), &format)?;
        self.splice(sheet, Some(index), &encoded)
    }

    /// Remove the classic conditional format at `index`. Its differential
    /// format stays in the styles part, as Excel and OfficeCLI leave it.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing position, a malformed stream or a
    /// changed edit of a signed workbook.
    pub fn remove(&mut self, sheet: usize, index: usize) -> Result<()> {
        self.splice(sheet, Some(index), &[]).map(|_| ())
    }

    /// Validate the staged package and publish it.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit_with_snapshot_limit(
            self.options.max_snapshot_bytes_value(),
        )?))
    }
}

fn map_authoring_error(error: super::Error) -> PackageError {
    match error {
        super::Error::UnsupportedFeature(_) | super::Error::Unrecognized { .. } => {
            PackageError::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_conditional_formats",
                "the conditional format contains unsupported authored content",
            )
        },
        error @ (super::Error::InvalidLength { .. }
        | super::Error::InvalidFormula(_)
        | super::Error::InvalidCellReference(_)
        | super::Error::Encoding(_)
        | super::Error::Wire(_)
        | super::Error::Formula(_)
        | super::Error::Io(_)) => error.into(),
    }
}

/// The top-left cell `(row, column)` of a format's first range.
fn first_cell(format: &Formatting) -> Result<(u32, u32)> {
    let first = format
        .ranges
        .first()
        .and_then(|ranges| ranges.split_whitespace().next())
        .ok_or_else(|| invalid("a conditional format needs a range"))?;
    let cell = first.split(':').next().unwrap_or(first).replace('$', "");
    let cell = litchi_core::selector::CellRef::parse(&cell)
        .ok_or_else(|| invalid("a conditional-format range is not an A1 range"))?;
    Ok((cell.row(), cell.column()))
}

/// Insert an empty `BrtBeginDXFs` collection at its style-sheet position:
/// before the table styles, the color palette, the future records or the
/// end of the style sheet.
fn insert_dxfs(styles: &[u8], limits: crate::raw::Limits) -> Result<Vec<u8>> {
    const COLOR_PALETTE: u16 = 0x01D9;
    let mut at = None;
    let mut records = Records::new(styles, limits);
    while let Some(record) = records.next() {
        let record = record?;
        if matches!(
            record.kind(),
            kind::BEGIN_TABLE_STYLES | kind::FRT_BEGIN | kind::END_STYLE_SHEET
        ) || record.kind().get() == COLOR_PALETTE
        {
            at = Some(record.offset());
            break;
        }
    }
    let at = at.ok_or_else(|| invalid("the styles part has no end"))?;
    let mut empty =
        crate::part_edit::record_with_limits(kind::BEGIN_DXFS, &0u32.to_le_bytes(), limits)?;
    empty.extend(crate::part_edit::record_with_limits(
        kind::END_DXFS,
        &[],
        limits,
    )?);
    Ok(splice(styles, at..at, &empty))
}

#[cfg(test)]
mod tests {
    use super::{DifferentialStyle, Underline};
    use crate::Package;
    use crate::conditional_formatting::{Formatting, Rule, RuleType};

    #[test]
    fn differential_styles_round_trip() {
        let style = DifferentialStyle {
            fill: Some(0x00FF_C7CE),
            font_color: Some(0x009C_0006),
            bold: Some(true),
            italic: Some(false),
            strike: Some(true),
            underline: Some(Underline::Double),
            size: Some(11.5),
            name: Some("Calibri".to_owned()),
        };
        let payload = style.encode(crate::raw::Limits::DEFAULT).unwrap();
        assert_eq!(
            DifferentialStyle::parse(&payload, crate::raw::Limits::DEFAULT).unwrap(),
            style
        );
    }

    #[test]
    fn add_set_remove_replay() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_conditional_formats().unwrap();
        let dxf = edit
            .add_differential(&DifferentialStyle {
                fill: Some(0x00FF_0000),
                ..DifferentialStyle::default()
            })
            .unwrap();
        let mut rule = Rule::new(RuleType::Expression, 1);
        rule.formula_texts = vec!["B2>5".to_owned()];
        rule.dxf_id = Some(dxf);
        let mut format = Formatting::new(vec!["B2:B9".to_owned()]);
        format.add_rule(rule);
        assert_eq!(edit.add(0, format.clone()).unwrap(), 0);
        // A second format reusing priority 1 is refused.
        assert!(edit.add(0, format.clone()).is_err());
        let (after, patch) = edit.commit().unwrap().into_parts();
        let read = after
            .edit_conditional_formats()
            .unwrap()
            .formats(0)
            .unwrap();
        assert_eq!(read.len(), 1);
        // The XLSB formula renderer parenthesizes every binary operation.
        assert_eq!(read[0].rules[0].formula_texts, ["(B2>5)"]);
        assert_eq!(read[0].rules[0].dxf_id, Some(dxf));
        assert_eq!(
            package
                .apply_conditional_formats(&patch)
                .map(|commit| commit.into_package())
                .unwrap()
                .to_bytes()
                .unwrap(),
            after.to_bytes().unwrap()
        );
        let workbook = after.workbook().unwrap();
        let grid = workbook.grid(0usize).unwrap().unwrap();
        assert_eq!(grid.worksheet().conditional_formattings().len(), 1);

        let mut edit = after.edit_conditional_formats().unwrap();
        let mut changed = read[0].clone();
        changed.ranges = vec!["C3:C4".to_owned()];
        changed.rules[0].formulas.clear();
        changed.rules[0].formula_extras.clear();
        changed.rules[0].formula_texts = vec!["C3<0".to_owned()];
        assert!(edit.set(0, 0, changed).unwrap());
        let set = edit.commit().unwrap().into_package();
        let read = set.edit_conditional_formats().unwrap().formats(0).unwrap();
        assert_eq!(read[0].ranges, ["C3:C4"]);
        assert_eq!(read[0].rules[0].formula_texts, ["(C3<0)"]);

        let mut edit = set.edit_conditional_formats().unwrap();
        edit.remove(0, 0).unwrap();
        let removed = edit.commit().unwrap().into_package();
        let worksheet = |package: &Package| {
            package
                .opc_package()
                .get_part(&litchi_opc::PackURI::new("/xl/worksheets/sheet1.bin").unwrap())
                .unwrap()
                .blob()
                .to_vec()
        };
        assert_eq!(worksheet(&removed), worksheet(&package));
    }
}
