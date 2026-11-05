//! Immutable conditional-formatting edits of a [`Package`] (unified-ops
//! record 0320).
//!
//! [`Package::edit_conditional_formats`] starts an [`Edit`] over an immutable
//! snapshot. [`Edit::formats`] reads a worksheet's complete ordered core
//! collection, [`Edit::set_formats`] replaces it (the worksheet bytes outside
//! the `conditionalFormatting` owners are copied exactly) and
//! [`Edit::add_differential`] appends a differential format to the styles
//! part for a new rule to reference. [`Edit::commit`] returns a [`Commit`]
//! with the new package, the reversible [`Patch`] of the parts it changed and
//! the [`Diagnostics`]; [`Package::apply_conditional_formats`] replays a
//! patch onto an equal package. An exact no-op keeps every byte, and a
//! changed edit of a signed workbook is refused with [`Error::Signed`].
//!
//! Office 2010 (`x14`) rules and markup-compatibility content are refused by
//! the core writer rather than dropped.
//!
//! ```
//! use litchi_xlsx::{Package, Rgb};
//! use litchi_xlsx::conditional_formatting::edit::DifferentialStyle;
//! use litchi_xlsx::conditional_formatting::{DifferentialRef, Formatting, Kind, Range, Rule};
//!
//! # fn main() -> Result<(), litchi_xlsx::Error> {
//! let package = Package::new()?;
//! let mut edit = package.edit_conditional_formats()?;
//! let style = DifferentialStyle {
//!     fill: Some(Rgb::from_hex("C6EFCE").unwrap()),
//!     ..DifferentialStyle::default()
//! };
//! let dxf = edit.add_differential(&style)?;
//! let mut rule = Rule::new(Kind::DuplicateValues, 1)?;
//! rule.differential_format = Some(DifferentialRef::StylesIndex(dxf));
//! edit.set_formats(0, vec![Formatting::new(vec![Range::new("A1:A10")?], vec![rule])?])?;
//! let after = edit.commit()?.into_package();
//! let sheet = after.workbook().sheet(0)?.unwrap();
//! assert_eq!(sheet.conditional_formattings()?.len(), 1);
//! assert!(package.workbook().sheet(0)?.unwrap().conditional_formattings()?.is_empty());
//! # Ok(())
//! # }
//! ```

use std::fmt::Write as _;

use litchi_opc::PackURI;
use quick_xml::events::Event;
use quick_xml::reader::Reader;

use super::{Differential, Formatting};
use crate::Package;
use crate::color::Rgb;
use crate::error::{Result, invalid};
use crate::part_edit::{Staged, family};

family!("conditional-formatting");

/// The most differential formats a styles part may carry after an edit.
pub const MAX_DIFFERENTIALS: usize = 65_536;
/// The most `conditionalFormatting` collections one worksheet may carry.
pub const MAX_FORMATS: usize = 65_536;

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Differential formats of the styles part ([`MAX_DIFFERENTIALS`]).
    Differentials,
    /// Collections of one worksheet ([`MAX_FORMATS`]).
    Formats,
}

impl LimitKind {
    const fn maximum(self) -> usize {
        match self {
            Self::Differentials => MAX_DIFFERENTIALS,
            Self::Formats => MAX_FORMATS,
        }
    }

    fn check(self, observed: usize) -> Result<()> {
        if observed > self.maximum() {
            return Err(Error::limit(self, observed, self.maximum()));
        }
        Ok(())
    }
}

/// The underline of a differential font.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Underline {
    /// `<u/>`.
    Single,
    /// `<u val="double"/>`.
    Double,
}

/// The authorable part of a differential format: the formatting a
/// conditional rule applies. [`DifferentialStyle::read`] projects an
/// existing differential format the same way.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct DifferentialStyle {
    /// The solid fill (`<fill><patternFill><bgColor/>`).
    pub fill: Option<Rgb>,
    /// The font color.
    pub font_color: Option<Rgb>,
    /// Bold.
    pub bold: Option<bool>,
    /// Italic.
    pub italic: Option<bool>,
    /// Strike-through.
    pub strike: Option<bool>,
    /// Underline.
    pub underline: Option<Underline>,
    /// The font size in points.
    pub size: Option<f64>,
    /// The font name.
    pub name: Option<String>,
}

impl DifferentialStyle {
    /// Whether the style sets nothing.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }

    /// Project an existing differential format's font and fill. Colors other
    /// than RGB and formatting this type does not model are not reported.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed component markup.
    pub fn read(value: &Differential) -> Result<Self> {
        let mut style = Self::default();
        if let Some(font) = &value.font {
            read_components(font.raw_xml(), |name, attributes| {
                let flag = || attributes.value("val").is_none_or(|value| truthy(&value));
                match name {
                    b"b" => style.bold = Some(flag()),
                    b"i" => style.italic = Some(flag()),
                    b"strike" => style.strike = Some(flag()),
                    b"u" => {
                        style.underline = match attributes.value("val").as_deref() {
                            Some("none") => None,
                            Some("double" | "doubleAccounting") => Some(Underline::Double),
                            _ => Some(Underline::Single),
                        };
                    },
                    b"sz" => style.size = attributes.value("val").and_then(|v| v.parse().ok()),
                    b"color" => style.font_color = attributes.rgb(),
                    b"name" => style.name = attributes.value("val"),
                    _ => {},
                }
            })?;
        }
        if let Some(fill) = &value.fill {
            let mut foreground = None;
            let mut background = None;
            read_components(fill.raw_xml(), |name, attributes| match name {
                b"fgColor" => foreground = attributes.rgb(),
                b"bgColor" => background = attributes.rgb(),
                _ => {},
            })?;
            style.fill = background.or(foreground);
        }
        Ok(style)
    }

    fn to_xml(&self) -> Result<String> {
        let mut out = String::from("<dxf>");
        let font = self.font_color.is_some()
            || self.bold.is_some()
            || self.italic.is_some()
            || self.strike.is_some()
            || self.underline.is_some()
            || self.size.is_some()
            || self.name.is_some();
        if font {
            out.push_str("<font>");
            for (tag, value) in [
                ("b", self.bold),
                ("i", self.italic),
                ("strike", self.strike),
            ] {
                match value {
                    Some(true) => {
                        let _ = write!(out, "<{tag}/>");
                    },
                    Some(false) => {
                        let _ = write!(out, "<{tag} val=\"0\"/>");
                    },
                    None => {},
                }
            }
            match self.underline {
                Some(Underline::Single) => out.push_str("<u/>"),
                Some(Underline::Double) => out.push_str("<u val=\"double\"/>"),
                None => {},
            }
            if let Some(size) = self.size {
                if !size.is_finite() || !(1.0..=409.0).contains(&size) {
                    return Err(invalid("a differential font size must be 1 to 409 points"));
                }
                let _ = write!(out, "<sz val=\"{size}\"/>");
            }
            if let Some(color) = self.font_color {
                let _ = write!(out, "<color rgb=\"{}\"/>", argb(color));
            }
            if let Some(name) = &self.name {
                if name.is_empty() || name.chars().count() > 31 || name.contains('\0') {
                    return Err(invalid(
                        "a differential font name must be 1 to 31 characters",
                    ));
                }
                let _ = write!(
                    out,
                    "<name val=\"{}\"/>",
                    quick_xml::escape::escape(name.as_str())
                );
            }
            out.push_str("</font>");
        }
        if let Some(fill) = self.fill {
            let _ = write!(
                out,
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{0}\"/><bgColor rgb=\"{0}\"/></patternFill></fill>",
                argb(fill)
            );
        }
        out.push_str("</dxf>");
        Ok(out)
    }
}

fn argb(color: Rgb) -> String {
    color.to_string()
}

fn truthy(value: &str) -> bool {
    !matches!(value, "0" | "false")
}

struct Attributes<'a>(quick_xml::events::BytesStart<'a>);

impl Attributes<'_> {
    fn value(&self, key: &str) -> Option<String> {
        litchi_ooxml_common::xml::attributes::first_wins(&self.0)
            .flatten()
            .find_map(|attribute| {
                (attribute.key.local_name().as_ref() == key.as_bytes())
                    .then(|| {
                        let raw = String::from_utf8_lossy(&attribute.value);
                        quick_xml::escape::unescape(&raw)
                            .ok()
                            .map(std::borrow::Cow::into_owned)
                    })
                    .flatten()
            })
    }

    fn rgb(&self) -> Option<Rgb> {
        self.value("rgb")
            .and_then(|value| Rgb::from_hex(&value).ok())
    }
}

/// Visit every element of a component fragment by local name.
fn read_components(xml: &[u8], mut visit: impl FnMut(&[u8], &Attributes<'_>)) -> Result<()> {
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader
            .read_event()
            .map_err(|error| invalid(format!("differential format markup: {error}")))?
        {
            Event::Start(start) | Event::Empty(start) => {
                let name = start.local_name().as_ref().to_vec();
                visit(&name, &Attributes(start.into_owned()));
            },
            Event::Eof => return Ok(()),
            _ => {},
        }
    }
}

/// A staged conditional-formatting edit of one [`Package`]. Dropping it
/// discards it.
#[derive(Debug)]
pub struct Edit {
    staged: Staged,
}

impl Package {
    /// Start a conditional-formatting edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused.
    pub fn edit_conditional_formats(&self) -> Result<Edit> {
        self.edit_conditional_formats_with_options(crate::EditOptions::default())
    }

    /// Start a conditional-formatting edit with an explicit retained-patch
    /// budget.
    pub fn edit_conditional_formats_with_options(
        &self,
        options: crate::EditOptions,
    ) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::new_with_options(self, "edit_conditional_formats", options)?,
        })
    }

    /// Apply a conditional-formatting patch to a new package; this one is
    /// unchanged. An empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_conditional_formats(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_conditional_formats")?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    fn differentials(&self) -> Result<(Option<PackURI>, usize)> {
        let Some(styles) = self.staged.styles() else {
            return Ok((None, 0));
        };
        let count = super::parse_differential_formats(self.staged.blob(&styles)?)?.len();
        Ok((Some(styles), count))
    }

    /// The staged complete ordered core collection of the worksheet at
    /// `sheet`.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, malformed
    /// markup, or content the core owner cannot edit (Office 2010 rules).
    pub fn formats(&self, sheet: usize) -> Result<Vec<Formatting>> {
        let worksheet = self.staged.worksheet(sheet)?;
        let (_, count) = self.differentials()?;
        super::package::parse_editable_conditional_formattings(self.staged.blob(&worksheet)?, count)
    }

    /// Replace the complete ordered core collection of the worksheet at
    /// `sheet`. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, an invalid
    /// collection, content the core owner cannot edit, or a changed edit of a
    /// signed workbook.
    pub fn set_formats(&mut self, sheet: usize, values: Vec<Formatting>) -> Result<bool> {
        LimitKind::Formats.check(values.len())?;
        let worksheet = self.staged.worksheet(sheet)?;
        let (_, count) = self.differentials()?;
        let xml =
            super::package::replace_in_root_prefix(self.staged.blob(&worksheet)?, &values, count)?;
        self.staged.replace_blob(&worksheet, xml)
    }

    /// Append a differential format to the styles part and return its index
    /// for [`super::DifferentialRef::StylesIndex`].
    ///
    /// # Errors
    ///
    /// Returns an error for a workbook without a styles part, an empty or
    /// invalid style, a full table, or a signed workbook.
    pub fn add_differential(&mut self, style: &DifferentialStyle) -> Result<u32> {
        if style.is_empty() {
            return Err(invalid("a differential format must set some formatting"));
        }
        let (styles, count) = self.differentials()?;
        let styles = styles.ok_or_else(|| invalid("the workbook has no styles part"))?;
        LimitKind::Differentials.check(count + 1)?;
        let fragment = style.to_xml()?;
        let xml = append_differential(self.staged.blob(&styles)?, &fragment, count)?;
        if super::parse_differential_formats(&xml)?.len() != count + 1 {
            return Err(invalid("the added differential format did not read back"));
        }
        self.staged.replace_blob(&styles, xml)?;
        u32::try_from(count).map_err(|_| invalid("differential format index overflow"))
    }

    /// Validate the staged package and publish it.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}

/// The stylesheet children that follow `dxfs` (ECMA-376 `CT_Stylesheet`).
const AFTER_DXFS: &[&[u8]] = &[b"tableStyles", b"colors", b"extLst"];

/// Append `fragment` to the `dxfs` of a styles part, creating the element
/// at its schema position when the part has none.
fn append_differential(xml: &[u8], fragment: &str, count: usize) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut prefix = Vec::new();
    // (byte offset where the new content goes, whether `dxfs` must be created,
    // the byte span of an empty `<dxfs/>` to replace)
    let mut target: Option<(usize, usize)> = None;
    let mut dxfs_empty: Option<(usize, usize)> = None;
    let mut dxfs_open: Option<usize> = None;
    let mut root_close = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("styles part offset overflow"))?;
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("styles part markup: {error}")))?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("styles part offset overflow"))?;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if let Some(position) = element.name().as_ref().iter().position(|b| *b == b':')
                    {
                        prefix = element.name().as_ref()[..=position].to_vec();
                    }
                }
                if depth == 1 {
                    let name = element.local_name();
                    if name.as_ref() == b"dxfs" {
                        dxfs_open = Some(start);
                    } else if target.is_none() && AFTER_DXFS.contains(&name.as_ref()) {
                        target = Some((start, start));
                    }
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if depth == 1 {
                    let name = element.local_name();
                    if name.as_ref() == b"dxfs" {
                        dxfs_empty = Some((start, end));
                    } else if target.is_none() && AFTER_DXFS.contains(&name.as_ref()) {
                        target = Some((start, start));
                    }
                }
            },
            Event::End(element) => {
                depth = depth.saturating_sub(1);
                if depth == 1 && element.local_name().as_ref() == b"dxfs" && dxfs_open.is_some() {
                    // Insert before `</dxfs>` and fix the count below.
                    let mut out = Vec::with_capacity(xml.len() + fragment.len() + 16);
                    out.extend_from_slice(&xml[..start]);
                    out.extend_from_slice(&prefixed(fragment, &prefix));
                    out.extend_from_slice(&xml[start..]);
                    return set_count(out, dxfs_open.unwrap_or(0), count + 1);
                }
                if depth == 0 {
                    root_close = Some(start);
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let element = |inner: &[u8]| {
        let name = [prefix.as_slice(), b"dxfs"].concat();
        let mut out = Vec::new();
        out.push(b'<');
        out.extend_from_slice(&name);
        out.extend_from_slice(format!(" count=\"{}\">", count + 1).as_bytes());
        out.extend_from_slice(inner);
        out.extend_from_slice(b"</");
        out.extend_from_slice(&name);
        out.push(b'>');
        out
    };
    let replacement = element(&prefixed(fragment, &prefix));
    let mut out = Vec::with_capacity(xml.len() + replacement.len());
    if let Some((start, end)) = dxfs_empty {
        out.extend_from_slice(&xml[..start]);
        out.extend_from_slice(&replacement);
        out.extend_from_slice(&xml[end..]);
        return Ok(out);
    }
    let at = match (target, root_close) {
        (Some((at, _)), _) | (None, Some(at)) => at,
        (None, None) => return Err(invalid("the styles part has no root element")),
    };
    out.extend_from_slice(&xml[..at]);
    out.extend_from_slice(&replacement);
    out.extend_from_slice(&xml[at..]);
    Ok(out)
}

/// The fragment with every element name given the root's prefix.
fn prefixed(fragment: &str, prefix: &[u8]) -> Vec<u8> {
    if prefix.is_empty() {
        return fragment.as_bytes().to_vec();
    }
    let prefix = String::from_utf8_lossy(prefix);
    fragment
        .replace("</", "\u{1}")
        .replace('<', &format!("<{prefix}"))
        .replace('\u{1}', &format!("</{prefix}"))
        .into_bytes()
}

/// Rewrite the `count` attribute of the `dxfs` start tag at `open`.
fn set_count(mut xml: Vec<u8>, open: usize, count: usize) -> Result<Vec<u8>> {
    let close = xml[open..]
        .iter()
        .position(|byte| *byte == b'>')
        .map(|offset| open + offset)
        .ok_or_else(|| invalid("unterminated dxfs start tag"))?;
    let tag = String::from_utf8_lossy(&xml[open..close]).into_owned();
    let rewritten = match tag.find(" count=\"") {
        Some(at) => {
            let value_start = at + " count=\"".len();
            let value_end = tag[value_start..]
                .find('"')
                .map(|offset| value_start + offset)
                .ok_or_else(|| invalid("unterminated dxfs count"))?;
            format!("{}{count}{}", &tag[..value_start], &tag[value_end..])
        },
        None => format!("{tag} count=\"{count}\""),
    };
    xml.splice(open..close, rewritten.into_bytes());
    Ok(xml)
}

#[cfg(test)]
mod tests {
    use super::{DifferentialStyle, Underline, append_differential};
    use crate::Package;
    use crate::color::Rgb;
    use crate::conditional_formatting::{
        DifferentialRef, Formatting, Kind, Operator, Range, Rule, parse_differential_formats,
    };

    #[test]
    fn differentials_append_to_existing_empty_and_absent_tables() {
        let style = DifferentialStyle {
            fill: Some(Rgb::from_hex("FFC7CE").unwrap()),
            bold: Some(true),
            ..DifferentialStyle::default()
        }
        .to_xml()
        .unwrap();
        for xml in [
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="1"><xf/></cellXfs></styleSheet>"#,
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dxfs count="0"/><tableStyles count="0"/></styleSheet>"#,
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dxfs count="1"><dxf><font><i/></font></dxf></dxfs></styleSheet>"#,
            r#"<x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:colors/></x:styleSheet>"#,
        ] {
            let before = parse_differential_formats(xml.as_bytes()).unwrap().len();
            let out = append_differential(xml.as_bytes(), &style, before).unwrap();
            let after = parse_differential_formats(&out).unwrap();
            assert_eq!(after.len(), before + 1, "{}", String::from_utf8_lossy(&out));
            let read = DifferentialStyle::read(&after[before]).unwrap();
            assert_eq!(read.bold, Some(true));
            assert_eq!(read.fill, Some(Rgb::from_hex("FFC7CE").unwrap()));
            assert!(String::from_utf8_lossy(&out).contains(&format!("count=\"{}\"", before + 1)));
        }
    }

    #[test]
    fn empty_rules_keep_their_priority_and_differential_format() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_conditional_formats().unwrap();
        let dxf = edit
            .add_differential(&DifferentialStyle {
                fill: Some(Rgb::from_hex("FFC7CE").unwrap()),
                ..DifferentialStyle::default()
            })
            .unwrap();
        let mut rule = Rule::new(Kind::DuplicateValues, 1).unwrap();
        rule.differential_format = Some(DifferentialRef::StylesIndex(dxf));
        let formats =
            vec![Formatting::new(vec![Range::new("A1:A10").unwrap()], vec![rule]).unwrap()];
        edit.set_formats(0, formats.clone()).unwrap();
        let after = edit.commit().unwrap().into_package();
        let sheet = after.workbook().sheet(0).unwrap().unwrap();
        assert_eq!(sheet.conditional_formattings().unwrap(), formats);
    }

    #[test]
    fn add_set_replay_and_no_op() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_conditional_formats().unwrap();
        let dxf = edit
            .add_differential(&DifferentialStyle {
                font_color: Some(Rgb::from_hex("9C0006").unwrap()),
                underline: Some(Underline::Double),
                size: Some(12.0),
                name: Some("Arial".into()),
                ..DifferentialStyle::default()
            })
            .unwrap();
        let mut rule = Rule::new(Kind::CellIs, 1).unwrap();
        rule.operator = Some(Operator::GreaterThan);
        rule.push_formula("80").unwrap();
        rule.differential_format = Some(DifferentialRef::StylesIndex(dxf));
        let formats =
            vec![Formatting::new(vec![Range::new("A2:A11").unwrap()], vec![rule]).unwrap()];
        assert!(edit.set_formats(0, formats.clone()).unwrap());
        let commit = edit.commit().unwrap();
        assert!(commit.diagnostics().changed());
        let (after, patch) = commit.into_parts();
        assert_eq!(
            after
                .edit_conditional_formats()
                .unwrap()
                .formats(0)
                .unwrap(),
            formats
        );
        let replayed = package
            .apply_conditional_formats(&patch)
            .unwrap()
            .into_package();
        assert!(replayed.same_content(&after).unwrap());
        let reverted = after
            .apply_conditional_formats(&patch.inverse())
            .unwrap()
            .into_package();
        assert!(reverted.same_content(&package).unwrap());
        assert!(after.apply_conditional_formats(&patch).is_err());

        let noop = package
            .edit_conditional_formats()
            .unwrap()
            .commit()
            .unwrap();
        assert!(!noop.diagnostics().changed());
        assert!(noop.patch().is_noop());
    }
}
