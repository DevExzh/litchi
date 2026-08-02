//! Typed, inert model of RTF 1.9.1 native math zones.
//!
//! The RTF 1.9.1 specification ("Mathematics") mirrors the OMML (`m:`)
//! vocabulary as RTF destinations: `\mmath` and `\mmathPara` zones contain
//! structure destinations (fractions, radicals, scripts, matrices, and so
//! on), argument destinations (`\me`, `\mnum`, `\mden`, ...), property
//! destinations (`\mfPr`, `\mradPr`, `\mctrlPr`, ...), and `\mr` math runs.
//!
//! The model is purely syntactic and inert: property values, characters, and
//! run text are exposed exactly as stored. Nothing here evaluates, lays out,
//! typesets, or renders an equation. The document-level math defaults
//! (`\mmathPr`) are modeled separately in [`crate::DocumentMathProperties`].

use crate::{RtfError, RtfResult};
use std::borrow::Cow;
use std::collections::HashSet;

pub(crate) const MAX_MATH_ZONES: usize = 65_536;
pub(crate) const MAX_MATH_DEPTH: usize = 64;
pub(crate) const MAX_MATH_OBJECTS_PER_CONTAINER: usize = 1_024;
pub(crate) const MAX_MATH_PROPERTIES_PER_DESTINATION: usize = 256;
pub(crate) const MAX_MATH_PROPERTY_VALUE_BYTES: usize = 1_024;
pub(crate) const MAX_MATH_RUN_TEXT_BYTES: usize = 65_536;
pub(crate) const MAX_MATH_TOTAL_TEXT_BYTES: usize = 16 * 1_048_576;

fn malformed(message: impl Into<String>) -> RtfError {
    RtfError::MalformedDocument(message.into())
}

/// The kind of an RTF math zone destination.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathZoneKind {
    /// Inline math zone (`\mmath`).
    Inline,
    /// Display math zone (`\mmathPara`).
    Display,
}

/// The kind of a math structure destination.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathStructureKind {
    /// Accent (`\macc`).
    Accent,
    /// Bar (`\mbar`).
    Bar,
    /// Border box (`\mborderBox`).
    BorderBox,
    /// Box (`\mbox`).
    Box,
    /// Delimiter (`\md`).
    Delimiter,
    /// Equation array (`\meqArr`).
    EquationArray,
    /// Fraction (`\mf`).
    Fraction,
    /// Function application (`\mfunc`).
    Function,
    /// Group character (`\mgroupChr`).
    GroupChar,
    /// Lower limit (`\mlimlow`).
    LimitLower,
    /// Upper limit (`\mlimupp`).
    LimitUpper,
    /// Matrix (`\mm`).
    Matrix,
    /// N-ary operator (`\mnary`).
    Nary,
    /// Phantom (`\mphant`).
    Phantom,
    /// Radical (`\mrad`).
    Radical,
    /// Pre-sub-superscript (`\msPre`).
    ScriptPre,
    /// Subscript (`\msSub`).
    ScriptSub,
    /// Sub-superscript (`\msSubSup`).
    ScriptSubSup,
    /// Superscript (`\msSup`).
    ScriptSup,
}

/// The role of a math argument destination.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathElementRole {
    /// Generic element (`\me`).
    Element,
    /// Fraction numerator (`\mnum`).
    Numerator,
    /// Fraction denominator (`\mden`).
    Denominator,
    /// Radical degree (`\mdeg`).
    Degree,
    /// Subscript argument (`\msub`).
    Subscript,
    /// Superscript argument (`\msup`).
    Superscript,
    /// Limit argument (`\mlim`).
    Limit,
    /// Function name (`\mfName`).
    FunctionName,
}

/// A named math property from a `\*Pr` destination.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MathPropertyName {
    /// Fraction type (`\mtype`).
    Type,
    /// Growth (`\mgrow`).
    Grow,
    /// Operator/accent character (`\mchr`).
    Char,
    /// Delimiter begin character (`\mbegChr`).
    BeginChar,
    /// Delimiter end character (`\mendChr`).
    EndChar,
    /// Delimiter separator character (`\msepChr`).
    SeparatorChar,
    /// Position (`\mpos`).
    Position,
    /// Vertical justification (`\mvertJc`).
    VerticalJustify,
    /// Base justification (`\mbaseJc`).
    BaseJustify,
    /// Paragraph justification (`\mjc`).
    Justify,
    /// Alignment (`\maln`).
    Align,
    /// Script alignment (`\malnScr`).
    AlignScript,
    /// Hidden degree (`\mdegHide`).
    DegreeHide,
    /// Differential (`\mdiff`).
    Differential,
    /// Differential style (`\mdiffSty`).
    DifferentialStyle,
    /// Hidden bottom border (`\mhideBot`).
    HideBottom,
    /// Hidden left border (`\mhideLeft`).
    HideLeft,
    /// Hidden right border (`\mhideRight`).
    HideRight,
    /// Hidden top border (`\mhideTop`).
    HideTop,
    /// Limit location (`\mlimLoc`).
    LimitLocation,
    /// Hidden placeholder (`\mplcHide`).
    PlaceholderHide,
    /// Hidden subscript (`\msubHide`).
    SubscriptHide,
    /// Hidden superscript (`\msupHide`).
    SuperscriptHide,
    /// Bottom-left to top-right strike (`\mstrikeBLTR`).
    StrikeBottomLeftToTopRight,
    /// Horizontal strike (`\mstrikeH`).
    StrikeHorizontal,
    /// Top-left to bottom-right strike (`\mstrikeTLBR`).
    StrikeTopLeftToBottomRight,
    /// Vertical strike (`\mstrikeV`).
    StrikeVertical,
    /// Style (`\msty`).
    Style,
    /// Script (`\mscr`).
    Script,
    /// Transparency (`\mtransp`).
    Transparent,
    /// Phantom show (`\mshow`).
    Show,
    /// Phantom shape (`\mshp`).
    Shape,
    /// Phantom zero ascent (`\mzeroAsc`).
    ZeroAscent,
    /// Phantom zero descent (`\mzeroDesc`).
    ZeroDescent,
    /// Phantom zero width (`\mzeroWid`).
    ZeroWidth,
    /// Operator emulator (`\mopEmu`).
    OperatorEmulator,
    /// No break (`\mnoBreak`).
    NoBreak,
    /// Normal text (`\mnor`).
    NormalText,
    /// Literal (`\mlit`).
    Literal,
    /// Matrix column gap (`\mcGp`).
    MatrixColumnGap,
    /// Matrix column gap rule (`\mcGpRule`).
    MatrixColumnGapRule,
    /// Matrix column spacing (`\mcSp`).
    MatrixColumnSpacing,
    /// Matrix cell count (`\mcount`).
    MatrixCellCount,
    /// Matrix cell justification (`\mmcJc`).
    MatrixCellJustify,
    /// Row spacing (`\mrSp`).
    RowSpacing,
    /// Row spacing rule (`\mrSpRule`).
    RowSpacingRule,
    /// Break (`\mbrk`).
    Break,
    /// Argument size (`\margSz`).
    ArgumentSize,
}

/// One inert math property: a name with its verbatim decoded value text.
///
/// On the wire a property is a group holding a control word followed by its
/// value (for example `{\mtype bar}` or `{\msubHide 1}`); the value is kept
/// exactly as stored and is never interpreted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathProperty<'a> {
    /// Property name.
    pub name: MathPropertyName,
    /// Verbatim decoded value text (may be empty).
    pub value: Cow<'a, str>,
}

impl<'a> MathProperty<'a> {
    /// Create a validated math property.
    pub fn new(name: MathPropertyName, value: Cow<'a, str>) -> RtfResult<Self> {
        let property = Self { name, value };
        property.validate()?;
        Ok(property)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.value.len() > MAX_MATH_PROPERTY_VALUE_BYTES {
            return Err(malformed(
                "RTF math property value exceeds the safety limit",
            ));
        }
        if self.value.contains(['\0', '\r', '\n']) {
            return Err(malformed(
                "RTF math property value contains a forbidden control character",
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> MathProperty<'static> {
        MathProperty {
            name: self.name,
            value: Cow::Owned(self.value.into_owned()),
        }
    }
}

/// The kind of a math property destination.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathPropertiesKind {
    /// Structure properties (`\maccPr`, `\mfPr`, and the like).
    Structure(MathStructureKind),
    /// Math run properties (`\mrPr`).
    Run,
    /// Math paragraph properties (`\mmathParaPr`).
    Paragraph,
    /// Control properties (`\mctrlPr`).
    Control,
    /// Argument properties (`\margPr`).
    Argument,
    /// Matrix column properties (`\mmcPr`).
    MatrixColumn,
}

/// One inert matrix column description (`\mmc`) from the `\mmcs` destination
/// of a matrix property destination.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathMatrixColumn<'a> {
    /// Column properties (`\mmcPr`): cell count (`\mcount`) and cell
    /// justification (`\mmcJc`).
    pub properties: Option<MathProperties<'a>>,
}

impl<'a> MathMatrixColumn<'a> {
    pub(crate) fn validate(&self) -> RtfResult<()> {
        if let Some(properties) = &self.properties {
            if properties.kind != MathPropertiesKind::MatrixColumn {
                return Err(malformed(
                    "RTF math matrix column properties must use the mmcPr destination",
                ));
            }
            properties.validate()?;
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> MathMatrixColumn<'static> {
        MathMatrixColumn {
            properties: self.properties.map(MathProperties::into_owned),
        }
    }
}

/// An inert math property destination.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathProperties<'a> {
    /// Which property destination this is.
    pub kind: MathPropertiesKind,
    /// Ordered properties declared in the destination.
    pub properties: Vec<MathProperty<'a>>,
    /// Matrix column descriptions (`\mmcs`); only meaningful for
    /// [`MathPropertiesKind::Structure`]`(`[`MathStructureKind::Matrix`]`)`.
    pub matrix_columns: Vec<MathMatrixColumn<'a>>,
    /// Nested control properties (`\mctrlPr`).
    pub control: Option<Box<MathProperties<'a>>>,
}

impl<'a> MathProperties<'a> {
    /// Create a validated math property destination.
    pub fn new(kind: MathPropertiesKind, properties: Vec<MathProperty<'a>>) -> RtfResult<Self> {
        let destination = Self {
            kind,
            properties,
            matrix_columns: Vec::new(),
            control: None,
        };
        destination.validate()?;
        Ok(destination)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.properties.len() > MAX_MATH_PROPERTIES_PER_DESTINATION {
            return Err(malformed(
                "RTF math property count exceeds the safety limit",
            ));
        }
        let mut names = HashSet::new();
        crate::error::try_reserve_set(&mut names, self.properties.len(), "math property names")?;
        for property in &self.properties {
            property.validate()?;
            if !names.insert(property.name) {
                return Err(malformed(
                    "RTF math property names must be unique within a destination",
                ));
            }
            let permitted = match self.kind {
                MathPropertiesKind::Argument => property.name == MathPropertyName::ArgumentSize,
                MathPropertiesKind::MatrixColumn => matches!(
                    property.name,
                    MathPropertyName::MatrixCellCount | MathPropertyName::MatrixCellJustify
                ),
                _ => property.name != MathPropertyName::ArgumentSize,
            };
            if !permitted {
                return Err(malformed(
                    "RTF math property is not permitted in this destination",
                ));
            }
        }
        if !self.matrix_columns.is_empty() {
            if self.kind != MathPropertiesKind::Structure(MathStructureKind::Matrix) {
                return Err(malformed(
                    "RTF math matrix columns may occur only inside matrix properties",
                ));
            }
            if self.matrix_columns.len() > MAX_MATH_OBJECTS_PER_CONTAINER {
                return Err(malformed(
                    "RTF math matrix column count exceeds the safety limit",
                ));
            }
            for column in &self.matrix_columns {
                column.validate()?;
            }
        }
        if let Some(control) = &self.control {
            if control.kind != MathPropertiesKind::Control {
                return Err(malformed(
                    "RTF math control properties must use the mctrlPr destination",
                ));
            }
            control.validate()?;
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> MathProperties<'static> {
        MathProperties {
            kind: self.kind,
            properties: self
                .properties
                .into_iter()
                .map(MathProperty::into_owned)
                .collect(),
            matrix_columns: self
                .matrix_columns
                .into_iter()
                .map(MathMatrixColumn::into_owned)
                .collect(),
            control: self.control.map(|control| Box::new(control.into_owned())),
        }
    }
}

/// A math argument destination (`\me`, `\mnum`, `\mden`, and the like).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathElement<'a> {
    /// Argument role.
    pub role: MathElementRole,
    /// Argument properties (`\margPr` with `\margSz`).
    pub argument_properties: Option<MathProperties<'a>>,
    /// Ordered objects contained in the argument.
    pub content: Vec<MathObject<'a>>,
}

/// One matrix row (`\mmr`) with its `\me` cell arguments.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathMatrixRow<'a> {
    /// Ordered cell arguments; each has [`MathElementRole::Element`].
    pub cells: Vec<MathElement<'a>>,
}

/// A direct child of a math structure destination.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MathStructureChild<'a> {
    /// An argument destination.
    Element(MathElement<'a>),
    /// A matrix row destination (only inside [`MathStructureKind::Matrix`]).
    MatrixRow(MathMatrixRow<'a>),
}

/// A math run (`\mr`) with its inert text.
///
/// Character formatting controls inside the run are passive and are not
/// retained; only the text, the `\mnor` flag, and `\mrPr` properties are.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathRun<'a> {
    /// Run properties (`\mrPr`).
    pub properties: Option<MathProperties<'a>>,
    /// Whether the run is marked as normal (non-math) text (`\mnor`).
    pub normal_text: bool,
    /// Decoded run text.
    pub text: Cow<'a, str>,
}

/// One object inside a math zone or argument.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MathObject<'a> {
    /// A structure destination (fraction, radical, script, ...).
    Structure(MathStructure<'a>),
    /// A math run.
    Run(MathRun<'a>),
}

/// A math structure destination with its properties and children.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathStructure<'a> {
    /// Structure kind.
    pub kind: MathStructureKind,
    /// Structure properties (`\mfPr` and the like).
    pub properties: Option<MathProperties<'a>>,
    /// Ordered argument children.
    pub children: Vec<MathStructureChild<'a>>,
}

/// An inert math zone anchored at a body-text position.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MathZone<'a> {
    /// Zone kind (inline or display).
    pub kind: MathZoneKind,
    /// Display-zone paragraph properties (`\mmathParaPr`).
    pub paragraph_properties: Option<MathProperties<'a>>,
    /// Ordered top-level objects.
    pub content: Vec<MathObject<'a>>,
    /// UTF-8 byte offset in the document body text where the zone is anchored.
    pub position: usize,
}

impl MathStructureKind {
    /// The property-destination kind matching this structure.
    pub const fn properties_kind(self) -> MathPropertiesKind {
        MathPropertiesKind::Structure(self)
    }

    /// Expected argument sequence as `(role, optional)` pairs, or `None` for
    /// kinds with variadic children (delimiter, equation array, matrix).
    fn expected_children(self) -> Option<&'static [(MathElementRole, bool)]> {
        const ELEMENT: &[(MathElementRole, bool)] = &[(MathElementRole::Element, false)];
        match self {
            Self::Accent
            | Self::Bar
            | Self::BorderBox
            | Self::Box
            | Self::GroupChar
            | Self::Phantom => Some(ELEMENT),
            Self::Radical => Some(&[
                (MathElementRole::Degree, true),
                (MathElementRole::Element, false),
            ]),
            Self::Fraction => Some(&[
                (MathElementRole::Numerator, false),
                (MathElementRole::Denominator, false),
            ]),
            Self::Function => Some(&[
                (MathElementRole::FunctionName, false),
                (MathElementRole::Element, false),
            ]),
            Self::LimitLower | Self::LimitUpper => Some(&[
                (MathElementRole::Element, false),
                (MathElementRole::Limit, false),
            ]),
            Self::Nary => Some(&[
                (MathElementRole::Subscript, true),
                (MathElementRole::Superscript, true),
                (MathElementRole::Element, false),
            ]),
            Self::ScriptPre | Self::ScriptSubSup => Some(&[
                (MathElementRole::Subscript, false),
                (MathElementRole::Superscript, false),
                (MathElementRole::Element, false),
            ]),
            Self::ScriptSub => Some(&[
                (MathElementRole::Subscript, false),
                (MathElementRole::Element, false),
            ]),
            Self::ScriptSup => Some(&[
                (MathElementRole::Superscript, false),
                (MathElementRole::Element, false),
            ]),
            Self::Delimiter | Self::EquationArray | Self::Matrix => None,
        }
    }
}

fn validate_objects(objects: &[MathObject<'_>], depth: usize) -> RtfResult<()> {
    if depth > MAX_MATH_DEPTH {
        return Err(malformed("RTF math nesting depth exceeds the safety limit"));
    }
    if objects.len() > MAX_MATH_OBJECTS_PER_CONTAINER {
        return Err(malformed("RTF math object count exceeds the safety limit"));
    }
    for object in objects {
        match object {
            MathObject::Structure(structure) => structure.validate_at(depth)?,
            MathObject::Run(run) => {
                if run.text.len() > MAX_MATH_RUN_TEXT_BYTES {
                    return Err(malformed("RTF math run text exceeds the safety limit"));
                }
                if let Some(properties) = &run.properties {
                    if properties.kind != MathPropertiesKind::Run {
                        return Err(malformed(
                            "RTF math run properties must use the mrPr destination",
                        ));
                    }
                    properties.validate()?;
                }
            },
        }
    }
    Ok(())
}

pub(crate) fn validate_element(element: &MathElement<'_>, depth: usize) -> RtfResult<()> {
    if let Some(properties) = &element.argument_properties {
        if properties.kind != MathPropertiesKind::Argument {
            return Err(malformed(
                "RTF math argument properties must use the margPr destination",
            ));
        }
        properties.validate()?;
    }
    validate_objects(&element.content, depth)
}

impl<'a> MathStructure<'a> {
    /// Create a validated math structure.
    pub fn new(
        kind: MathStructureKind,
        properties: Option<MathProperties<'a>>,
        children: Vec<MathStructureChild<'a>>,
    ) -> RtfResult<Self> {
        let structure = Self {
            kind,
            properties,
            children,
        };
        structure.validate_at(1)?;
        Ok(structure)
    }

    pub(crate) fn validate_at(&self, depth: usize) -> RtfResult<()> {
        if let Some(properties) = &self.properties {
            if properties.kind != self.kind.properties_kind() {
                return Err(malformed(
                    "RTF math structure properties do not match the structure kind",
                ));
            }
            properties.validate()?;
        }
        match self.kind.expected_children() {
            Some(expected) => {
                let mut index = 0usize;
                for &(role, optional) in expected {
                    match self.children.get(index) {
                        Some(MathStructureChild::Element(element)) if element.role == role => {
                            index += 1;
                        },
                        _ if optional => {},
                        _ => {
                            return Err(malformed(
                                "RTF math structure children do not match the required argument sequence",
                            ));
                        },
                    }
                }
                if index != self.children.len() {
                    return Err(malformed(
                        "RTF math structure has unexpected argument children",
                    ));
                }
                for child in &self.children {
                    let MathStructureChild::Element(element) = child else {
                        return Err(malformed("RTF matrix rows may occur only inside a matrix"));
                    };
                    validate_element(element, depth + 1)?;
                }
            },
            None if self.kind == MathStructureKind::Matrix => {
                if self.children.is_empty() || self.children.len() > MAX_MATH_OBJECTS_PER_CONTAINER
                {
                    return Err(malformed(
                        "RTF math matrix must contain at least one row within the safety limit",
                    ));
                }
                for child in &self.children {
                    let MathStructureChild::MatrixRow(row) = child else {
                        return Err(malformed("RTF math matrix may contain only matrix rows"));
                    };
                    if row.cells.is_empty() || row.cells.len() > MAX_MATH_OBJECTS_PER_CONTAINER {
                        return Err(malformed(
                            "RTF math matrix row must contain at least one cell within the safety limit",
                        ));
                    }
                    for cell in &row.cells {
                        if cell.role != MathElementRole::Element {
                            return Err(malformed(
                                "RTF math matrix cells must use the me destination",
                            ));
                        }
                        validate_element(cell, depth + 1)?;
                    }
                }
            },
            None => {
                // Delimiter and equation array: one or more \me arguments.
                if self.children.is_empty() || self.children.len() > MAX_MATH_OBJECTS_PER_CONTAINER
                {
                    return Err(malformed(
                        "RTF math structure must contain at least one argument within the safety limit",
                    ));
                }
                for child in &self.children {
                    let MathStructureChild::Element(element) = child else {
                        return Err(malformed("RTF matrix rows may occur only inside a matrix"));
                    };
                    if element.role != MathElementRole::Element {
                        return Err(malformed(
                            "RTF math structure arguments must use the me destination",
                        ));
                    }
                    validate_element(element, depth + 1)?;
                }
            },
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> MathStructure<'static> {
        MathStructure {
            kind: self.kind,
            properties: self.properties.map(MathProperties::into_owned),
            children: self
                .children
                .into_iter()
                .map(MathStructureChild::into_owned)
                .collect(),
        }
    }
}

impl<'a> MathStructureChild<'a> {
    pub(crate) fn into_owned(self) -> MathStructureChild<'static> {
        match self {
            Self::Element(element) => MathStructureChild::Element(element.into_owned()),
            Self::MatrixRow(row) => MathStructureChild::MatrixRow(row.into_owned()),
        }
    }
}

impl<'a> MathElement<'a> {
    pub(crate) fn into_owned(self) -> MathElement<'static> {
        MathElement {
            role: self.role,
            argument_properties: self.argument_properties.map(MathProperties::into_owned),
            content: self
                .content
                .into_iter()
                .map(MathObject::into_owned)
                .collect(),
        }
    }
}

impl<'a> MathMatrixRow<'a> {
    pub(crate) fn into_owned(self) -> MathMatrixRow<'static> {
        MathMatrixRow {
            cells: self
                .cells
                .into_iter()
                .map(MathElement::into_owned)
                .collect(),
        }
    }
}

impl<'a> MathObject<'a> {
    pub(crate) fn into_owned(self) -> MathObject<'static> {
        match self {
            Self::Structure(structure) => MathObject::Structure(structure.into_owned()),
            Self::Run(run) => MathObject::Run(run.into_owned()),
        }
    }
}

impl<'a> MathRun<'a> {
    pub(crate) fn into_owned(self) -> MathRun<'static> {
        MathRun {
            properties: self.properties.map(MathProperties::into_owned),
            normal_text: self.normal_text,
            text: Cow::Owned(self.text.into_owned()),
        }
    }
}

impl<'a> MathZone<'a> {
    /// Create a validated math zone.
    pub fn new(
        kind: MathZoneKind,
        paragraph_properties: Option<MathProperties<'a>>,
        content: Vec<MathObject<'a>>,
        position: usize,
    ) -> RtfResult<Self> {
        let zone = Self {
            kind,
            paragraph_properties,
            content,
            position,
        };
        zone.validate()?;
        Ok(zone)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        match (&self.kind, &self.paragraph_properties) {
            (MathZoneKind::Inline, Some(_)) => {
                return Err(malformed(
                    "RTF inline math zones cannot carry paragraph properties",
                ));
            },
            (MathZoneKind::Display, Some(properties))
                if properties.kind != MathPropertiesKind::Paragraph =>
            {
                return Err(malformed(
                    "RTF display math zones must use the mmathParaPr destination",
                ));
            },
            _ => {},
        }
        if let Some(properties) = &self.paragraph_properties {
            properties.validate()?;
        }
        validate_objects(&self.content, 1)
    }

    pub(crate) fn into_owned(self) -> MathZone<'static> {
        MathZone {
            kind: self.kind,
            paragraph_properties: self.paragraph_properties.map(MathProperties::into_owned),
            content: self
                .content
                .into_iter()
                .map(MathObject::into_owned)
                .collect(),
            position: self.position,
        }
    }
}
