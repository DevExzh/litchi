//! Rich text formatting on shapes, paragraphs, runs and table cells, staged
//! through the native `set_text_properties` edits of both PPTX owners.
use super::{Commit, Editable, Pptx, Proof, allocation, bound, error, validate_text};
use crate::format::pptx::Nav;
use crate::format::pptx::node::Item;
use crate::format::pptx::spec::{
    CELL_WRITE_TABLE, PARAGRAPH_WRITE_TABLE, RUN_WRITE_TABLE, TextSetter,
};
use crate::registry::{Operation, Table};
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::{Base, Scalar};
use crate::{Code, Props, Result};
use litchi_core::Position;
use litchi_drawingml::geom::Preset;
use litchi_drawingml::text::{Alignment, Anchor, Strike, Underline};
use litchi_pptx::opened::{
    CellFill, CellMargin, Rgb, ShapeFill, TextAssignment, TextProperties, TextTarget,
};
use std::collections::HashSet;
use std::sync::Arc;

/// The text element a path or query addresses.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum TextKind {
    Paragraph,
    Run,
    Cell,
}

fn kind_of(name: &str) -> Option<TextKind> {
    if name.eq_ignore_ascii_case("paragraph") || name.eq_ignore_ascii_case("p") {
        Some(TextKind::Paragraph)
    } else if ["run", "r", "linebreak", "br"]
        .iter()
        .any(|alias| name.eq_ignore_ascii_case(alias))
    {
        Some(TextKind::Run)
    } else if name.eq_ignore_ascii_case("tc") || name.eq_ignore_ascii_case("cell") {
        Some(TextKind::Cell)
    } else {
        None
    }
}

/// Classify a mutation target by its last path step or its query subjects,
/// without resolving it.
pub(super) fn classify(target: &TargetArg<'_>) -> Option<TextKind> {
    if let Some(path) = target.path() {
        return path
            .path()
            .steps()
            .last()
            .and_then(|step| kind_of(step.name()));
    }
    let query = target.query()?.query()?;
    let mut kinds = query
        .alternatives()
        .iter()
        .map(|alternative| alternative.subject().element().name().and_then(kind_of));
    let first = kinds.next()??;
    kinds.all(|kind| kind == Some(first)).then_some(first)
}

/// Admitted rich text assignments, owned so they outlive the props borrow.
#[derive(Clone, Default)]
#[allow(
    clippy::option_option,
    reason = "omission, removal and replacement are distinct assignments"
)]
pub(in crate::format::pptx) struct TextValues {
    text: Option<String>,
    bold: Option<Option<bool>>,
    italic: Option<Option<bool>>,
    size: Option<Option<u32>>,
    color: Option<Option<Rgb>>,
    font: Option<Option<String>>,
    latin: Option<Option<String>>,
    align: Option<Option<Alignment>>,
    underline: Option<Option<Underline>>,
    strike: Option<Option<Strike>>,
    fill: Option<Option<CellFill>>,
    valign: Option<Option<Anchor>>,
    margins: [Option<Option<i64>>; 4],
    shape_fill: Option<Option<ShapeFill>>,
    geometry: Option<Preset>,
}

impl TextValues {
    pub(in crate::format::pptx) fn is_empty(&self) -> bool {
        self.text.is_none()
            && self.bold.is_none()
            && self.italic.is_none()
            && self.size.is_none()
            && self.color.is_none()
            && self.font.is_none()
            && self.latin.is_none()
            && self.align.is_none()
            && self.underline.is_none()
            && self.strike.is_none()
            && self.fill.is_none()
            && self.valign.is_none()
            && self.margins.iter().all(Option::is_none)
            && self.shape_fill.is_none()
            && self.geometry.is_none()
    }

    /// Whether this assignment changes only the target text.  The signed
    /// cancellation proof can track text plans by their native value; rich
    /// formatting remains eligible for the native patch's own no-op proof.
    pub(in crate::format::pptx) fn is_text_only(&self) -> bool {
        self.text.is_some()
            && self.bold.is_none()
            && self.italic.is_none()
            && self.size.is_none()
            && self.color.is_none()
            && self.font.is_none()
            && self.latin.is_none()
            && self.align.is_none()
            && self.underline.is_none()
            && self.strike.is_none()
            && self.fill.is_none()
            && self.valign.is_none()
            && self.margins.iter().all(Option::is_none)
            && self.shape_fill.is_none()
            && self.geometry.is_none()
    }

    /// Whether a coalesced formatting run can stage these assignments and
    /// tell their change from the reader's model (record 0333): run and
    /// paragraph character formatting, alignment and run text, but no
    /// fills, cell properties or geometry.
    pub(in crate::format::pptx) fn coalescible(&self) -> bool {
        !self.is_empty()
            && self.fill.is_none()
            && self.valign.is_none()
            && self.margins.iter().all(Option::is_none)
            && self.shape_fill.is_none()
            && self.geometry.is_none()
    }

    /// Whether a coalesced run can stage these table-cell assignments
    /// (record 0346): [`Self::coalescible`]'s and the cell's fill, anchor and
    /// insets, whose change the leaf reports (record 0348).
    pub(in crate::format::pptx) fn cell_coalescible(&self) -> bool {
        !self.is_empty() && self.shape_fill.is_none() && self.geometry.is_none()
    }

    /// Whether the assignments change a run, compared with the local
    /// properties the reader gives.
    pub(in crate::format::pptx) fn run_changed(
        &self,
        run: &litchi_drawingml::text::body::Run,
    ) -> bool {
        let properties = &run.properties;
        fn face(font: &Option<litchi_drawingml::text::body::Font>) -> Option<&str> {
            font.as_ref().map(|font| font.typeface.as_str())
        }
        let color = properties
            .fill
            .as_ref()
            .map(|fill| match fill.solid_color() {
                Some(litchi_drawingml::color::Value::Rgb(rgb)) => {
                    Some(Rgb([rgb.red(), rgb.green(), rgb.blue()]))
                },
                _ => None,
            });
        self.text.as_deref().is_some_and(|text| text != run.text)
            || self.bold.is_some_and(|value| value != properties.bold)
            || self.italic.is_some_and(|value| value != properties.italic)
            || self
                .size
                .is_some_and(|value| value != properties.size.map(|size| size.get()))
            || self.color.is_some_and(|value| match (value, color) {
                (None, None) => false,
                (Some(wanted), Some(Some(current))) => wanted != current,
                _ => true,
            })
            || self
                .underline
                .is_some_and(|value| value != properties.underline)
            || self.strike.is_some_and(|value| value != properties.strike)
            || self
                .latin
                .as_ref()
                .is_some_and(|value| value.as_deref() != face(&properties.latin))
            || self.font.as_ref().is_some_and(|value| {
                let value = value.as_deref();
                value != face(&properties.latin)
                    || value != face(&properties.east_asian)
                    || value != face(&properties.complex_script)
            })
    }

    /// Whether the assignments change a paragraph: its alignment, or any of
    /// its runs. Paragraph text is not coalesced.
    pub(in crate::format::pptx) fn paragraph_changed(
        &self,
        paragraph: &litchi_drawingml::text::body::Paragraph,
    ) -> bool {
        let runs = Self {
            text: None,
            align: None,
            ..self.clone()
        };
        self.align
            .is_some_and(|value| value != paragraph.properties.alignment)
            || (!runs.is_empty() && paragraph.runs.iter().any(|run| runs.run_changed(run)))
    }

    /// The assigned text, if any.
    pub(in crate::format::pptx) fn text(&self) -> Option<&str> {
        self.text.as_deref()
    }

    pub(super) fn charge(&self) -> usize {
        let text = |value: &Option<Option<String>>| {
            value
                .as_ref()
                .and_then(Option::as_ref)
                .map_or(0, String::len)
        };
        self.text
            .as_ref()
            .map_or(0, String::len)
            .saturating_add(text(&self.font))
            .saturating_add(text(&self.latin))
            .saturating_add(64)
    }

    pub(in crate::format::pptx) fn properties(&self) -> TextProperties<'_> {
        let mut properties = TextProperties::default();
        if let Some(text) = &self.text {
            properties = properties.with_text(text);
        }
        if let Some(value) = self.bold {
            properties = properties.with_bold(value);
        }
        if let Some(value) = self.italic {
            properties = properties.with_italic(value);
        }
        if let Some(value) = self.size {
            properties = properties.with_size(value);
        }
        if let Some(value) = self.color {
            properties = properties.with_color(value);
        }
        if let Some(value) = &self.font {
            properties = properties.with_font(value.as_deref());
        }
        if let Some(value) = &self.latin {
            properties = properties.with_latin(value.as_deref());
        }
        if let Some(value) = self.align {
            properties = properties.with_align(value);
        }
        if let Some(value) = self.underline {
            properties = properties.with_underline(value);
        }
        if let Some(value) = self.strike {
            properties = properties.with_strike(value);
        }
        if let Some(value) = self.fill {
            properties = properties.with_fill(value);
        }
        if let Some(value) = self.valign {
            properties = properties.with_cell_anchor(value);
        }
        for (margin, value) in CellMargin::ALL.into_iter().zip(self.margins) {
            if let Some(value) = value {
                properties = properties.with_cell_margin(margin, value);
            }
        }
        if let Some(value) = self.shape_fill {
            properties = properties.with_shape_fill(value);
        }
        if let Some(value) = self.geometry {
            properties = properties.with_geometry(value);
        }
        properties
    }

    /// Whether `setter` already has an assignment.
    pub(in crate::format::pptx) const fn has(&self, setter: TextSetter) -> bool {
        match setter {
            TextSetter::Text => self.text.is_some(),
            TextSetter::Bold => self.bold.is_some(),
            TextSetter::Italic => self.italic.is_some(),
            TextSetter::Size => self.size.is_some(),
            TextSetter::Color => self.color.is_some(),
            TextSetter::Font => self.font.is_some(),
            TextSetter::FontLatin => self.latin.is_some(),
            TextSetter::Align => self.align.is_some(),
            TextSetter::Fill => self.fill.is_some(),
            TextSetter::Underline => self.underline.is_some(),
            TextSetter::Strike => self.strike.is_some(),
            TextSetter::Valign => self.valign.is_some(),
            TextSetter::Margin(margin) => self.margins[margin as usize].is_some(),
            TextSetter::Padding => {
                self.margins[0].is_some()
                    || self.margins[1].is_some()
                    || self.margins[2].is_some()
                    || self.margins[3].is_some()
            },
            TextSetter::ShapeFill => self.shape_fill.is_some(),
            TextSetter::Geometry => self.geometry.is_some(),
            TextSetter::Click(_) => false,
        }
    }

    /// Record one coerced assignment.
    pub(in crate::format::pptx) fn stage(
        &mut self,
        setter: TextSetter,
        value: Option<Scalar<'_>>,
        limits: &SemanticLimits,
    ) -> Result<()> {
        let owned = |value: &str| -> Result<String> {
            bound(
                LimitKind::StagedBytes,
                value.len(),
                limits.max_staged_bytes(),
            )?;
            validate_text(value)?;
            let mut owned = String::new();
            owned.try_reserve_exact(value.len()).map_err(allocation)?;
            owned.push_str(value);
            Ok(owned)
        };
        match setter {
            // Click actions are split off before text admission (record
            // 0325) and written through the hyperlink verbs.
            TextSetter::Click(_) => {
                return Err(crate::Error::new(
                    Code::InternalError,
                    "a click action reached the text edit",
                ));
            },
            TextSetter::Text => {
                let text = match value {
                    Some(Scalar::Text(text)) => text,
                    None => "",
                    _ => unreachable!("text coercion"),
                };
                self.text = Some(owned(text)?);
            },
            TextSetter::Bold | TextSetter::Italic => {
                let flag = match value {
                    Some(Scalar::Bool(flag)) => Some(flag),
                    None => None,
                    _ => unreachable!("boolean coercion"),
                };
                if setter == TextSetter::Bold {
                    self.bold = Some(flag);
                } else {
                    self.italic = Some(flag);
                }
            },
            TextSetter::Size => {
                self.size = Some(match value {
                    Some(Scalar::Length(length)) => Some(font_size(length.points())?),
                    None => None,
                    _ => unreachable!("font size coercion"),
                });
            },
            TextSetter::Color => {
                self.color = Some(match value {
                    Some(Scalar::Color(color)) => Some(rgb(&color)?),
                    None => None,
                    _ => unreachable!("colour coercion"),
                });
            },
            TextSetter::Fill => {
                self.fill = Some(match value {
                    Some(Scalar::Color(color)) => Some(cell_fill(&color)?),
                    None => None,
                    _ => unreachable!("colour coercion"),
                });
            },
            TextSetter::Underline => {
                self.underline = Some(match value {
                    Some(Scalar::Token(token)) => Some(underline(token)?),
                    None => None,
                    _ => unreachable!("underline coercion"),
                });
            },
            TextSetter::Strike => {
                self.strike = Some(match value {
                    Some(Scalar::Token(token)) => Some(match token {
                        "single" => Strike::Single,
                        "double" => Strike::Double,
                        "none" => Strike::None,
                        _ => {
                            return Err(
                                error(Code::InvalidValue, "unknown strike").with_property("strike")
                            );
                        },
                    }),
                    None => None,
                    _ => unreachable!("strike coercion"),
                });
            },
            TextSetter::Valign => {
                self.valign = Some(match value {
                    Some(Scalar::Token(token)) => Some(match token {
                        "top" => Anchor::Top,
                        "center" => Anchor::Center,
                        "bottom" => Anchor::Bottom,
                        _ => {
                            return Err(error(Code::InvalidValue, "unknown vertical alignment")
                                .with_property("valign"));
                        },
                    }),
                    None => None,
                    _ => unreachable!("anchor coercion"),
                });
            },
            TextSetter::Margin(_) | TextSetter::Padding => {
                let inset = match value {
                    Some(Scalar::Length(length)) => {
                        let emu = length.emu();
                        if emu < 0 {
                            return Err(error(
                                Code::InvalidValue,
                                "a cell inset cannot be negative",
                            )
                            .with_property("padding"));
                        }
                        Some(emu)
                    },
                    None => None,
                    _ => unreachable!("inset coercion"),
                };
                // `padding` and one side's key name the same inset.
                if self.has(setter) {
                    return Err(error(
                        Code::InvalidArgument,
                        "a cell inset was assigned more than once",
                    )
                    .with_property("padding"));
                }
                match setter {
                    TextSetter::Margin(margin) => self.margins[margin as usize] = Some(inset),
                    _ => self.margins = [Some(inset); 4],
                }
            },
            TextSetter::Font | TextSetter::FontLatin => {
                let face = match value {
                    Some(Scalar::Text(face)) => Some(owned(face)?),
                    None => None,
                    _ => unreachable!("typeface coercion"),
                };
                if setter == TextSetter::Font {
                    self.font = Some(face);
                } else {
                    self.latin = Some(face);
                }
            },
            TextSetter::ShapeFill => {
                self.shape_fill = Some(match value {
                    Some(Scalar::Text(text)) => Some(shape_fill(text)?),
                    None => None,
                    _ => unreachable!("shape fill coercion"),
                });
            },
            TextSetter::Geometry => {
                let Some(Scalar::Text(token)) = value else {
                    unreachable!("geometry coercion admits no clear")
                };
                self.geometry = Some(geometry(token)?);
            },
            TextSetter::Align => {
                self.align = Some(match value {
                    Some(Scalar::Token(token)) => Some(alignment(token)?),
                    None => None,
                    _ => unreachable!("alignment coercion"),
                });
            },
        }
        Ok(())
    }
}

/// A shape fill: `none`, or a colour in the drawing vocabulary that the
/// native writer can store (opaque RGB).
fn shape_fill(text: &str) -> Result<ShapeFill> {
    if text.trim().eq_ignore_ascii_case("none") {
        return Ok(ShapeFill::None);
    }
    let color =
        crate::value::Color::parse(text, crate::value::ColorContext::Drawing).map_err(|_| {
            error(Code::InvalidValue, "a shape fill is `none` or a colour")
                .with_property("fill")
                .with_suggestion("use none, a hex colour such as 1F4E79, or a colour name")
        })?;
    Ok(ShapeFill::Solid(
        rgb(&color).map_err(|error| error.with_property("fill"))?,
    ))
}

/// A preset geometry token (`ST_ShapeType`), matched as OfficeCLI matches
/// it: exactly, then ignoring ASCII case.
fn geometry(token: &str) -> Result<Preset> {
    let token = token.trim();
    token
        .parse::<Preset>()
        .ok()
        .or_else(|| {
            Preset::ALL
                .into_iter()
                .find(|preset| preset.token().eq_ignore_ascii_case(token))
        })
        .ok_or_else(|| {
            let error =
                error(Code::InvalidValue, "unknown preset geometry").with_property("geometry");
            match crate::value::suggest::nearest_keyed(
                token,
                Preset::ALL
                    .into_iter()
                    .map(|preset| (preset.token(), preset.token())),
            ) {
                Some(name) => error.with_suggestion(format!("did you mean '{name}'?")),
                None => error,
            }
        })
}

fn font_size(points: f64) -> Result<u32> {
    let hundredths = (points * 100.0).round();
    if !(f64::from(litchi_pptx::opened::MIN_FONT_SIZE)
        ..=f64::from(litchi_pptx::opened::MAX_FONT_SIZE))
        .contains(&hundredths)
    {
        return Err(error(
            Code::InvalidValue,
            "font size must be between 1pt and 4000pt",
        )
        .with_property("size"));
    }
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "the value was range-checked against the u32 font-size domain"
    )]
    Ok(hundredths as u32)
}

fn rgb(color: &crate::value::Color) -> Result<Rgb> {
    match color.base() {
        Base::Rgb(value) if value.alpha() == 255 && color.transforms().is_empty() => {
            Ok(Rgb([value.red(), value.green(), value.blue()]))
        },
        _ => Err(error(
            Code::UnsupportedCapability,
            "only opaque RGB colours without transforms are writable",
        )
        .with_suggestion("use a hex colour such as 1F4E79")),
    }
}

/// A table cell fill colour: opaque RGB, or a theme colour without
/// transforms (record 0333).
fn cell_fill(color: &crate::value::Color) -> Result<CellFill> {
    if let Base::Scheme(scheme) = color.base()
        && color.transforms().is_empty()
    {
        return Ok(CellFill::Scheme(drawing_scheme(scheme)));
    }
    rgb(color).map(CellFill::Rgb).map_err(|error| {
        error.with_property("fill").with_suggestion(
            "use a hex colour such as 1F4E79, a theme colour such as accent1, or none",
        )
    })
}

/// The `DrawingML` token of a theme colour.
fn drawing_scheme(scheme: crate::value::Scheme) -> litchi_drawingml::color::Scheme {
    use crate::value::Scheme;
    use litchi_drawingml::color::Scheme as Drawing;
    match scheme {
        Scheme::Background1 => Drawing::Background,
        Scheme::Text1 => Drawing::Text,
        Scheme::Background2 => Drawing::Background2,
        Scheme::Text2 => Drawing::Text2,
        Scheme::Accent1 => Drawing::Accent1,
        Scheme::Accent2 => Drawing::Accent2,
        Scheme::Accent3 => Drawing::Accent3,
        Scheme::Accent4 => Drawing::Accent4,
        Scheme::Accent5 => Drawing::Accent5,
        Scheme::Accent6 => Drawing::Accent6,
        Scheme::Hyperlink => Drawing::Hyperlink,
        Scheme::FollowedHyperlink => Drawing::FollowedHyperlink,
        Scheme::Placeholder => Drawing::Placeholder,
        Scheme::Dark1 => Drawing::Dark1,
        Scheme::Light1 => Drawing::Light1,
        Scheme::Dark2 => Drawing::Dark2,
        Scheme::Light2 => Drawing::Light2,
    }
}

/// An underline token, as OfficeCLI names it.
fn underline(token: &str) -> Result<Underline> {
    Ok(match token {
        "none" => Underline::None,
        "single" => Underline::Single,
        "double" => Underline::Double,
        "words" => Underline::Words,
        "heavy" => Underline::Heavy,
        "dotted" => Underline::Dotted,
        "dottedHeavy" => Underline::DottedHeavy,
        "dash" => Underline::Dash,
        "dashHeavy" => Underline::DashHeavy,
        "dashLong" => Underline::DashLong,
        "dashLongHeavy" => Underline::DashLongHeavy,
        "dotDash" => Underline::DotDash,
        "dotDashHeavy" => Underline::DotDashHeavy,
        "dotDotDash" => Underline::DotDotDash,
        "dotDotDashHeavy" => Underline::DotDotDashHeavy,
        "wavy" => Underline::Wavy,
        "wavyHeavy" => Underline::WavyHeavy,
        "wavyDbl" => Underline::WavyDouble,
        _ => return Err(error(Code::InvalidValue, "unknown underline").with_property("underline")),
    })
}

fn alignment(token: &str) -> Result<Alignment> {
    Ok(match token {
        "left" => Alignment::Left,
        "center" => Alignment::Center,
        "right" => Alignment::Right,
        "justify" => Alignment::Justified,
        "justLow" => Alignment::JustifiedLow,
        "dist" => Alignment::Distributed,
        "thaiDist" => Alignment::ThaiDistributed,
        _ => return Err(error(Code::InvalidValue, "unknown alignment").with_property("align")),
    })
}

pub(super) fn admit<const N: usize, const A: usize>(
    table: &Table<TextSetter, N, A>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<TextValues> {
    let values = admit_for(table, props, Operation::Set, limits)?;
    if values.is_empty() {
        return Err(error(
            Code::InvalidArgument,
            "no text property was assigned",
        ));
    }
    Ok(values)
}

/// Admit `props` for `operation` through `table`: `set` or, for a new
/// paragraph or run, `add` (record 0333). An empty payload is admitted.
pub(in crate::format::pptx) fn admit_for<const N: usize, const A: usize>(
    table: &Table<TextSetter, N, A>,
    props: &Props<'_>,
    operation: Operation,
    limits: &SemanticLimits,
) -> Result<TextValues> {
    let mut values = TextValues::default();
    let mut seen = 0_u64;
    for (key, value) in props.iter() {
        let (index, entry) = table.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(litchi_core::FileFormat::Pptx)
                .with_property(key.as_str().to_owned())
        })?;
        let bit = 1_u64 << index;
        if seen & bit != 0 {
            return Err(error(
                Code::InvalidArgument,
                "a text property was assigned more than once",
            )
            .with_property(entry.spec.name));
        }
        seen |= bit;
        // OfficeCLI's `fill=none` is an explicit no-fill, not a colour.
        if matches!(entry.access, crate::registry::Access::Set(TextSetter::Fill))
            && matches!(value, crate::PropValue::Text(text) if text.trim().eq_ignore_ascii_case("none"))
        {
            values.fill = Some(Some(CellFill::None));
            continue;
        }
        entry.prepare(operation, value, |setter, value| {
            values.stage(setter, value, limits)
        })?;
    }
    Ok(values)
}

impl Pptx {
    /// Set rich text properties on paragraphs, runs or table cells.
    pub(super) fn set_text(
        &self,
        target: TargetArg<'_>,
        kind: TextKind,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let values = match kind {
            TextKind::Paragraph => admit(&PARAGRAPH_WRITE_TABLE, props, limits)?,
            TextKind::Run => admit(&RUN_WRITE_TABLE, props, limits)?,
            TextKind::Cell => admit(&CELL_WRITE_TABLE, props, limits)?,
        };
        if kind == TextKind::Run
            && values
                .text
                .as_deref()
                .is_some_and(|text| text.contains(['\n', '\r']))
        {
            return Err(
                error(Code::InvalidValue, "run text cannot contain a line break")
                    .with_property("text"),
            );
        }
        let per_target = values.charge();
        // `(notes page, slide, target)`: a notes body's text is edited in
        // its notes part (record 0333).
        let mut targets: Vec<(bool, usize, TextTarget)> = Vec::new();
        let mut seen = HashSet::new();
        let mut append = |node: crate::format::pptx::Node<'_>| -> Result<()> {
            let (container, target) = match (kind, node.item()) {
                (TextKind::Paragraph, Item::Paragraph(text)) => (
                    text.shape.container,
                    match text.cell.get() {
                        Some((row, column)) => TextTarget::CellParagraph {
                            shape: text.shape.handle.position(),
                            row: Position::new(row as usize),
                            column: Position::new(column as usize),
                            paragraph: Position::new(text.paragraph as usize),
                        },
                        None => TextTarget::Paragraph {
                            shape: text.shape.handle.position(),
                            paragraph: Position::new(text.paragraph as usize),
                        },
                    },
                ),
                (TextKind::Run, Item::Run(text) | Item::LineBreak(text)) => (
                    text.shape.container,
                    match text.cell.get() {
                        Some((row, column)) => TextTarget::CellRun {
                            shape: text.shape.handle.position(),
                            row: Position::new(row as usize),
                            column: Position::new(column as usize),
                            paragraph: Position::new(text.paragraph as usize),
                            run: Position::new(text.run as usize),
                        },
                        None => TextTarget::Run {
                            shape: text.shape.handle.position(),
                            paragraph: Position::new(text.paragraph as usize),
                            run: Position::new(text.run as usize),
                        },
                    },
                ),
                (TextKind::Cell, Item::Cell(cell)) => (
                    cell.table.container,
                    TextTarget::Cell {
                        shape: cell.table.handle.position(),
                        row: Position::new(cell.row as usize),
                        column: Position::new(cell.column as usize),
                    },
                ),
                _ => {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "the selected element does not take these text properties",
                    ));
                },
            };
            let (notes, slide) = match (container.slide(), container.notes()) {
                (Some(slide), _) => (false, slide),
                (None, Some(slide)) => (true, slide),
                (None, None) => {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "master and layout text writes are not integrated",
                    ));
                },
            };
            let key = (notes, slide, target);
            if seen.contains(&key) {
                return Ok(());
            }
            let count = targets.len().saturating_add(1);
            bound(
                LimitKind::ExpandedTargets,
                count,
                limits.max_expanded_targets(),
            )?;
            bound(
                LimitKind::StagedBytes,
                count.saturating_mul(per_target),
                limits.max_staged_bytes(),
            )?;
            targets.try_reserve(1).map_err(allocation)?;
            seen.try_reserve(1).map_err(allocation)?;
            seen.insert(key);
            targets.push((notes, slide as usize, target));
            Ok(())
        };
        let nav = Nav::new(self);
        if let Some(path) = target.path() {
            append(crate::path::resolve::get(&nav, path)?)?;
        } else if let Some(query) = target.query() {
            let walk = crate::query::run(nav, query)?;
            if !walk.plan().diagnostics().is_empty() {
                return Err(error(
                    Code::InvalidSelector,
                    "mutation selector contains an unsupported predicate",
                ));
            }
            for node in walk {
                append(node?)?;
            }
        } else {
            return Err(error(Code::InvalidArgument, "mutation target is empty"));
        }
        if targets.is_empty() {
            return Err(error(
                Code::ZeroMatches,
                "mutation selector matched no element",
            ));
        }
        targets.sort_by_key(|(notes, slide, _)| (*notes, *slide));
        let properties = values.properties();
        let mut groups: Vec<(usize, Vec<TextAssignment<'_>>)> = Vec::new();
        let mut notes_groups: Vec<(usize, Vec<TextAssignment<'_>>)> = Vec::new();
        for (notes, slide, target) in &targets {
            let groups = if *notes {
                &mut notes_groups
            } else {
                &mut groups
            };
            if groups.last().is_none_or(|(current, _)| current != slide) {
                groups.try_reserve(1).map_err(allocation)?;
                groups.push((*slide, Vec::new()));
            }
            let Some((_, edits)) = groups.last_mut() else {
                unreachable!("a group was just pushed")
            };
            edits.try_reserve(1).map_err(allocation)?;
            edits.push(TextAssignment::new(*target, properties));
        }
        self.commit_story_text_edits(&groups, &notes_groups, limits)
    }

    /// Publish slide-grouped text edits through one native edit.
    pub(super) fn commit_text_edits(
        &self,
        groups: &[(usize, Vec<TextAssignment<'_>>)],
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.commit_story_text_edits(groups, &[], limits)
    }

    /// Publish slide-grouped text edits and notes-page text edits through
    /// one native edit. Notes pages are edited on the eager package: a
    /// source-backed presentation with notes edits is materialized first,
    /// as a notes text write is (record 0285).
    pub(super) fn commit_story_text_edits(
        &self,
        groups: &[(usize, Vec<TextAssignment<'_>>)],
        notes: &[(usize, Vec<TextAssignment<'_>>)],
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let capacity = litchi_pptx::opened::MAX_SHAPE_PROPERTIES_REPLACEMENTS;
        let source = self.editable_owner_for_write()?;
        let native_limit = super::native_snapshot_budget(limits);
        let mut operations = 0;
        let materialized = match (&source, notes.is_empty()) {
            (Editable::Source(package), false) => {
                let bytes = package.to_bytes()?;
                Some(super::super::eager_package_from_source(
                    bytes,
                    self.0.ooxml_limits,
                )?)
            },
            _ => None,
        };
        let owner = match &materialized {
            Some(package) => Editable::Eager(Arc::clone(package)),
            None => source.clone(),
        };
        let (target, proof, changed) = match &owner {
            Editable::Source(package) => {
                let mut edit = package
                    .edit_with_options(super::source_edit_options(native_limit))
                    .map_err(super::native_error)?;
                for (slide, edits) in groups {
                    let chunks: Vec<&[TextAssignment<'_>]> = edits.chunks(capacity).collect();
                    operations += edit.set_text_property_chunks(Position::new(*slide), &chunks)?;
                }
                let (target, patch) = edit.commit()?.into_parts();
                let changed = !patch.is_noop();
                (Editable::Source(target), Proof::Source(patch), changed)
            },
            Editable::Eager(package) => {
                let mut edit = package
                    .edit_with_options(super::opened_edit_options(native_limit))
                    .map_err(super::native_error)?;
                for (slide, edits) in groups {
                    let chunks: Vec<&[TextAssignment<'_>]> = edits.chunks(capacity).collect();
                    operations += edit.set_text_property_chunks(*slide, &chunks)?;
                }
                for (slide, edits) in notes {
                    // One pass over the notes part (record 0348).
                    let chunks: Vec<&[TextAssignment<'_>]> = edits.chunks(capacity).collect();
                    operations += edit.set_notes_text_property_chunks(*slide, &chunks)?;
                }
                let (target, patch) = edit.commit()?.into_parts();
                let changed = !patch.is_noop();
                (
                    Editable::Eager(Arc::new(target)),
                    Proof::Eager(patch),
                    changed,
                )
            },
        };
        // Native text chunks may report zero physical changes for an equal
        // assignment, but the unified setter still authored one command.
        operations = operations.max(1);
        let after = if changed {
            match &target {
                Editable::Source(package) => self.adopt_source(package.clone())?,
                Editable::Eager(package) => self.adopt_eager(Arc::clone(package)),
            }
        } else {
            self.clone()
        };
        Ok(Commit::new(after, source, target, proof, operations))
    }
}
