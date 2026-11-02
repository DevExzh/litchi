//! Pages body text-box character and paragraph property parity (wave W4-P,
//! record 0413): the host `set_text_box_*`, `reset_text_box_*` and
//! `text_box_*` property families against
//! `litchi_pages::Package::edit_body_text_box_text`.

use litchi_iwa::pages;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_protos::tswp;
use litchi_pages::drawable::text_box_text::{Property, PropertyKind};
use litchi_pages::{DrawableSelector, Package};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::decode_as;
use crate::leaf_bytes;
use crate::text_style_cases::{sample, second_sample};

const BASIC: &str = "iwork/pages/basic.pages";

/// The properties the host's text-box families cover: every engine
/// property.
pub const KINDS: [PropertyKind; 24] = [
    PropertyKind::TextStyle,
    PropertyKind::TextFont,
    PropertyKind::TextDecorations,
    PropertyKind::TextColor,
    PropertyKind::TextCapitalization,
    PropertyKind::TextScript,
    PropertyKind::TextBaselineShift,
    PropertyKind::TextCharacterSpacing,
    PropertyKind::TextLigatures,
    PropertyKind::TextOutline,
    PropertyKind::TextShadow,
    PropertyKind::TextBackground,
    PropertyKind::ParagraphBackground,
    PropertyKind::Borders,
    PropertyKind::Flow,
    PropertyKind::WritingDirection,
    PropertyKind::FollowingStyle,
    PropertyKind::Alignment,
    PropertyKind::LineSpacing,
    PropertyKind::Spacing,
    PropertyKind::Indents,
    PropertyKind::DecimalTabCharacter,
    PropertyKind::DefaultTabInterval,
    PropertyKind::TabStops,
];

/// `basic.pages` with two host text boxes appended to the body
/// (`Draft::add_text_box`), the first with two paragraphs.
pub(crate) fn text_box_fixture() -> TestResult<Fixture> {
    let mut editor = pages::Draft::from_bytes(&fixture(BASIC)?.bytes)?;
    for (text, x) in [("First box\nSecond line", 72.0), ("Second box", 320.0)] {
        let anchor = editor.body_text()?.encode_utf16().count();
        editor.add_text_box(
            anchor,
            text,
            DrawablePoint { x, y: 96.0 },
            DrawableSize {
                width: 200.0,
                height: 80.0,
            },
        )?;
    }
    Ok(Fixture {
        path: "basic.pages + two add_text_box",
        bytes: editor.to_bytes()?,
    })
}

/// The host identifier of body text box `index`, in body order (the host
/// adds them in that order, so identifier order is body order here).
pub(crate) fn host_text_box(editor: &pages::Draft, index: usize) -> TestResult<u64> {
    let mut boxes = editor
        .drawable_text_storages()?
        .into_iter()
        .map(|item| item.drawable_object_id)
        .filter(|id| editor.text_box_geometry(*id).is_ok())
        .collect::<Vec<_>>();
    boxes.sort_unstable();
    Ok(*boxes
        .get(index)
        .ok_or_else(|| format!("no host text box {index}"))?)
}

/// One change, expressed once for both routes.
#[derive(Debug, Clone)]
pub enum Change {
    Set(Property),
    Reset(PropertyKind),
}

/// The host route: the per-property `set_text_box_*` / `reset_text_box_*`
/// method.
#[allow(clippy::too_many_lines, reason = "one arm per host method")]
pub fn host_change(source: &[u8], index: usize, change: &Change) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    let d = host_text_box(&e, index)?;
    macro_rules! set {
        ($($variant:ident => $method:ident),* $(,)?) => {
            match change {
                $(Change::Set(Property::$variant(v)) => e.$method(d, v.clone())?,)*
                _ => return Err(format!("unmapped change {change:?}").into()),
            }
        };
    }
    macro_rules! reset {
        ($($variant:ident => $method:ident),* $(,)?) => {
            match change {
                $(Change::Reset(PropertyKind::$variant) => drop(e.$method(d)?),)*
                _ => return Err(format!("unmapped change {change:?}").into()),
            }
        };
    }
    match change {
        Change::Set(_) => set!(
            TextStyle => set_text_box_text_style,
            TextFont => set_text_box_text_font,
            TextDecorations => set_text_box_text_decorations,
            TextColor => set_text_box_text_color,
            TextCapitalization => set_text_box_text_capitalization,
            TextScript => set_text_box_text_script,
            TextBaselineShift => set_text_box_text_baseline_shift,
            TextCharacterSpacing => set_text_box_text_character_spacing,
            TextLigatures => set_text_box_text_ligatures,
            TextOutline => set_text_box_text_outline,
            TextShadow => set_text_box_text_shadow,
            TextBackground => set_text_box_text_background,
            ParagraphBackground => set_text_box_paragraph_background,
            Borders => set_text_box_paragraph_borders,
            Flow => set_text_box_paragraph_flow,
            WritingDirection => set_text_box_paragraph_writing_direction,
            FollowingStyle => set_text_box_paragraph_following_style,
            Alignment => set_text_box_paragraph_alignment,
            LineSpacing => set_text_box_paragraph_line_spacing,
            Spacing => set_text_box_paragraph_spacing,
            Indents => set_text_box_paragraph_indents,
            DecimalTabCharacter => set_text_box_paragraph_decimal_tab_character,
            DefaultTabInterval => set_text_box_paragraph_default_tab_interval,
            TabStops => set_text_box_paragraph_tab_stops,
        ),
        Change::Reset(_) => reset!(
            TextStyle => reset_text_box_text_style,
            TextFont => reset_text_box_text_font,
            TextDecorations => reset_text_box_text_decorations,
            TextColor => reset_text_box_text_color,
            TextCapitalization => reset_text_box_text_capitalization,
            TextScript => reset_text_box_text_script,
            TextBaselineShift => reset_text_box_text_baseline_shift,
            TextCharacterSpacing => reset_text_box_text_character_spacing,
            TextLigatures => reset_text_box_text_ligatures,
            TextOutline => reset_text_box_text_outline,
            TextShadow => reset_text_box_text_shadow,
            TextBackground => reset_text_box_text_background,
            ParagraphBackground => reset_text_box_paragraph_background,
            Borders => reset_text_box_paragraph_borders,
            Flow => reset_text_box_paragraph_flow,
            WritingDirection => reset_text_box_paragraph_writing_direction,
            FollowingStyle => reset_text_box_paragraph_following_style,
            Alignment => reset_text_box_paragraph_alignment,
            LineSpacing => reset_text_box_paragraph_line_spacing,
            Spacing => reset_text_box_paragraph_spacing,
            Indents => reset_text_box_paragraph_indents,
            DecimalTabCharacter => reset_text_box_paragraph_decimal_tab_character,
            DefaultTabInterval => reset_text_box_paragraph_default_tab_interval,
            TabStops => reset_text_box_paragraph_tab_stops,
        ),
    }
    Ok(e.to_bytes()?)
}

/// The host read: the per-property `text_box_*` method.
pub fn host_read(bytes: &[u8], index: usize, kind: PropertyKind) -> TestResult<Property> {
    let e = pages::Draft::from_bytes(bytes)?;
    let d = host_text_box(&e, index)?;
    macro_rules! read {
        ($($variant:ident => $method:ident),* $(,)?) => {
            match kind {
                $(PropertyKind::$variant => Property::$variant(e.$method(d)?),)*
                other => return Err(format!("unmapped kind {other:?}").into()),
            }
        };
    }
    Ok(read!(
        TextStyle => text_box_text_style,
        TextFont => text_box_text_font,
        TextDecorations => text_box_text_decorations,
        TextColor => text_box_text_color,
        TextCapitalization => text_box_text_capitalization,
        TextScript => text_box_text_script,
        TextBaselineShift => text_box_text_baseline_shift,
        TextCharacterSpacing => text_box_text_character_spacing,
        TextLigatures => text_box_text_ligatures,
        TextOutline => text_box_text_outline,
        TextShadow => text_box_text_shadow,
        TextBackground => text_box_text_background,
        ParagraphBackground => text_box_paragraph_background,
        Borders => text_box_paragraph_borders,
        Flow => text_box_paragraph_flow,
        WritingDirection => text_box_paragraph_writing_direction,
        FollowingStyle => text_box_paragraph_following_style,
        Alignment => text_box_paragraph_alignment,
        LineSpacing => text_box_paragraph_line_spacing,
        Spacing => text_box_paragraph_spacing,
        Indents => text_box_paragraph_indents,
        DecimalTabCharacter => text_box_paragraph_decimal_tab_character,
        DefaultTabInterval => text_box_paragraph_default_tab_interval,
        TabStops => text_box_paragraph_tab_stops,
    ))
}

/// The leaf route through the text-box text family.
pub fn leaf_change(source: &[u8], index: usize, change: &Change) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_text_box_text(DrawableSelector::text_box(index))?;
    match change {
        Change::Set(property) => {
            edit.set(property.clone())?;
        },
        Change::Reset(kind) => {
            edit.reset(*kind)?;
        },
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_text_box_text(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

pub fn leaf_read(bytes: &[u8], index: usize, kind: PropertyKind) -> TestResult<Property> {
    Ok(Package::from_bytes(bytes)?.body_text_box_text(DrawableSelector::text_box(index), kind)?)
}

fn change_kind(change: &Change) -> PropertyKind {
    match change {
        Change::Set(property) => property.kind(),
        Change::Reset(kind) => *kind,
    }
}

/// A sample value of `kind` for text box `index` of `bytes`: the following
/// style names the last named paragraph style of the document.
fn sample_on(bytes: &[u8], index: usize, kind: PropertyKind) -> TestResult<Property> {
    if kind != PropertyKind::FollowingStyle {
        return sample(kind);
    }
    let editor = pages::Draft::from_bytes(bytes)?;
    let drawable = host_text_box(&editor, index)?;
    let last = editor
        .text_box_named_paragraph_styles(drawable)?
        .pop()
        .ok_or("no named paragraph style")?;
    Ok(Property::FollowingStyle(
        litchi_iwa_text::paragraph::style::ParagraphFollowingStyle::Named(last.id()),
    ))
}

/// One case on text box `index` of `input`.
pub fn run(
    name: &'static str,
    tier: Tier,
    input: &Fixture,
    index: usize,
    change: &Change,
) -> TestResult {
    let kind = change_kind(change);
    let mut case = Case::new(name, Format::Pages, tier, Evidence::Queued, input)
        .host(|source| host_change(source, index, change))
        .leaf(|source| leaf_change(source, index, change))
        .read("host property", move |bytes| host_read(bytes, index, kind))
        .read("leaf property", move |bytes| leaf_read(bytes, index, kind))
        .known_fields(|known| {
            known.with(
                2_022,
                "TSWP.ParagraphStyleArchive",
                decode_as::<tswp::ParagraphStyleArchive>,
            )
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

/// The host output of setting a sample of `kind` on text box `index`: a
/// fixture whose text box has an exclusive managed variation.
pub fn derived(input: &Fixture, index: usize, kind: PropertyKind) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_change(
            &input.bytes,
            index,
            &Change::Set(sample_on(&input.bytes, index, kind)?),
        )?,
    })
}

#[test]
fn pages_text_box_property_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = text_box_fixture()?;
    for index in [0, 1] {
        for kind in KINDS {
            run(
                "pages_text_box_property_set_parity",
                Tier::B,
                &input,
                index,
                &Change::Set(sample_on(&input.bytes, index, kind)?),
            )
            .map_err(|e| format!("text box {index} {kind:?}: {e}"))?;
        }
    }
    Ok(())
}

#[test]
fn pages_text_box_property_rewrite_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = text_box_fixture()?;
    for kind in KINDS {
        let derived = derived(&input, 0, kind)?;
        for change in [
            Change::Set(second_sample(kind)?),
            Change::Set(sample(if kind == PropertyKind::Alignment {
                PropertyKind::Indents
            } else {
                PropertyKind::Alignment
            })?),
        ] {
            run(
                "pages_text_box_property_rewrite_parity",
                Tier::B,
                &derived,
                0,
                &change,
            )
            .map_err(|e| format!("{kind:?} {change:?}: {e}"))?;
        }
    }
    Ok(())
}

#[test]
fn pages_text_box_property_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = text_box_fixture()?;
    for kind in KINDS {
        let derived = derived(&input, 0, kind)?;
        run(
            "pages_text_box_property_reset_parity",
            Tier::B,
            &derived,
            0,
            &Change::Reset(kind),
        )
        .map_err(|e| format!("{kind:?}: {e}"))?;
    }
    Ok(())
}

#[test]
fn pages_text_box_property_read_parity() -> TestResult {
    let input = text_box_fixture()?;
    for kind in KINDS {
        for bytes in [input.bytes.clone(), derived(&input, 1, kind)?.bytes] {
            for index in [0, 1] {
                assert_eq!(
                    host_read(&bytes, index, kind)?,
                    leaf_read(&bytes, index, kind)?,
                    "text box {index} {kind:?}"
                );
            }
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

// ---- frame, properties, text layout, columns and plain text ----------------

use litchi_iwa_common::text::layout::{AutoSize, Inset, Insets, VerticalAlignment};
use litchi_iwa_text::columns::{Columns, Count, Gap};
use litchi_pages::drawable::style::{Layout, Property as StyleProperty, PropertyKind as StyleKind};
use litchi_pages::drawable::{Geometry, GeometryFlags, Point, Size};

use crate::pages_drawable_cases::{host_geometry, host_properties, properties};

/// One frame or style edit of text box 0, expressed once for both routes.
#[derive(Clone)]
enum FrameOp {
    Geometry(Geometry),
    Properties(litchi_pages::drawable::Properties),
    Style(StyleProperty),
    Reset(StyleKind),
    Text(TextOp),
}

#[derive(Clone, Copy)]
enum TextOp {
    Replace(usize, usize, &'static str),
    Set(&'static str),
    Clear,
}

fn layout() -> TestResult<Layout> {
    Ok(Layout::new(
        VerticalAlignment::Middle,
        Insets::new(
            Inset::from_points(4.0)?,
            Inset::from_points(6.0)?,
            Inset::from_points(8.0)?,
            Inset::from_points(10.0)?,
        ),
        AutoSize::ShrinkToFit,
    ))
}

fn columns() -> TestResult<Columns> {
    Ok(Columns::equal(
        Count::new(2)?,
        Some(Gap::from_points(12.0)?),
    ))
}

fn host_frame(source: &[u8], op: &FrameOp) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    let d = host_text_box(&e, 0)?;
    match op {
        FrameOp::Geometry(geometry) => e.set_text_box_geometry(d, host_geometry(*geometry))?,
        FrameOp::Properties(value) => e.set_text_box_properties(d, host_properties(value))?,
        FrameOp::Style(StyleProperty::TextLayout(layout)) => {
            e.set_text_box_text_layout(d, *layout)?;
        },
        FrameOp::Style(StyleProperty::TextColumns(columns)) => {
            e.set_text_box_columns(d, columns)?;
        },
        FrameOp::Reset(StyleKind::TextLayout) => drop(e.reset_text_box_text_layout(d)?),
        FrameOp::Reset(StyleKind::TextColumns) => drop(e.reset_text_box_columns(d)?),
        FrameOp::Text(TextOp::Replace(start, end, text)) => {
            e.replace_drawable_text(d, *start..*end, text)?;
        },
        FrameOp::Text(TextOp::Set(text)) => e.set_drawable_text(d, text)?,
        FrameOp::Text(TextOp::Clear) => e.clear_drawable_text(d)?,
        FrameOp::Style(_) | FrameOp::Reset(_) => return Err("no host text-box route".into()),
    }
    Ok(e.to_bytes()?)
}

fn leaf_frame(source: &[u8], op: &FrameOp) -> TestResult<LeafRun> {
    leaf_frame_with(source, op, true)
}

fn leaf_frame_with(source: &[u8], op: &FrameOp, inverse: bool) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let selector = DrawableSelector::text_box(0);
    let (output, restored) = match op {
        FrameOp::Geometry(_) | FrameOp::Properties(_) => {
            let mut edit = package.edit_body_drawable_properties(selector)?;
            match op {
                FrameOp::Geometry(geometry) => drop(edit.set_geometry(*geometry)?),
                FrameOp::Properties(value) => drop(edit.set_properties(value.clone())),
                _ => unreachable!(),
            }
            let commit = edit.commit()?;
            let restored = if inverse {
                let restored = commit
                    .package()
                    .apply_body_drawable_properties(&commit.patch().inverse())?;
                Some(leaf_bytes!(restored.package()))
            } else {
                None
            };
            (leaf_bytes!(commit.package()), restored)
        },
        FrameOp::Style(_) | FrameOp::Reset(_) => {
            let mut edit = package.edit_body_drawable_style(selector)?;
            match op {
                FrameOp::Style(property) => drop(edit.set(property.clone())?),
                FrameOp::Reset(kind) => drop(edit.reset(*kind)),
                _ => unreachable!(),
            }
            let commit = edit.commit()?;
            let restored = if inverse {
                let restored = commit
                    .package()
                    .apply_body_drawable_style(&commit.patch().inverse())?;
                Some(leaf_bytes!(restored.package()))
            } else {
                None
            };
            (leaf_bytes!(commit.package()), restored)
        },
        FrameOp::Text(text) => {
            let mut edit = package.edit_body_drawable_text(selector)?;
            match *text {
                TextOp::Replace(start, end, text) => drop(edit.replace(
                    litchi_pages::TextSpan::new(
                        litchi_pages::TextPosition::from_utf16_index(start)?,
                        litchi_pages::TextPosition::from_utf16_index(end)?,
                    )?,
                    text,
                )?),
                TextOp::Set(text) => drop(edit.set(text)?),
                TextOp::Clear => drop(edit.clear()?),
            }
            let commit = edit.commit()?;
            let restored = if inverse {
                let restored = commit
                    .package()
                    .apply_body_drawable_text(&commit.patch().inverse())?;
                Some(leaf_bytes!(restored.package()))
            } else {
                None
            };
            (leaf_bytes!(commit.package()), restored)
        },
    };
    Ok(match restored {
        Some(restored) => LeafRun::with_inverse(output, restored),
        None => LeafRun::output(output),
    })
}

/// Every frame, property, layout, column and text read of both text boxes.
fn host_frame_read(bytes: &[u8]) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let mut out = String::new();
    for index in [0, 1] {
        let d = host_text_box(&e, index)?;
        let text = e
            .drawable_text_storages()?
            .into_iter()
            .find(|item| item.drawable_object_id == d)
            .ok_or("no storage")?
            .storage
            .storage
            .into_text();
        out.push_str(&format!(
            "{:?}\n{:?}\n{:?}\n{:?}\n{text:?}\n",
            e.text_box_geometry(d)?,
            e.text_box_properties(d)?,
            e.text_box_text_layout(d)?,
            e.text_box_columns(d)?,
        ));
    }
    Ok(out)
}

fn leaf_frame_read(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let mut out = String::new();
    for index in [0, 1] {
        let selector = DrawableSelector::text_box(index);
        let snapshot = package.body_drawable_properties(selector)?;
        let geometry = snapshot.geometry().ok_or("no frame")?;
        let style = |kind| package.body_drawable_style(selector, kind);
        let StyleProperty::TextLayout(layout) = style(StyleKind::TextLayout)? else {
            return Err("layout".into());
        };
        let StyleProperty::TextColumns(columns) = style(StyleKind::TextColumns)? else {
            return Err("columns".into());
        };
        out.push_str(&format!(
            "{:?}\n{:?}\n{:?}\n{:?}\n{:?}\n",
            host_geometry(geometry),
            host_properties(snapshot.properties()),
            layout,
            columns,
            package.body_drawable_text(selector)?,
        ));
    }
    Ok(out)
}

fn frame_case(name: &'static str, tier: Tier, input: &Fixture, op: FrameOp) -> TestResult {
    let op = &op;
    let mut case = Case::new(name, Format::Pages, tier, Evidence::Queued, input)
        .host(|source| host_frame(source, op))
        .leaf(|source| leaf_frame(source, op))
        .read("host text box", host_frame_read)
        .read("leaf text box", leaf_frame_read);
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

#[test]
fn pages_text_box_geometry_parity() -> TestResult {
    let _ = Evidence::Queued;
    frame_case(
        "pages_text_box_geometry_parity",
        Tier::A,
        &text_box_fixture()?,
        FrameOp::Geometry(Geometry {
            position: Some(Point { x: 30.0, y: 40.0 }),
            size: Some(Size {
                width: 150.0,
                height: 60.0,
            }),
            flags: Some(GeometryFlags::from_bits(3)),
            angle: Some(15.0),
        }),
    )
}

#[test]
fn pages_text_box_properties_parity() -> TestResult {
    let _ = Evidence::Queued;
    frame_case(
        "pages_text_box_properties_parity",
        Tier::A,
        &text_box_fixture()?,
        FrameOp::Properties(properties()),
    )
}

/// A fixture whose first text box carries a host layout and columns.
fn styled_fixture() -> TestResult<Fixture> {
    let input = text_box_fixture()?;
    let bytes = host_frame(
        &input.bytes,
        &FrameOp::Style(StyleProperty::TextLayout(layout()?)),
    )?;
    let bytes = host_frame(
        &bytes,
        &FrameOp::Style(StyleProperty::TextColumns(columns()?)),
    )?;
    Ok(Fixture {
        path: "basic.pages + two add_text_box + host layout and columns",
        bytes,
    })
}

#[test]
fn pages_text_box_layout_and_columns_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = text_box_fixture()?;
    for op in [
        FrameOp::Style(StyleProperty::TextLayout(layout()?)),
        FrameOp::Style(StyleProperty::TextColumns(columns()?)),
    ] {
        frame_case(
            "pages_text_box_layout_and_columns_set_parity",
            Tier::B,
            &input,
            op,
        )?;
    }
    Ok(())
}

/// In-place rewrites of the text box's own variation.
#[test]
fn pages_text_box_layout_and_columns_rewrite_parity() -> TestResult {
    let _ = Evidence::Queued;
    let styled = styled_fixture()?;
    let rewritten = Columns::equal(Count::new(3)?, Some(Gap::from_points(6.0)?));
    let moved = Layout::new(
        VerticalAlignment::Bottom,
        Insets::uniform(Inset::from_points(2.0)?),
        AutoSize::Fixed,
    );
    for op in [
        FrameOp::Style(StyleProperty::TextColumns(rewritten)),
        FrameOp::Style(StyleProperty::TextLayout(moved)),
    ] {
        frame_case(
            "pages_text_box_layout_and_columns_rewrite_parity",
            Tier::A,
            &styled,
            op,
        )?;
    }
    Ok(())
}

#[test]
fn pages_text_box_layout_and_columns_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    let styled = styled_fixture()?;
    for kind in [StyleKind::TextLayout, StyleKind::TextColumns] {
        frame_case(
            "pages_text_box_layout_and_columns_reset_parity",
            Tier::A,
            &styled,
            FrameOp::Reset(kind),
        )?;
    }
    Ok(())
}

#[test]
fn pages_text_box_frame_read_parity() -> TestResult {
    for input in [text_box_fixture()?, styled_fixture()?] {
        assert_eq!(
            host_frame_read(&input.bytes)?,
            leaf_frame_read(&input.bytes)?
        );
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

#[test]
fn pages_text_box_plain_text_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = text_box_fixture()?;
    for op in [
        TextOp::Replace(0, 5, "Ω"),
        TextOp::Set("New box text"),
        TextOp::Clear,
    ] {
        frame_case(
            "pages_text_box_plain_text_parity",
            Tier::A,
            &input,
            FrameOp::Text(op),
        )?;
    }
    Ok(())
}

// ---- cost probes (record 0413 §5) ------------------------------------------

fn host_bytes(source: &[u8], op: &FrameOp) -> TestResult<Vec<u8>> {
    host_frame(source, op)
}

fn leaf_bytes_of(source: &[u8], op: &FrameOp) -> TestResult<Vec<u8>> {
    Ok(leaf_frame_with(source, op, false)?.output)
}

#[test]
#[ignore = "callgrind cost probe (record 0413)"]
fn text_box_alignment_set_cost_abba() -> TestResult {
    let change = Change::Set(sample(PropertyKind::Alignment)?);
    crate::pages_cost_probes::abba(
        &text_box_fixture()?,
        |source| host_change(source, 0, &change),
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_text_box_text(DrawableSelector::text_box(0))?;
            edit.set(sample(PropertyKind::Alignment)?)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0413)"]
fn text_box_columns_set_cost_abba() -> TestResult {
    let op = FrameOp::Style(StyleProperty::TextColumns(columns()?));
    crate::pages_cost_probes::abba(
        &text_box_fixture()?,
        |source| host_bytes(source, &op),
        |source| leaf_bytes_of(source, &op),
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0413)"]
fn text_box_geometry_cost_abba() -> TestResult {
    let op = FrameOp::Geometry(Geometry {
        position: Some(Point { x: 30.0, y: 40.0 }),
        size: Some(Size {
            width: 150.0,
            height: 60.0,
        }),
        flags: Some(GeometryFlags::from_bits(3)),
        angle: Some(15.0),
    });
    crate::pages_cost_probes::abba(
        &text_box_fixture()?,
        |source| host_bytes(source, &op),
        |source| leaf_bytes_of(source, &op),
    )
}
