//! Keynote text-box character and paragraph property parity (wave W1-C,
//! record 0372): the host `set_slide_text_box_*` / `reset_slide_text_box_*`
//! / `slide_text_box_*` families against
//! `slide::drawable::Edit::{set_text_property, reset_text_property}` and
//! `Package::slide_drawable_text_property` on the shared W1-D engine.

use litchi_iwa::keynote;
use litchi_iwa_protos::tswp;
use litchi_iwa_text_wire::paragraph_style::{self, Property, PropertyKind};
use litchi_keynote::slide::drawable::Kind;
use litchi_keynote::{DrawableSelector, Package, SlideSelector};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::keynote_drawable_cases::{host_id, leaf_position};
use crate::known_fields::decode_as;
use crate::leaf_bytes;
use crate::text_style_cases::{sample, second_sample};

const SLIDE: usize = 0;
const TEXT_BOX: &str = "iwork/keynote/basic.key";

/// Every property with a Keynote text-box route.
const KINDS: [PropertyKind; 23] = [
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
    PropertyKind::Alignment,
    PropertyKind::LineSpacing,
    PropertyKind::Spacing,
    PropertyKind::Indents,
    PropertyKind::DecimalTabCharacter,
    PropertyKind::DefaultTabInterval,
    PropertyKind::TabStops,
];

#[derive(Clone)]
enum Change {
    Set(Property),
    Reset(PropertyKind),
}

#[allow(clippy::too_many_lines, reason = "one arm per host method")]
fn host_set(editor: &mut keynote::Draft, id: u64, property: Property) -> TestResult {
    match property {
        Property::TextStyle(v) => editor.set_slide_text_box_text_style(SLIDE, id, v)?,
        Property::TextFont(v) => editor.set_slide_text_box_text_font(SLIDE, id, v)?,
        Property::TextDecorations(v) => editor.set_slide_text_box_text_decorations(SLIDE, id, v)?,
        Property::TextColor(v) => editor.set_slide_text_box_text_color(SLIDE, id, v)?,
        Property::TextCapitalization(v) => {
            editor.set_slide_text_box_text_capitalization(SLIDE, id, v)?;
        },
        Property::TextScript(v) => editor.set_slide_text_box_text_script(SLIDE, id, v)?,
        Property::TextBaselineShift(v) => {
            editor.set_slide_text_box_text_baseline_shift(SLIDE, id, v)?;
        },
        Property::TextCharacterSpacing(v) => {
            editor.set_slide_text_box_text_character_spacing(SLIDE, id, v)?;
        },
        Property::TextLigatures(v) => editor.set_slide_text_box_text_ligatures(SLIDE, id, v)?,
        Property::TextOutline(v) => editor.set_slide_text_box_text_outline(SLIDE, id, v)?,
        Property::TextShadow(v) => editor.set_slide_text_box_text_shadow(SLIDE, id, v)?,
        Property::TextBackground(v) => editor.set_slide_text_box_text_background(SLIDE, id, v)?,
        Property::ParagraphBackground(v) => {
            editor.set_slide_text_box_paragraph_background(SLIDE, id, v)?;
        },
        Property::Borders(v) => editor.set_slide_text_box_paragraph_borders(SLIDE, id, v)?,
        Property::Flow(v) => editor.set_slide_text_box_paragraph_flow(SLIDE, id, v)?,
        Property::WritingDirection(v) => {
            editor.set_slide_text_box_paragraph_writing_direction(SLIDE, id, v)?;
        },
        Property::Alignment(v) => editor.set_slide_text_box_paragraph_alignment(SLIDE, id, v)?,
        Property::LineSpacing(v) => {
            editor.set_slide_text_box_paragraph_line_spacing(SLIDE, id, v)?;
        },
        Property::Spacing(v) => editor.set_slide_text_box_paragraph_spacing(SLIDE, id, v)?,
        Property::Indents(v) => editor.set_slide_text_box_paragraph_indents(SLIDE, id, v)?,
        Property::DecimalTabCharacter(v) => {
            editor.set_slide_text_box_paragraph_decimal_tab_character(SLIDE, id, v)?;
        },
        Property::DefaultTabInterval(v) => {
            editor.set_slide_text_box_paragraph_default_tab_interval(SLIDE, id, v)?;
        },
        Property::TabStops(v) => editor.set_slide_text_box_paragraph_tab_stops(SLIDE, id, v)?,
        other => return Err(format!("unmapped property {other:?}").into()),
    }
    Ok(())
}

#[allow(clippy::too_many_lines, reason = "one arm per host method")]
fn host_reset(editor: &mut keynote::Draft, id: u64, kind: PropertyKind) -> TestResult {
    let _changed = match kind {
        PropertyKind::TextStyle => editor.reset_slide_text_box_text_style(SLIDE, id)?,
        PropertyKind::TextFont => editor.reset_slide_text_box_text_font(SLIDE, id)?,
        PropertyKind::TextDecorations => editor.reset_slide_text_box_text_decorations(SLIDE, id)?,
        PropertyKind::TextColor => editor.reset_slide_text_box_text_color(SLIDE, id)?,
        PropertyKind::TextCapitalization => {
            editor.reset_slide_text_box_text_capitalization(SLIDE, id)?
        },
        PropertyKind::TextScript => editor.reset_slide_text_box_text_script(SLIDE, id)?,
        PropertyKind::TextBaselineShift => {
            editor.reset_slide_text_box_text_baseline_shift(SLIDE, id)?
        },
        PropertyKind::TextCharacterSpacing => {
            editor.reset_slide_text_box_text_character_spacing(SLIDE, id)?
        },
        PropertyKind::TextLigatures => editor.reset_slide_text_box_text_ligatures(SLIDE, id)?,
        PropertyKind::TextOutline => editor.reset_slide_text_box_text_outline(SLIDE, id)?,
        PropertyKind::TextShadow => editor.reset_slide_text_box_text_shadow(SLIDE, id)?,
        PropertyKind::TextBackground => editor.reset_slide_text_box_text_background(SLIDE, id)?,
        PropertyKind::ParagraphBackground => {
            editor.reset_slide_text_box_paragraph_background(SLIDE, id)?
        },
        PropertyKind::Borders => editor.reset_slide_text_box_paragraph_borders(SLIDE, id)?,
        PropertyKind::Flow => editor.reset_slide_text_box_paragraph_flow(SLIDE, id)?,
        PropertyKind::WritingDirection => {
            editor.reset_slide_text_box_paragraph_writing_direction(SLIDE, id)?
        },
        PropertyKind::Alignment => editor.reset_slide_text_box_paragraph_alignment(SLIDE, id)?,
        PropertyKind::LineSpacing => {
            editor.reset_slide_text_box_paragraph_line_spacing(SLIDE, id)?
        },
        PropertyKind::Spacing => editor.reset_slide_text_box_paragraph_spacing(SLIDE, id)?,
        PropertyKind::Indents => editor.reset_slide_text_box_paragraph_indents(SLIDE, id)?,
        PropertyKind::DecimalTabCharacter => {
            editor.reset_slide_text_box_paragraph_decimal_tab_character(SLIDE, id)?
        },
        PropertyKind::DefaultTabInterval => {
            editor.reset_slide_text_box_paragraph_default_tab_interval(SLIDE, id)?
        },
        PropertyKind::TabStops => editor.reset_slide_text_box_paragraph_tab_stops(SLIDE, id)?,
        other => return Err(format!("unmapped kind {other:?}").into()),
    };
    Ok(())
}

#[allow(clippy::too_many_lines, reason = "one arm per host method")]
fn host_read(bytes: &[u8], kind: PropertyKind) -> TestResult<Property> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let id = host_id(&editor, bytes, leaf_position(bytes, Kind::TextBox)?)?;
    Ok(match kind {
        PropertyKind::TextStyle => {
            Property::TextStyle(editor.slide_text_box_text_style(SLIDE, id)?)
        },
        PropertyKind::TextFont => Property::TextFont(editor.slide_text_box_text_font(SLIDE, id)?),
        PropertyKind::TextDecorations => {
            Property::TextDecorations(editor.slide_text_box_text_decorations(SLIDE, id)?)
        },
        PropertyKind::TextColor => {
            Property::TextColor(editor.slide_text_box_text_color(SLIDE, id)?)
        },
        PropertyKind::TextCapitalization => {
            Property::TextCapitalization(editor.slide_text_box_text_capitalization(SLIDE, id)?)
        },
        PropertyKind::TextScript => {
            Property::TextScript(editor.slide_text_box_text_script(SLIDE, id)?)
        },
        PropertyKind::TextBaselineShift => {
            Property::TextBaselineShift(editor.slide_text_box_text_baseline_shift(SLIDE, id)?)
        },
        PropertyKind::TextCharacterSpacing => {
            Property::TextCharacterSpacing(editor.slide_text_box_text_character_spacing(SLIDE, id)?)
        },
        PropertyKind::TextLigatures => {
            Property::TextLigatures(editor.slide_text_box_text_ligatures(SLIDE, id)?)
        },
        PropertyKind::TextOutline => {
            Property::TextOutline(editor.slide_text_box_text_outline(SLIDE, id)?)
        },
        PropertyKind::TextShadow => {
            Property::TextShadow(editor.slide_text_box_text_shadow(SLIDE, id)?)
        },
        PropertyKind::TextBackground => {
            Property::TextBackground(editor.slide_text_box_text_background(SLIDE, id)?)
        },
        PropertyKind::ParagraphBackground => {
            Property::ParagraphBackground(editor.slide_text_box_paragraph_background(SLIDE, id)?)
        },
        PropertyKind::Borders => {
            Property::Borders(editor.slide_text_box_paragraph_borders(SLIDE, id)?)
        },
        PropertyKind::Flow => Property::Flow(editor.slide_text_box_paragraph_flow(SLIDE, id)?),
        PropertyKind::WritingDirection => Property::WritingDirection(
            editor.slide_text_box_paragraph_writing_direction(SLIDE, id)?,
        ),
        PropertyKind::Alignment => {
            Property::Alignment(editor.slide_text_box_paragraph_alignment(SLIDE, id)?)
        },
        PropertyKind::LineSpacing => {
            Property::LineSpacing(editor.slide_text_box_paragraph_line_spacing(SLIDE, id)?)
        },
        PropertyKind::Spacing => {
            Property::Spacing(editor.slide_text_box_paragraph_spacing(SLIDE, id)?)
        },
        PropertyKind::Indents => {
            Property::Indents(editor.slide_text_box_paragraph_indents(SLIDE, id)?)
        },
        PropertyKind::DecimalTabCharacter => Property::DecimalTabCharacter(
            editor.slide_text_box_paragraph_decimal_tab_character(SLIDE, id)?,
        ),
        PropertyKind::DefaultTabInterval => Property::DefaultTabInterval(
            editor.slide_text_box_paragraph_default_tab_interval(SLIDE, id)?,
        ),
        PropertyKind::TabStops => {
            Property::TabStops(editor.slide_text_box_paragraph_tab_stops(SLIDE, id)?)
        },
        other => return Err(format!("unmapped kind {other:?}").into()),
    })
}

fn leaf_read(bytes: &[u8], kind: PropertyKind) -> TestResult<Property> {
    let position = leaf_position(bytes, Kind::TextBox)?;
    Ok(Package::from_bytes(bytes)?.slide_drawable_text_property(
        SlideSelector::index(SLIDE),
        DrawableSelector::index(position),
        kind,
    )?)
}

fn host_change(source: &[u8], change: &Change) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let id = host_id(&editor, source, leaf_position(source, Kind::TextBox)?)?;
    match change.clone() {
        Change::Set(property) => host_set(&mut editor, id, property)?,
        Change::Reset(kind) => host_reset(&mut editor, id, kind)?,
    }
    Ok(editor.to_bytes()?)
}

fn leaf_change(source: &[u8], change: &Change) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let position = leaf_position(source, Kind::TextBox)?;
    let (slide, drawable) = (
        SlideSelector::index(SLIDE),
        DrawableSelector::index(position),
    );
    let mut edit = package.edit_slide_drawables();
    match change.clone() {
        Change::Set(property) => edit.set_text_property(slide, drawable, property)?,
        Change::Reset(kind) => edit.reset_text_property(slide, drawable, kind)?,
    };
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_drawables(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn run(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    change: &Change,
) -> TestResult {
    let kind = match change {
        Change::Set(property) => property.kind(),
        Change::Reset(kind) => *kind,
    };
    Case::new(name, Format::Keynote, tier, evidence, input)
        .host(|source| host_change(source, change))
        .leaf(|source| leaf_change(source, change))
        .read("host property", move |bytes| host_read(bytes, kind))
        .read("leaf property", move |bytes| leaf_read(bytes, kind))
        .known_fields(|known| {
            known.with(
                paragraph_style::MESSAGE_TYPE,
                "TSWP.ParagraphStyleArchive",
                decode_as::<tswp::ParagraphStyleArchive>,
            )
        })
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .run()?;
    Ok(())
}

/// The host output of setting `sample(kind)` on the text box: its storage
/// has an exclusive managed variation.
fn derived(kind: PropertyKind) -> TestResult<Fixture> {
    let input = fixture(TEXT_BOX)?;
    Ok(Fixture {
        path: input.path,
        bytes: host_change(&input.bytes, &Change::Set(sample(kind)?))?,
    })
}

/// Set on the native text box: the engine derives a new variation.
#[test]
fn keynote_slide_text_box_property_set_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    for kind in KINDS {
        run(
            "keynote_slide_text_box_property_set_parity",
            Tier::B,
            Evidence::Queued,
            &input,
            &Change::Set(sample(kind)?),
        )?;
    }
    Ok(())
}

/// Set over an exclusive variation: the engine rewrites it in place.
#[test]
fn keynote_slide_text_box_property_rewrite_parity() -> TestResult {
    for kind in KINDS {
        run(
            "keynote_slide_text_box_property_rewrite_parity",
            Tier::B,
            Evidence::Queued,
            &derived(kind)?,
            &Change::Set(second_sample(kind)?),
        )?;
    }
    Ok(())
}

/// Reset of the only override: the engine collapses the variation.
#[test]
fn keynote_slide_text_box_property_reset_parity() -> TestResult {
    for kind in KINDS {
        run(
            "keynote_slide_text_box_property_reset_parity",
            Tier::B,
            Evidence::Queued,
            &derived(kind)?,
            &Change::Reset(kind),
        )?;
    }
    Ok(())
}

/// Host and leaf reads agree on every property of the native and the
/// derived text box.
#[test]
fn keynote_slide_text_box_property_read_parity() -> TestResult {
    let native = fixture(TEXT_BOX)?;
    for kind in KINDS {
        for input in [&native, &derived(kind)?] {
            assert_eq!(
                host_read(&input.bytes, kind)?,
                leaf_read(&input.bytes, kind)?
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
