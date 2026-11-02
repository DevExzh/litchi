//! Keynote slide-table cell text-property parity (record 0412): the host
//! editor's `set_slide_table_cell_*`, `reset_slide_table_cell_*` and
//! `slide_table_cell_*` families against the leaf's one property family
//! `edit_slide_table_cell_text` / `slide_table_cell_text`, which runs the
//! shared cell rich-text engine (`litchi_iwa_table::cell_rich_text::text`)
//! lifted from the Numbers family of record 0370.

use litchi_iwa::keynote;
use litchi_keynote::slide::table::TableSelector;
use litchi_keynote::slide::table::cell_text::{At, Property, PropertyKind};
use litchi_keynote::{Package, SlideSelector};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::decode_as;
use crate::leaf_bytes;
use crate::text_style_cases::sample;

const FIXTURE: &str = "iwork/keynote/slide-table-read-native.key";
const SLIDE: usize = 0;
const TABLE: usize = 0;
const ROW: usize = 1;
const COLUMN: usize = 1;

/// Keynote fixtures with a slide table.
pub const CORPUS: &[&str] = &[
    "iwork/keynote/slide-table-read-native.key",
    "iwork/keynote/slide-table-merges-native.key",
    "iwork/keynote/table-discovery.key",
];

/// The cell properties the host editor sets, resets and reads.
pub const KINDS: [PropertyKind; 17] = [
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
    PropertyKind::Alignment,
    PropertyKind::LineSpacing,
    PropertyKind::Spacing,
    PropertyKind::Indents,
    PropertyKind::TabStops,
];

fn host_model(editor: &keynote::Draft) -> TestResult<u64> {
    Ok(editor
        .slide_tables(SLIDE)?
        .get(TABLE)
        .ok_or("slide table missing")?
        .model_object_id)
}

/// The host's set route for `property`.
pub fn host_set(editor: &mut keynote::Draft, id: u64, property: Property) -> TestResult {
    let (r, c) = (ROW, COLUMN);
    match property {
        Property::TextStyle(v) => editor.set_slide_table_cell_text_style(SLIDE, id, r, c, v)?,
        Property::TextFont(v) => editor.set_slide_table_cell_text_font(SLIDE, id, r, c, v)?,
        Property::TextDecorations(v) => {
            editor.set_slide_table_cell_text_decorations(SLIDE, id, r, c, v)?
        },
        Property::TextColor(v) => editor.set_slide_table_cell_text_color(SLIDE, id, r, c, v)?,
        Property::TextCapitalization(v) => {
            editor.set_slide_table_cell_text_capitalization(SLIDE, id, r, c, v)?
        },
        Property::TextScript(v) => editor.set_slide_table_cell_text_script(SLIDE, id, r, c, v)?,
        Property::TextBaselineShift(v) => {
            editor.set_slide_table_cell_text_baseline_shift(SLIDE, id, r, c, v)?
        },
        Property::TextCharacterSpacing(v) => {
            editor.set_slide_table_cell_text_character_spacing(SLIDE, id, r, c, v)?;
        },
        Property::TextLigatures(v) => {
            editor.set_slide_table_cell_text_ligatures(SLIDE, id, r, c, v)?
        },
        Property::TextOutline(v) => editor.set_slide_table_cell_text_outline(SLIDE, id, r, c, v)?,
        Property::TextShadow(v) => editor.set_slide_table_cell_text_shadow(SLIDE, id, r, c, v)?,
        Property::TextBackground(v) => {
            editor.set_slide_table_cell_text_background(SLIDE, id, r, c, v)?
        },
        Property::Alignment(v) => editor.set_slide_table_cell_text_alignment(SLIDE, id, r, c, v)?,
        Property::LineSpacing(v) => {
            editor.set_slide_table_cell_paragraph_line_spacing(SLIDE, id, r, c, v)?
        },
        Property::Spacing(v) => {
            editor.set_slide_table_cell_paragraph_spacing(SLIDE, id, r, c, v)?
        },
        Property::Indents(v) => {
            editor.set_slide_table_cell_paragraph_indents(SLIDE, id, r, c, v)?
        },
        Property::TabStops(v) => {
            editor.set_slide_table_cell_paragraph_tab_stops(SLIDE, id, r, c, v)?
        },
        other => return Err(format!("the host has no cell setter for {other:?}").into()),
    }
    Ok(())
}

/// The host's reset route for `kind`.
pub fn host_reset(editor: &mut keynote::Draft, id: u64, kind: PropertyKind) -> TestResult<bool> {
    let (r, c) = (ROW, COLUMN);
    Ok(match kind {
        PropertyKind::TextStyle => editor.reset_slide_table_cell_text_style(SLIDE, id, r, c)?,
        PropertyKind::TextFont => editor.reset_slide_table_cell_text_font(SLIDE, id, r, c)?,
        PropertyKind::TextDecorations => {
            editor.reset_slide_table_cell_text_decorations(SLIDE, id, r, c)?
        },
        PropertyKind::TextColor => editor.reset_slide_table_cell_text_color(SLIDE, id, r, c)?,
        PropertyKind::TextCapitalization => {
            editor.reset_slide_table_cell_text_capitalization(SLIDE, id, r, c)?
        },
        PropertyKind::TextScript => editor.reset_slide_table_cell_text_script(SLIDE, id, r, c)?,
        PropertyKind::TextBaselineShift => {
            editor.reset_slide_table_cell_text_baseline_shift(SLIDE, id, r, c)?
        },
        PropertyKind::TextCharacterSpacing => {
            editor.reset_slide_table_cell_text_character_spacing(SLIDE, id, r, c)?
        },
        PropertyKind::TextLigatures => {
            editor.reset_slide_table_cell_text_ligatures(SLIDE, id, r, c)?
        },
        PropertyKind::TextOutline => editor.reset_slide_table_cell_text_outline(SLIDE, id, r, c)?,
        PropertyKind::TextShadow => editor.reset_slide_table_cell_text_shadow(SLIDE, id, r, c)?,
        PropertyKind::TextBackground => {
            editor.reset_slide_table_cell_text_background(SLIDE, id, r, c)?
        },
        PropertyKind::Alignment => editor.reset_slide_table_cell_text_alignment(SLIDE, id, r, c)?,
        PropertyKind::LineSpacing => {
            editor.reset_slide_table_cell_paragraph_line_spacing(SLIDE, id, r, c)?
        },
        PropertyKind::Spacing => {
            editor.reset_slide_table_cell_paragraph_spacing(SLIDE, id, r, c)?
        },
        PropertyKind::Indents => {
            editor.reset_slide_table_cell_paragraph_indents(SLIDE, id, r, c)?
        },
        PropertyKind::TabStops => {
            editor.reset_slide_table_cell_paragraph_tab_stops(SLIDE, id, r, c)?
        },
        other => return Err(format!("the host has no cell reset for {other:?}").into()),
    })
}

/// The host's read route for `kind`.
pub fn host_read(bytes: &[u8], (r, c): (usize, usize), kind: PropertyKind) -> TestResult<Property> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let id = host_model(&editor)?;

    Ok(match kind {
        PropertyKind::TextStyle => {
            Property::TextStyle(editor.slide_table_cell_text_style(SLIDE, id, r, c)?)
        },
        PropertyKind::TextFont => {
            Property::TextFont(editor.slide_table_cell_text_font(SLIDE, id, r, c)?)
        },
        PropertyKind::TextDecorations => {
            Property::TextDecorations(editor.slide_table_cell_text_decorations(SLIDE, id, r, c)?)
        },
        PropertyKind::TextColor => {
            Property::TextColor(editor.slide_table_cell_text_color(SLIDE, id, r, c)?)
        },
        PropertyKind::TextCapitalization => Property::TextCapitalization(
            editor.slide_table_cell_text_capitalization(SLIDE, id, r, c)?,
        ),
        PropertyKind::TextScript => {
            Property::TextScript(editor.slide_table_cell_text_script(SLIDE, id, r, c)?)
        },
        PropertyKind::TextBaselineShift => Property::TextBaselineShift(
            editor.slide_table_cell_text_baseline_shift(SLIDE, id, r, c)?,
        ),
        PropertyKind::TextCharacterSpacing => Property::TextCharacterSpacing(
            editor.slide_table_cell_text_character_spacing(SLIDE, id, r, c)?,
        ),
        PropertyKind::TextLigatures => {
            Property::TextLigatures(editor.slide_table_cell_text_ligatures(SLIDE, id, r, c)?)
        },
        PropertyKind::TextOutline => {
            Property::TextOutline(editor.slide_table_cell_text_outline(SLIDE, id, r, c)?)
        },
        PropertyKind::TextShadow => {
            Property::TextShadow(editor.slide_table_cell_text_shadow(SLIDE, id, r, c)?)
        },
        PropertyKind::TextBackground => {
            Property::TextBackground(editor.slide_table_cell_text_background(SLIDE, id, r, c)?)
        },
        PropertyKind::Alignment => {
            Property::Alignment(editor.slide_table_cell_text_alignment(SLIDE, id, r, c)?)
        },
        PropertyKind::LineSpacing => {
            Property::LineSpacing(editor.slide_table_cell_paragraph_line_spacing(SLIDE, id, r, c)?)
        },
        PropertyKind::Spacing => {
            Property::Spacing(editor.slide_table_cell_paragraph_spacing(SLIDE, id, r, c)?)
        },
        PropertyKind::Indents => {
            Property::Indents(editor.slide_table_cell_paragraph_indents(SLIDE, id, r, c)?)
        },
        PropertyKind::TabStops => {
            Property::TabStops(editor.slide_table_cell_paragraph_tab_stops(SLIDE, id, r, c)?)
        },
        other => return Err(format!("the host has no cell read for {other:?}").into()),
    })
}

fn leaf_read(bytes: &[u8], (r, c): (usize, usize), kind: PropertyKind) -> TestResult<Property> {
    Ok(Package::from_bytes(bytes)?.slide_table_cell_text(
        SlideSelector::index(SLIDE),
        TableSelector::index(TABLE),
        At::new(r as u32, c as u32),
        kind,
    )?)
}

fn case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    change: Result<Property, PropertyKind>,
) -> TestResult {
    let kind = match &change {
        Ok(property) => property.kind(),
        Err(kind) => *kind,
    };
    let change = &change;
    Case::new(name, Format::Keynote, tier, evidence, input)
        .leaf_structural_metadata()
        .known_fields(|known| {
            known.with(
                2_022,
                "TSWP.ParagraphStyleArchive",
                decode_as::<litchi_iwa_protos::tswp::ParagraphStyleArchive>,
            )
        })
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let id = host_model(&editor)?;
            match change {
                Ok(property) => host_set(&mut editor, id, property.clone())?,
                Err(kind) => {
                    if !host_reset(&mut editor, id, *kind)? {
                        return Err("the host reset changed nothing".into());
                    }
                },
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let edit = package.edit_slide_table_cell_text(
                SlideSelector::index(SLIDE),
                TableSelector::index(TABLE),
            )?;
            let at = At::new(ROW as u32, COLUMN as u32);
            let edit = match change {
                Ok(property) => edit.set(at, property.clone())?,
                Err(kind) => edit.reset(at, *kind)?,
            };
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_slide_table_cell_text(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host read", move |bytes| {
            host_read(bytes, (ROW, COLUMN), kind)
        })
        .read("leaf read", move |bytes| {
            leaf_read(bytes, (ROW, COLUMN), kind)
        })
        .run()?;
    Ok(())
}

/// Every host cell property: set its sample on the fixture's cell; on
/// the host's output, reset it and set the next property (the cell's own
/// variation is rewritten in place); the effective values read equal on
/// both routes.
#[test]
fn keynote_slide_table_cell_text_property_parity() -> TestResult {
    let input = fixture(FIXTURE)?;
    let mut failures = Vec::new();
    for kind in KINDS {
        let property = sample(kind)?;
        if let Err(error) = case(
            "keynote_slide_table_cell_text_property_parity",
            Tier::B,
            Evidence::Queued,
            &input,
            Ok(property.clone()),
        ) {
            failures.push(format!("set {kind:?}: {error}"));
        }
        // Reset on the host's output.
        let mut editor = keynote::Draft::from_bytes(&input.bytes)?;
        let id = host_model(&editor)?;
        host_set(&mut editor, id, property)?;
        let set = Fixture {
            path: FIXTURE,
            bytes: editor.to_bytes()?,
        };
        if let Err(error) = case(
            "keynote_slide_table_cell_text_property_parity",
            Tier::B,
            Evidence::Queued,
            &set,
            Err(kind),
        ) {
            failures.push(format!("reset {kind:?}: {error}"));
        }
        // A second property on the cell's own variation (rewrite in place).
        let next =
            KINDS[(KINDS.iter().position(|other| *other == kind).unwrap_or(0) + 1) % KINDS.len()];
        if let Err(error) = case(
            "keynote_slide_table_cell_text_property_parity",
            Tier::B,
            Evidence::Queued,
            &set,
            Ok(sample(next)?),
        ) {
            failures.push(format!("set {next:?} after {kind:?}: {error}"));
        }
    }
    if failures.is_empty() {
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}

/// Every host cell property reads the same effective value on the first
/// cells of every fixture's table (when the host reads it).
#[test]
fn keynote_slide_table_cell_text_property_read_parity() -> TestResult {
    let (_tier, _evidence) = (Tier::A, Evidence::Queued);
    let mut compared = 0_usize;
    for path in CORPUS {
        let input = fixture(path)?;
        for cell in [(0, 0), (1, 1), (2, 1), (1, 2), (3, 0)] {
            for kind in KINDS {
                let Ok(host) = host_read(&input.bytes, cell, kind) else {
                    continue;
                };
                assert_eq!(
                    leaf_read(&input.bytes, cell, kind)?,
                    host,
                    "{path} {cell:?} {kind:?}"
                );
                compared += 1;
            }
        }
    }
    eprintln!("cell text reads compared: {compared}");
    assert!(compared >= 100, "only {compared} reads compared");
    Ok(())
}

/// Callgrind probe (record 0412): paragraph alignment set on the fixture's
/// cell (a derived variation), host against leaf in ABBA order.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_cell_text_cost_abba() -> TestResult {
    let input = fixture(FIXTURE)?;
    let property = sample(PropertyKind::Alignment)?;
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let id = host_model(&editor)?;
            host_set(&mut editor, id, property.clone())?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_slide_table_cell_text(
                    SlideSelector::index(SLIDE),
                    TableSelector::index(TABLE),
                )?
                .set(At::new(ROW as u32, COLUMN as u32), property.clone())?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}
