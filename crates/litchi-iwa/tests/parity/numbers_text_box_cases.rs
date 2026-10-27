//! Numbers sheet text-box character and paragraph property parity (wave
//! W2-A, record 0378): the host `set_sheet_text_box_*`,
//! `reset_sheet_text_box_*` and `sheet_text_box_*` property families
//! against `litchi_numbers::Package::edit_sheet_text_box_text`.

use litchi_iwa::numbers;
use litchi_iwa_protos::tswp;
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::DrawableSelector;
use litchi_numbers::sheet::drawable::text::{Property, PropertyKind};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::known_fields::decode_as;
use crate::leaf_bytes;
use crate::numbers_drawable_cases::{authored_fixture, host_drawable};
use crate::text_style_cases::{sample, second_sample};
use litchi_numbers::sheet::drawable::Kind;

/// The properties the host's text-box families cover (every engine property
/// but the following style, which Numbers text boxes do not expose).
pub const KINDS: [PropertyKind; 23] = [
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

/// One change, expressed once for both routes.
#[derive(Debug, Clone)]
pub enum Change {
    Set(Property),
    Reset(PropertyKind),
}

/// The host route: the per-property `set_sheet_text_box_*` /
/// `reset_sheet_text_box_*` method.
#[allow(clippy::too_many_lines, reason = "one arm per host method")]
pub fn host_change(source: &[u8], index: usize, change: &Change) -> TestResult<Vec<u8>> {
    let mut e = numbers::Draft::from_bytes(source)?;
    let (s, d) = host_drawable(&e, Kind::TextBox, index)?;
    macro_rules! set {
        ($($variant:ident => $method:ident),* $(,)?) => {
            match change {
                $(Change::Set(Property::$variant(v)) => e.$method(s, d, v.clone())?,)*
                _ => return Err(format!("unmapped change {change:?}").into()),
            }
        };
    }
    macro_rules! reset {
        ($($variant:ident => $method:ident),* $(,)?) => {
            match change {
                $(Change::Reset(PropertyKind::$variant) => drop(e.$method(s, d)?),)*
                _ => return Err(format!("unmapped change {change:?}").into()),
            }
        };
    }
    match change {
        Change::Set(_) => set!(
            TextStyle => set_sheet_text_box_text_style,
            TextFont => set_sheet_text_box_text_font,
            TextDecorations => set_sheet_text_box_text_decorations,
            TextColor => set_sheet_text_box_text_color,
            TextCapitalization => set_sheet_text_box_text_capitalization,
            TextScript => set_sheet_text_box_text_script,
            TextBaselineShift => set_sheet_text_box_text_baseline_shift,
            TextCharacterSpacing => set_sheet_text_box_text_character_spacing,
            TextLigatures => set_sheet_text_box_text_ligatures,
            TextOutline => set_sheet_text_box_text_outline,
            TextShadow => set_sheet_text_box_text_shadow,
            TextBackground => set_sheet_text_box_text_background,
            ParagraphBackground => set_sheet_text_box_paragraph_background,
            Borders => set_sheet_text_box_paragraph_borders,
            Flow => set_sheet_text_box_paragraph_flow,
            WritingDirection => set_sheet_text_box_paragraph_writing_direction,
            Alignment => set_sheet_text_box_paragraph_alignment,
            LineSpacing => set_sheet_text_box_paragraph_line_spacing,
            Spacing => set_sheet_text_box_paragraph_spacing,
            Indents => set_sheet_text_box_paragraph_indents,
            DecimalTabCharacter => set_sheet_text_box_paragraph_decimal_tab_character,
            DefaultTabInterval => set_sheet_text_box_paragraph_default_tab_interval,
            TabStops => set_sheet_text_box_paragraph_tab_stops,
        ),
        Change::Reset(_) => reset!(
            TextStyle => reset_sheet_text_box_text_style,
            TextFont => reset_sheet_text_box_text_font,
            TextDecorations => reset_sheet_text_box_text_decorations,
            TextColor => reset_sheet_text_box_text_color,
            TextCapitalization => reset_sheet_text_box_text_capitalization,
            TextScript => reset_sheet_text_box_text_script,
            TextBaselineShift => reset_sheet_text_box_text_baseline_shift,
            TextCharacterSpacing => reset_sheet_text_box_text_character_spacing,
            TextLigatures => reset_sheet_text_box_text_ligatures,
            TextOutline => reset_sheet_text_box_text_outline,
            TextShadow => reset_sheet_text_box_text_shadow,
            TextBackground => reset_sheet_text_box_text_background,
            ParagraphBackground => reset_sheet_text_box_paragraph_background,
            Borders => reset_sheet_text_box_paragraph_borders,
            Flow => reset_sheet_text_box_paragraph_flow,
            WritingDirection => reset_sheet_text_box_paragraph_writing_direction,
            Alignment => reset_sheet_text_box_paragraph_alignment,
            LineSpacing => reset_sheet_text_box_paragraph_line_spacing,
            Spacing => reset_sheet_text_box_paragraph_spacing,
            Indents => reset_sheet_text_box_paragraph_indents,
            DecimalTabCharacter => reset_sheet_text_box_paragraph_decimal_tab_character,
            DefaultTabInterval => reset_sheet_text_box_paragraph_default_tab_interval,
            TabStops => reset_sheet_text_box_paragraph_tab_stops,
        ),
    }
    Ok(e.to_bytes()?)
}

/// The host read: the per-property `sheet_text_box_*` method.
pub fn host_read(bytes: &[u8], index: usize, kind: PropertyKind) -> TestResult<Property> {
    let e = numbers::Draft::from_bytes(bytes)?;
    let (s, d) = host_drawable(&e, Kind::TextBox, index)?;
    macro_rules! read {
        ($($variant:ident => $method:ident),* $(,)?) => {
            match kind {
                $(PropertyKind::$variant => Property::$variant(e.$method(s, d)?),)*
                other => return Err(format!("unmapped kind {other:?}").into()),
            }
        };
    }
    Ok(read!(
        TextStyle => sheet_text_box_text_style,
        TextFont => sheet_text_box_text_font,
        TextDecorations => sheet_text_box_text_decorations,
        TextColor => sheet_text_box_text_color,
        TextCapitalization => sheet_text_box_text_capitalization,
        TextScript => sheet_text_box_text_script,
        TextBaselineShift => sheet_text_box_text_baseline_shift,
        TextCharacterSpacing => sheet_text_box_text_character_spacing,
        TextLigatures => sheet_text_box_text_ligatures,
        TextOutline => sheet_text_box_text_outline,
        TextShadow => sheet_text_box_text_shadow,
        TextBackground => sheet_text_box_text_background,
        ParagraphBackground => sheet_text_box_paragraph_background,
        Borders => sheet_text_box_paragraph_borders,
        Flow => sheet_text_box_paragraph_flow,
        WritingDirection => sheet_text_box_paragraph_writing_direction,
        Alignment => sheet_text_box_paragraph_alignment,
        LineSpacing => sheet_text_box_paragraph_line_spacing,
        Spacing => sheet_text_box_paragraph_spacing,
        Indents => sheet_text_box_paragraph_indents,
        DecimalTabCharacter => sheet_text_box_paragraph_decimal_tab_character,
        DefaultTabInterval => sheet_text_box_paragraph_default_tab_interval,
        TabStops => sheet_text_box_paragraph_tab_stops,
    ))
}

/// The leaf route through the text-box text family.
pub fn leaf_change(source: &[u8], index: usize, change: &Change) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_text_box_text(0, DrawableSelector::text_box(index))?;
    match change {
        Change::Set(property) => drop(edit.set(property.clone())?),
        Change::Reset(kind) => drop(edit.reset(*kind)?),
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_sheet_text_box_text(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

pub fn leaf_read(bytes: &[u8], index: usize, kind: PropertyKind) -> TestResult<Property> {
    Ok(Package::from_bytes(bytes)?.sheet_text_box_text(
        0,
        DrawableSelector::text_box(index),
        kind,
    )?)
}

fn change_kind(change: &Change) -> PropertyKind {
    match change {
        Change::Set(property) => property.kind(),
        Change::Reset(kind) => *kind,
    }
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
    let mut case = Case::new(name, Format::Numbers, tier, Evidence::Queued, input)
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

/// The host output of setting `sample(kind)` on text box `index`: a fixture
/// whose text box has an exclusive managed variation.
pub fn derived(input: &Fixture, index: usize, kind: PropertyKind) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_change(&input.bytes, index, &Change::Set(sample(kind)?))?,
    })
}

#[test]
fn numbers_sheet_text_box_property_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = authored_fixture()?;
    for index in [0, 1] {
        for kind in KINDS {
            run(
                "numbers_sheet_text_box_property_set_parity",
                Tier::B,
                &input,
                index,
                &Change::Set(sample(kind)?),
            )
            .map_err(|e| format!("text box {index} {kind:?}: {e}"))?;
        }
    }
    Ok(())
}

#[test]
fn numbers_sheet_text_box_property_rewrite_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = authored_fixture()?;
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
                "numbers_sheet_text_box_property_rewrite_parity",
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
fn numbers_sheet_text_box_property_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = authored_fixture()?;
    for kind in KINDS {
        let derived = derived(&input, 0, kind)?;
        run(
            "numbers_sheet_text_box_property_reset_parity",
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
fn numbers_sheet_text_box_property_read_parity() -> TestResult {
    let input = authored_fixture()?;
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

// ---- languages -----------------------------------------------------------

use litchi_numbers::sheet::drawable::language::{TextLanguage, TextPosition};

#[derive(Clone, Debug)]
enum LanguageChange {
    Set(usize, &'static str),
    Remove(usize),
    Reset,
}

fn host_languages(bytes: &[u8]) -> TestResult<String> {
    let e = numbers::Draft::from_bytes(bytes)?;
    let (s, d) = host_drawable(&e, Kind::TextBox, 0)?;
    Ok(format!("{:?}", e.sheet_text_box_text_languages(s, d)?))
}

fn leaf_languages(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:?}",
        Package::from_bytes(bytes)?
            .sheet_drawable_text_languages(0, DrawableSelector::text_box(0))?
    ))
}

fn host_language_change(source: &[u8], change: &LanguageChange) -> TestResult<Vec<u8>> {
    let mut e = numbers::Draft::from_bytes(source)?;
    let (s, d) = host_drawable(&e, Kind::TextBox, 0)?;
    match change {
        LanguageChange::Set(position, tag) => e.set_sheet_text_box_text_language(
            s,
            d,
            TextPosition::from_utf16_index(*position)?,
            TextLanguage::tag(*tag)?,
        )?,
        LanguageChange::Remove(position) => drop(e.remove_sheet_text_box_text_language_boundary(
            s,
            d,
            TextPosition::from_utf16_index(*position)?,
        )?),
        LanguageChange::Reset => drop(e.reset_sheet_text_box_text_languages(s, d)?),
    }
    Ok(e.to_bytes()?)
}

fn language_case(name: &'static str, input: &Fixture, change: LanguageChange) -> TestResult {
    let change = &change;
    Case::new(name, Format::Numbers, Tier::A, Evidence::Queued, input)
        .host(|source| host_language_change(source, change))
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit =
                package.edit_sheet_drawable_text_languages(0, DrawableSelector::text_box(0))?;
            match change {
                LanguageChange::Set(position, tag) => {
                    edit.set(
                        TextPosition::from_utf16_index(*position)?,
                        TextLanguage::tag(*tag)?,
                    )?;
                },
                LanguageChange::Remove(position) => {
                    edit.remove_boundary(TextPosition::from_utf16_index(*position)?)?;
                },
                LanguageChange::Reset => {
                    edit.reset()?;
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_drawable_text_languages(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host languages", host_languages)
        .read("leaf languages", leaf_languages)
        .run()?;
    Ok(())
}

/// The authored fixture with French from 6 set by the host.
fn with_boundary() -> TestResult<Fixture> {
    let input = authored_fixture()?;
    Ok(Fixture {
        path: "authored + host French boundary at 6",
        bytes: host_language_change(&input.bytes, &LanguageChange::Set(6, "fr"))?,
    })
}

#[test]
fn numbers_sheet_text_box_language_set_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    language_case(
        "numbers_sheet_text_box_language_set_parity",
        &authored_fixture()?,
        LanguageChange::Set(6, "fr"),
    )?;
    language_case(
        "numbers_sheet_text_box_language_set_parity",
        &with_boundary()?,
        LanguageChange::Set(0, "de"),
    )
}

#[test]
fn numbers_sheet_text_box_language_remove_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    language_case(
        "numbers_sheet_text_box_language_remove_parity",
        &with_boundary()?,
        LanguageChange::Remove(6),
    )
}

#[test]
fn numbers_sheet_text_box_language_reset_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    language_case(
        "numbers_sheet_text_box_language_reset_parity",
        &with_boundary()?,
        LanguageChange::Reset,
    )
}

#[test]
fn numbers_sheet_text_box_language_read_parity() -> TestResult {
    for input in [authored_fixture()?, with_boundary()?] {
        assert_eq!(host_languages(&input.bytes)?, leaf_languages(&input.bytes)?);
        let e = numbers::Draft::from_bytes(&input.bytes)?;
        let (s, d) = host_drawable(&e, Kind::TextBox, 0)?;
        for position in [0, 3, 6, 8] {
            let position = TextPosition::from_utf16_index(position)?;
            assert_eq!(
                e.sheet_text_box_text_language(s, d, position)?,
                Package::from_bytes(&input.bytes)?.sheet_drawable_text_language(
                    0,
                    DrawableSelector::text_box(0),
                    position
                )?
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

// ---- Drop Caps -------------------------------------------------------------

use litchi_numbers::sheet::drawable::drop_cap::{
    CharacterCount, DropCap, LineCount, Outdent, Padding, RaisedLines, Wrap,
};

/// `chart-grid-authored-native.numbers` (its theme has a Drop Cap base) with
/// one host text box of two paragraphs.
fn drop_cap_fixture() -> TestResult<Fixture> {
    let mut editor = numbers::Draft::from_bytes(
        &crate::harness::fixture("iwork/numbers/chart-grid-authored-native.numbers")?.bytes,
    )?;
    let sheet = editor.sheets()?.first().ok_or("no sheet")?.id();
    editor.add_sheet_text_box(
        sheet,
        "First paragraph\nSecond paragraph",
        litchi_iwa::shapes::DrawablePoint { x: 300.0, y: 200.0 },
        litchi_iwa::shapes::DrawableSize {
            width: 220.0,
            height: 90.0,
        },
    )?;
    Ok(Fixture {
        path: "chart-grid-authored-native.numbers + add_sheet_text_box",
        bytes: editor.to_bytes()?,
    })
}

fn drop_cap_samples() -> TestResult<[DropCap; 2]> {
    Ok([
        DropCap::new(LineCount::new(3)?, CharacterCount::new(1)?),
        DropCap::new(LineCount::new(4)?, CharacterCount::new(2)?)
            .with_raised_lines(RaisedLines::new(1)?)
            .with_wrap(Wrap::Contour)
            .with_padding(Padding::from_points(6.0)?)
            .with_outdent(Outdent::from_ratio(0.25)?),
    ])
}

fn host_drop_caps(bytes: &[u8]) -> TestResult<String> {
    let e = numbers::Draft::from_bytes(bytes)?;
    let (s, d) = host_drawable(&e, Kind::TextBox, 0)?;
    Ok(format!(
        "{:?}",
        e.sheet_text_box_paragraph_drop_caps(s, d)?
            .into_iter()
            .map(|placement| (placement.paragraph, placement.drop_cap))
            .collect::<Vec<_>>()
    ))
}

fn leaf_drop_caps(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:?}",
        Package::from_bytes(bytes)?.sheet_drawable_drop_caps(0, DrawableSelector::text_box(0))?
    ))
}

fn host_drop_cap(
    source: &[u8],
    paragraph: TextPosition,
    value: Option<DropCap>,
) -> TestResult<Vec<u8>> {
    let mut e = numbers::Draft::from_bytes(source)?;
    let (s, d) = host_drawable(&e, Kind::TextBox, 0)?;
    match value {
        Some(value) => e.set_sheet_text_box_paragraph_drop_cap(s, d, paragraph, value)?,
        None => drop(e.remove_sheet_text_box_paragraph_drop_cap(s, d, paragraph)?),
    }
    Ok(e.to_bytes()?)
}

fn drop_cap_case(
    name: &'static str,
    tier: Tier,
    input: &Fixture,
    paragraph: TextPosition,
    value: Option<DropCap>,
) -> TestResult {
    let mut case = Case::new(name, Format::Numbers, tier, Evidence::Queued, input)
        .host(|source| host_drop_cap(source, paragraph, value))
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit =
                package.edit_sheet_drawable_drop_caps(0, DrawableSelector::text_box(0))?;
            match value {
                Some(value) => {
                    edit.set(paragraph, value)?;
                },
                None => {
                    edit.remove(paragraph)?;
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_drawable_drop_caps(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host drop caps", host_drop_caps)
        .read("leaf drop caps", leaf_drop_caps)
        .known_fields(|known| {
            known.with(
                10_024,
                "TSWP.DropCapStyleArchive",
                decode_as::<tswp::DropCapStyleArchive>,
            )
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

fn second_paragraph() -> TestResult<TextPosition> {
    Ok(TextPosition::from_utf16_index(
        "First paragraph\n".encode_utf16().count(),
    )?)
}

#[test]
fn numbers_sheet_text_box_drop_cap_set_parity() -> TestResult {
    let input = drop_cap_fixture()?;
    let [first, second] = drop_cap_samples()?;
    for paragraph in [TextPosition::ZERO, second_paragraph()?] {
        drop_cap_case(
            "numbers_sheet_text_box_drop_cap_set_parity",
            Tier::B,
            &input,
            paragraph,
            Some(first),
        )?;
    }
    // A second value on the paragraph's own variation rewrites it in place.
    let set = Fixture {
        path: input.path,
        bytes: host_drop_cap(&input.bytes, TextPosition::ZERO, Some(first))?,
    };
    drop_cap_case(
        "numbers_sheet_text_box_drop_cap_set_parity",
        Tier::B,
        &set,
        TextPosition::ZERO,
        Some(second),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn numbers_sheet_text_box_drop_cap_remove_parity() -> TestResult {
    let input = drop_cap_fixture()?;
    let [first, _] = drop_cap_samples()?;
    let set = Fixture {
        path: input.path,
        bytes: host_drop_cap(&input.bytes, TextPosition::ZERO, Some(first))?,
    };
    drop_cap_case(
        "numbers_sheet_text_box_drop_cap_remove_parity",
        Tier::B,
        &set,
        TextPosition::ZERO,
        None,
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn numbers_sheet_text_box_drop_cap_read_parity() -> TestResult {
    let input = drop_cap_fixture()?;
    let [first, _] = drop_cap_samples()?;
    let set = host_drop_cap(&input.bytes, second_paragraph()?, Some(first))?;
    for bytes in [input.bytes.clone(), set] {
        assert_eq!(host_drop_caps(&bytes)?, leaf_drop_caps(&bytes)?);
        let e = numbers::Draft::from_bytes(&bytes)?;
        let (s, d) = host_drawable(&e, Kind::TextBox, 0)?;
        let leaf = Package::from_bytes(&bytes)?
            .sheet_drawable_drop_caps(0, DrawableSelector::text_box(0))?;
        for paragraph in [TextPosition::ZERO, second_paragraph()?] {
            assert_eq!(
                e.sheet_text_box_paragraph_drop_cap(s, d, paragraph)?,
                leaf.iter()
                    .find(|(at, _)| *at == paragraph)
                    .map(|(_, cap)| *cap)
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

// ---- hyperlinks ------------------------------------------------------------

use litchi_numbers::sheet::drawable::hyperlink::{TextHyperlinkTarget, TextRange};

#[derive(Clone, Debug)]
enum LinkChange {
    Add(TextRange, &'static str),
    Update(usize, TextRange, &'static str),
    Remove(usize),
}

fn range(start: usize, end: usize) -> TestResult<TextRange> {
    Ok(TextRange::new(
        TextPosition::from_utf16_index(start)?,
        TextPosition::from_utf16_index(end)?,
    )?)
}

fn host_links(bytes: &[u8]) -> TestResult<String> {
    let e = numbers::Draft::from_bytes(bytes)?;
    let (s, d) = host_drawable(&e, Kind::TextBox, 0)?;
    Ok(format!(
        "{:?}",
        e.sheet_text_box_hyperlinks(s, d)?
            .into_iter()
            .map(|link| (link.range, link.target))
            .collect::<Vec<_>>()
    ))
}

fn leaf_links(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:?}",
        Package::from_bytes(bytes)?
            .sheet_drawable_hyperlinks(0, DrawableSelector::text_box(0))?
            .into_iter()
            .map(|link| (link.range(), link.target().clone()))
            .collect::<Vec<_>>()
    ))
}

fn host_link(source: &[u8], change: &LinkChange) -> TestResult<Vec<u8>> {
    let mut e = numbers::Draft::from_bytes(source)?;
    let (s, d) = host_drawable(&e, Kind::TextBox, 0)?;
    let links = e.sheet_text_box_hyperlinks(s, d)?;
    match change {
        LinkChange::Add(range, target) => {
            drop(e.add_sheet_text_box_hyperlink(
                s,
                d,
                *range,
                TextHyperlinkTarget::new(target)?,
            )?);
        },
        LinkChange::Update(position, range, target) => drop(e.update_sheet_text_box_hyperlink(
            s,
            d,
            links.get(*position).ok_or("host link")?.id,
            *range,
            TextHyperlinkTarget::new(target)?,
        )?),
        LinkChange::Remove(position) => drop(e.remove_sheet_text_box_hyperlink(
            s,
            d,
            links.get(*position).ok_or("host link")?.id,
        )?),
    }
    Ok(e.to_bytes()?)
}

fn link_case(name: &'static str, tier: Tier, input: &Fixture, change: LinkChange) -> TestResult {
    let change = &change;
    let mut case = Case::new(name, Format::Numbers, tier, Evidence::Queued, input)
        .host(|source| host_link(source, change))
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit =
                package.edit_sheet_drawable_hyperlinks(0, DrawableSelector::text_box(0))?;
            match change {
                LinkChange::Add(range, target) => {
                    edit.add(*range, TextHyperlinkTarget::new(target)?)?;
                },
                LinkChange::Update(position, range, target) => {
                    edit.update(*position, *range, TextHyperlinkTarget::new(target)?)?;
                },
                LinkChange::Remove(position) => {
                    edit.remove(*position)?;
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_drawable_hyperlinks(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host hyperlinks", host_links)
        .read("leaf hyperlinks", leaf_links);
    if tier == Tier::B {
        case = case
            .known_fields(crate::text_annotation_cases::masked)
            .leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

fn with_link() -> TestResult<Fixture> {
    let input = authored_fixture()?;
    Ok(Fixture {
        path: "authored + host hyperlink over 0..5",
        bytes: host_link(
            &input.bytes,
            &LinkChange::Add(range(0, 5)?, "https://example.com/litchi"),
        )?,
    })
}

#[test]
fn numbers_sheet_text_box_hyperlink_add_parity() -> TestResult {
    link_case(
        "numbers_sheet_text_box_hyperlink_add_parity",
        Tier::B,
        &authored_fixture()?,
        LinkChange::Add(range(0, 5)?, "https://example.com/litchi"),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn numbers_sheet_text_box_hyperlink_update_parity() -> TestResult {
    link_case(
        "numbers_sheet_text_box_hyperlink_update_parity",
        Tier::A,
        &with_link()?,
        LinkChange::Update(0, range(2, 8)?, "mailto:a@example.com"),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn numbers_sheet_text_box_hyperlink_remove_parity() -> TestResult {
    link_case(
        "numbers_sheet_text_box_hyperlink_remove_parity",
        Tier::A,
        &with_link()?,
        LinkChange::Remove(0),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn numbers_sheet_text_box_hyperlink_read_parity() -> TestResult {
    for input in [authored_fixture()?, with_link()?] {
        assert_eq!(host_links(&input.bytes)?, leaf_links(&input.bytes)?);
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

// ---- Date & Time fields ------------------------------------------------------

use litchi_numbers::sheet::drawable::date_time::DisplayText;

#[derive(Clone, Debug)]
enum FieldChange {
    Add(TextRange, usize),
    Insert(TextPosition, &'static str, usize),
    Update(usize, TextRange, usize),
    Remove(usize),
}

fn host_dates(bytes: &[u8]) -> TestResult<String> {
    let e = numbers::Draft::from_bytes(bytes)?;
    let (s, d) = host_drawable(&e, Kind::TextBox, 0)?;
    Ok(format!(
        "{:?}",
        e.sheet_text_box_date_time_fields(s, d)?
            .into_iter()
            .map(|field| (field.range, field.settings))
            .collect::<Vec<_>>()
    ))
}

fn leaf_dates(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:?}",
        Package::from_bytes(bytes)?
            .sheet_drawable_date_time_fields(0, DrawableSelector::text_box(0))?
            .into_iter()
            .map(|field| (field.range(), field.settings().clone()))
            .collect::<Vec<_>>()
    ))
}

fn host_date(source: &[u8], change: &FieldChange) -> TestResult<Vec<u8>> {
    let samples = crate::text_field_cases::settings()?;
    let mut e = numbers::Draft::from_bytes(source)?;
    let (s, d) = host_drawable(&e, Kind::TextBox, 0)?;
    let fields = e.sheet_text_box_date_time_fields(s, d)?;
    match change {
        FieldChange::Add(range, value) => {
            drop(e.add_sheet_text_box_date_time_field(s, d, *range, samples[*value].clone())?)
        },
        FieldChange::Insert(position, display, value) => {
            drop(e.insert_sheet_text_box_date_time_field(
                s,
                d,
                *position,
                DisplayText::new(display)?,
                samples[*value].clone(),
            )?);
        },
        FieldChange::Update(position, range, value) => {
            drop(e.update_sheet_text_box_date_time_field(
                s,
                d,
                fields.get(*position).ok_or("host field")?.id,
                *range,
                samples[*value].clone(),
            )?);
        },
        FieldChange::Remove(position) => drop(e.remove_sheet_text_box_date_time_field(
            s,
            d,
            fields.get(*position).ok_or("host field")?.id,
        )?),
    }
    Ok(e.to_bytes()?)
}

fn date_case(name: &'static str, tier: Tier, input: &Fixture, change: FieldChange) -> TestResult {
    let change = &change;
    let mut case = Case::new(name, Format::Numbers, tier, Evidence::Queued, input)
        .host(|source| host_date(source, change))
        .leaf(|source| {
            let samples = crate::text_field_cases::settings()?;
            let package = Package::from_bytes(source)?;
            let mut edit =
                package.edit_sheet_drawable_date_time_fields(0, DrawableSelector::text_box(0))?;
            match change {
                FieldChange::Add(range, value) => {
                    edit.add(*range, samples[*value].clone())?;
                },
                FieldChange::Insert(position, display, value) => {
                    edit.insert(
                        *position,
                        DisplayText::new(display)?,
                        samples[*value].clone(),
                    )?;
                },
                FieldChange::Update(position, range, value) => {
                    edit.update(*position, *range, samples[*value].clone())?;
                },
                FieldChange::Remove(position) => {
                    edit.remove(*position)?;
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_drawable_date_time_fields(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host fields", host_dates)
        .read("leaf fields", leaf_dates);
    if tier == Tier::B {
        case = case
            .known_fields(crate::text_field_cases::masked)
            .leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

fn with_date() -> TestResult<Fixture> {
    let input = authored_fixture()?;
    Ok(Fixture {
        path: "authored + host date-time field over 0..5",
        bytes: host_date(&input.bytes, &FieldChange::Add(range(0, 5)?, 0))?,
    })
}

#[test]
fn numbers_sheet_text_box_date_time_add_parity() -> TestResult {
    for value in [0, 1] {
        date_case(
            "numbers_sheet_text_box_date_time_add_parity",
            Tier::B,
            &authored_fixture()?,
            FieldChange::Add(range(0, 5)?, value),
        )?;
    }
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn numbers_sheet_text_box_date_time_insert_parity() -> TestResult {
    for position in [0, 9] {
        date_case(
            "numbers_sheet_text_box_date_time_insert_parity",
            Tier::B,
            &authored_fixture()?,
            FieldChange::Insert(TextPosition::from_utf16_index(position)?, "2026-10-26", 0),
        )?;
    }
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn numbers_sheet_text_box_date_time_update_parity() -> TestResult {
    date_case(
        "numbers_sheet_text_box_date_time_update_parity",
        Tier::A,
        &with_date()?,
        FieldChange::Update(0, range(2, 8)?, 1),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn numbers_sheet_text_box_date_time_remove_parity() -> TestResult {
    date_case(
        "numbers_sheet_text_box_date_time_remove_parity",
        Tier::A,
        &with_date()?,
        FieldChange::Remove(0),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn numbers_sheet_text_box_date_time_read_parity() -> TestResult {
    for input in [authored_fixture()?, with_date()?] {
        assert_eq!(host_dates(&input.bytes)?, leaf_dates(&input.bytes)?);
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
