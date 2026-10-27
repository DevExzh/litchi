//! Storage-scoped paragraph-style property parity (wave W1-D, record 0374):
//! the host `text::Draft` `set_text_*`/`reset_text_*`/`set_paragraph_*`/
//! `reset_paragraph_*` families against
//! `litchi_iwa_text_wire::paragraph_style` published by [`text_glue`].
//!
//! [`text_glue`]: crate::text_glue

use litchi_iwa::text::{Draft, TextStorageId};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_common::shape::stroke::{Pattern, Width};
use litchi_iwa_protos::tswp;
use litchi_iwa_text::appearance::{Background, Outline, ParagraphBackground, Shadow};
use litchi_iwa_text::character::{
    TextBaselineShift, TextCapitalization, TextCharacterSpacing, TextDecorations, TextLigatures,
    TextPointSize, TextScript, TextStrikethrough, TextStyle, TextUnderline,
};
use litchi_iwa_text::font::Font;
use litchi_iwa_text::paragraph::border::{Offset as BorderOffset, Sides};
use litchi_iwa_text::paragraph::direction::WritingDirection;
use litchi_iwa_text::paragraph::flow::{Flow, Hyphenation};
use litchi_iwa_text::paragraph::format::{
    Alignment, Border, Borders, IndentPoints, Indents, LineSpacing, LineSpacingPoints, Spacing,
    SpacingPoints,
};
use litchi_iwa_text::paragraph::style::ParagraphFollowingStyle;
use litchi_iwa_text::paragraph::tabs::{
    Alignment as TabAlignment, DecimalCharacter, DefaultInterval, Leader, Position, Stop, Stops,
};
use litchi_iwa_text_wire::paragraph_style::{
    self, Change, MESSAGE_VERSION, Plan, Property, PropertyKind, Target,
};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::decode_as;
use crate::text_glue::{
    STYLESHEET_TYPE, Workspace, exclusive, managed_field_info, new_object, paragraph_styles,
    sync_field_info,
};

const LIMITS: WireLimits = WireLimits::DEFAULT;

/// Fixtures with uniformly styled text storages, one per format.
pub const FIXTURES: [(&str, Format); 4] = [
    ("iwork/keynote/basic.key", Format::Keynote),
    ("iwork/keynote/multi-slide.key", Format::Keynote),
    ("iwork/pages/basic.pages", Format::Pages),
    ("iwork/numbers/basic.numbers", Format::Numbers),
];

const MANAGED_PATHS: [[u32; 2]; 4] = [[11, 46], [11, 44], [11, 21], [11, 26]];

fn managed(variation: &paragraph_style::Variation) -> [bool; 4] {
    let fields = variation.managed_fields();
    [
        fields.text_fill(),
        fields.text_stroke(),
        fields.text_shadow(),
        fields.text_background(),
    ]
}

fn engine(error: paragraph_style::Error) -> litchi_iwa::Error {
    litchi_iwa::Error::InvalidFormat(error.to_string())
}

/// The leaf route: plan with the engine, publish through the glue.
pub fn leaf_change(source: &[u8], storage: u64, change: &Change) -> TestResult<Vec<u8>> {
    let mut workspace = Workspace::open(source)?;
    apply_change(&mut workspace, storage, change)?;
    workspace.into_bytes()
}

/// Plan `change` with the engine and publish it into `workspace`.
pub fn apply_change(workspace: &mut Workspace, storage: u64, change: &Change) -> TestResult {
    let located = workspace.storage(storage)?;
    let style = paragraph_style::uniform_style(&located.payload, LIMITS)?;
    let (styles, entries) = paragraph_styles(workspace)?;
    let target = Target::new(style, exclusive(workspace, style)?);
    let plan = paragraph_style::plan(&styles, target, change, LIMITS)?;
    let style_entry = entries.get(&style).ok_or("style entry")?.clone();
    match plan {
        Plan::Unchanged => {},
        Plan::Rewrite(variation) => {
            let message = workspace.message(style, paragraph_style::MESSAGE_TYPE)?;
            let flags = managed(&variation);
            workspace.update(&style_entry, |archive| {
                let object = archive
                    .object_mut(style)
                    .ok_or_else(|| engine(paragraph_style::Error::Internal))?;
                object.replace_message(
                    message.index,
                    litchi_iwa_archive::iwa::RawMessage {
                        type_: paragraph_style::MESSAGE_TYPE,
                        data: variation.into_payload(),
                    },
                    Default::default(),
                )?;
                let info = &mut object.archive_info.message_infos.as_mut_slice()[message.index];
                for (present, path) in flags.into_iter().zip(MANAGED_PATHS) {
                    sync_field_info(&mut info.field_infos, present, path);
                }
                Ok(())
            })?;
        },
        Plan::Derive(variation) => {
            let created = workspace.next_identifier()?;
            let stylesheet = variation.stylesheet();
            let sheet = workspace.message(stylesheet, STYLESHEET_TYPE)?;
            if sheet.entry != style_entry {
                return Err("stylesheet is not stored with the style".into());
            }
            let repointed = paragraph_style::repoint(&located.payload, style, created, LIMITS)?;
            workspace.replace(&located, repointed, Some((style, created)))?;
            let field_infos = managed(&variation)
                .into_iter()
                .zip(MANAGED_PATHS)
                .filter(|(present, _)| *present)
                .map(|(_, path)| managed_field_info(path))
                .collect();
            let object = new_object(
                created,
                paragraph_style::MESSAGE_TYPE,
                variation.payload().to_vec(),
                &MESSAGE_VERSION,
                variation.references(),
                field_infos,
            )?;
            workspace.update(&style_entry, |archive| {
                let sheet_object = archive
                    .object_mut(stylesheet)
                    .ok_or_else(|| engine(paragraph_style::Error::Internal))?;
                let data = paragraph_style::stylesheet_with_variation(
                    &sheet_object.messages[sheet.index].data,
                    style,
                    created,
                    LIMITS,
                )
                .map_err(engine)?;
                sheet_object.replace_message(
                    sheet.index,
                    litchi_iwa_archive::iwa::RawMessage {
                        type_: STYLESHEET_TYPE,
                        data,
                    },
                    Default::default(),
                )?;
                sheet_object.archive_info.message_infos.as_mut_slice()[sheet.index]
                    .object_references
                    .push(created);
                archive.insert_object(object, Default::default())?;
                Ok(())
            })?;
            if let Some(style_component) = workspace.component(&style_entry)? {
                workspace.add_object_uuid(style_component, created)?;
                if let Some(owner) = workspace.component(&located.entry)?
                    && owner != style_component
                {
                    workspace.add_external_reference(owner, style_component, created)?;
                }
            }
            workspace.raise_watermark(created)?;
        },
        Plan::Collapse { parent, stylesheet } => {
            let sheet = workspace.message(stylesheet, STYLESHEET_TYPE)?;
            let repointed = paragraph_style::repoint(&located.payload, style, parent, LIMITS)?;
            workspace.replace(&located, repointed, Some((style, parent)))?;
            workspace.update(&style_entry, |archive| {
                let sheet_object = archive
                    .object_mut(stylesheet)
                    .ok_or_else(|| engine(paragraph_style::Error::Internal))?;
                let data = paragraph_style::stylesheet_without_variation(
                    &sheet_object.messages[sheet.index].data,
                    parent,
                    style,
                    LIMITS,
                )
                .map_err(engine)?;
                sheet_object.replace_message(
                    sheet.index,
                    litchi_iwa_archive::iwa::RawMessage {
                        type_: STYLESHEET_TYPE,
                        data,
                    },
                    Default::default(),
                )?;
                let info = &mut sheet_object.archive_info.message_infos.as_mut_slice()[sheet.index];
                info.object_references
                    .retain(|&reference| reference != style);
                for field in &mut info.field_infos {
                    field
                        .object_references
                        .retain(|&reference| reference != style);
                }
                archive
                    .remove_object(style)
                    .ok_or_else(|| engine(paragraph_style::Error::Internal))?;
                Ok(())
            })?;
            if let Some(style_component) = workspace.component(&style_entry)? {
                workspace.remove_object_uuid(style_component, style)?;
                workspace.remove_external_references_to(style_component, style)?;
                if let Some(owner) = workspace.component(&located.entry)?
                    && owner != style_component
                {
                    workspace.add_external_reference(owner, style_component, parent)?;
                }
            }
            workspace.release_watermark(&[style])?;
        },
        other => return Err(format!("unmapped plan {other:?}").into()),
    }
    Ok(())
}

/// The host route for one change.
#[allow(clippy::too_many_lines, reason = "one arm per host method")]
pub fn host_change(source: &[u8], storage: u64, change: &Change) -> TestResult<Vec<u8>> {
    let mut draft = Draft::from_bytes(source)?;
    let id = TextStorageId::try_from(storage)?;
    match change.clone() {
        Change::Set(property) => match property {
            Property::TextStyle(v) => draft.set_text_style(id, v)?,
            Property::TextFont(v) => draft.set_text_font(id, v)?,
            Property::TextDecorations(v) => draft.set_text_decorations(id, v)?,
            Property::TextColor(v) => draft.set_text_color(id, v)?,
            Property::TextCapitalization(v) => draft.set_text_capitalization(id, v)?,
            Property::TextScript(v) => draft.set_text_script(id, v)?,
            Property::TextBaselineShift(v) => draft.set_text_baseline_shift(id, v)?,
            Property::TextCharacterSpacing(v) => draft.set_text_character_spacing(id, v)?,
            Property::TextLigatures(v) => draft.set_text_ligatures(id, v)?,
            Property::TextOutline(v) => draft.set_text_outline(id, v)?,
            Property::TextShadow(v) => draft.set_text_shadow(id, v)?,
            Property::TextBackground(v) => draft.set_text_background(id, v)?,
            Property::ParagraphBackground(v) => draft.set_paragraph_background(id, v)?,
            Property::Borders(v) => draft.set_paragraph_borders(id, v)?,
            Property::Flow(v) => draft.set_paragraph_flow(id, v)?,
            Property::WritingDirection(v) => draft.set_paragraph_writing_direction(id, v)?,
            Property::FollowingStyle(v) => draft.set_paragraph_following_style(id, v)?,
            Property::Alignment(v) => draft.set_paragraph_alignment(id, v)?,
            Property::LineSpacing(v) => draft.set_paragraph_line_spacing(id, v)?,
            Property::Spacing(v) => draft.set_paragraph_spacing(id, v)?,
            Property::Indents(v) => draft.set_paragraph_indents(id, v)?,
            Property::DecimalTabCharacter(v) => draft.set_paragraph_decimal_tab_character(id, v)?,
            Property::DefaultTabInterval(v) => draft.set_paragraph_default_tab_interval(id, v)?,
            Property::TabStops(v) => draft.set_paragraph_tab_stops(id, v)?,
            other => return Err(format!("unmapped property {other:?}").into()),
        },
        Change::Reset(kind) => {
            match kind {
                PropertyKind::TextStyle => draft.reset_text_style(id)?,
                PropertyKind::TextFont => draft.reset_text_font(id)?,
                PropertyKind::TextDecorations => draft.reset_text_decorations(id)?,
                PropertyKind::TextColor => draft.reset_text_color(id)?,
                PropertyKind::TextCapitalization => draft.reset_text_capitalization(id)?,
                PropertyKind::TextScript => draft.reset_text_script(id)?,
                PropertyKind::TextBaselineShift => draft.reset_text_baseline_shift(id)?,
                PropertyKind::TextCharacterSpacing => draft.reset_text_character_spacing(id)?,
                PropertyKind::TextLigatures => draft.reset_text_ligatures(id)?,
                PropertyKind::TextOutline => draft.reset_text_outline(id)?,
                PropertyKind::TextShadow => draft.reset_text_shadow(id)?,
                PropertyKind::TextBackground => draft.reset_text_background(id)?,
                PropertyKind::ParagraphBackground => draft.reset_paragraph_background(id)?,
                PropertyKind::Borders => draft.reset_paragraph_borders(id)?,
                PropertyKind::Flow => draft.reset_paragraph_flow(id)?,
                PropertyKind::WritingDirection => draft.reset_paragraph_writing_direction(id)?,
                PropertyKind::FollowingStyle => draft.reset_paragraph_following_style(id)?,
                PropertyKind::Alignment => draft.reset_paragraph_alignment(id)?,
                PropertyKind::LineSpacing => draft.reset_paragraph_line_spacing(id)?,
                PropertyKind::Spacing => draft.reset_paragraph_spacing(id)?,
                PropertyKind::Indents => draft.reset_paragraph_indents(id)?,
                PropertyKind::DecimalTabCharacter => {
                    draft.reset_paragraph_decimal_tab_character(id)?
                },
                PropertyKind::DefaultTabInterval => {
                    draft.reset_paragraph_default_tab_interval(id)?
                },
                PropertyKind::TabStops => draft.reset_paragraph_tab_stops(id)?,
                other => return Err(format!("unmapped kind {other:?}").into()),
            };
        },
        other => return Err(format!("unmapped change {other:?}").into()),
    }
    Ok(draft.to_bytes()?)
}

/// The host read of one property.
#[allow(clippy::too_many_lines, reason = "one arm per host method")]
pub fn host_read(bytes: &[u8], storage: u64, kind: PropertyKind) -> TestResult<Property> {
    let draft = Draft::from_bytes(bytes)?;
    let id = TextStorageId::try_from(storage)?;
    Ok(match kind {
        PropertyKind::TextStyle => Property::TextStyle(draft.text_style(id)?),
        PropertyKind::TextFont => Property::TextFont(draft.text_font(id)?),
        PropertyKind::TextDecorations => Property::TextDecorations(draft.text_decorations(id)?),
        PropertyKind::TextColor => Property::TextColor(draft.text_color(id)?),
        PropertyKind::TextCapitalization => {
            Property::TextCapitalization(draft.text_capitalization(id)?)
        },
        PropertyKind::TextScript => Property::TextScript(draft.text_script(id)?),
        PropertyKind::TextBaselineShift => {
            Property::TextBaselineShift(draft.text_baseline_shift(id)?)
        },
        PropertyKind::TextCharacterSpacing => {
            Property::TextCharacterSpacing(draft.text_character_spacing(id)?)
        },
        PropertyKind::TextLigatures => Property::TextLigatures(draft.text_ligatures(id)?),
        PropertyKind::TextOutline => Property::TextOutline(draft.text_outline(id)?),
        PropertyKind::TextShadow => Property::TextShadow(draft.text_shadow(id)?),
        PropertyKind::TextBackground => Property::TextBackground(draft.text_background(id)?),
        PropertyKind::ParagraphBackground => {
            Property::ParagraphBackground(draft.paragraph_background(id)?)
        },
        PropertyKind::Borders => Property::Borders(draft.paragraph_borders(id)?),
        PropertyKind::Flow => Property::Flow(draft.paragraph_flow(id)?),
        PropertyKind::WritingDirection => {
            Property::WritingDirection(draft.paragraph_writing_direction(id)?)
        },
        PropertyKind::FollowingStyle => {
            Property::FollowingStyle(draft.paragraph_following_style(id)?)
        },
        PropertyKind::Alignment => Property::Alignment(draft.paragraph_alignment(id)?),
        PropertyKind::LineSpacing => Property::LineSpacing(draft.paragraph_line_spacing(id)?),
        PropertyKind::Spacing => Property::Spacing(draft.paragraph_spacing(id)?),
        PropertyKind::Indents => Property::Indents(draft.paragraph_indents(id)?),
        PropertyKind::DecimalTabCharacter => {
            Property::DecimalTabCharacter(draft.paragraph_decimal_tab_character(id)?)
        },
        PropertyKind::DefaultTabInterval => {
            Property::DefaultTabInterval(draft.paragraph_default_tab_interval(id)?)
        },
        PropertyKind::TabStops => Property::TabStops(draft.paragraph_tab_stops(id)?),
        other => return Err(format!("unmapped kind {other:?}").into()),
    })
}

/// The leaf read of one property.
pub fn leaf_read(bytes: &[u8], storage: u64, kind: PropertyKind) -> TestResult<Property> {
    let workspace = Workspace::open(bytes)?;
    let located = workspace.storage(storage)?;
    let style = paragraph_style::uniform_style(&located.payload, LIMITS)?;
    let (styles, _) = paragraph_styles(&workspace)?;
    Ok(paragraph_style::effective(&styles, style, kind, LIMITS)?)
}

/// A distinct value per property: different from the native defaults.
#[allow(clippy::too_many_lines, reason = "one value per property")]
pub fn sample(kind: PropertyKind) -> TestResult<Property> {
    let color = Rgba::new(0.2, 0.4, 0.6, 1.0, RgbColorSpace::Srgb)?;
    Ok(match kind {
        PropertyKind::TextStyle => Property::TextStyle(
            TextStyle::new(TextPointSize::from_points(27.0)?)
                .with_bold(true)
                .with_italic(true),
        ),
        PropertyKind::TextFont => Property::TextFont(Font::named("Courier")?),
        PropertyKind::TextDecorations => Property::TextDecorations(TextDecorations::new(
            TextUnderline::Double,
            TextStrikethrough::Single,
        )),
        PropertyKind::TextColor => Property::TextColor(color),
        PropertyKind::TextCapitalization => {
            Property::TextCapitalization(TextCapitalization::TitleCase)
        },
        PropertyKind::TextScript => Property::TextScript(TextScript::Superscript),
        PropertyKind::TextBaselineShift => {
            Property::TextBaselineShift(TextBaselineShift::from_points(3.0)?)
        },
        PropertyKind::TextCharacterSpacing => {
            Property::TextCharacterSpacing(TextCharacterSpacing::from_percent(25.0)?)
        },
        PropertyKind::TextLigatures => Property::TextLigatures(TextLigatures::All),
        PropertyKind::TextOutline => Property::TextOutline(Outline::standard()),
        PropertyKind::TextShadow => Property::TextShadow(Shadow::standard()),
        PropertyKind::TextBackground => Property::TextBackground(Background::Color(color)),
        PropertyKind::ParagraphBackground => {
            Property::ParagraphBackground(ParagraphBackground::Color(color))
        },
        PropertyKind::Borders => Property::Borders(Borders::Bordered(Border::new(
            color,
            Width::new(2.0)?,
            Pattern::ShortDash,
            Sides::TOP | Sides::BOTTOM,
            BorderOffset::from_points(9.0)?,
            false,
        )?)),
        PropertyKind::Flow => Property::Flow(
            Flow::new()
                .with_hyphenation(Hyphenation::Automatic)
                .with_keep_with_next(true)
                .with_start_on_new_page(true),
        ),
        PropertyKind::WritingDirection => Property::WritingDirection(WritingDirection::RightToLeft),
        PropertyKind::FollowingStyle => Property::FollowingStyle(ParagraphFollowingStyle::Same),
        PropertyKind::Alignment => Property::Alignment(Alignment::Justified),
        PropertyKind::LineSpacing => {
            Property::LineSpacing(LineSpacing::Exactly(LineSpacingPoints::from_points(18.0)?))
        },
        PropertyKind::Spacing => Property::Spacing(Spacing::new(
            SpacingPoints::from_points(6.0)?,
            SpacingPoints::from_points(12.0)?,
        )),
        PropertyKind::Indents => Property::Indents(Indents::new(
            IndentPoints::from_points(10.0)?,
            IndentPoints::from_points(20.0)?,
            IndentPoints::from_points(5.0)?,
        )),
        PropertyKind::DecimalTabCharacter => {
            Property::DecimalTabCharacter(DecimalCharacter::new(',')?)
        },
        PropertyKind::DefaultTabInterval => {
            Property::DefaultTabInterval(DefaultInterval::from_points(48.0)?)
        },
        PropertyKind::TabStops => Property::TabStops(Stops::new(vec![
            Stop::new(Position::from_points(36.0)?, TabAlignment::Left),
            Stop::new(Position::from_points(144.0)?, TabAlignment::Decimal)
                .with_leader(Leader::new(".")?),
        ])?),
        other => return Err(format!("no sample for {other:?}").into()),
    })
}

/// [`sample`], with the following style naming the last named paragraph
/// style of the storage's stylesheet (the default `Same` is a no-op).
pub fn sample_for(bytes: &[u8], storage: u64, kind: PropertyKind) -> TestResult<Property> {
    if kind != PropertyKind::FollowingStyle {
        return sample(kind);
    }
    let draft = Draft::from_bytes(bytes)?;
    let named = draft.named_paragraph_styles(TextStorageId::try_from(storage)?)?;
    let last = named
        .last()
        .ok_or("the stylesheet has no named paragraph style")?;
    Ok(Property::FollowingStyle(ParagraphFollowingStyle::Named(
        last.id(),
    )))
}

/// A second distinct value per property (for in-place rewrites).
pub fn second_sample(kind: PropertyKind) -> TestResult<Property> {
    let color = Rgba::new(0.9, 0.1, 0.3, 1.0, RgbColorSpace::DisplayP3)?;
    Ok(match kind {
        PropertyKind::TextStyle => {
            Property::TextStyle(TextStyle::new(TextPointSize::from_points(9.0)?))
        },
        PropertyKind::TextFont => Property::TextFont(Font::named("Avenir")?),
        PropertyKind::TextDecorations => Property::TextDecorations(TextDecorations::new(
            TextUnderline::Wavy,
            TextStrikethrough::None,
        )),
        PropertyKind::TextColor => Property::TextColor(color),
        PropertyKind::TextCapitalization => {
            Property::TextCapitalization(TextCapitalization::SmallCaps)
        },
        PropertyKind::TextScript => Property::TextScript(TextScript::Subscript),
        PropertyKind::TextBaselineShift => {
            Property::TextBaselineShift(TextBaselineShift::from_points(-2.0)?)
        },
        PropertyKind::TextCharacterSpacing => {
            Property::TextCharacterSpacing(TextCharacterSpacing::from_percent(-10.0)?)
        },
        PropertyKind::TextLigatures => Property::TextLigatures(TextLigatures::RequiredOnly),
        PropertyKind::TextOutline => Property::TextOutline(Outline::None),
        PropertyKind::TextShadow => Property::TextShadow(Shadow::None),
        PropertyKind::TextBackground => Property::TextBackground(Background::None),
        PropertyKind::ParagraphBackground => {
            Property::ParagraphBackground(ParagraphBackground::None)
        },
        PropertyKind::Borders => Property::Borders(Borders::None),
        PropertyKind::Flow => Property::Flow(Flow::new().with_keep_lines_together(true)),
        PropertyKind::WritingDirection => Property::WritingDirection(WritingDirection::LeftToRight),
        PropertyKind::FollowingStyle => Property::FollowingStyle(ParagraphFollowingStyle::Same),
        PropertyKind::Alignment => Property::Alignment(Alignment::Center),
        PropertyKind::LineSpacing => {
            Property::LineSpacing(LineSpacing::AtLeast(LineSpacingPoints::from_points(30.0)?))
        },
        PropertyKind::Spacing => Property::Spacing(Spacing::new(
            SpacingPoints::from_points(0.0)?,
            SpacingPoints::from_points(3.0)?,
        )),
        PropertyKind::Indents => Property::Indents(Indents::new(
            IndentPoints::from_points(0.0)?,
            IndentPoints::from_points(40.0)?,
            IndentPoints::from_points(0.0)?,
        )),
        PropertyKind::DecimalTabCharacter => {
            Property::DecimalTabCharacter(DecimalCharacter::new('\u{066B}')?)
        },
        PropertyKind::DefaultTabInterval => {
            Property::DefaultTabInterval(DefaultInterval::from_points(72.0)?)
        },
        PropertyKind::TabStops => Property::TabStops(Stops::new(vec![Stop::new(
            Position::from_points(90.0)?,
            TabAlignment::Right,
        )])?),
        other => return Err(format!("no second sample for {other:?}").into()),
    })
}

/// The uniformly styled storages of a fixture (both routes must locate a
/// uniform paragraph style), at most `limit`.
pub fn uniform_storages(bytes: &[u8], limit: usize) -> TestResult<Vec<u64>> {
    let draft = Draft::from_bytes(bytes)?;
    let mut found = Vec::new();
    for storage in draft.storages()? {
        let id: u64 = storage.id.to_string().parse()?;
        if draft.paragraph_alignment(storage.id).is_ok()
            && leaf_read(bytes, id, PropertyKind::Alignment).is_ok()
        {
            found.push(id);
            if found.len() == limit {
                break;
            }
        }
    }
    Ok(found)
}

fn with_paragraph_styles(
    known: crate::known_fields::KnownFields,
) -> crate::known_fields::KnownFields {
    known.with(
        paragraph_style::MESSAGE_TYPE,
        "TSWP.ParagraphStyleArchive",
        decode_as::<tswp::ParagraphStyleArchive>,
    )
}

/// One case: `change` on `storage` of `input`, both routes, the effective
/// value read back through both.
pub fn run(
    name: &'static str,
    format: Format,
    tier: Tier,
    input: &Fixture,
    storage: u64,
    change: &Change,
) -> TestResult {
    let kind = match change {
        Change::Set(property) => property.kind(),
        Change::Reset(kind) => *kind,
        _ => return Err("unmapped change".into()),
    };
    let mut case = Case::new(name, format, tier, Evidence::Queued, input)
        .host(|source| host_change(source, storage, change))
        .leaf(|source| Ok(LeafRun::output(leaf_change(source, storage, change)?)))
        .read("host property", move |bytes| {
            host_read(bytes, storage, kind)
        })
        .read("leaf property", move |bytes| {
            leaf_read(bytes, storage, kind)
        })
        .known_fields(with_paragraph_styles);
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

/// The host output of setting `sample(kind)`: a fixture whose storage has
/// an exclusive managed variation.
pub fn derived(input: &Fixture, storage: u64, kind: PropertyKind) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_change(
            &input.bytes,
            storage,
            &Change::Set(sample_for(&input.bytes, storage, kind)?),
        )?,
    })
}

#[test]
fn text_paragraph_property_set_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for storage in uniform_storages(&input.bytes, 2)? {
            for kind in PropertyKind::ALL {
                run(
                    "text_paragraph_property_set_parity",
                    format,
                    Tier::B,
                    &input,
                    storage,
                    &Change::Set(sample_for(&input.bytes, storage, kind)?),
                )
                .map_err(|e| format!("{path} storage {storage} {kind:?}: {e}"))?;
                cases += 1;
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 4 * 24, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_paragraph_property_rewrite_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for storage in uniform_storages(&input.bytes, 1)? {
            for kind in PropertyKind::ALL {
                let derived = derived(&input, storage, kind)?;
                // The same property again (in place) and a second property
                // folded into the managed variation.
                for change in [
                    Change::Set(second_sample(kind)?),
                    Change::Set(sample(if kind == PropertyKind::Alignment {
                        PropertyKind::Indents
                    } else {
                        PropertyKind::Alignment
                    })?),
                ] {
                    run(
                        "text_paragraph_property_rewrite_parity",
                        format,
                        Tier::A,
                        &derived,
                        storage,
                        &change,
                    )
                    .map_err(|e| format!("{path} storage {storage} {kind:?} {change:?}: {e}"))?;
                    cases += 1;
                }
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 4 * 24, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_paragraph_property_reset_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for storage in uniform_storages(&input.bytes, 1)? {
            for kind in PropertyKind::ALL {
                let derived = derived(&input, storage, kind)?;
                run(
                    "text_paragraph_property_reset_parity",
                    format,
                    Tier::A,
                    &derived,
                    storage,
                    &Change::Reset(kind),
                )
                .map_err(|e| format!("{path} storage {storage} {kind:?}: {e}"))?;
                cases += 1;
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 4 * 24, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_paragraph_property_read_parity() -> TestResult {
    let mut reads = 0;
    for (path, _format) in FIXTURES {
        let input = fixture(path)?;
        for storage in uniform_storages(&input.bytes, 8)? {
            for kind in PropertyKind::ALL {
                for bytes in [input.bytes.clone(), derived(&input, storage, kind)?.bytes] {
                    let host = host_read(&bytes, storage, kind);
                    let leaf = leaf_read(&bytes, storage, kind);
                    match (host, leaf) {
                        (Ok(host), Ok(leaf)) => {
                            assert_eq!(host, leaf, "{path} storage {storage} {kind:?}");
                        },
                        (Err(_), Err(_)) => {},
                        (host, leaf) => {
                            return Err(format!(
                                "{path} storage {storage} {kind:?}: host {host:?}, leaf {leaf:?}"
                            )
                            .into());
                        },
                    }
                    reads += 1;
                }
            }
        }
    }
    assert!(reads >= 4 * 48, "only {reads} reads ran");
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
