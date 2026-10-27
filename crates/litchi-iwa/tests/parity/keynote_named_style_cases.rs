//! Keynote text-box named paragraph-style parity (wave W1-C, record 0372):
//! the host `*_slide_text_box_named_paragraph_style*` families against
//! `slide::drawable::Edit::{apply,create,rename,delete,delete_applied,
//! redefine}_named_style` and the named-style reads on the shared W1-D
//! engine.

use litchi_iwa::keynote;
use litchi_iwa_protos::tswp;
use litchi_iwa_text::paragraph::format::Alignment;
use litchi_iwa_text::paragraph::style::{ParagraphStyleId, ParagraphStyleName};
use litchi_iwa_text_wire::paragraph_style;
use litchi_keynote::slide::drawable::Kind;
use litchi_keynote::{DrawableSelector, Package, SlideSelector};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::keynote_drawable_cases::{host_id, leaf_position};
use crate::known_fields::decode_as;
use crate::leaf_bytes;

const SLIDE: usize = 0;
const TEXT_BOX: &str = "iwork/keynote/basic.key";

#[derive(Clone, Debug)]
enum Operation {
    Apply(ParagraphStyleId),
    Create(ParagraphStyleId, &'static str),
    Rename(ParagraphStyleId, &'static str),
    Delete(ParagraphStyleId),
    DeleteApplied(ParagraphStyleId, ParagraphStyleId),
    Redefine,
}

fn name(text: &str) -> TestResult<ParagraphStyleName> {
    Ok(ParagraphStyleName::new(text)?)
}

fn host_route(source: &[u8], operation: &Operation) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let id = host_id(&editor, source, leaf_position(source, Kind::TextBox)?)?;
    match operation {
        Operation::Apply(target) => {
            editor.apply_slide_text_box_named_paragraph_style(SLIDE, id, *target)?;
        },
        Operation::Create(source, text) => {
            editor.create_slide_text_box_named_paragraph_style(SLIDE, id, *source, name(text)?)?;
        },
        Operation::Rename(target, text) => {
            editor.rename_slide_text_box_named_paragraph_style(SLIDE, id, *target, name(text)?)?;
        },
        Operation::Delete(target) => {
            editor.delete_slide_text_box_named_paragraph_style(SLIDE, id, *target)?;
        },
        Operation::DeleteApplied(target, replacement) => {
            editor.delete_applied_slide_text_box_named_paragraph_style_with_replacement(
                SLIDE,
                id,
                *target,
                *replacement,
            )?;
        },
        Operation::Redefine => {
            editor.redefine_applied_slide_text_box_named_paragraph_style(SLIDE, id)?;
        },
    }
    Ok(editor.to_bytes()?)
}

fn leaf_route(source: &[u8], operation: &Operation) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let (slide, drawable) = (
        SlideSelector::index(SLIDE),
        DrawableSelector::index(leaf_position(source, Kind::TextBox)?),
    );
    let mut edit = package.edit_slide_drawables();
    match operation {
        Operation::Apply(target) => edit.apply_named_style(slide, drawable, *target)?,
        Operation::Create(source, text) => {
            edit.create_named_style(slide, drawable, *source, name(text)?)?
        },
        Operation::Rename(target, text) => {
            edit.rename_named_style(slide, drawable, *target, name(text)?)?
        },
        Operation::Delete(target) => edit.delete_named_style(slide, drawable, *target)?,
        Operation::DeleteApplied(target, replacement) => {
            edit.delete_applied_named_style(slide, drawable, *target, *replacement)?
        },
        Operation::Redefine => edit.redefine_named_style(slide, drawable)?,
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

fn host_reads(bytes: &[u8]) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let id = host_id(&editor, bytes, leaf_position(bytes, Kind::TextBox)?)?;
    Ok(format!(
        "{:?} {:?}",
        editor.slide_text_box_named_paragraph_styles(SLIDE, id).ok(),
        editor
            .slide_text_box_applied_named_paragraph_style(SLIDE, id)
            .ok()
            .map(|applied| (applied.style().clone(), applied.has_overrides()))
    ))
}

fn leaf_reads(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let (slide, drawable) = (
        SlideSelector::index(SLIDE),
        DrawableSelector::index(leaf_position(bytes, Kind::TextBox)?),
    );
    Ok(format!(
        "{:?} {:?}",
        package.slide_drawable_named_styles(slide, drawable).ok(),
        package
            .slide_drawable_applied_named_style(slide, drawable)
            .ok()
    ))
}

fn run(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    operation: &Operation,
) -> TestResult {
    Case::new(name, Format::Keynote, tier, evidence, input)
        .host(|source| host_route(source, operation))
        .leaf(|source| leaf_route(source, operation))
        .read("host named styles", host_reads)
        .read("leaf named styles", leaf_reads)
        .known_fields(|known| {
            known.with(
                paragraph_style::MESSAGE_TYPE,
                "TSWP.ParagraphStyleArchive",
                decode_as::<tswp::ParagraphStyleArchive>,
            )
        })
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .run()
        .map_err(|error| format!("{operation:?}: {error}"))?;
    Ok(())
}

fn with(input: &Fixture, operation: &Operation) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_route(&input.bytes, operation)?,
    })
}

/// The text box's catalog and applied style.
fn catalog(bytes: &[u8]) -> TestResult<(Vec<ParagraphStyleId>, ParagraphStyleId)> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let id = host_id(&editor, bytes, leaf_position(bytes, Kind::TextBox)?)?;
    Ok((
        editor
            .slide_text_box_named_paragraph_styles(SLIDE, id)?
            .iter()
            .map(|style| style.id())
            .collect(),
        editor
            .slide_text_box_applied_named_paragraph_style(SLIDE, id)?
            .style()
            .id(),
    ))
}

/// The text box with an override of its named style.
fn overridden(input: &Fixture) -> TestResult<Fixture> {
    let mut editor = keynote::Draft::from_bytes(&input.bytes)?;
    let id = host_id(
        &editor,
        &input.bytes,
        leaf_position(&input.bytes, Kind::TextBox)?,
    )?;
    editor.set_slide_text_box_paragraph_alignment(SLIDE, id, Alignment::Center)?;
    Ok(Fixture {
        path: input.path,
        bytes: editor.to_bytes()?,
    })
}

fn with_created(
    input: &Fixture,
    applied: ParagraphStyleId,
) -> TestResult<(Fixture, ParagraphStyleId)> {
    let created = with(input, &Operation::Create(applied, "Litchi Copy"))?;
    let editor = keynote::Draft::from_bytes(&created.bytes)?;
    let id = host_id(
        &editor,
        &created.bytes,
        leaf_position(&created.bytes, Kind::TextBox)?,
    )?;
    let style = editor
        .slide_text_box_named_paragraph_styles(SLIDE, id)?
        .into_iter()
        .find(|style| style.name().as_str() == "Litchi Copy")
        .ok_or("created style")?;
    Ok((created, style.id()))
}

#[test]
fn keynote_slide_text_box_named_style_apply_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    let (styles, applied) = catalog(&input.bytes)?;
    let derived = overridden(&input)?;
    for &target in styles.iter().take(3) {
        for base in [&input, &derived] {
            if target == applied && std::ptr::eq(base, &input) {
                continue;
            }
            run(
                "keynote_slide_text_box_named_style_apply_parity",
                Tier::B,
                Evidence::Queued,
                base,
                &Operation::Apply(target),
            )?;
        }
    }
    Ok(())
}

#[test]
fn keynote_slide_text_box_named_style_create_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    let (_, applied) = catalog(&input.bytes)?;
    run(
        "keynote_slide_text_box_named_style_create_parity",
        Tier::B,
        Evidence::Queued,
        &input,
        &Operation::Create(applied, "Litchi Heading"),
    )
}

#[test]
fn keynote_slide_text_box_named_style_rename_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    let (styles, _) = catalog(&input.bytes)?;
    for &target in styles.iter().take(2) {
        run(
            "keynote_slide_text_box_named_style_rename_parity",
            Tier::B,
            Evidence::Queued,
            &input,
            &Operation::Rename(target, "Litchi Renamed"),
        )?;
    }
    Ok(())
}

#[test]
fn keynote_slide_text_box_named_style_delete_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    let (_, applied) = catalog(&input.bytes)?;
    let (created, copy) = with_created(&input, applied)?;
    run(
        "keynote_slide_text_box_named_style_delete_parity",
        Tier::B,
        Evidence::Queued,
        &created,
        &Operation::Delete(copy),
    )?;
    let applied_copy = with(&created, &Operation::Apply(copy))?;
    run(
        "keynote_slide_text_box_named_style_delete_parity",
        Tier::B,
        Evidence::Queued,
        &applied_copy,
        &Operation::DeleteApplied(copy, applied),
    )
}

#[test]
fn keynote_slide_text_box_named_style_redefine_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    run(
        "keynote_slide_text_box_named_style_redefine_parity",
        Tier::B,
        Evidence::Queued,
        &overridden(&input)?,
        &Operation::Redefine,
    )
}

#[test]
fn keynote_slide_text_box_named_style_read_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    let (_, applied) = catalog(&input.bytes)?;
    let derived = overridden(&input)?;
    let (created, _) = with_created(&input, applied)?;
    for bytes in [&input.bytes, &derived.bytes, &created.bytes] {
        assert_eq!(host_reads(bytes)?, leaf_reads(bytes)?);
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
