//! Pages body text-box named paragraph-style parity (wave W4-P, record
//! 0413): the host `*_text_box_named_paragraph_style*` families against
//! `litchi_pages::Package::edit_body_text_box_named_styles` and
//! the named-style reads on the shared W1-D engine.

use litchi_iwa::pages;
use litchi_iwa_protos::tswp;
use litchi_iwa_text::paragraph::format::Alignment;
use litchi_iwa_text_wire::paragraph_style;
use litchi_pages::DrawableSelector;
use litchi_pages::Package;
use litchi_pages::drawable::named_style::{ParagraphStyleId, ParagraphStyleName};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::known_fields::decode_as;
use crate::leaf_bytes;
use crate::pages_text_box_cases::{host_text_box, text_box_fixture as authored_fixture};

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
    let mut e = pages::Draft::from_bytes(source)?;
    let d = host_text_box(&e, 0)?;
    match operation {
        Operation::Apply(target) => {
            e.apply_text_box_named_paragraph_style(d, *target)?;
        },
        Operation::Create(source, text) => {
            e.create_text_box_named_paragraph_style(d, *source, name(text)?)?;
        },
        Operation::Rename(target, text) => {
            e.rename_text_box_named_paragraph_style(d, *target, name(text)?)?;
        },
        Operation::Delete(target) => {
            e.delete_text_box_named_paragraph_style(d, *target)?;
        },
        Operation::DeleteApplied(target, replacement) => {
            e.delete_applied_text_box_named_paragraph_style_with_replacement(
                d,
                *target,
                *replacement,
            )?;
        },
        Operation::Redefine => {
            e.redefine_applied_text_box_named_paragraph_style(d)?;
        },
    }
    Ok(e.to_bytes()?)
}

fn leaf_route(source: &[u8], operation: &Operation) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_text_box_named_styles(DrawableSelector::text_box(0))?;
    match operation {
        Operation::Apply(target) => edit.apply(*target)?,
        Operation::Create(source, text) => edit.create(*source, name(text)?)?,
        Operation::Rename(target, text) => edit.rename_style(*target, name(text)?)?,
        Operation::Delete(target) => edit.delete(*target)?,
        Operation::DeleteApplied(target, replacement) => {
            edit.delete_applied(*target, *replacement)?
        },
        Operation::Redefine => edit.redefine()?,
    };
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_text_box_named_styles(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn host_reads(bytes: &[u8]) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let d = host_text_box(&e, 0)?;
    Ok(format!(
        "{:?} {:?}",
        e.text_box_named_paragraph_styles(d).ok(),
        e.text_box_applied_named_paragraph_style(d)
            .ok()
            .map(|applied| (applied.style().clone(), applied.has_overrides()))
    ))
}

fn leaf_reads(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let drawable = DrawableSelector::text_box(0);
    Ok(format!(
        "{:?} {:?}",
        package.body_text_box_named_styles(drawable).ok(),
        package.body_text_box_applied_named_style(drawable).ok()
    ))
}

fn run(name: &'static str, tier: Tier, input: &Fixture, operation: &Operation) -> TestResult {
    let mut case = Case::new(name, Format::Pages, tier, Evidence::Queued, input)
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
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()
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
    let e = pages::Draft::from_bytes(bytes)?;
    let d = host_text_box(&e, 0)?;
    Ok((
        e.text_box_named_paragraph_styles(d)?
            .iter()
            .map(|style| style.id())
            .collect(),
        e.text_box_applied_named_paragraph_style(d)?.style().id(),
    ))
}

/// The text box with an override of its named style.
fn overridden(input: &Fixture) -> TestResult<Fixture> {
    let mut e = pages::Draft::from_bytes(&input.bytes)?;
    let d = host_text_box(&e, 0)?;
    e.set_text_box_paragraph_alignment(d, Alignment::Center)?;
    Ok(Fixture {
        path: input.path,
        bytes: e.to_bytes()?,
    })
}

fn with_created(
    input: &Fixture,
    applied: ParagraphStyleId,
) -> TestResult<(Fixture, ParagraphStyleId)> {
    let created = with(input, &Operation::Create(applied, "Litchi Copy"))?;
    let e = pages::Draft::from_bytes(&created.bytes)?;
    let d = host_text_box(&e, 0)?;
    let style = e
        .text_box_named_paragraph_styles(d)?
        .into_iter()
        .find(|style| style.name().as_str() == "Litchi Copy")
        .ok_or("created style")?;
    Ok((created, style.id()))
}

#[test]
fn pages_text_box_named_style_apply_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = authored_fixture()?;
    let (styles, applied) = catalog(&input.bytes)?;
    assert!(
        styles.len() >= 3,
        "the fixture lists {} named styles",
        styles.len()
    );
    let derived = overridden(&input)?;
    for &target in styles.iter().take(3) {
        for base in [&input, &derived] {
            if target == applied && std::ptr::eq(base, &input) {
                continue;
            }
            run(
                "pages_text_box_named_style_apply_parity",
                Tier::B,
                base,
                &Operation::Apply(target),
            )?;
        }
    }
    Ok(())
}

#[test]
fn pages_text_box_named_style_create_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = authored_fixture()?;
    let (_, applied) = catalog(&input.bytes)?;
    run(
        "pages_text_box_named_style_create_parity",
        Tier::B,
        &input,
        &Operation::Create(applied, "Litchi Heading"),
    )
}

#[test]
fn pages_text_box_named_style_rename_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = authored_fixture()?;
    let (styles, _) = catalog(&input.bytes)?;
    for &target in styles.iter().take(2) {
        run(
            "pages_text_box_named_style_rename_parity",
            Tier::A,
            &input,
            &Operation::Rename(target, "Litchi Renamed"),
        )?;
    }
    Ok(())
}

#[test]
fn pages_text_box_named_style_delete_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = authored_fixture()?;
    let (_, applied) = catalog(&input.bytes)?;
    let (created, copy) = with_created(&input, applied)?;
    run(
        "pages_text_box_named_style_delete_parity",
        Tier::B,
        &created,
        &Operation::Delete(copy),
    )?;
    let applied_copy = with(&created, &Operation::Apply(copy))?;
    run(
        "pages_text_box_named_style_delete_parity",
        Tier::B,
        &applied_copy,
        &Operation::DeleteApplied(copy, applied),
    )
}

#[test]
fn pages_text_box_named_style_redefine_parity() -> TestResult {
    let _ = Evidence::Queued;
    let input = authored_fixture()?;
    run(
        "pages_text_box_named_style_redefine_parity",
        Tier::B,
        &overridden(&input)?,
        &Operation::Redefine,
    )
}

#[test]
fn pages_text_box_named_style_read_parity() -> TestResult {
    let input = authored_fixture()?;
    let (_, applied) = catalog(&input.bytes)?;
    let derived = overridden(&input)?;
    let (created, _) = with_created(&input, applied)?;
    for bytes in [&input.bytes, &derived.bytes, &created.bytes] {
        assert_eq!(host_reads(bytes)?, leaf_reads(bytes)?);
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe (record 0413)"]
fn text_box_named_style_apply_cost_abba() -> TestResult {
    let input = overridden(&authored_fixture()?)?;
    let (styles, _) = catalog(&input.bytes)?;
    let target = *styles.first().ok_or("no named style")?;
    crate::pages_cost_probes::abba(
        &input,
        |source| host_route(source, &Operation::Apply(target)),
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit =
                package.edit_body_text_box_named_styles(DrawableSelector::text_box(0))?;
            edit.apply(target)?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}
