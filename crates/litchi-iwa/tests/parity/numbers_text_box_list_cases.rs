//! Numbers sheet text-box paragraph-list parity (wave W2-A, record 0378):
//! the host `set_/reset_/sheet_text_box_paragraph_list*` families against
//! `litchi_numbers::Package::edit_sheet_text_box_lists` (W1-D's list engine
//! published by the shared storage-list publisher).

use litchi_iwa::numbers;
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::list::{
    ParagraphList, ParagraphListLevel, ParagraphListNumbering, ParagraphListStart, Property,
    PropertyKind, TextPosition,
};
use litchi_numbers::sheet::drawable::{DrawableSelector, Kind};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::known_fields::decode_as;
use crate::leaf_bytes;
use crate::numbers_drawable_cases::{authored_fixture, host_drawable};
use crate::text_list_cases::{list_second, preset_for, samples};

/// One text-box list change.
#[derive(Clone, Debug)]
enum Operation {
    SetList(ParagraphList),
    ResetList,
    SetLevel(TextPosition, ParagraphListLevel),
    ResetLevel(TextPosition),
    SetNumbering(TextPosition, ParagraphListNumbering),
    Set(TextPosition, Property),
    Reset(TextPosition, PropertyKind),
}

#[allow(clippy::too_many_lines, reason = "one arm per host method")]
fn host_route(source: &[u8], operation: &Operation) -> TestResult<Vec<u8>> {
    let mut e = numbers::Draft::from_bytes(source)?;
    let (s, d) = host_drawable(&e, Kind::TextBox, 0)?;
    match operation.clone() {
        Operation::SetList(list) => e.set_sheet_text_box_paragraph_list(s, d, list)?,
        Operation::ResetList => drop(e.reset_sheet_text_box_paragraph_list(s, d)?),
        Operation::SetLevel(p, level) => {
            e.set_sheet_text_box_paragraph_list_level(s, d, p, level)?
        },
        Operation::ResetLevel(p) => drop(e.reset_sheet_text_box_paragraph_list_level(s, d, p)?),
        Operation::SetNumbering(p, numbering) => {
            e.set_sheet_text_box_paragraph_list_numbering(s, d, p, numbering)?;
        },
        Operation::Set(p, property) => match property {
            Property::Bullet(v) => e.set_sheet_text_box_paragraph_list_bullet(s, d, p, &v)?,
            Property::BulletGeometry(v) => {
                e.set_sheet_text_box_paragraph_list_bullet_geometry(s, d, p, v)?;
            },
            Property::Indentation(v) => {
                e.set_sheet_text_box_paragraph_list_indentation(s, d, p, v)?;
            },
            Property::LabelColor(v) => {
                e.set_sheet_text_box_paragraph_list_label_color(s, d, p, v)?
            },
            Property::NumberFormat(v) => {
                e.set_sheet_text_box_paragraph_list_number_format(s, d, p, v)?;
            },
            Property::NumberScale(v) => {
                e.set_sheet_text_box_paragraph_list_number_scale(s, d, p, v)?;
            },
            Property::NumberTiering(v) => {
                e.set_sheet_text_box_paragraph_list_number_tiering(s, d, p, v)?;
            },
            other => return Err(format!("unmapped property {other:?}").into()),
        },
        Operation::Reset(p, kind) => drop(match kind {
            PropertyKind::Bullet => e.reset_sheet_text_box_paragraph_list_bullet(s, d, p)?,
            PropertyKind::BulletGeometry => {
                e.reset_sheet_text_box_paragraph_list_bullet_geometry(s, d, p)?
            },
            PropertyKind::Indentation => {
                e.reset_sheet_text_box_paragraph_list_indentation(s, d, p)?
            },
            PropertyKind::LabelColor => {
                e.reset_sheet_text_box_paragraph_list_label_color(s, d, p)?
            },
            PropertyKind::NumberFormat => {
                e.reset_sheet_text_box_paragraph_list_number_format(s, d, p)?
            },
            PropertyKind::NumberScale => {
                e.reset_sheet_text_box_paragraph_list_number_scale(s, d, p)?
            },
            PropertyKind::NumberTiering => {
                e.reset_sheet_text_box_paragraph_list_number_tiering(s, d, p)?
            },
            other => return Err(format!("unmapped kind {other:?}").into()),
        }),
    }
    Ok(e.to_bytes()?)
}

fn leaf_route(source: &[u8], operation: &Operation) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_text_box_lists(0, DrawableSelector::text_box(0))?;
    match operation.clone() {
        Operation::SetList(list) => {
            edit.set_list(list)?;
        },
        Operation::ResetList => {
            edit.set_list(ParagraphList::None)?;
        },
        Operation::SetLevel(p, level) => {
            edit.set_level(p, level)?;
        },
        Operation::ResetLevel(p) => {
            edit.set_level(p, ParagraphListLevel::ZERO)?;
        },
        Operation::SetNumbering(p, numbering) => {
            edit.set_numbering(p, numbering)?;
        },
        Operation::Set(p, property) => {
            edit.set(p, property)?;
        },
        Operation::Reset(p, kind) => {
            edit.reset(p, kind)?;
        },
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_sheet_text_box_lists(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

const PROBES: u32 = 12;

fn host_reads(bytes: &[u8]) -> TestResult<Vec<String>> {
    let e = numbers::Draft::from_bytes(bytes)?;
    let (s, d) = host_drawable(&e, Kind::TextBox, 0)?;
    let mut reads = vec![
        format!("{:?}", e.sheet_text_box_paragraph_list(s, d).ok()),
        format!("{:?}", e.sheet_text_box_paragraph_list_levels(s, d).ok()),
    ];
    for index in 0..PROBES {
        let p = TextPosition::from_utf16_code_units(index);
        reads.push(format!(
            "{p:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?}",
            e.sheet_text_box_paragraph_list_level(s, d, p).ok(),
            e.sheet_text_box_paragraph_list_numbering(s, d, p).ok(),
            e.sheet_text_box_paragraph_list_bullet(s, d, p)
                .ok()
                .map(Property::Bullet),
            e.sheet_text_box_paragraph_list_bullet_geometry(s, d, p)
                .ok()
                .map(Property::BulletGeometry),
            e.sheet_text_box_paragraph_list_indentation(s, d, p)
                .ok()
                .map(Property::Indentation),
            e.sheet_text_box_paragraph_list_label_color(s, d, p)
                .ok()
                .map(Property::LabelColor),
            e.sheet_text_box_paragraph_list_number_format(s, d, p)
                .ok()
                .map(Property::NumberFormat),
            e.sheet_text_box_paragraph_list_number_scale(s, d, p)
                .ok()
                .map(Property::NumberScale),
            e.sheet_text_box_paragraph_list_number_tiering(s, d, p)
                .ok()
                .map(Property::NumberTiering),
        ));
    }
    Ok(reads)
}

fn leaf_reads(bytes: &[u8]) -> TestResult<Vec<String>> {
    let package = Package::from_bytes(bytes)?;
    let t = DrawableSelector::text_box(0);
    let mut reads = vec![
        format!("{:?}", package.sheet_text_box_list(0, t).ok()),
        format!("{:?}", package.sheet_text_box_list_levels(0, t).ok()),
    ];
    for index in 0..PROBES {
        let p = TextPosition::from_utf16_code_units(index);
        let property = |kind| package.sheet_text_box_list_property(0, t, p, kind).ok();
        reads.push(format!(
            "{p:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?}",
            package.sheet_text_box_list_level(0, t, p).ok(),
            package.sheet_text_box_list_numbering(0, t, p).ok(),
            property(PropertyKind::Bullet),
            property(PropertyKind::BulletGeometry),
            property(PropertyKind::Indentation),
            property(PropertyKind::LabelColor),
            property(PropertyKind::NumberFormat),
            property(PropertyKind::NumberScale),
            property(PropertyKind::NumberTiering),
        ));
    }
    Ok(reads)
}

fn run(name: &'static str, tier: Tier, input: &Fixture, operation: &Operation) -> TestResult {
    Case::new(name, Format::Numbers, tier, Evidence::Queued, input)
        .leaf_structural_metadata()
        .known_fields(|known| {
            known
                .with(
                    2_001,
                    "TSWP.StorageArchive",
                    decode_as::<litchi_iwa_protos::tswp::StorageArchive>,
                )
                .with(
                    2_023,
                    "TSWP.ListStyleArchive",
                    decode_as::<litchi_iwa_protos::tswp::ListStyleArchive>,
                )
        })
        .host(|source| host_route(source, operation))
        .leaf(|source| leaf_route(source, operation))
        .read("host reads", host_reads)
        .read("leaf reads", leaf_reads)
        .run()?;
    Ok(())
}

fn with(input: &Fixture, operation: &Operation) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_route(&input.bytes, operation)?,
    })
}

#[test]
fn numbers_sheet_text_box_list_preset_parity() -> TestResult {
    let input = authored_fixture()?;
    for preset in [ParagraphList::Bullet, ParagraphList::Numbered] {
        run(
            "numbers_sheet_text_box_list_preset_parity",
            Tier::B,
            &input,
            &Operation::SetList(preset),
        )
        .map_err(|e| format!("set {preset:?}: {e}"))?;
        let base = with(&input, &Operation::SetList(preset))?;
        let other = if preset == ParagraphList::Bullet {
            ParagraphList::Numbered
        } else {
            ParagraphList::Bullet
        };
        for operation in [Operation::ResetList, Operation::SetList(other)] {
            run(
                "numbers_sheet_text_box_list_preset_parity",
                Tier::B,
                &base,
                &operation,
            )
            .map_err(|e| format!("{preset:?} then {operation:?}: {e}"))?;
        }
    }
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn numbers_sheet_text_box_list_level_numbering_parity() -> TestResult {
    let input = authored_fixture()?;
    let two = ParagraphListLevel::new(2)?;
    let base = with(&input, &Operation::SetList(ParagraphList::Numbered))?;
    for (fixture, operation) in [
        (&base, Operation::SetLevel(TextPosition::ZERO, two)),
        (
            &base,
            Operation::SetNumbering(
                TextPosition::ZERO,
                ParagraphListNumbering::StartAt(ParagraphListStart::new(3)?),
            ),
        ),
    ] {
        run(
            "numbers_sheet_text_box_list_level_numbering_parity",
            Tier::A,
            fixture,
            &operation,
        )
        .map_err(|e| format!("{operation:?}: {e}"))?;
    }
    let leveled = with(&base, &Operation::SetLevel(TextPosition::ZERO, two))?;
    run(
        "numbers_sheet_text_box_list_level_numbering_parity",
        Tier::A,
        &leveled,
        &Operation::ResetLevel(TextPosition::ZERO),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

fn property_cases(
    name: &'static str,
    prepare: impl Fn(&Fixture, &Property) -> TestResult<(Fixture, Operation)>,
) -> TestResult {
    let input = authored_fixture()?;
    for property in samples()? {
        let base = with(&input, &Operation::SetList(preset_for(property.kind())))?;
        let (fixture, operation) = prepare(&base, &property)?;
        run(name, Tier::B, &fixture, &operation).map_err(|e| format!("{operation:?}: {e}"))?;
    }
    Ok(())
}

#[test]
fn numbers_sheet_text_box_list_property_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    property_cases(
        "numbers_sheet_text_box_list_property_set_parity",
        |base, property| {
            Ok((
                Fixture {
                    path: base.path,
                    bytes: base.bytes.clone(),
                },
                Operation::Set(TextPosition::ZERO, property.clone()),
            ))
        },
    )
}

#[test]
fn numbers_sheet_text_box_list_property_rewrite_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    property_cases(
        "numbers_sheet_text_box_list_property_rewrite_parity",
        |base, property| {
            let set = with(base, &Operation::Set(TextPosition::ZERO, property.clone()))?;
            Ok((
                set,
                Operation::Set(TextPosition::ZERO, list_second(property)?),
            ))
        },
    )
}

#[test]
fn numbers_sheet_text_box_list_property_reset_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    property_cases(
        "numbers_sheet_text_box_list_property_reset_parity",
        |base, property| {
            let set = with(base, &Operation::Set(TextPosition::ZERO, property.clone()))?;
            Ok((set, Operation::Reset(TextPosition::ZERO, property.kind())))
        },
    )
}

#[test]
fn numbers_sheet_text_box_list_read_parity() -> TestResult {
    let input = authored_fixture()?;
    let mut fixtures = vec![Fixture {
        path: input.path,
        bytes: input.bytes.clone(),
    }];
    for preset in [ParagraphList::Bullet, ParagraphList::Numbered] {
        fixtures.push(with(&input, &Operation::SetList(preset))?);
    }
    for property in samples()? {
        let base = with(&input, &Operation::SetList(preset_for(property.kind())))?;
        fixtures.push(with(&base, &Operation::Set(TextPosition::ZERO, property))?);
    }
    for fixture in &fixtures {
        assert_eq!(host_reads(&fixture.bytes)?, leaf_reads(&fixture.bytes)?);
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
