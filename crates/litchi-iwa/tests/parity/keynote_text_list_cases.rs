//! Keynote text-box paragraph-list parity (wave W1-C, record 0372): the
//! host `set_/reset_/slide_text_box_paragraph_list*` families against
//! `slide::drawable::Edit::{set_list, set_list_level, set_list_numbering,
//! set_list_property, reset_list_property}` and the list reads, on the
//! shared W1-D list engine.

use litchi_iwa::keynote;
use litchi_iwa_protos::tswp;
use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListLevel, ParagraphListNumbering, ParagraphListStart,
};
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::list::{self, Property, PropertyKind};
use litchi_keynote::slide::drawable::Kind;
use litchi_keynote::{DrawableSelector, Package, SlideSelector};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::keynote_drawable_cases::{host_id, leaf_position};
use crate::known_fields::decode_as;
use crate::leaf_bytes;
use crate::text_list_cases::{list_second, preset_for, samples};

const SLIDE: usize = 0;
const TEXT_BOX: &str = "iwork/keynote/basic.key";

#[derive(Clone, Debug)]
enum Operation {
    SetList(ParagraphList),
    ResetList,
    SetLevel(ParagraphListLevel),
    ResetLevel,
    SetNumbering(ParagraphListNumbering),
    Set(Property),
    Reset(PropertyKind),
}

const P: TextPosition = TextPosition::ZERO;

fn host_route(source: &[u8], operation: &Operation) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let id = host_id(&editor, source, leaf_position(source, Kind::TextBox)?)?;
    match operation.clone() {
        Operation::SetList(preset) => {
            editor.set_slide_text_box_paragraph_list(SLIDE, id, preset)?
        },
        Operation::ResetList => {
            editor.reset_slide_text_box_paragraph_list(SLIDE, id)?;
        },
        Operation::SetLevel(level) => {
            editor.set_slide_text_box_paragraph_list_level(SLIDE, id, P, level)?;
        },
        Operation::ResetLevel => {
            editor.reset_slide_text_box_paragraph_list_level(SLIDE, id, P)?;
        },
        Operation::SetNumbering(numbering) => {
            editor.set_slide_text_box_paragraph_list_numbering(SLIDE, id, P, numbering)?;
        },
        Operation::Set(property) => match property {
            Property::Bullet(v) => {
                editor.set_slide_text_box_paragraph_list_bullet(SLIDE, id, P, &v)?
            },
            Property::BulletGeometry(v) => {
                editor.set_slide_text_box_paragraph_list_bullet_geometry(SLIDE, id, P, v)?;
            },
            Property::Indentation(v) => {
                editor.set_slide_text_box_paragraph_list_indentation(SLIDE, id, P, v)?;
            },
            Property::LabelColor(v) => {
                editor.set_slide_text_box_paragraph_list_label_color(SLIDE, id, P, v)?;
            },
            Property::NumberFormat(v) => {
                editor.set_slide_text_box_paragraph_list_number_format(SLIDE, id, P, v)?;
            },
            Property::NumberScale(v) => {
                editor.set_slide_text_box_paragraph_list_number_scale(SLIDE, id, P, v)?;
            },
            Property::NumberTiering(v) => {
                editor.set_slide_text_box_paragraph_list_number_tiering(SLIDE, id, P, v)?;
            },
            other => return Err(format!("unmapped property {other:?}").into()),
        },
        Operation::Reset(kind) => {
            let _ = match kind {
                PropertyKind::Bullet => {
                    editor.reset_slide_text_box_paragraph_list_bullet(SLIDE, id, P)?
                },
                PropertyKind::BulletGeometry => {
                    editor.reset_slide_text_box_paragraph_list_bullet_geometry(SLIDE, id, P)?
                },
                PropertyKind::Indentation => {
                    editor.reset_slide_text_box_paragraph_list_indentation(SLIDE, id, P)?
                },
                PropertyKind::LabelColor => {
                    editor.reset_slide_text_box_paragraph_list_label_color(SLIDE, id, P)?
                },
                PropertyKind::NumberFormat => {
                    editor.reset_slide_text_box_paragraph_list_number_format(SLIDE, id, P)?
                },
                PropertyKind::NumberScale => {
                    editor.reset_slide_text_box_paragraph_list_number_scale(SLIDE, id, P)?
                },
                PropertyKind::NumberTiering => {
                    editor.reset_slide_text_box_paragraph_list_number_tiering(SLIDE, id, P)?
                },
                other => return Err(format!("unmapped kind {other:?}").into()),
            };
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
    match operation.clone() {
        Operation::SetList(preset) => edit.set_list(slide, drawable, preset)?,
        Operation::ResetList => edit.set_list(slide, drawable, ParagraphList::None)?,
        Operation::SetLevel(level) => edit.set_list_level(slide, drawable, P, level)?,
        Operation::ResetLevel => {
            edit.set_list_level(slide, drawable, P, ParagraphListLevel::ZERO)?
        },
        Operation::SetNumbering(numbering) => {
            edit.set_list_numbering(slide, drawable, P, numbering)?
        },
        Operation::Set(property) => edit.set_list_property(slide, drawable, P, property)?,
        Operation::Reset(kind) => edit.reset_list_property(slide, drawable, P, kind)?,
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
        "{:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?}",
        editor.slide_text_box_paragraph_list(SLIDE, id).ok(),
        editor.slide_text_box_paragraph_list_levels(SLIDE, id).ok(),
        editor
            .slide_text_box_paragraph_list_level(SLIDE, id, P)
            .ok(),
        editor
            .slide_text_box_paragraph_list_numbering(SLIDE, id, P)
            .ok(),
        editor
            .slide_text_box_paragraph_list_bullet(SLIDE, id, P)
            .ok(),
        editor
            .slide_text_box_paragraph_list_bullet_geometry(SLIDE, id, P)
            .ok(),
        editor
            .slide_text_box_paragraph_list_indentation(SLIDE, id, P)
            .ok(),
        editor
            .slide_text_box_paragraph_list_label_color(SLIDE, id, P)
            .ok(),
        editor
            .slide_text_box_paragraph_list_number_format(SLIDE, id, P)
            .ok(),
        editor
            .slide_text_box_paragraph_list_number_scale(SLIDE, id, P)
            .ok(),
        editor
            .slide_text_box_paragraph_list_number_tiering(SLIDE, id, P)
            .ok(),
    ))
}

fn leaf_reads(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let (slide, drawable) = (
        SlideSelector::index(SLIDE),
        DrawableSelector::index(leaf_position(bytes, Kind::TextBox)?),
    );
    let read = |kind| {
        package
            .slide_drawable_list_property(slide, drawable, P, kind)
            .ok()
    };
    macro_rules! value {
        ($kind:ident) => {
            match read(PropertyKind::$kind) {
                Some(Property::$kind(value)) => Some(value),
                _ => None,
            }
        };
    }
    Ok(format!(
        "{:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?}",
        package.slide_drawable_paragraph_list(slide, drawable).ok(),
        package.slide_drawable_list_levels(slide, drawable).ok(),
        package.slide_drawable_list_level(slide, drawable, P).ok(),
        package
            .slide_drawable_list_numbering(slide, drawable, P)
            .ok(),
        value!(Bullet),
        value!(BulletGeometry),
        value!(Indentation),
        value!(LabelColor),
        value!(NumberFormat),
        value!(NumberScale),
        value!(NumberTiering),
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
        .read("host list reads", host_reads)
        .read("leaf list reads", leaf_reads)
        .known_fields(|known| {
            known.with(
                list::MESSAGE_TYPE,
                "TSWP.ListStyleArchive",
                decode_as::<tswp::ListStyleArchive>,
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

#[test]
fn keynote_slide_text_box_list_preset_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    for preset in [ParagraphList::Bullet, ParagraphList::Numbered] {
        run(
            "keynote_slide_text_box_list_preset_parity",
            Tier::B,
            Evidence::Queued,
            &input,
            &Operation::SetList(preset),
        )?;
        let base = with(&input, &Operation::SetList(preset))?;
        run(
            "keynote_slide_text_box_list_preset_parity",
            Tier::B,
            Evidence::Queued,
            &base,
            &Operation::ResetList,
        )?;
    }
    Ok(())
}

#[test]
fn keynote_slide_text_box_list_level_parity() -> TestResult {
    let base = with(
        &fixture(TEXT_BOX)?,
        &Operation::SetList(ParagraphList::Bullet),
    )?;
    let level = ParagraphListLevel::new(2)?;
    run(
        "keynote_slide_text_box_list_level_parity",
        Tier::B,
        Evidence::Queued,
        &base,
        &Operation::SetLevel(level),
    )?;
    run(
        "keynote_slide_text_box_list_level_parity",
        Tier::B,
        Evidence::Queued,
        &with(&base, &Operation::SetLevel(level))?,
        &Operation::ResetLevel,
    )?;
    Ok(())
}

#[test]
fn keynote_slide_text_box_list_numbering_parity() -> TestResult {
    let base = with(
        &fixture(TEXT_BOX)?,
        &Operation::SetList(ParagraphList::Numbered),
    )?;
    run(
        "keynote_slide_text_box_list_numbering_parity",
        Tier::B,
        Evidence::Queued,
        &base,
        &Operation::SetNumbering(ParagraphListNumbering::StartAt(ParagraphListStart::new(5)?)),
    )?;
    Ok(())
}

#[test]
fn keynote_slide_text_box_list_property_set_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    for property in samples()? {
        let base = with(&input, &Operation::SetList(preset_for(property.kind())))?;
        run(
            "keynote_slide_text_box_list_property_set_parity",
            Tier::B,
            Evidence::Queued,
            &base,
            &Operation::Set(property.clone()),
        )?;
        let set = with(&base, &Operation::Set(property.clone()))?;
        run(
            "keynote_slide_text_box_list_property_set_parity",
            Tier::B,
            Evidence::Queued,
            &set,
            &Operation::Set(list_second(&property)?),
        )?;
    }
    Ok(())
}

#[test]
fn keynote_slide_text_box_list_property_reset_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    for property in samples()? {
        let base = with(&input, &Operation::SetList(preset_for(property.kind())))?;
        let set = with(&base, &Operation::Set(property.clone()))?;
        run(
            "keynote_slide_text_box_list_property_reset_parity",
            Tier::B,
            Evidence::Queued,
            &set,
            &Operation::Reset(property.kind()),
        )?;
    }
    Ok(())
}

#[test]
fn keynote_slide_text_box_list_read_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    let mut inputs = vec![input.bytes.clone()];
    for preset in [ParagraphList::Bullet, ParagraphList::Numbered] {
        let base = with(&input, &Operation::SetList(preset))?;
        inputs.push(base.bytes.clone());
        for property in samples()? {
            if preset_for(property.kind()) == preset {
                inputs.push(with(&base, &Operation::Set(property))?.bytes);
            }
        }
    }
    for bytes in &inputs {
        assert_eq!(host_reads(bytes)?, leaf_reads(bytes)?);
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
