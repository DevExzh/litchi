//! Pages body-table cell paragraph-list parity (record 0412): the host
//! editor's `set_table_cell_paragraph_list*`, `reset_table_cell_paragraph_list*`
//! and `table_cell_paragraph_list*` families on body tables against the
//! leaf's one family `edit_body_table_cell_lists` / `body_table_cell_list*`
//! over the shared cell rich-text engine
//! (`litchi_iwa_table::cell_rich_text::list`), lifted from the Numbers
//! family of record 0370b. The leaf promotes a plain cell to rich text as
//! the host does.
//!
//! A change that stores a new list style registers a fresh UUID (random on
//! the host), so it is tier B; promotion itself registers none.

use litchi_iwa::pages;
use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListLevel, ParagraphListNumbering, ParagraphListPlacement,
    ParagraphListStart,
};
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::list::{Property, PropertyKind};
use litchi_pages::Package;
use litchi_pages::table::cell_list::At;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::decode_as;
use crate::leaf_bytes;
use crate::text_list_cases::{list_second, preset_for, samples};

pub use crate::numbers_cell_list_cases::Operation;

/// The fixtures whose cells the write cases change.
const WRITE_FIXTURES: [&str; 2] = [
    "iwork/pages/body-table-cells-native.pages",
    "iwork/pages/body-table-read-native.pages",
];
/// The body table the cases edit.
const TABLE: usize = 0;

/// The positions the reads probe in each cell's text.
const PROBES: u32 = 8;

/// The model of the leaf's first body table.
fn host_first_model(bytes: &[u8]) -> TestResult<u64> {
    Ok(pages::Draft::from_bytes(bytes)?
        .tables()?
        .get(TABLE)
        .ok_or("body table missing")?
        .model_object_id)
}

#[allow(clippy::too_many_lines, reason = "one arm per host method")]
fn host_apply(
    editor: &mut pages::Draft,
    id: u64,
    (r, c): (usize, usize),
    operation: &Operation,
) -> TestResult<bool> {
    Ok(match operation.clone() {
        Operation::SetList(list) => {
            editor.set_table_cell_paragraph_list(id, r, c, list)?;
            true
        },
        Operation::ResetList => editor.reset_table_cell_paragraph_list(id, r, c)?,
        Operation::SetLists(placements) => {
            editor.set_table_cell_paragraph_lists(id, r, c, &placements)?;
            true
        },
        Operation::SetLevel(p, level) => {
            editor.set_table_cell_paragraph_list_level(id, r, c, p, level)?;
            true
        },
        Operation::ResetLevel(p) => editor.reset_table_cell_paragraph_list_level(id, r, c, p)?,
        Operation::SetNumbering(p, numbering) => {
            editor.set_table_cell_paragraph_list_numbering(id, r, c, p, numbering)?;
            true
        },
        Operation::Set(p, property) => {
            match property {
                Property::Bullet(v) => {
                    editor.set_table_cell_paragraph_list_bullet(id, r, c, p, &v)?
                },
                Property::BulletGeometry(v) => {
                    editor.set_table_cell_paragraph_list_bullet_geometry(id, r, c, p, v)?;
                },
                Property::Indentation(v) => {
                    editor.set_table_cell_paragraph_list_indentation(id, r, c, p, v)?;
                },
                Property::LabelColor(v) => {
                    editor.set_table_cell_paragraph_list_label_color(id, r, c, p, v)?;
                },
                Property::NumberFormat(v) => {
                    editor.set_table_cell_paragraph_list_number_format(id, r, c, p, v)?;
                },
                Property::NumberScale(v) => {
                    editor.set_table_cell_paragraph_list_number_scale(id, r, c, p, v)?;
                },
                Property::NumberTiering(v) => {
                    editor.set_table_cell_paragraph_list_number_tiering(id, r, c, p, v)?;
                },
                other => return Err(format!("unmapped property {other:?}").into()),
            }
            true
        },
        Operation::Reset(p, kind) => match kind {
            PropertyKind::Bullet => editor.reset_table_cell_paragraph_list_bullet(id, r, c, p)?,
            PropertyKind::BulletGeometry => {
                editor.reset_table_cell_paragraph_list_bullet_geometry(id, r, c, p)?
            },
            PropertyKind::Indentation => {
                editor.reset_table_cell_paragraph_list_indentation(id, r, c, p)?
            },
            PropertyKind::LabelColor => {
                editor.reset_table_cell_paragraph_list_label_color(id, r, c, p)?
            },
            PropertyKind::NumberFormat => {
                editor.reset_table_cell_paragraph_list_number_format(id, r, c, p)?
            },
            PropertyKind::NumberScale => {
                editor.reset_table_cell_paragraph_list_number_scale(id, r, c, p)?
            },
            PropertyKind::NumberTiering => {
                editor.reset_table_cell_paragraph_list_number_tiering(id, r, c, p)?
            },
            other => return Err(format!("unmapped kind {other:?}").into()),
        },
    })
}

fn host_route(source: &[u8], cell: (usize, usize), operation: &Operation) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    let id = host_first_model(source)?;
    host_apply(&mut editor, id, cell, operation)?;
    Ok(editor.to_bytes()?)
}

fn leaf_route(source: &[u8], cell: (usize, usize), operation: &Operation) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let at = At::new(cell.0 as u32, cell.1 as u32);
    let edit = package.edit_body_table_cell_lists(TABLE)?;
    let edit = match operation.clone() {
        Operation::SetList(list) => edit.set_list(at, list)?,
        Operation::ResetList => edit.set_list(at, ParagraphList::None)?,
        Operation::SetLists(placements) => edit.set_lists(at, placements)?,
        Operation::SetLevel(p, level) => edit.set_level(at, p, level)?,
        Operation::ResetLevel(p) => edit.set_level(at, p, ParagraphListLevel::ZERO)?,
        Operation::SetNumbering(p, numbering) => edit.set_numbering(at, p, numbering)?,
        Operation::Set(p, property) => edit.set(at, p, property)?,
        Operation::Reset(p, kind) => edit.reset(at, p, kind)?,
    };
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_table_cell_lists(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

/// Every list read of one cell, through the host.
pub fn host_reads(bytes: &[u8], (r, c): (usize, usize)) -> TestResult<Vec<String>> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let id = host_first_model(bytes)?;
    let mut reads = vec![
        format!("{:?}", editor.table_cell_paragraph_list(id, r, c).ok()),
        format!("{:?}", editor.table_cell_paragraph_lists(id, r, c).ok()),
        format!(
            "{:?}",
            editor.table_cell_paragraph_list_levels(id, r, c).ok()
        ),
    ];
    for index in 0..PROBES {
        let p = TextPosition::from_utf16_code_units(index);
        reads.push(format!(
            "{p:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?}",
            editor.table_cell_paragraph_list_numbering(id, r, c, p).ok(),
            editor
                .table_cell_paragraph_list_bullet(id, r, c, p)
                .ok()
                .map(Property::Bullet),
            editor
                .table_cell_paragraph_list_bullet_geometry(id, r, c, p)
                .ok()
                .map(Property::BulletGeometry),
            editor
                .table_cell_paragraph_list_indentation(id, r, c, p)
                .ok()
                .map(Property::Indentation),
            editor
                .table_cell_paragraph_list_label_color(id, r, c, p)
                .ok()
                .map(Property::LabelColor),
            editor
                .table_cell_paragraph_list_number_format(id, r, c, p)
                .ok()
                .map(Property::NumberFormat),
            editor
                .table_cell_paragraph_list_number_scale(id, r, c, p)
                .ok()
                .map(Property::NumberScale),
            editor
                .table_cell_paragraph_list_number_tiering(id, r, c, p)
                .ok()
                .map(Property::NumberTiering),
        ));
    }
    Ok(reads)
}

/// Every list read of one cell, through the leaf.
pub fn leaf_reads(bytes: &[u8], (r, c): (usize, usize)) -> TestResult<Vec<String>> {
    let package = Package::from_bytes(bytes)?;
    let at = At::new(r as u32, c as u32);
    let mut reads = vec![
        format!("{:?}", package.body_table_cell_list(TABLE, at).ok()),
        format!("{:?}", package.body_table_cell_lists(TABLE, at).ok()),
        format!("{:?}", package.body_table_cell_list_levels(TABLE, at).ok()),
    ];
    for index in 0..PROBES {
        let p = TextPosition::from_utf16_code_units(index);
        let property = |kind| {
            package
                .body_table_cell_list_property(TABLE, at, p, kind)
                .ok()
        };
        reads.push(format!(
            "{p:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?}",
            package.body_table_cell_list_numbering(TABLE, at, p).ok(),
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

fn run(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    cell: (usize, usize),
    operation: &Operation,
) -> TestResult {
    Case::new(name, Format::Pages, tier, evidence, input)
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
                .with(
                    6_218,
                    "TST.RichTextPayloadArchive",
                    decode_as::<litchi_iwa_protos::tst::RichTextPayloadArchive>,
                )
        })
        .host(|source| host_route(source, cell, operation))
        .leaf(|source| leaf_route(source, cell, operation))
        .read("host reads", move |bytes| host_reads(bytes, cell))
        .read("leaf reads", move |bytes| leaf_reads(bytes, cell))
        .run()?;
    Ok(())
}

fn with(input: &Fixture, cell: (usize, usize), operation: &Operation) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_route(&input.bytes, cell, operation)?,
    })
}

/// The write targets of one fixture: its first table's text and empty
/// cells the host reads lists of (up to `limit`), first row and column
/// included.
fn targets(input: &Fixture, limit: usize) -> TestResult<Vec<(usize, usize)>> {
    let (rows, columns) = (8, 6);
    let editor = pages::Draft::from_bytes(&input.bytes)?;
    let id = host_first_model(&input.bytes)?;
    let mut found = Vec::new();
    for r in 0..rows.min(8) {
        for c in 0..columns.min(6) {
            if editor.table_cell_paragraph_list(id, r, c).is_ok() {
                found.push((r, c));
            }
        }
    }
    // Prefer a spread: the first, the last and one in between.
    let mut chosen = Vec::new();
    for index in [0, found.len() / 2, found.len().saturating_sub(1)] {
        if let Some(cell) = found.get(index)
            && !chosen.contains(cell)
            && chosen.len() < limit
        {
            chosen.push(*cell);
        }
    }
    Ok(chosen)
}

/// Bullet and numbered presets on plain cells (promotion), then a reset to
/// `None` and per-paragraph placements on the host's output.
#[test]
fn pages_body_table_cell_list_preset_parity() -> TestResult {
    let mut cases = 0;
    let mut failures = Vec::new();
    for path in WRITE_FIXTURES {
        let input = fixture(path)?;
        for cell in targets(&input, 1)? {
            for preset in [ParagraphList::Bullet, ParagraphList::Numbered] {
                let mut check = |input: &Fixture, operation: Operation| {
                    cases += 1;
                    if let Err(error) = run(
                        "pages_body_table_cell_list_preset_parity",
                        Tier::B,
                        Evidence::Queued,
                        input,
                        cell,
                        &operation,
                    ) {
                        failures.push(format!("{path} {cell:?} {operation:?}: {error}"));
                    }
                };
                check(&input, Operation::SetList(preset));
                check(
                    &input,
                    Operation::SetLists(vec![ParagraphListPlacement::new(
                        TextPosition::ZERO,
                        preset,
                    )]),
                );
                let Ok(base) = with(&input, cell, &Operation::SetList(preset)) else {
                    continue;
                };
                check(&base, Operation::ResetList);
                check(&base, Operation::SetList(ParagraphList::None));
                let other = if preset == ParagraphList::Bullet {
                    ParagraphList::Numbered
                } else {
                    ParagraphList::Bullet
                };
                check(&base, Operation::SetList(other));
            }
        }
    }
    eprintln!("cell list preset cases: {cases}");
    if failures.is_empty() {
        assert!(cases >= 20, "only {cases} cases ran");
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}

/// Level and numbering of a listed cell's first paragraph, and the same
/// changes on a plain cell (promotion).
#[test]
fn pages_body_table_cell_list_level_numbering_parity() -> TestResult {
    let mut cases = 0;
    let mut failures = Vec::new();
    let two = ParagraphListLevel::new(2)?;
    for path in WRITE_FIXTURES {
        let input = fixture(path)?;
        for cell in targets(&input, 1)? {
            let mut check = |input: &Fixture, operation: Operation| {
                cases += 1;
                if let Err(error) = run(
                    "pages_body_table_cell_list_level_numbering_parity",
                    Tier::A,
                    Evidence::Queued,
                    input,
                    cell,
                    &operation,
                ) {
                    failures.push(format!("{path} {cell:?} {operation:?}: {error}"));
                }
            };
            check(&input, Operation::SetLevel(TextPosition::ZERO, two));
            let Ok(base) = with(&input, cell, &Operation::SetList(ParagraphList::Numbered)) else {
                continue;
            };
            check(&base, Operation::SetLevel(TextPosition::ZERO, two));
            check(
                &base,
                Operation::SetNumbering(
                    TextPosition::ZERO,
                    ParagraphListNumbering::StartAt(ParagraphListStart::new(3)?),
                ),
            );
            let leveled = with(&base, cell, &Operation::SetLevel(TextPosition::ZERO, two))?;
            check(&leveled, Operation::ResetLevel(TextPosition::ZERO));
        }
    }
    eprintln!("cell list level/numbering cases: {cases}");
    if failures.is_empty() {
        assert!(cases >= 8, "only {cases} cases ran");
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}

/// Each list property set on a listed cell (a new paragraph variation).
#[test]
fn pages_body_table_cell_list_property_set_parity() -> TestResult {
    let _ = Evidence::Queued;
    property_cases(
        "pages_body_table_cell_list_property_set_parity",
        Tier::B,
        |base, _cell, property| {
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

/// A second value of each property on the cell's own variation (rewritten
/// in place).
#[test]
fn pages_body_table_cell_list_property_rewrite_parity() -> TestResult {
    let _ = Evidence::Queued;
    property_cases(
        "pages_body_table_cell_list_property_rewrite_parity",
        Tier::A,
        |base, cell, property| {
            let set = with(
                base,
                cell,
                &Operation::Set(TextPosition::ZERO, property.clone()),
            )?;
            Ok((
                set,
                Operation::Set(TextPosition::ZERO, list_second(property)?),
            ))
        },
    )
}

/// Each property reset on the cell's own variation (collapsed onto its
/// parent).
#[test]
fn pages_body_table_cell_list_property_reset_parity() -> TestResult {
    let _ = Evidence::Queued;
    property_cases(
        "pages_body_table_cell_list_property_reset_parity",
        Tier::B,
        |base, cell, property| {
            let set = with(
                base,
                cell,
                &Operation::Set(TextPosition::ZERO, property.clone()),
            )?;
            Ok((set, Operation::Reset(TextPosition::ZERO, property.kind())))
        },
    )
}

fn property_cases(
    name: &'static str,
    tier: Tier,
    prepare: impl Fn(&Fixture, (usize, usize), &Property) -> TestResult<(Fixture, Operation)>,
) -> TestResult {
    let mut cases = 0;
    let mut failures = Vec::new();
    for path in WRITE_FIXTURES {
        let input = fixture(path)?;
        for cell in targets(&input, 1)? {
            for property in samples()? {
                let base = with(
                    &input,
                    cell,
                    &Operation::SetList(preset_for(property.kind())),
                )?;
                let (input, operation) = prepare(&base, cell, &property)?;
                cases += 1;
                if let Err(error) = run(name, tier, Evidence::Queued, &input, cell, &operation) {
                    failures.push(format!("{path} {cell:?} {operation:?}: {error}"));
                }
            }
        }
    }
    eprintln!("{name}: {cases} cases");
    if failures.is_empty() {
        assert!(cases >= 14, "only {cases} cases ran");
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}

/// Every list read agrees on corpus cells and on listed cells.
#[test]
fn pages_body_table_cell_list_read_parity() -> TestResult {
    let (_tier, _evidence) = (Tier::A, Evidence::Queued);
    let mut compared = 0_usize;
    for path in crate::pages_cell_text_cases::CORPUS {
        let input = fixture(path)?;
        let Ok(cells) = targets(&input, 3) else {
            continue;
        };
        for cell in cells {
            assert_eq!(
                leaf_reads(&input.bytes, cell)?,
                host_reads(&input.bytes, cell)?,
                "{path} {cell:?}"
            );
            compared += 1;
        }
    }
    for path in WRITE_FIXTURES {
        let input = fixture(path)?;
        for cell in targets(&input, 1)? {
            for preset in [ParagraphList::Bullet, ParagraphList::Numbered] {
                let base = with(&input, cell, &Operation::SetList(preset))?;
                assert_eq!(
                    leaf_reads(&base.bytes, cell)?,
                    host_reads(&base.bytes, cell)?,
                    "{path} {cell:?} {preset:?}"
                );
                compared += 1;
                for property in samples()? {
                    if preset_for(property.kind()) != preset {
                        continue;
                    }
                    let set = with(&base, cell, &Operation::Set(TextPosition::ZERO, property))?;
                    assert_eq!(
                        leaf_reads(&set.bytes, cell)?,
                        host_reads(&set.bytes, cell)?,
                        "{path} {cell:?} {preset:?}"
                    );
                    compared += 1;
                }
            }
        }
    }
    eprintln!("cell list reads compared: {compared}");
    assert!(compared >= 40, "only {compared} reads compared");
    Ok(())
}

/// A two-paragraph plain cell (written through the leaf's cell writer):
/// per-paragraph placements, then level, numbering and a property of the
/// second paragraph.
#[test]
fn pages_body_table_cell_list_paragraph_parity() -> TestResult {
    let source = fixture(WRITE_FIXTURES[0])?;
    let package = Package::from_bytes(&source.bytes)?;
    let mut edit = package.edit_body_table_values(TABLE)?;
    edit.set(
        1,
        1,
        litchi_pages::table::values::Value::Text("One\nTwo".into()),
    )?;
    let written = edit.commit()?.into_package();
    let input = Fixture {
        path: source.path,
        bytes: leaf_bytes!(written),
    };
    let cell = (1, 1);
    let second = TextPosition::from_utf16_code_units(4);
    let mut failures = Vec::new();
    let mut check = |input: &Fixture, tier: Tier, operation: Operation| {
        if let Err(error) = run(
            "pages_body_table_cell_list_paragraph_parity",
            tier,
            Evidence::Queued,
            input,
            cell,
            &operation,
        ) {
            failures.push(format!("{operation:?}: {error}"));
        }
    };
    let mixed = vec![
        ParagraphListPlacement::new(TextPosition::ZERO, ParagraphList::Bullet),
        ParagraphListPlacement::new(second, ParagraphList::Numbered),
    ];
    check(&input, Tier::B, Operation::SetLists(mixed.clone()));
    check(&input, Tier::B, Operation::SetList(ParagraphList::Numbered));
    let listed = with(&input, cell, &Operation::SetLists(mixed))?;
    check(
        &listed,
        Tier::A,
        Operation::SetLevel(second, ParagraphListLevel::new(1)?),
    );
    check(
        &listed,
        Tier::A,
        Operation::SetNumbering(
            second,
            ParagraphListNumbering::StartAt(ParagraphListStart::new(5)?),
        ),
    );
    for property in samples()? {
        if preset_for(property.kind()) == ParagraphList::Numbered {
            check(&listed, Tier::B, Operation::Set(second, property));
        }
    }
    // One preset for paragraphs of different lists: both routes refuse.
    let whole = Operation::SetList(ParagraphList::None);
    assert!(host_route(&listed.bytes, cell, &whole).is_err());
    assert!(matches!(
        leaf_route(&listed.bytes, cell, &whole)
            .err()
            .and_then(|error| error
                .downcast::<litchi_pages::table::cell_list::Error>()
                .ok())
            .map(|error| *error),
        Some(litchi_pages::table::cell_list::Error::Refused(
            litchi_pages::table::cell_list::Refusal::NotUniform
        ))
    ));
    check(
        &listed,
        Tier::B,
        Operation::SetLists(vec![ParagraphListPlacement::new(
            TextPosition::ZERO,
            ParagraphList::None,
        )]),
    );
    assert_eq!(
        leaf_reads(&listed.bytes, cell)?,
        host_reads(&listed.bytes, cell)?
    );
    if failures.is_empty() {
        Ok(())
    } else {
        Err(failures.join("\n").into())
    }
}

/// A plain cell has no list properties, and a list change needs a
/// paragraph start: both routes refuse.
#[test]
fn pages_body_table_cell_list_refusal_parity() -> TestResult {
    let (_tier, _evidence) = (Tier::A, Evidence::Queued);
    let input = fixture(WRITE_FIXTURES[0])?;
    let cell = (1, 1);
    let bullet = samples()?
        .into_iter()
        .find(|property| property.kind() == PropertyKind::Bullet)
        .ok_or("no bullet sample")?;
    for operation in [
        Operation::Set(TextPosition::ZERO, bullet),
        Operation::SetLevel(
            TextPosition::from_utf16_code_units(1),
            ParagraphListLevel::new(1)?,
        ),
    ] {
        assert!(
            host_route(&input.bytes, cell, &operation).is_err(),
            "host {operation:?}"
        );
        assert!(
            leaf_route(&input.bytes, cell, &operation).is_err(),
            "leaf {operation:?}"
        );
    }
    Ok(())
}

/// Callgrind probe (record 0412): a bullet list on a plain cell (promotion),
/// host against leaf in ABBA order.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_cell_list_cost_abba() -> TestResult {
    let input = fixture(WRITE_FIXTURES[0])?;
    let cell = *targets(&input, 1)?.first().ok_or("no list target")?;
    let operation = Operation::SetList(ParagraphList::Bullet);
    crate::pages_cost_probes::abba(
        &input,
        |source| host_route(source, cell, &operation),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_body_table_cell_lists(TABLE)?
                .set_list(At::new(cell.0 as u32, cell.1 as u32), ParagraphList::Bullet)?
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}
