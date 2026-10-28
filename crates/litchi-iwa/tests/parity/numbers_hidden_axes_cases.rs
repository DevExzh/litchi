//! Numbers table hidden axes of wave W1-G (record 0379): the host's
//! `table_hidden_axes` / `set_table_hidden_axes` against the leaf's
//! `edit_table_hidden_axes` over the shared engine
//! (`litchi_iwa_table::hidden_axes`).

use litchi_iwa::numbers;
use litchi_iwa_common::table::axis::{AxisIndex, HiddenAxes};
use litchi_numbers::{Package, TableSelector};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

/// Native workbooks whose first table both routes edit. The engine refuses
/// tables whose formula owner carries dependencies (record 0379 §5).
const FIXTURES: [&str; 2] = [
    "iwork/numbers/basic.numbers",
    "iwork/numbers/custom-number-native.numbers",
];

/// A host-built workbook: its table owner records the table's extents.
fn built() -> TestResult<Fixture> {
    Ok(Fixture {
        path: "builder: host 6 × 4 workbook",
        bytes: numbers::Builder::new()
            .table_dimensions(6, 4)
            .build()?
            .to_bytes()?,
    })
}

const SHEET: usize = 0;
const TABLE: usize = 0;

fn inputs() -> TestResult<Vec<Fixture>> {
    let mut inputs = FIXTURES
        .iter()
        .map(|path| fixture(path))
        .collect::<TestResult<Vec<_>>>()?;
    inputs.push(built()?);
    Ok(inputs)
}

fn host_read(bytes: &[u8]) -> TestResult<HiddenAxes> {
    Ok(numbers::Draft::from_bytes(bytes)?.table_hidden_axes(TableSelector::index(TABLE))?)
}

fn leaf_read(bytes: &[u8]) -> TestResult<HiddenAxes> {
    Ok(Package::from_bytes(bytes)?.table_hidden_axes(SHEET, TABLE)?)
}

fn host_set(source: &[u8], axes: &HiddenAxes) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    editor.set_table_hidden_axes(TableSelector::index(TABLE), axes)?;
    Ok(editor.to_bytes()?)
}

fn leaf_set(source: &[u8], axes: &HiddenAxes) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let commit = package
        .edit_table_hidden_axes(SHEET, TABLE)?
        .set(axes.clone())
        .commit()?;
    let restored = commit
        .package()
        .apply_table_hidden_axes(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn axes() -> TestResult<HiddenAxes> {
    Ok(HiddenAxes::new([AxisIndex::row(1), AxisIndex::column(0)])?)
}

fn set_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    axes: &HiddenAxes,
) -> TestResult {
    Case::new(name, Format::Numbers, tier, evidence, input)
        .host(|source| host_set(source, axes))
        .leaf(|source| leaf_set(source, axes))
        .read("host table_hidden_axes", host_read)
        .read("leaf table_hidden_axes", leaf_read)
        .run()?;
    Ok(())
}

/// A superset of [`axes`]: the replacement only hides more (an unhidden
/// axis is the clear divergence below).
fn other_axes() -> TestResult<HiddenAxes> {
    Ok(HiddenAxes::new([
        AxisIndex::row(1),
        AxisIndex::row(2),
        AxisIndex::column(0),
        AxisIndex::column(1),
    ])?)
}

/// Hide a row and a column of each native workbook's first table, then
/// replace the hidden axes the host stored (the existing-owner path) on the
/// native workbooks and on a host-built one whose owner records extents.
#[test]
fn numbers_table_hidden_axes_set_parity() -> TestResult {
    for path in FIXTURES {
        let input = fixture(path)?;
        set_case(
            "numbers_table_hidden_axes_set_parity",
            Tier::B,
            Evidence::Queued,
            &input,
            &axes()?,
        )?;
    }
    for input in inputs()? {
        let hidden = Fixture {
            path: input.path,
            bytes: host_set(&input.bytes, &axes()?)?,
        };
        set_case(
            "numbers_table_hidden_axes_set_parity",
            Tier::B,
            Evidence::Queued,
            &hidden,
            &other_axes()?,
        )?;
    }
    Ok(())
}

/// Hiding axes of a host-built table, which has no hidden-state owner,
/// differs (record 0379 §5): both routes create the two filter sets and the
/// two formula owners with the same identifiers, but the engine registers
/// the four new objects' UUIDs in the package metadata and writes empty
/// filter offsets, where the host registers none and writes one zero offset.
/// Both outputs read the same hidden axes on both routes.
#[test]
fn numbers_table_hidden_axes_creation_divergence_is_reported() -> TestResult {
    {
        let input = built()?;
        let result = set_case(
            "numbers_table_hidden_axes_creation_divergence_is_reported",
            Tier::B,
            Evidence::Queued,
            &input,
            &axes()?,
        );
        let message = result
            .err()
            .ok_or_else(|| format!("the creation routes unexpectedly agree on {}", input.path))?
            .to_string();
        assert!(message.contains("filter_offsets"), "{message}");
        let leaf = leaf_set(&input.bytes, &axes()?)?.output;
        assert_eq!(leaf_read(&leaf)?, axes()?);
        assert_eq!(host_read(&leaf)?, axes()?);
    }
    Ok(())
}

/// Clearing the hidden axes the host stored differs (record 0379 §5): the
/// host drops the cleared row and column state records, the shared engine
/// keeps them with no user-hidden flag, as the Pages owner's native cleared
/// documents store them. Both outputs read as unhidden on both routes.
#[test]
fn numbers_table_hidden_axes_clear_divergence_is_reported() -> TestResult {
    for input in inputs()? {
        let hidden = Fixture {
            path: input.path,
            bytes: host_set(&input.bytes, &axes()?)?,
        };
        let result = set_case(
            "numbers_table_hidden_axes_clear_divergence_is_reported",
            Tier::B,
            Evidence::Queued,
            &hidden,
            &HiddenAxes::empty(),
        );
        let message = result
            .err()
            .ok_or("the clear routes unexpectedly agree")?
            .to_string();
        assert!(message.contains("base_hidden_states"), "{message}");
        let leaf = leaf_set(&hidden.bytes, &HiddenAxes::empty())?.output;
        assert_eq!(leaf_read(&leaf)?, HiddenAxes::empty());
        assert_eq!(host_read(&leaf)?, HiddenAxes::empty());
    }
    Ok(())
}

/// The host and leaf reads agree on the source and on the host's output.
#[test]
fn numbers_table_hidden_axes_read_parity() -> TestResult {
    for input in inputs()? {
        let hidden = Fixture {
            path: input.path,
            bytes: host_set(&input.bytes, &axes()?)?,
        };
        for source in [&input, &hidden] {
            Case::new(
                "numbers_table_hidden_axes_read_parity",
                Format::Numbers,
                Tier::A,
                Evidence::Queued,
                source,
            )
            .read_only()
            .host(|bytes| Ok(numbers::Draft::from_bytes(bytes)?.to_bytes()?))
            .leaf(|bytes| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(bytes)?))))
            .read("host table_hidden_axes", host_read)
            .read("leaf table_hidden_axes", leaf_read)
            .run()?;
            assert_eq!(host_read(&source.bytes)?, leaf_read(&source.bytes)?);
        }
    }
    Ok(())
}

/// Callgrind probe (record 0379): hide a row and a column of the first
/// table of `custom-number-native.numbers`, host against leaf, ABBA
/// (`--toggle-collect=*probe_host*` / `*probe_leaf*`).
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn numbers_hidden_axes_cost_abba() -> TestResult {
    let input = fixture(FIXTURES[1])?;
    let axes = axes()?;
    crate::pages_cost_probes::abba(
        &input,
        |source| host_set(source, &axes),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_table_hidden_axes(SHEET, TABLE)?
                .set(axes.clone())
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}
