//! Keynote slide-table hidden axes of wave W1-G (record 0379): the host's
//! `slide_table_hidden_axes` / `set_slide_table_hidden_axes` against the
//! leaf's `edit_slide_table_hidden_axes` over the shared engine
//! (`litchi_iwa_table::hidden_axes`).

use litchi_iwa::keynote;
use litchi_iwa_common::table::axis::{AxisIndex, HiddenAxes};
use litchi_keynote::Package;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

/// The two native decks whose first slide table both routes edit. The
/// engine refuses `slide-table-read-native.key` (record 0379 §5).
const FIXTURES: [&str; 2] = [
    "iwork/keynote/table-discovery.key",
    "iwork/keynote/slide-table-merges-native.key",
];
const SLIDE: usize = 0;
const TABLE: usize = 0;

fn host_model(editor: &keynote::Draft) -> TestResult<u64> {
    Ok(editor
        .slide_tables(SLIDE)?
        .first()
        .ok_or("no slide table")?
        .model_object_id)
}

fn host_read(bytes: &[u8]) -> TestResult<HiddenAxes> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    Ok(editor.slide_table_hidden_axes(SLIDE, model)?)
}

fn leaf_read(bytes: &[u8]) -> TestResult<HiddenAxes> {
    Ok(Package::from_bytes(bytes)?.slide_table_hidden_axes(SLIDE, TABLE)?)
}

fn host_set(source: &[u8], axes: &HiddenAxes) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let model = host_model(&editor)?;
    editor.set_slide_table_hidden_axes(SLIDE, model, axes)?;
    Ok(editor.to_bytes()?)
}

fn leaf_set(source: &[u8], axes: &HiddenAxes) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let commit = package
        .edit_slide_table_hidden_axes(SLIDE, TABLE)?
        .set(axes.clone())
        .commit()?;
    let restored = commit
        .package()
        .apply_slide_table_hidden_axes(&commit.patch().inverse())?;
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
    Case::new(name, Format::Keynote, tier, evidence, input)
        .host(|source| host_set(source, axes))
        .leaf(|source| leaf_set(source, axes))
        .read("host slide_table_hidden_axes", host_read)
        .read("leaf slide_table_hidden_axes", leaf_read)
        .run()?;
    Ok(())
}

/// Hide a row and a column of the first slide table, which hides nothing
/// yet (the host's owner-creation path).
#[test]
fn keynote_slide_table_hidden_axes_set_parity() -> TestResult {
    for path in FIXTURES {
        let input = fixture(path)?;
        set_case(
            "keynote_slide_table_hidden_axes_set_parity",
            Tier::B,
            Evidence::Queued,
            &input,
            &axes()?,
        )?;
    }
    Ok(())
}

/// Clearing the hidden axes the host stored differs (record 0379 §5): the
/// host drops the cleared row and column state records, the shared engine
/// keeps them with no user-hidden flag, as the Pages owner's native cleared
/// documents store them. Both outputs read as unhidden on both routes.
#[test]
fn keynote_slide_table_hidden_axes_clear_divergence_is_reported() -> TestResult {
    for path in FIXTURES {
        let input = fixture(path)?;
        let hidden = Fixture {
            path: input.path,
            bytes: host_set(&input.bytes, &axes()?)?,
        };
        let result = set_case(
            "keynote_slide_table_hidden_axes_clear_divergence_is_reported",
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
fn keynote_slide_table_hidden_axes_read_parity() -> TestResult {
    for path in FIXTURES {
        let input = fixture(path)?;
        let hidden = Fixture {
            path: input.path,
            bytes: host_set(&input.bytes, &axes()?)?,
        };
        for source in [&input, &hidden] {
            Case::new(
                "keynote_slide_table_hidden_axes_read_parity",
                Format::Keynote,
                Tier::A,
                Evidence::Queued,
                source,
            )
            .read_only()
            .host(|bytes| Ok(keynote::Draft::from_bytes(bytes)?.to_bytes()?))
            .leaf(|bytes| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(bytes)?))))
            .read("host slide_table_hidden_axes", host_read)
            .read("leaf slide_table_hidden_axes", leaf_read)
            .run()?;
            assert_eq!(host_read(&source.bytes)?, leaf_read(&source.bytes)?);
        }
    }
    Ok(())
}

/// Callgrind probe (record 0379): hide a row and a column of the first
/// slide table of `table-discovery.key`, host against leaf, ABBA
/// (`--toggle-collect=*probe_host*` / `*probe_leaf*`).
#[test]
#[ignore = "callgrind probe; run explicitly"]
fn keynote_hidden_axes_cost_abba() -> TestResult {
    let input = fixture(FIXTURES[0])?;
    let axes = axes()?;
    crate::pages_cost_probes::abba(
        &input,
        |source| host_set(source, &axes),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_slide_table_hidden_axes(SLIDE, TABLE)?
                .set(axes.clone())
                .commit()?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}
