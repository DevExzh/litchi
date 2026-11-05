#![cfg(feature = "rtf")]
use litchi_core::{
    Position,
    selector::{Path, Scope},
};
use litchi_ops::error::Source;
use litchi_ops::query::Query;
use litchi_ops::transaction::{Error, LimitKind, SemanticLimits};
use litchi_ops::{BatchItem, BatchMode, Code, File, open_bytes, props};
use litchi_rtf::edit::RunTextReplacement;
const INPUT: &[u8]=br"{\rtf1\ansi{\fonttbl{\f0\fnil Example Sans;}}{\colortbl;\red255\green0\blue0;}{\info{\title Keep title}}{\header Keep header}\f0\cf1\qr Plain{\i Italic}\b Bold\par\pard\qc Tail}";
fn query(value: &str) -> Query<'_> {
    Query::parse(value)
        .unwrap()
        .with_scope(Scope::Path(Box::new(Path::parse("/body").unwrap())))
}
fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap()
        .into_owned()
}
fn property(file: &File, path: &str, key: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|p| p.key() == key)
        .unwrap()
        .to_string()
}
fn command(path: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("remove");
    item.path = Some(path.into());
    item
}

#[test]
fn direct_run_removal_matches_native_and_reopens_with_exact_inverse() {
    for (path, p, r) in [
        ("/body/p[1]/r[1]", 0, 0),
        ("/body/p[1]/run[2]", 0, 1),
        ("/body/p[1]/r[last()]", 0, 2),
        ("/body/p[last()]/r[1]", 1, 0),
    ] {
        let source = open_bytes(INPUT).unwrap();
        let commit = source.edit(|edit| edit.remove(path)).unwrap();
        let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
        let mut edit = native.edit();
        edit.replace_body_run_texts(&[RunTextReplacement::new(
            Position::new(p),
            Position::new(r),
            "",
        )])
        .unwrap();
        let expected = edit.commit().unwrap().package().to_bytes().unwrap();
        assert_eq!(commit.file().to_bytes().unwrap(), expected);
        assert_eq!(commit.diagnostics().operations(), 1);
        assert_eq!(commit.file().query("p").unwrap().count(), 2);
        assert_eq!(commit.file().query("r").unwrap().count(), 3);
        assert_eq!(
            commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
            expected
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            INPUT
        );
        assert_eq!(open_bytes(expected).unwrap().query("r").unwrap().count(), 3);
        assert_eq!(source.to_bytes().unwrap(), INPUT);
    }
}

#[test]
fn query_union_resolves_original_runs_once_and_retains_empty_paragraphs() {
    let source = open_bytes(INPUT).unwrap();
    let changed = source
        .remove(query("r[text=Tail], r[italic=true], r[text=Tail]"))
        .unwrap();
    assert_eq!(text(&changed, "/body/p[1]"), "PlainBold");
    assert_eq!(text(&changed, "/body/p[2]"), "");
    assert_eq!(changed.query("r").unwrap().count(), 2);
    assert_eq!(property(&changed, "/body/p[1]/r[2]", "bold"), "true");
    assert_eq!(property(&changed, "/body/p[1]/r[2]", "color"), "#FF0000");
    assert_eq!(property(&changed, "/body/p[1]", "align"), "right");
    assert_eq!(property(&changed, "/body/p[2]", "align"), "center");
    let all = source.remove(query("r")).unwrap();
    assert_eq!(all.query("p").unwrap().count(), 2);
    assert_eq!(all.query("r").unwrap().count(), 0);
    assert_eq!(text(&all, "/body/p[1]"), "");
    assert_eq!(text(&all, "/body/p[2]"), "");
    for marker in [
        b"Keep title".as_slice(),
        b"Keep header".as_slice(),
        b"Example Sans".as_slice(),
    ] {
        assert!(
            all.to_bytes()
                .unwrap()
                .windows(marker.len())
                .any(|window| window == marker)
        );
    }
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn empty_paragraphs_and_terminal_boundaries_survive_removing_every_run() {
    for input in [
        br"{\rtf1 Text}".as_slice(),
        br"{\rtf1 One\par Two\par}".as_slice(),
        br"{\rtf1 One\par\par Three}".as_slice(),
        br"{\rtf1\par One\par\par}".as_slice(),
    ] {
        let source = open_bytes(input).unwrap();
        let count = source.query("p").unwrap().count();
        let changed = source.remove(query("r")).unwrap();
        assert_eq!(changed.query("p").unwrap().count(), count);
        assert_eq!(changed.query("r").unwrap().count(), 0);
        assert_eq!(source.to_bytes().unwrap(), input);
    }
}

#[test]
fn run_removal_limits_include_retained_before_text_prior_to_native_allocation() {
    let bytes = format!(r"{{\rtf1 {}}}", "x".repeat(4096)).into_bytes();
    let source = open_bytes(bytes.as_slice()).unwrap();
    let limits = SemanticLimits::builder()
        .max_staged_bytes(1024)
        .unwrap()
        .build()
        .unwrap();
    for remove in [false, true] {
        let error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |edit| {
                    if remove {
                        edit.remove("/body/p[1]/r[1]")
                    } else {
                        edit.set("/body/p[1]/r[1]", props! {"text"=>"x"})
                    }
                },
            )
            .unwrap_err();
        assert!(matches!(
            error.leaf(),
            Source::Transaction(Error::Limit {
                kind: LimitKind::StagedBytes,
                ..
            })
        ));
    }
    // Format-only operations do not retain a cloned before-text payload.
    source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.set("/body/p[1]/r[1]", props! {"bold"=>true}),
        )
        .unwrap();
    assert_eq!(source.to_bytes().unwrap(), bytes);
    let source = open_bytes(INPUT).unwrap();
    for (limits, kind) in [
        (
            SemanticLimits::builder()
                .max_expanded_targets(1)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::ExpandedTargets,
        ),
        (
            SemanticLimits::builder()
                .max_operations(1)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::Operations,
        ),
        (
            SemanticLimits::builder()
                .max_staged_bytes(1)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::StagedBytes,
        ),
        (
            SemanticLimits::builder()
                .max_snapshot_bytes(1)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::SnapshotBytes,
        ),
    ] {
        let error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |edit| edit.remove(query("r")),
            )
            .unwrap_err();
        assert!(
            matches!(error.leaf(),Source::Transaction(Error::Limit{kind:actual,..}) if *actual==kind)
        );
    }
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn mixed_targets_misses_and_dependent_sources_refuse_atomically() {
    let source = open_bytes(INPUT).unwrap();
    assert_eq!(
        source.remove(query("p, r")).unwrap_err().code(),
        Code::UnsupportedType
    );
    assert_eq!(
        source.remove(query("r[text=missing]")).unwrap_err().code(),
        Code::ZeroMatches
    );
    assert_eq!(
        source.remove("r").unwrap_err().code(),
        Code::BareSelectorRejected
    );
    assert!(source.remove("/body/p[1]/r[99]").is_err());
    assert_eq!(source.to_bytes().unwrap(), INPUT);
    for input in [
        br"{\rtf1 Text\line tail}".as_slice(),
        br"{\rtf1{\stylesheet{\s1 Styled;}}\s1 Text}".as_slice(),
    ] {
        let source = open_bytes(input).unwrap();
        assert_eq!(
            source.remove("/body/p[1]/r[1]").unwrap_err().code(),
            Code::UnsupportedCapability
        );
        assert_eq!(source.to_bytes().unwrap(), input);
    }
    let input = br"{\rtf1\readprot Text}";
    let source = open_bytes(input).unwrap();
    assert_eq!(
        source.remove("/body/p[1]/r[1]").unwrap_err().code(),
        Code::DocumentProtected
    );
    for lzfu in [false, true] {
        let bytes = litchi_rtf::transport::compress(
            INPUT,
            lzfu,
            litchi_rtf::transport::CompressionLimits::default(),
        )
        .unwrap();
        let source = open_bytes(bytes.as_slice()).unwrap();
        assert_eq!(
            source.remove(query("r")).unwrap_err().code(),
            Code::UnsupportedCapability
        );
        assert_eq!(source.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn sequential_batch_removals_resolve_current_positions_and_rollback_exactly() {
    let source = open_bytes(INPUT).unwrap();
    let commit = source
        .edit(|edit| {
            edit.remove("/body/p[1]/r[1]")?;
            edit.remove("/body/p[1]/r[1]")
        })
        .unwrap();
    assert_eq!(text(commit.file(), "/body/p[1]"), "Bold");
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        INPUT
    );
    assert!(
        source
            .batch(
                [command("/body/p[1]/r[2]"), command("/body/p[1]/r[99]")],
                BatchMode::Atomic
            )
            .is_err()
    );
    let best = source
        .batch(
            [command("/body/p[1]/r[2]"), command("/body/p[1]/r[99]")],
            BatchMode::BestEffort,
        )
        .unwrap();
    assert_eq!(text(best.file(), "/body/p[1]"), "PlainBold");
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn unicode_fragments_are_native_positional_targets_not_whole_words() {
    let input = br"{\rtf1\ansi\uc1 A\u233?B{\i C\u38634?D}\tab end}";
    let native = litchi_rtf::Package::from_bytes(input).unwrap();
    let source = open_bytes(input).unwrap();
    for (index, run) in native
        .body()
        .paragraphs()
        .next()
        .unwrap()
        .runs()
        .enumerate()
    {
        let mut edit = native.edit();
        edit.replace_run_text(Position::new(0), Position::new(index), "")
            .unwrap();
        let changed = source
            .remove(&format!("/body/p[1]/r[{}]", index + 1))
            .unwrap();
        assert_eq!(
            changed.to_bytes().unwrap(),
            edit.commit().unwrap().package().to_bytes().unwrap(),
            "{:?}",
            run.text()
        );
    }
    assert_eq!(source.to_bytes().unwrap(), input);
}

#[test]
fn run_removal_advertises_native_source_gap_and_preserves_atomic_save_origin() {
    let source = open_bytes(INPUT).unwrap();
    assert!(
        source
            .capabilities()
            .element("r")
            .unwrap()
            .operations()
            .remove
    );
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("run.rtf");
    std::fs::write(&path, INPUT).unwrap();
    litchi_ops::open(&path)
        .unwrap()
        .remove("/body/p[1]/r[2]")
        .unwrap()
        .save()
        .unwrap();
    assert_eq!(
        text(&litchi_ops::open(&path).unwrap(), "/body/p[1]"),
        "PlainBold"
    );
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
}

#[test]
fn opaque_bytes_and_shifted_anchors_survive_run_removal() {
    for input in [
        br"{\rtf1{\*\vendor opaque}Text}".as_slice(),
        br"{\rtf1 Text{\*\vendor opaque}}".as_slice(),
        br"{\rtf1{\*\vendor before}{\i One}{\*\vendor between}{\b Two}{\*\vendor after}}"
            .as_slice(),
    ] {
        let source = open_bytes(input).unwrap();
        let commit = source.edit(|edit| edit.remove(query("r"))).unwrap();
        assert_eq!(commit.file().query("p").unwrap().count(), 1);
        assert_eq!(commit.file().query("r").unwrap().count(), 0);
        let bytes = commit.file().to_bytes().unwrap();
        let native_source = litchi_rtf::Package::from_bytes(input).unwrap();
        let native_target = litchi_rtf::Package::from_bytes(&bytes).unwrap();
        let opaque = |package: &litchi_rtf::Package| {
            package
                .opaque()
                .iter()
                .map(|node| node.source().to_vec())
                .collect::<Vec<_>>()
        };
        assert_eq!(opaque(&native_source), opaque(&native_target));
        assert_eq!(
            commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
            bytes
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .unwrap()
                .to_bytes()
                .unwrap(),
            input
        );
        assert_eq!(source.to_bytes().unwrap(), input);
    }
}
