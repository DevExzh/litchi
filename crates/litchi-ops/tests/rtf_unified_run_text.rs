#![cfg(feature = "rtf")]

use litchi_core::{
    Position,
    selector::{Path, Scope},
};
use litchi_ops::error::Source;
use litchi_ops::query::Query;
use litchi_ops::transaction::{Error, LimitKind, SemanticLimits};
use litchi_ops::{BatchItem, BatchMode, Code, File, Props, open_bytes, props};
use litchi_rtf::edit::{RunFormatPatch, RunTextReplacement};

const INPUT: &[u8] = br"{\rtf1\ansi{\fonttbl{\f0\fnil Example Sans;}}{\colortbl;\red255\green0\blue0;}{\info{\title Keep title}}{\header Keep header}\f0\cf1\qr Plain{\i Italic}\b Bold\par\pard\qc Tail}";
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
fn item(path: &'static str, props: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    item
}

#[test]
fn run_text_and_format_publish_once_with_exact_native_bytes_and_inverse() {
    let source = open_bytes(INPUT).unwrap();
    let value = "new \\ {text}\tvalue";
    let limits = SemanticLimits::builder()
        .max_operations(1)
        .unwrap()
        .build()
        .unwrap();
    let changed = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| {
                edit.set(
                    "/body/p[1]/r[2]",
                    props! {"text"=>value,"bold"=>true,"size"=>"14pt","underline"=>"double"},
                )
            },
        )
        .unwrap();
    let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
    let mut edit = native.edit();
    edit.replace_body_run_texts(&[RunTextReplacement::new(
        Position::new(0),
        Position::new(1),
        value,
    )
    .with_format(
        RunFormatPatch::new()
            .with_bold(true)
            .with_size(std::num::NonZeroU16::new(28).unwrap())
            .with_underline(litchi_rtf::text::Underline::Double),
    )])
    .unwrap();
    let bytes = changed.file().to_bytes().unwrap();
    assert_eq!(bytes, edit.commit().unwrap().package().to_bytes().unwrap());
    assert_eq!(text(changed.file(), "/body/p[1]/r[2]"), value);
    assert_eq!(
        text(changed.file(), "/body/p[1]"),
        format!("Plain{value}Bold")
    );
    for (key, expected) in [
        ("bold", "true"),
        ("italic", "true"),
        ("size", "14pt"),
        ("underline", "double"),
        ("font", "Example Sans"),
        ("color", "#FF0000"),
    ] {
        assert_eq!(property(changed.file(), "/body/p[1]/r[2]", key), expected);
    }
    assert_eq!(text(changed.file(), "/body/p[1]/r[1]"), "Plain");
    assert_eq!(text(changed.file(), "/body/p[1]/r[3]"), "Bold");
    assert_eq!(property(changed.file(), "/body/p[1]", "align"), "right");
    assert!(
        bytes
            .windows(b"Keep title".len())
            .any(|b| b == b"Keep title")
    );
    assert!(
        bytes
            .windows(b"Keep header".len())
            .any(|b| b == b"Keep header")
    );
    assert_eq!(
        changed.patch().apply(&source).unwrap().to_bytes().unwrap(),
        bytes
    );
    assert_eq!(
        changed
            .patch()
            .inverse()
            .apply(changed.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        INPUT
    );
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn fanout_uses_original_positions_and_empty_text_removes_only_selected_runs() {
    let source = open_bytes(INPUT).unwrap();
    let changed = source
        .set(
            query("r[italic=true], r[text=Tail]"),
            props! {"text"=>"new"},
        )
        .unwrap();
    assert_eq!(text(&changed, "/body/p[1]"), "PlainnewBold");
    assert_eq!(text(&changed, "/body/p[2]"), "new");
    assert_eq!(property(&changed, "/body/p[1]/r[2]", "italic"), "true");
    let empty = source
        .set(query("r[italic=true], r[text=Tail]"), props! {"text"=>""})
        .unwrap();
    assert_eq!(text(&empty, "/body/p[1]"), "PlainBold");
    assert_eq!(text(&empty, "/body/p[2]"), "");
    assert_eq!(empty.query("p").unwrap().count(), 2);
    assert_eq!(empty.query("r").unwrap().count(), 2);
    assert_eq!(property(&empty, "/body/p[1]/r[2]", "bold"), "true");
    let all_empty = source.set(query("r"), props! {"text"=>""}).unwrap();
    assert_eq!(all_empty.query("p").unwrap().count(), 2);
    assert_eq!(all_empty.query("r").unwrap().count(), 0);
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn replacement_respects_unicode_fallback_state_and_group_boundaries() {
    for input in [
        br"{\rtf1\ansi\uc0 \u38634{\i inner} tail}".as_slice(),
        br"{\rtf1\ansi\uc2 \u233??{\i inner} tail}".as_slice(),
        br"{\rtf1\ansi\uc1 \u-10179?\u-8704?{\i inner} tail}".as_slice(),
        br"{\rtf1\ansi first{same}last}".as_slice(),
        br"{\rtf1\ansi first{\b same}last}".as_slice(),
    ] {
        let source = open_bytes(input).unwrap();
        let count = source.query("r").unwrap().count();
        let value = "雪 😀 \\ {}\tX";
        let changed = source
            .set(query("r"), props! {"text"=>value,"outline"=>true})
            .unwrap();
        // The native parser exposes Unicode sequences as separate run fragments.
        // Semantic text and formatting coverage, rather than fragment count,
        // define preservation after a text replacement.
        assert_eq!(text(&changed, "/body/p[1]"), value.repeat(count));
        for node in changed.query("r").unwrap() {
            let node = node.unwrap();
            let props = node.view().format().unwrap().collect::<Vec<_>>();
            assert_eq!(
                props
                    .iter()
                    .find(|p| p.key() == "outline")
                    .unwrap()
                    .to_string(),
                "true"
            );
        }
        assert_eq!(source.to_bytes().unwrap(), input);
    }
}

#[test]
fn run_text_bounds_reject_before_publication_and_compound_cost_is_one_target() {
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
                |edit| edit.set(query("r"), props! {"text"=>"replacement","bold"=>true}),
            )
            .unwrap_err();
        assert!(
            matches!(error.leaf(), Source::Transaction(Error::Limit{kind:actual,..}) if *actual==kind)
        );
    }
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn invalid_values_and_read_only_resources_do_not_publish_partial_text() {
    let source = open_bytes(INPUT).unwrap();
    let error = source
        .set("/body/p[1]/r[2]", props! {"text"=>true})
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    assert_eq!(error.property(), Some("text"));
    assert_eq!(error.format(), Some(litchi_core::FileFormat::Rtf));
    for props in [
        Props::new().with("text", "new").with("size", "12.25pt"),
        Props::new().with("text", "new").with("font", "Other"),
        Props::new().with("text", "new").with("color", "#008800"),
    ] {
        assert!(source.set("/body/p[1]/r[2]", props).is_err());
    }
    assert_eq!(
        source
            .set(query("r[text=missing]"), props! {"text"=>"new"})
            .unwrap_err()
            .code(),
        Code::ZeroMatches
    );
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn atomic_and_best_effort_batches_keep_run_text_history_and_query_visibility() {
    let source = open_bytes(INPUT).unwrap();
    let commands = || {
        [
            item("/body/p[1]/r[2]", Props::new().with("text", "new")),
            item("/body/p[1]/r[3]", Props::new().with("text", true)),
        ]
    };
    assert!(source.batch(commands(), BatchMode::Atomic).is_err());
    assert_eq!(source.to_bytes().unwrap(), INPUT);
    let changed = source.batch(commands(), BatchMode::BestEffort).unwrap();
    assert_eq!(text(changed.file(), "/body/p[1]/r[2]"), "new");
    let changed = source
        .edit(|edit| {
            edit.set("/body/p[1]/r[2]", props! {"text"=>"new"})?;
            edit.set(query("r[text=new]"), props! {"bold"=>true})
        })
        .unwrap();
    assert_eq!(property(changed.file(), "/body/p[1]/r[2]", "bold"), "true");
    assert_eq!(
        changed
            .patch()
            .inverse()
            .apply(changed.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        INPUT
    );
}

#[test]
fn unchanged_text_preserves_protected_source_and_atomic_save_origin() {
    let input = br"{\rtf1\ansi\readprot Text}";
    let source = open_bytes(input).unwrap();
    assert_eq!(
        source
            .set("/body/p[1]/r[1]", props! {"text"=>"Text"})
            .unwrap()
            .to_bytes()
            .unwrap(),
        input
    );
    assert_eq!(
        source
            .set("/body/p[1]/r[1]", props! {"text"=>"changed"})
            .unwrap_err()
            .code(),
        Code::DocumentProtected
    );
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("runs.rtf");
    std::fs::write(&path, INPUT).unwrap();
    litchi_ops::open(&path)
        .unwrap()
        .set("/body/p[1]/r[2]", props! {"text"=>"saved"})
        .unwrap()
        .save()
        .unwrap();
    assert_eq!(
        text(&litchi_ops::open(&path).unwrap(), "/body/p[1]/r[2]"),
        "saved"
    );
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
}

#[test]
fn structural_text_and_dependent_sources_are_explicit_atomic_refusals() {
    let source = open_bytes(INPUT).unwrap();
    for value in ["first\nsecond", "first\r\nsecond"] {
        let error = source
            .set("/body/p[1]/r[2]", props! {"text"=>value})
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert!(matches!(error.leaf(), Source::RtfEdit(_)));
    }
    // Literal CR/VT/NUL are hex-encoded native run text, not RTF break controls.
    for value in ["first\rsecond", "first\u{b}second", "first\0second"] {
        let changed = source
            .set("/body/p[1]/r[2]", props! {"text"=>value})
            .unwrap();
        assert_eq!(text(&changed, "/body/p[1]"), format!("Plain{value}Bold"));
    }
    for input in [
        br"{\rtf1\ansi First\line second}".as_slice(),
        br"{\rtf1\ansi{\stylesheet{\s1 Styled;}}\s1 Text}".as_slice(),
    ] {
        let source = open_bytes(input).unwrap();
        assert_eq!(
            source
                .set("/body/p[1]/r[1]", props! {"text"=>"new"})
                .unwrap_err()
                .code(),
            Code::UnsupportedCapability
        );
        assert_eq!(source.to_bytes().unwrap(), input);
    }
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn exact_noops_bypass_changed_source_closure_after_target_validation() {
    let plain = br"{\rtf1\ansi Text}";
    let mut inputs = vec![
        br"{\rtf1\ansi Text\line tail}".to_vec(),
        br"{\rtf1\ansi{\stylesheet{\s1 Style;}}\s1 Text}".to_vec(),
    ];
    for lzfu in [false, true] {
        inputs.push(
            litchi_rtf::transport::compress(
                plain,
                lzfu,
                litchi_rtf::transport::CompressionLimits::default(),
            )
            .unwrap(),
        );
    }
    for bytes in inputs {
        let source = open_bytes(bytes.as_slice()).unwrap();
        let changed = source
            .set("/body/p[1]/r[1]", props! {"text"=>"Text"})
            .unwrap();
        assert_eq!(changed.to_bytes().unwrap(), bytes);
        assert_eq!(
            source
                .set("/body/p[1]/r[1]", props! {"text"=>"changed"})
                .unwrap_err()
                .code(),
            Code::UnsupportedCapability
        );
        assert!(
            source
                .set("/body/p[99]/r[1]", props! {"text"=>"Text"})
                .is_err()
        );
    }
}

#[test]
fn run_text_fanout_projects_opaque_anchors_and_preserves_exact_history() {
    let input = br"{\rtf1{\*\vendor before}{\i First}{\*\vendor middle}{\b Second}{\*\vendor after}\par Tail{\*\vendor tail}}";
    for replacement in ["x", "expanded text", "é雪", ""] {
        let source = open_bytes(input).unwrap();
        let commit = source
            .edit(|edit| {
                edit.set(
                    query("r[italic=true], r[bold=true]"),
                    props! {"text"=>replacement},
                )
            })
            .unwrap();
        let bytes = commit.file().to_bytes().unwrap();
        let reopened = open_bytes(bytes.as_slice()).unwrap();
        assert_eq!(text(&reopened, "/body/p[1]"), replacement.repeat(2));
        assert_eq!(text(&reopened, "/body/p[2]"), "Tail");
        let native = litchi_rtf::Package::from_bytes(&bytes).unwrap();
        let original = litchi_rtf::Package::from_bytes(input).unwrap();
        assert_eq!(
            native
                .opaque()
                .iter()
                .map(|node| node.source())
                .collect::<Vec<_>>(),
            original
                .opaque()
                .iter()
                .map(|node| node.source())
                .collect::<Vec<_>>()
        );
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

#[test]
fn opaque_prefix_anchors_stay_fixed_when_later_run_changes() {
    let input = br"{\rtf1{\*\vendor zero}Prefix{\i Target}{\*\vendor end}\par{\*\vendor next}Tail}";
    let source = open_bytes(input).unwrap();
    let changed = source
        .set("/body/p[1]/r[2]", props! {"text"=>"é雪"})
        .unwrap();
    let bytes = changed.to_bytes().unwrap();
    let native = litchi_rtf::Package::from_bytes(&bytes).unwrap();
    let anchors = native
        .opaque()
        .iter()
        .map(|node| node.anchor())
        .collect::<Vec<_>>();
    assert_eq!(
        anchors,
        [
            litchi_rtf::opaque::Anchor::Body(0),
            litchi_rtf::opaque::Anchor::Body(11),
            litchi_rtf::opaque::Anchor::Body(12),
        ]
    );
    assert_eq!(text(&changed, "/body/p[1]"), "Prefixé雪");
    assert_eq!(text(&changed, "/body/p[2]"), "Tail");
    assert_eq!(source.to_bytes().unwrap(), input);
}
