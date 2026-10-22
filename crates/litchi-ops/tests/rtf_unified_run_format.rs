#![cfg(feature = "rtf")]

use litchi_core::Position;
use litchi_ops::error::Source;
use litchi_ops::query::Query;
use litchi_ops::transaction::{Error, LimitKind, SemanticLimits};
use litchi_ops::{BatchItem, BatchMode, Code, File, Props, open_bytes, props};
use litchi_rtf::edit::{RunFormatPatch, RunFormatUpdate};

const INPUT: &[u8] = br"{\rtf1\ansi{\fonttbl{\f0\fnil Example Sans;}}{\colortbl;\red255\green0\blue0;}{\info{\title Keep title}}{\header Keep header}\f0\cf1\qr\sb120 Plain{\i Italic}\b Bold\par\pard\qc Tail}";
fn file() -> File {
    open_bytes(INPUT).unwrap()
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
fn query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .unwrap()
        .with_scope(litchi_ops::path::Scope::Path(Box::new(
            litchi_ops::path::Path::parse("/body").unwrap(),
        )))
}
fn all_props() -> Props<'static> {
    props! {"bold"=>true,"italic"=>true,"underline"=>"double","size"=>"14pt","strike"=>true,"dstrike"=>true,"vanish"=>true,"caps"=>true,"smallcaps"=>true,"outline"=>true}.into_iter().collect()
}
fn native_delta() -> RunFormatPatch {
    RunFormatPatch::new()
        .with_bold(true)
        .with_italic(true)
        .with_underline(litchi_rtf::text::Underline::Double)
        .with_size(std::num::NonZeroU16::new(28).unwrap())
        .with_strike(true)
        .with_double_strike(true)
        .with_hidden(true)
        .with_all_caps(true)
        .with_small_caps(true)
        .with_outline(true)
}
fn item(
    path: &'static str,
    props: impl IntoIterator<Item = (litchi_ops::Key<'static>, litchi_ops::PropValue<'static>)>,
) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props.into_iter().collect());
    item
}

#[test]
fn compound_run_format_matches_native_and_preserves_unselected_content() {
    let source = file();
    let changed = source.set("/body/p[1]/r[2]", all_props()).unwrap();
    let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
    let mut edit = native.edit();
    edit.patch_body_run_formats(&[RunFormatUpdate::new(
        Position::new(0),
        Position::new(1),
        native_delta(),
    )])
    .unwrap();
    assert_eq!(
        changed.to_bytes().unwrap(),
        edit.commit().unwrap().package().to_bytes().unwrap()
    );
    let reopened = open_bytes(changed.to_bytes().unwrap()).unwrap();
    for key in [
        "bold",
        "italic",
        "strike",
        "dstrike",
        "vanish",
        "caps",
        "smallcaps",
        "outline",
    ] {
        assert_eq!(property(&reopened, "/body/p[1]/r[2]", key), "true", "{key}");
    }
    assert_eq!(property(&reopened, "/body/p[1]/r[2]", "size"), "14pt");
    assert_eq!(
        property(&reopened, "/body/p[1]/r[2]", "underline"),
        "double"
    );
    for path in ["/body/p[1]/r[1]", "/body/p[1]/r[3]", "/body/p[2]/r[1]"] {
        for key in [
            "font",
            "color",
            "bold",
            "italic",
            "size",
            "underline",
            "strike",
            "dstrike",
            "vanish",
            "caps",
            "smallcaps",
            "outline",
        ] {
            assert_eq!(
                property(&reopened, path, key),
                property(&source, path, key),
                "{path} {key}"
            );
        }
    }
    assert_eq!(
        property(&reopened, "/body/p[1]/r[2]", "font"),
        "Example Sans"
    );
    assert_eq!(property(&reopened, "/body/p[1]/r[2]", "color"), "#FF0000");
    assert_eq!(property(&reopened, "/body/p[1]", "align"), "right");
    assert_eq!(property(&reopened, "/body/p[1]", "spaceBefore"), "6pt");
    let bytes = changed.to_bytes().unwrap();
    assert!(
        bytes
            .windows(b"Keep header".len())
            .any(|w| w == b"Keep header")
    );
    assert!(
        bytes
            .windows(b"Keep title".len())
            .any(|w| w == b"Keep title")
    );
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn query_fanout_uses_one_native_batch_and_history_has_exact_endpoints() {
    let source = file();
    let selector = query("r[italic=true], r[text=Tail]");
    let commit = source.edit(|edit| edit.set(selector, all_props())).unwrap();
    assert_eq!(commit.diagnostics().operations(), 1);
    let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
    let mut edit = native.edit();
    edit.patch_body_run_formats(&[
        RunFormatUpdate::new(Position::new(0), Position::new(1), native_delta()),
        RunFormatUpdate::new(Position::new(1), Position::new(0), native_delta()),
    ])
    .unwrap();
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        edit.commit().unwrap().package().to_bytes().unwrap()
    );
    assert_eq!(commit.file().query("r[outline=true]").unwrap().count(), 2);
    assert_eq!(
        commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
        commit.file().to_bytes().unwrap()
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
    assert_eq!(
        commit.patch().apply(&file()).unwrap().to_bytes().unwrap(),
        commit.file().to_bytes().unwrap()
    );
    let reopened = open_bytes(commit.file().to_bytes().unwrap()).unwrap();
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(&reopened)
            .unwrap()
            .to_bytes()
            .unwrap(),
        INPUT
    );
    assert!(matches!(
        commit.patch().apply(commit.file()).unwrap_err().leaf(),
        Source::Transaction(Error::PatchConflict)
    ));
}

#[test]
fn clearing_flags_preserves_runs_and_repeated_assignments_are_exact_noops() {
    let source = file().set("/body/p[1]/r[2]", all_props()).unwrap();
    let before = source.to_bytes().unwrap();
    let noop = source
        .edit(|edit| edit.set("/body/p[1]/r[2]", all_props()))
        .unwrap();
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.file().to_bytes().unwrap(), before);
    let cleared=source.set("/body/p[1]/r[2]",props! {"bold"=>false,"italic"=>false,"underline"=>"none","strike"=>false,"dstrike"=>false,"vanish"=>false,"caps"=>false,"smallcaps"=>false,"outline"=>false}).unwrap();
    for key in [
        "bold",
        "italic",
        "strike",
        "dstrike",
        "vanish",
        "caps",
        "smallcaps",
        "outline",
    ] {
        assert_eq!(property(&cleared, "/body/p[1]/r[2]", key), "false", "{key}");
    }
    assert_eq!(property(&cleared, "/body/p[1]/r[2]", "size"), "14pt");
    assert_eq!(property(&cleared, "/body/p[1]/r[2]", "underline"), "none");
    assert_eq!(
        cleared
            .get("/body/p[1]")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("PlainItalicBold")
    );
}

#[test]
fn every_underline_style_can_be_assigned_and_queried_without_collapsing() {
    for token in [
        "none",
        "single",
        "double",
        "dotted",
        "dashed",
        "dashDot",
        "dashDotDot",
        "words",
        "thick",
        "wave",
        "hairline",
        "thickDotted",
        "thickDashed",
        "thickDashDot",
        "thickDashDotDot",
        "thickLongDash",
        "longDash",
        "heavyWave",
        "doubleWave",
    ] {
        let source = open_bytes(br"{\rtf1\ansi Text}").unwrap();
        let changed = source
            .set("/body/p[1]/r[1]", Props::new().with("underline", token))
            .unwrap();
        let reopened = open_bytes(changed.to_bytes().unwrap()).unwrap();
        assert_eq!(property(&reopened, "/body/p[1]/r[1]", "underline"), token);
        assert_eq!(
            reopened
                .query(format!("r[underline={token}]").as_str())
                .unwrap()
                .count(),
            1
        );
    }
}

#[test]
fn invalid_values_and_unsupported_run_properties_leave_source_unchanged() {
    let source = file();
    for props in [
        props! {"size"=>"0pt"},
        props! {"size"=>"12.25pt"},
        props! {"size"=>"-1pt"},
        props! {"size"=>"999999999pt"},
        props! {"bold"=>"perhaps"},
        props! {"underline"=>"invented"},
    ] {
        let key = props[0].0.as_str();
        let error = source.set("/body/p[1]/r[2]", props.clone()).unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
        assert_eq!(error.property(), Some(key));
        assert_eq!(error.format(), Some(litchi_core::FileFormat::Rtf));
    }
    let half_point = source
        .set("/body/p[1]/r[2]", props! {"size"=>"12.5pt"})
        .unwrap();
    assert_eq!(property(&half_point, "/body/p[1]/r[2]", "size"), "12.5pt");
    for props in [
        props! {"bold"=>true,"font"=>"replacement"}
            .into_iter()
            .collect::<Props<'_>>(),
        props! {"font"=>"New font"}
            .into_iter()
            .collect::<Props<'_>>(),
        props! {"color"=>"#008800"}
            .into_iter()
            .collect::<Props<'_>>(),
        props! {"align"=>"center"}
            .into_iter()
            .collect::<Props<'_>>(),
    ] {
        assert_eq!(
            source.set("/body/p[1]/r[2]", props).unwrap_err().code(),
            Code::UnsupportedCapability
        );
    }
    assert!(source.set(query("p, r"), props! {"bold"=>true}).is_err());
    assert_eq!(
        source
            .set(query("r[text=missing]"), props! {"bold"=>true})
            .unwrap_err()
            .code(),
        Code::ZeroMatches
    );
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn fanout_limits_count_compound_run_updates_and_bound_publication() {
    let source = file();
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
        let err = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |edit| edit.set(query("r"), all_props()),
            )
            .unwrap_err();
        assert!(
            matches!(err.leaf(),Source::Transaction(Error::Limit{kind:actual,..}) if *actual==kind),
            "{err:?}"
        );
    }
    let limits = SemanticLimits::builder()
        .max_operations(1)
        .unwrap()
        .build()
        .unwrap();
    let commit = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.set("/body/p[1]/r[2]", all_props()),
        )
        .unwrap();
    assert_eq!(property(commit.file(), "/body/p[1]/r[2]", "size"), "14pt");
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn atomic_and_best_effort_batches_and_sequential_queries_keep_their_contract() {
    let source = file();
    let command = || item("/body/p[1]/r[2]", props! {"bold"=>true});
    let invalid = || item("/body/p[1]/r[2]", props! {"size"=>"invalid"});
    assert_eq!(
        source
            .batch([command(), invalid()], BatchMode::Atomic)
            .unwrap_err()
            .item(),
        Some(1)
    );
    let best = source
        .batch([command(), invalid()], BatchMode::BestEffort)
        .unwrap();
    assert_eq!(property(best.file(), "/body/p[1]/r[2]", "bold"), "true");
    let commit = source
        .edit(|edit| {
            edit.set("/body/p[1]/r[2]", props! {"outline"=>true})?;
            edit.set(query("r[outline=true]"), props! {"strike"=>true})
        })
        .unwrap();
    assert_eq!(commit.file().query("r[strike=true]").unwrap().count(), 1);
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn protection_and_origin_are_preserved_for_run_format_publication() {
    let input = br"{\rtf1\ansi\readprot\enforceprot1 Text}";
    let protected = open_bytes(input.as_slice()).unwrap();
    assert_eq!(
        protected
            .set("/body/p[1]/r[1]", props! {"bold"=>false})
            .unwrap()
            .to_bytes()
            .unwrap(),
        input
    );
    assert_eq!(
        protected
            .set("/body/p[1]/r[1]", props! {"bold"=>true})
            .unwrap_err()
            .code(),
        Code::DocumentProtected
    );
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("runs.rtf");
    std::fs::write(&path, INPUT).unwrap();
    litchi_ops::open(&path)
        .unwrap()
        .set("/body/p[1]/r[2]", all_props())
        .unwrap()
        .save()
        .unwrap();
    assert_eq!(
        property(&litchi_ops::open(&path).unwrap(), "/body/p[1]/r[2]", "size"),
        "14pt"
    );
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
}

#[test]
fn unsupported_source_closures_are_typed_atomic_refusals() {
    for input in [
        br"{\rtf1\ansi First\line second}".as_slice(),
        br"{\rtf1\ansi{\stylesheet{\s1 Styled;}}\s1 Text}".as_slice(),
    ] {
        let source = open_bytes(input).unwrap();
        let err = source
            .set("/body/p[1]/r[1]", props! {"bold"=>true})
            .unwrap_err();
        assert_eq!(err.code(), Code::UnsupportedCapability);
        assert!(matches!(err.leaf(), Source::RtfEdit(_)));
        assert_eq!(source.to_bytes().unwrap(), input);
    }
}

#[test]
fn escaped_unicode_run_text_survives_compound_formatting() {
    let input = br"{\rtf1\ansi\uc1 Caf\u233?{\i \u38634? \\ \{\}} tail}";
    let source = open_bytes(input.as_slice()).unwrap();
    let before = source
        .get("/body/p[1]")
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap()
        .into_owned();
    let changed = source.set(query("r[italic=true]"), all_props()).unwrap();
    let reopened = open_bytes(changed.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .get("/body/p[1]")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some(before.as_str())
    );
    assert!(
        reopened
            .query("r[italic=true][bold=true][underline=double]")
            .unwrap()
            .count()
            > 0
    );
    assert_eq!(source.to_bytes().unwrap(), input);
}

#[test]
fn empty_assignments_and_retained_native_source_policy_remain_consistent() {
    let source =
        litchi_ops::open_bytes_with(INPUT, &litchi_ops::OpenOptions::new().retain_source(false))
            .unwrap();
    let noop = source
        .edit(|edit| edit.set("/body/p[1]/r[2]", Props::new()))
        .unwrap();
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.file().to_bytes().unwrap(), INPUT);
    let changed = source.set("/body/p[1]/r[2]", all_props()).unwrap();
    assert_eq!(property(&changed, "/body/p[1]/r[2]", "bold"), "true");
    let limits = litchi_rtf::SourceLimits::builder()
        .max_input_bytes(INPUT.len())
        .unwrap()
        .build()
        .unwrap();
    let bounded = litchi_ops::open_bytes_with(
        INPUT,
        &litchi_ops::OpenOptions::new()
            .rtf_options(litchi_rtf::OpenOptions::new().source_limits(limits)),
    )
    .unwrap();
    let err = bounded.set("/body/p[1]/r[2]", all_props()).unwrap_err();
    assert_eq!(err.code(), Code::LimitExceeded);
    assert!(matches!(err.leaf(), Source::RtfEdit(_)));
    assert_eq!(bounded.to_bytes().unwrap(), INPUT);
}

#[test]
fn changing_local_flags_preserves_other_native_character_facets() {
    let input=br"{\rtf1\ansi{\fonttbl{\f0\fnil Example Sans;}}{\colortbl;\red255\green0\blue0;}\f0\cf1\cb1\highlight1\ulc1\up6\expndtw24\charscalex150\kerning28 Text}";
    let source = open_bytes(input.as_slice()).unwrap();
    let before = litchi_rtf::Package::from_bytes(input).unwrap();
    let changed = source
        .set("/body/p[1]/r[1]", props! {"bold"=>true,"italic"=>true})
        .unwrap();
    let after = litchi_rtf::Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();
    let old = before.body().runs().next().unwrap().format();
    let new = after.body().runs().next().unwrap().format();
    assert_eq!(new.baseline(), old.baseline());
    assert_eq!(new.expansion(), old.expansion());
    assert_eq!(
        new.horizontal_scale_percent(),
        old.horizontal_scale_percent()
    );
    assert_eq!(new.kerning_half_points(), old.kerning_half_points());
    assert_eq!(new.foreground(), old.foreground());
    assert_eq!(new.background(), old.background());
    assert_eq!(new.highlight(), old.highlight());
    assert_eq!(new.underline_paint(), old.underline_paint());
    assert!(new.bold() && new.italic());
    assert_eq!(source.to_bytes().unwrap(), input);
}

#[test]
fn unicode_fallback_widths_surrogates_and_control_symbols_keep_native_text() {
    for input in [
        br"{\rtf1\ansi\uc1 \u-10179?\u-8704? tail}".as_slice(),
        br"{\rtf1\ansi\uc2 \u233??tail}".as_slice(),
        br"{\rtf1\ansi\uc0 \u38634 tail}".as_slice(),
        br"{\rtf1\ansi a\emdash b\~c}".as_slice(),
        br"{\rtf1\ansi Caf\'e9 tail}".as_slice(),
    ] {
        let native = litchi_rtf::Package::from_bytes(input).unwrap();
        let source = open_bytes(input).unwrap();
        let changed = source
            .set(query("r"), props! {"bold"=>true,"italic"=>true})
            .unwrap_or_else(|error| {
                panic!(
                    "input={} error={error} leaf={:?}",
                    String::from_utf8_lossy(input),
                    error.leaf()
                )
            });
        let reopened = litchi_rtf::Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.text(), native.text());
        assert!(
            reopened
                .body()
                .runs()
                .all(|run| run.format().bold() && run.format().italic())
        );
        assert_eq!(source.to_bytes().unwrap(), input);
    }
}
