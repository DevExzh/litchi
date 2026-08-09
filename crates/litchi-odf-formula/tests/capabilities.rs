#![allow(
    clippy::expect_used,
    reason = "integration tests deliberately panic on unexpected failures"
)]

use litchi_odf_formula::authoring::{self, Display};
use litchi_odf_formula::{
    ChangeKind, DependencyConflictKind, Formula, History, MAX_COMMIT_HISTORY, NodePath,
    OpaqueStarMath, Patch, StarMathVersion, codec,
};

const NS: &str = "http://www.w3.org/1998/Math/MathML";

#[test]
fn mathml_two_content_arity_and_value_domains_are_enforced() {
    let malformed = [
        format!(r#"<math xmlns="{NS}"><mfrac><mi>x</mi></mfrac></math>"#),
        format!(r#"<math xmlns="{NS}"><mroot><mi>x</mi><mn>2</mn><mn>3</mn></mroot></math>"#),
        format!(r#"<math xmlns="{NS}"><mi mathvariant="initial">x</mi></math>"#),
        format!(r#"<math xmlns="{NS}"><mo stretchy="yes">+</mo></math>"#),
        format!(r#"<math xmlns="{NS}"><semantics><mi>x</mi><mi>y</mi></semantics></math>"#),
        format!(r#"<math xmlns="{NS}"><mtable><mtd><mi>x</mi></mtd></mtable></math>"#),
        format!(r#"<math xmlns="{NS}"><unknown/></math>"#),
    ];
    for xml in malformed {
        assert!(Formula::create(xml).is_err());
    }

    let valid = format!(
        r#"<math xmlns="{NS}" display="block"><mrow><mfrac linethickness="thin"><mi>x</mi><mn>2</mn></mfrac><mmultiscripts><mi>A</mi><none/><mi>n</mi><mprescripts/><mi>i</mi><none/></mmultiscripts></mrow></math>"#
    );
    assert!(Formula::create(valid).is_ok());
}

#[test]
fn checked_in_libreoffice_mathml_samples_validate() {
    let samples = [
        include_str!("../../../3rdparty/libreoffice-core/starmath/qa/extras/data/simple.mml"),
        include_str!("../../../3rdparty/libreoffice-core/starmath/qa/extras/data/mspace.mml"),
        include_str!("../../../3rdparty/libreoffice-core/starmath/qa/extras/data/color.mml"),
        include_str!("../../../3rdparty/libreoffice-core/starmath/qa/extras/data/tdf103430.mml"),
        include_str!("../../../3rdparty/libreoffice-core/starmath/qa/extras/data/tdf103500.mml"),
    ];
    for sample in samples {
        let root = codec::parse(sample).expect("checked-in MathML fixture");
        codec::validate(&root).expect("schema projection");
    }
}

#[test]
fn granular_edits_record_paths_and_round_trip_durable_history() {
    let xml = format!(
        r#"<math xmlns="{NS}"><semantics><mrow><mi>x</mi></mrow><annotation encoding="StarMath 5.0">x</annotation></semantics></math>"#
    );
    let source = Formula::create(xml).expect("source");
    let row_path = NodePath::new([0, 0]);
    let identifier_path = row_path.child(0);
    let mut edit = source.edit();
    edit.set_text(&identifier_path, "y").expect("token text");
    edit.insert_child(&row_path, 1, authoring::operator("+"))
        .expect("operator");
    edit.insert_child(&row_path, 2, authoring::number("1"))
        .expect("number");
    edit.set_starmath_source(StarMathVersion::V6, "y + 1")
        .expect("StarMath");
    let commit = edit.commit().expect("commit");

    assert_eq!(commit.patch().changes().len(), 4);
    assert_eq!(commit.patch().changes()[0].kind(), ChangeKind::SetText);
    assert_eq!(commit.patch().changes()[1].path().indices(), &[0, 0, 1]);
    assert_eq!(
        commit.formula().starmath_annotations()[0].version(),
        StarMathVersion::V6
    );

    let durable = commit.patch().to_bytes().expect("encode patch");
    let reopened = Patch::from_bytes(&durable).expect("decode patch");
    let target = reopened.apply(&source).expect("apply durable patch");
    assert_eq!(target.as_bytes(), commit.formula().as_bytes());
    assert_eq!(
        reopened
            .inverse()
            .apply(&target)
            .expect("inverse")
            .as_bytes(),
        source.as_bytes()
    );

    let mut history = History::new();
    history.push(reopened).expect("history entry");
    let history_bytes = history.to_bytes().expect("encode history");
    let reopened_history = History::from_bytes(&history_bytes).expect("decode history");
    assert_eq!(
        reopened_history.apply(&source).expect("replay").as_bytes(),
        target.as_bytes()
    );
}

#[test]
fn malformed_and_fuzz_like_inputs_fail_without_panics() {
    let malformed_patch = b"LITCHI-ODF-PATCH\0\x01\xff\xff\xff";
    assert!(Patch::from_bytes(malformed_patch).is_err());

    for count in 0..64 {
        let children = "<mi>x</mi>".repeat(count);
        let xml = format!(r#"<math xmlns="{NS}"><mfrac>{children}</mfrac></math>"#);
        if count == 2 {
            assert!(Formula::create(xml).is_ok());
        } else {
            assert!(Formula::create(xml).is_err());
        }
    }

    let mut deep = authoring::identifier("x");
    for _index in 0..140 {
        deep = authoring::row(vec![deep]);
    }
    let deep_document = authoring::document_root(deep, Display::Inline);
    assert!(Formula::create(codec::serialize(&deep_document)).is_err());
}

#[test]
fn content_mathml_schema_corpus_has_checked_structure_and_values() {
    let valid = [
        format!(r#"<math xmlns="{NS}"><apply><plus/><ci>x</ci><cn>2</cn></apply></math>"#),
        format!(
            r#"<math xmlns="{NS}"><declare><ci>basis</ci><vector><cn>1</cn><cn>0</cn></vector></declare><lambda><bvar><ci>x</ci><degree><cn>2</cn></degree></bvar><domainofapplication><integers/></domainofapplication><apply><power/><ci>x</ci><cn>2</cn></apply></lambda></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><piecewise><piece><cn>1</cn><true/></piece><otherwise><cn>0</cn></otherwise></piecewise></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><matrix><matrixrow><cn>1</cn><cn>2</cn></matrixrow></matrix></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><cn type="rational">1<sep/>2</cn><cn type="complex-polar">1<sep/>3.14</cn></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><ci><msub><mi>x</mi><mn>1</mn></msub></ci><csymbol>external</csymbol></math>"#
        ),
    ];
    for xml in valid {
        assert!(Formula::create(xml).is_ok());
    }

    let malformed = [
        format!(r#"<math xmlns="{NS}"><apply/></math>"#),
        format!(r#"<math xmlns="{NS}"><apply>text<plus/></apply></math>"#),
        format!(r#"<math xmlns="{NS}"><ci><cn>1</cn></ci></math>"#),
        format!(r#"<math xmlns="{NS}"><cn type="rational">1</cn></math>"#),
        format!(r#"<math xmlns="{NS}"><cn>1<sep/>2</cn></math>"#),
        format!(r#"<math xmlns="{NS}"><interval><cn>1</cn></interval></math>"#),
        format!(
            r#"<math xmlns="{NS}"><interval closure="half"><cn>1</cn><cn>2</cn></interval></math>"#
        ),
        format!(r#"<math xmlns="{NS}"><bvar/></math>"#),
        format!(r#"<math xmlns="{NS}"><lambda><bvar><ci>x</ci></bvar></lambda></math>"#),
        format!(r#"<math xmlns="{NS}"><lambda><ci>x</ci><ci>y</ci></lambda></math>"#),
        format!(r#"<math xmlns="{NS}"><piece><cn>1</cn></piece></math>"#),
        format!(
            r#"<math xmlns="{NS}"><piecewise><otherwise><cn>0</cn></otherwise><piece><cn>1</cn><true/></piece></piecewise></math>"#
        ),
        format!(r#"<math xmlns="{NS}"><matrix><ci>x</ci></matrix></math>"#),
        format!(r#"<math xmlns="{NS}"><matrix><matrixrow><mi>x</mi></matrixrow></matrix></math>"#),
        format!(r#"<math xmlns="{NS}"><declare><cn>1</cn></declare></math>"#),
        format!(r#"<math xmlns="{NS}"><declare><ci>x</ci><cn>1</cn></declare></math>"#),
        format!(r#"<math xmlns="{NS}"><ci>x</ci><declare><ci>y</ci></declare></math>"#),
        format!(r#"<math xmlns="{NS}"><declare occurrence="2"><ci>x</ci></declare></math>"#),
        format!(r#"<math xmlns="{NS}"><declare nargs="-1"><ci>x</ci></declare></math>"#),
        format!(r#"<math xmlns="{NS}"><mi><mglyph alt="glyph" index="0"/></mi></math>"#),
        format!(
            r#"<math xmlns="{NS}"><mi><mglyph alt="glyph" fontfamily="serif" index="-1"/></mi></math>"#
        ),
    ];
    for xml in malformed {
        assert!(
            Formula::create(&xml).is_err(),
            "unexpectedly accepted {xml}"
        );
    }
}

#[test]
fn opaque_starmath_boundary_is_versioned_bounded_and_publishable() {
    let source = Formula::create(format!(
        r#"<math xmlns="{NS}"><semantics><mi>x</mi><annotation encoding="StarMath 5.0">x</annotation></semantics></math>"#
    ))
    .expect("source");
    assert!(OpaqueStarMath::new(StarMathVersion::V6, "bad\0source").is_err());

    let opaque = OpaqueStarMath::new(StarMathVersion::V6, "x + 1").expect("opaque source");
    let mut edit = source.edit();
    edit.set_starmath(&opaque).expect("stage StarMath");
    let commit = edit.commit().expect("publish StarMath");
    let reopened = Formula::from_bytes(commit.formula().to_bytes()).expect("full reopen");
    let annotation = reopened
        .starmath_annotations()
        .into_iter()
        .next()
        .expect("StarMath annotation");
    assert_eq!(annotation.to_opaque().expect("opaque readback"), opaque);
}

#[test]
fn independent_sub_edits_transfer_paths_join_atomically_and_report_conflicts() {
    let source = Formula::create(format!(
        r#"<math xmlns="{NS}"><mrow><mi>a</mi><mi>b</mi></mrow></math>"#
    ))
    .expect("source");
    let row = NodePath::new([0]);

    let mut insertion = source.edit();
    insertion
        .insert_child(&row, 0, authoring::operator("-"))
        .expect("insert");
    let mut independent = source.edit();
    independent
        .set_text(&NodePath::new([0, 1]), "y")
        .expect("independent text");
    let transfer = insertion.plan_join(&independent).expect("transfer plan");
    assert!(transfer.is_complete());
    assert_eq!(transfer.operations()[0].path().indices(), &[0, 2]);
    insertion.join(&independent).expect("atomic join");
    let joined = insertion.commit().expect("publish join");
    assert_eq!(joined.formula().root().all_text(), "-ay");
    assert_eq!(joined.patch().changes().len(), 2);

    let mut left = source.edit();
    left.set_text(&NodePath::new([0, 0]), "x")
        .expect("left text");
    let left_root = left.root().clone();
    let left_changes = left.changes().to_vec();
    let mut right = source.edit();
    right
        .set_text(&NodePath::new([0, 0]), "z")
        .expect("right text");
    let conflict = left.plan_join(&right).expect("conflict plan");
    assert_eq!(
        conflict.conflicts()[0].kind(),
        DependencyConflictKind::SameTarget
    );
    assert!(left.join(&right).is_err());
    assert_eq!(left.root(), &left_root);
    assert_eq!(left.changes(), left_changes);
}

#[test]
fn three_way_plan_is_non_mutating_and_publishes_reopenable_reversible_bytes() {
    let base = Formula::create(format!(
        r#"<math xmlns="{NS}"><mrow><mi>a</mi><mi>b</mi></mrow></math>"#
    ))
    .expect("base");
    let base_bytes = base.to_bytes();
    let mut left_edit = base.edit();
    left_edit
        .set_text(&NodePath::new([0, 0]), "x")
        .expect("left");
    let left_commit = left_edit.commit().expect("left commit");
    let mut right_edit = base.edit();
    right_edit
        .set_text(&NodePath::new([0, 1]), "y")
        .expect("right");
    let right_commit = right_edit.commit().expect("right commit");

    let plan =
        Patch::plan_three_way(&base, left_commit.patch(), right_commit.patch()).expect("plan");
    assert!(plan.is_publishable());
    assert!(plan.conflicts().is_empty());
    assert_eq!(base.as_bytes(), base_bytes);
    let commit = plan.publish().expect("publish");
    assert!(commit.diagnostics().candidate_reopened());
    assert_eq!(commit.formula().root().all_text(), "xy");
    assert_eq!(commit.record().expect("record").changes().len(), 2);

    let durable = Patch::from_bytes(&commit.patch().to_bytes().expect("durable bytes"))
        .expect("durable reopen");
    let published = durable.apply(&base).expect("durable apply");
    let reopened = Formula::from_bytes(published.to_bytes()).expect("package reopen");
    assert_eq!(reopened.root().all_text(), "xy");
    assert!(durable.apply(left_commit.formula()).is_err());
    let restored = durable.inverse().apply(&published).expect("inverse");
    assert_eq!(restored.as_bytes(), base_bytes);
}

#[test]
fn published_commit_history_is_coupled_and_bounded() {
    let mut formula =
        Formula::create(format!(r#"<math xmlns="{NS}"><mi>0</mi></math>"#)).expect("source");
    for index in 1..=MAX_COMMIT_HISTORY + 2 {
        let source_revision = formula.revision();
        let mut edit = formula.edit();
        edit.set_text(&NodePath::new([0]), &index.to_string())
            .expect("stage history step");
        let commit = edit.commit().expect("publish history step");
        let record = commit.record().expect("commit record");
        assert_eq!(record.source_revision(), source_revision);
        assert_eq!(record.target_revision(), commit.formula().revision());
        formula = commit.into_formula();
    }
    assert_eq!(formula.commit_history().len(), MAX_COMMIT_HISTORY);
    assert!(
        formula
            .commit_history()
            .windows(2)
            .all(|window| window[0].target_revision() == window[1].source_revision())
    );
}
