#![allow(
    clippy::expect_used,
    reason = "integration tests deliberately panic on unexpected failures"
)]

use litchi_odf_formula::authoring::{self, Display};
use litchi_odf_formula::{ChangeKind, Formula, History, NodePath, Patch, StarMathVersion, codec};

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
