#![allow(
    clippy::expect_used,
    reason = "integration tests deliberately panic on unexpected failures"
)]

use litchi_odf_common::constants::ODF_FORMULA;
use litchi_odf_formula::authoring::content::{self, Closure, NumberType, Qualifier};
use litchi_odf_formula::{
    ContentKind, ContentSymbol, Formula, Kind, MAX_STARMATH_SOURCE_BYTES, NodePath, OpaqueStarMath,
    StarMathVersion, codec,
};
use soapberry_zip::office::StreamingArchiveWriter;

const NS: &str = "http://www.w3.org/1998/Math/MathML";

#[test]
fn complete_content_model_and_checked_constructor_families_full_reopen() {
    let integer = content::number("2", NumberType::Integer).expect("integer");
    let rational = content::number_pair("1", "3", NumberType::Rational).expect("rational number");
    let x = content::identifier("x");
    let power =
        content::apply(ContentSymbol::POWER, vec![x.clone(), integer.clone()]).expect("power");
    let relation = content::relation(
        ContentSymbol::LT,
        vec![
            x.clone(),
            content::number("10", NumberType::Real).expect("real"),
        ],
    )
    .expect("relation");
    let bound = content::bound_variable(x.clone(), Some(integer.clone())).expect("bound variable");
    let domain = content::qualifier(
        Qualifier::DomainOfApplication,
        content::symbol(ContentSymbol::INTEGERS),
    )
    .expect("domain");
    let lambda = content::lambda(vec![bound], Some(domain), power.clone()).expect("lambda");
    let interval = content::interval(
        content::number("0", NumberType::Integer).expect("lower"),
        content::number("1", NumberType::Integer).expect("upper"),
        Closure::ClosedOpen,
    )
    .expect("interval");
    let matrix = content::matrix(vec![
        content::matrix_row(vec![integer.clone(), rational.clone()]).expect("row one"),
        content::matrix_row(vec![rational.clone(), integer.clone()]).expect("row two"),
    ])
    .expect("matrix");
    let piecewise = content::piecewise(
        vec![(integer.clone(), relation.clone())],
        Some(rational.clone()),
    )
    .expect("piecewise");
    let vector = content::vector(vec![integer.clone(), rational.clone()]).expect("vector");
    let list = content::list(vec![lambda.clone(), interval.clone()]).expect("list");
    let set = content::set(vec![integer.clone(), rational.clone()]).expect("set");

    let mut root = litchi_odf_formula::Element::new("math").expect("math root");
    for child in [
        power, relation, lambda, interval, matrix, piecewise, vector, list, set,
    ] {
        root.push_child(child);
    }
    codec::validate(&root).expect("complete checked tree");

    let formula = Formula::create(codec::serialize(&root)).expect("authored formula");
    let reopened = Formula::from_bytes(formula.to_bytes()).expect("full package reopen");
    let kinds: Vec<_> = reopened
        .root()
        .children()
        .map(litchi_odf_formula::Element::kind)
        .collect();
    assert_eq!(kinds[0], Kind::Content(ContentKind::Application));
    assert_eq!(kinds[1], Kind::Content(ContentKind::Relation));
    assert_eq!(kinds[2], Kind::Content(ContentKind::Lambda));
    assert_eq!(kinds[3], Kind::Content(ContentKind::Interval));
    assert_eq!(kinds[4], Kind::Content(ContentKind::Matrix));
    assert_eq!(kinds[5], Kind::Content(ContentKind::Piecewise));
    assert_eq!(kinds[6], Kind::Content(ContentKind::Vector));
    assert_eq!(kinds[7], Kind::Content(ContentKind::List));
    assert_eq!(kinds[8], Kind::Content(ContentKind::Set));

    assert!(content::apply(ContentSymbol::DIVIDE, vec![integer.clone()]).is_err());
    assert!(content::relation(ContentSymbol::PLUS, vec![integer.clone(), rational]).is_err());
    assert!(content::number("1", NumberType::Rational).is_err());
    assert!(content::number_pair("1", "2", NumberType::Integer).is_err());
}

#[test]
fn every_named_content_symbol_is_constructible_and_classified() {
    assert!(ContentSymbol::ALL.len() >= 100);
    for symbol in ContentSymbol::ALL {
        let element = content::symbol(*symbol);
        assert_eq!(element.kind(), Kind::ContentSymbol(*symbol));
        assert_eq!(
            ContentSymbol::from_local_name(symbol.as_str()),
            Some(*symbol)
        );
        let mut root = litchi_odf_formula::Element::new("math").expect("math");
        root.push_child(element);
        codec::validate(&root).expect("named symbol validates");
    }
    assert!(ContentSymbol::from_local_name("vendor-symbol").is_none());
}

#[test]
fn recognized_starmath_is_opaque_at_open_authoring_read_and_edit_boundaries() {
    let opaque = OpaqueStarMath::new(StarMathVersion::V6, "x + 1").expect("opaque source");
    let semantics = litchi_odf_formula::authoring::semantics(
        litchi_odf_formula::authoring::identifier("x"),
        Some(&opaque),
    );
    let root = litchi_odf_formula::authoring::document_root(
        semantics,
        litchi_odf_formula::authoring::Display::Block,
    );
    let source = Formula::create(codec::serialize(&root)).expect("opaque authored source");
    assert_eq!(source.starmath().expect("StarMath").opaque(), &opaque);

    let replacement = OpaqueStarMath::new(StarMathVersion::V5, "y").expect("replacement");
    let mut edit = source.edit();
    edit.set_starmath(&replacement).expect("opaque edit");
    let target = edit.commit().expect("commit").into_formula();
    let reopened = Formula::from_bytes(target.to_bytes()).expect("full reopen");
    assert_eq!(
        reopened.starmath().expect("StarMath").opaque(),
        &replacement
    );

    let oversized = "x".repeat(MAX_STARMATH_SOURCE_BYTES + 1);
    let oversized_xml = format!(
        r#"<math xmlns="{NS}"><semantics><mi>x</mi><annotation encoding="StarMath 6">{oversized}</annotation></semantics></math>"#
    );
    assert!(Formula::create(&oversized_xml).is_err());
    assert!(Formula::from_bytes(raw_formula_package(&oversized_xml)).is_err());

    let mut generic_edit = source.edit();
    assert!(
        generic_edit
            .set_text(&NodePath::new([0, 1]), &oversized)
            .is_err()
    );
}

#[test]
fn raw_malformed_and_prettified_packages_prove_ingress_and_exact_provenance() {
    let malformed = format!(r#"<math xmlns="{NS}"><mfrac><mi>x</mi></mfrac></math>"#);
    assert!(Formula::from_bytes(raw_formula_package(&malformed)).is_err());

    let pretty = format!("<?xml version=\"1.0\"?>\n<math xmlns=\"{NS}\">\n  <mi>x</mi>\n</math>\n");
    let pretty_bytes = raw_formula_package(&pretty);
    let source = Formula::from_bytes(pretty_bytes.clone()).expect("prettified ingress");
    assert_eq!(source.as_bytes(), pretty_bytes);

    let mut edit = source.edit();
    edit.set_text(&NodePath::new([0]), "y").expect("text edit");
    assert!(edit.commit().is_err());
    assert_eq!(source.as_bytes(), pretty_bytes);
}

#[test]
fn deterministic_mutation_corpus_never_panics_or_changes_classification() {
    let baseline =
        format!(r#"<math xmlns="{NS}"><apply><plus/><ci>x</ci><cn>1</cn></apply></math>"#);
    let mut state = 0xA5A5_1F2D_u32;
    for case in 0..512_u32 {
        state = state.wrapping_mul(1_664_525).wrapping_add(1_013_904_223);
        let mut bytes = baseline.as_bytes().to_vec();
        let index = usize::try_from(state).unwrap_or(0) % bytes.len();
        bytes[index] = b"<>/'x= \n"[usize::try_from(state >> 24).unwrap_or(0) % 8];
        if case % 5 == 0 {
            bytes.truncate(
                bytes
                    .len()
                    .saturating_sub(usize::try_from(case % 17).unwrap_or(0)),
            );
        }
        let candidate = String::from_utf8(bytes).expect("ASCII mutation");
        let first = Formula::create(&candidate).is_ok();
        let second = Formula::create(&candidate).is_ok();
        assert_eq!(first, second, "case {case}");
        if first {
            let formula = Formula::create(&candidate).expect("classified valid");
            assert!(Formula::from_bytes(formula.to_bytes()).is_ok());
        }
    }
}

fn raw_formula_package(content: &str) -> Vec<u8> {
    const MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.formula"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#;
    let mut archive = StreamingArchiveWriter::new();
    archive
        .write_stored("mimetype", ODF_FORMULA.as_bytes())
        .expect("mimetype");
    archive
        .write_deflated("content.xml", content.as_bytes())
        .expect("raw test content");
    archive
        .write_deflated("META-INF/manifest.xml", MANIFEST)
        .expect("manifest");
    archive.finish_to_bytes().expect("raw test package")
}
