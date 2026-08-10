#![allow(
    clippy::expect_used,
    reason = "integration tests deliberately panic on unexpected failures"
)]

use litchi_odf_common::constants::ODF_FORMULA;
use litchi_odf_common::core::OwnedPackage;
use litchi_odf_formula::authoring::content::{self, Closure, NumberType, Qualifier};
use litchi_odf_formula::authoring::{self, Display, Variant};
use litchi_odf_formula::{
    ContentSymbol, DependencyConflictKind, Element, Formula, History, NodePath, OpaqueStarMath,
    StarMathVersion, codec,
};
use soapberry_zip::office::StreamingArchiveWriter;

const NS: &str = "http://www.w3.org/1998/Math/MathML";
const RESOURCE_PATH: &str = "Resources/formula-proof.bin";
const RESOURCE_BYTES: &[u8] = b"\0formula-resource\xff";

#[test]
fn every_checked_constructor_family_builds_compact_reopenable_mathml() {
    let mut children = complete_content_corpus();
    children.extend(complete_presentation_corpus());
    let mut root = Element::new("math").expect("math root");
    root.set_attribute(None, "display", Display::Block.as_str())
        .expect("display");
    for child in children {
        root.push_child(child);
    }
    let xml = codec::serialize(&root);

    assert!(!xml.contains('\n'));
    assert!(!xml.contains("> <"));
    let formula = Formula::create(&xml).expect("complete constructor corpus");
    let reopened = Formula::from_bytes(formula.to_bytes()).expect("full package reopen");
    assert_eq!(reopened.root(), &root);

    let bad_identifier = content::number("1", NumberType::Integer).expect("number");
    assert!(content::declaration(bad_identifier, None).is_err());
    let bad_definition = content::apply(
        ContentSymbol::PLUS,
        vec![content::identifier("x"), content::identifier("y")],
    )
    .expect("application");
    assert!(content::declaration(content::identifier("f"), Some(bad_definition)).is_err());
}

#[test]
fn deterministic_xml_property_corpus_covers_every_constructor_and_raw_ingress() {
    let mut corpus = complete_content_corpus();
    corpus.extend(complete_presentation_corpus());
    let mut malformed = 0_usize;

    for (document_index, element) in corpus.into_iter().enumerate() {
        let root = if element.local_name() == "declare" {
            let mut root = Element::new("math").expect("math root");
            root.push_child(element);
            root
        } else {
            authoring::document_root(element, Display::Inline)
        };
        let baseline = codec::serialize(&root);
        let baseline_formula = Formula::create(&baseline).expect("constructor baseline");
        assert!(Formula::from_bytes(baseline_formula.to_bytes()).is_ok());

        for mutation in 0..12_usize {
            let mut bytes = baseline.as_bytes().to_vec();
            let state = deterministic_state(document_index, mutation);
            let index = usize::try_from(state).unwrap_or(0) % bytes.len();
            bytes[index] = b"<>/'x= \n"[usize::try_from(state >> 29).unwrap_or(0) % 8];
            if mutation.is_multiple_of(7) {
                bytes.truncate(bytes.len().saturating_sub((mutation % 5).saturating_add(1)));
            }
            let candidate = String::from_utf8(bytes).expect("ASCII corpus");
            let first = Formula::create(&candidate).is_ok();
            let second = Formula::create(&candidate).is_ok();
            let raw = Formula::from_bytes(raw_formula_package(&candidate, None)).is_ok();
            let parsed = codec::parse(&candidate).is_ok();
            assert_eq!(
                first, second,
                "document {document_index}, mutation {mutation}"
            );
            assert_eq!(
                parsed, raw,
                "raw ingress diverged from the XML validator for document {document_index}, mutation {mutation}: {candidate:?}"
            );
            assert!(!first || raw, "authored ingress accepted rejected raw XML");
            if first {
                let accepted = Formula::create(&candidate).expect("accepted mutation");
                assert!(Formula::from_bytes(accepted.to_bytes()).is_ok());
            }
            if !raw {
                malformed += 1;
            }
        }
    }

    assert!(
        malformed > 100,
        "mutation corpus did not exercise enough rejects"
    );
}

#[test]
fn presentation_content_and_annotation_boundaries_are_deterministic() {
    let valid = [
        format!(
            r#"<math xmlns="{NS}"><mrow><semantics><mi>x</mi><annotation encoding="text/plain">identifier</annotation></semantics></mrow></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><apply><plus/><semantics><ci>x</ci><annotation encoding="application/x-private">opaque</annotation></semantics><cn>1</cn></apply></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><semantics><ci>x</ci><annotation-xml encoding="application/xml"><v:proof xmlns:v="urn:vendor:proof"><apply/></v:proof></annotation-xml></semantics></math>"#
        ),
    ];
    let malformed = [
        format!(r#"<math xmlns="{NS}"><mrow><ci>x</ci></mrow></math>"#),
        format!(r#"<math xmlns="{NS}"><vector><mi>x</mi></vector></math>"#),
        format!(r#"<math xmlns="{NS}"><apply><plus/><mi>x</mi><cn>1</cn></apply></math>"#),
        format!(r#"<math xmlns="{NS}"><semantics/></math>"#),
        format!(
            r#"<math xmlns="{NS}"><semantics><annotation encoding="text/plain">early</annotation><mi>x</mi></semantics></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><semantics><mi>x</mi><annotation encoding="StarMath 6"><mi>not opaque text</mi></annotation></semantics></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><semantics><mi>x</mi><annotation encoding="text/plain">ok</annotation><ci>late</ci></semantics></math>"#
        ),
        format!(
            r#"<math xmlns="{NS}"><annotation-xml><v:x xmlns:v="urn:v"/></annotation-xml></math>"#
        ),
        format!(r#"<math xmlns="{NS}"><fn><ci>x</ci><ci>y</ci></fn></math>"#),
        format!(
            r#"<math xmlns="{NS}"><declare><ci>f</ci><apply><plus/><ci>x</ci><ci>y</ci></apply></declare></math>"#
        ),
    ];

    for xml in valid {
        assert!(
            Formula::create(&xml).is_ok(),
            "rejected valid boundary: {xml}"
        );
        assert!(Formula::from_bytes(raw_formula_package(&xml, None)).is_ok());
    }
    for xml in malformed {
        let first = Formula::create(&xml).is_err();
        let second = Formula::create(&xml).is_err();
        assert!(first && second, "accepted malformed boundary: {xml}");
        assert!(Formula::from_bytes(raw_formula_package(&xml, None)).is_err());
    }
}

#[test]
fn annotations_and_resources_retain_exact_provenance_through_opaque_edit_history() {
    let xml = format!(
        r#"<math xmlns="{NS}"><semantics><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow><annotation encoding="StarMath 5.0">x + 1</annotation><annotation encoding="text/plain">human note</annotation><annotation-xml encoding="application/xml"><v:proof xmlns:v="urn:vendor:proof" id="7"/></annotation-xml><annotation encoding="StarMath 6">x plus 1</annotation></semantics></math>"#
    );
    let raw = raw_formula_package(&xml, Some((RESOURCE_PATH, RESOURCE_BYTES)));
    let source = Formula::from_bytes(raw.clone()).expect("raw resource package");
    assert_eq!(source.as_bytes(), raw);
    assert_eq!(source.annotations().len(), 4);
    let starmath = source.starmath_annotations();
    assert_eq!(starmath.len(), 2);
    assert_eq!(starmath[0].opaque().source(), "x + 1");
    assert_eq!(starmath[1].opaque().source(), "x plus 1");

    let mut token_edit = source.edit();
    token_edit
        .set_text(&NodePath::new([0, 0, 0]), "y")
        .expect("presentation edit");
    let mut opaque_edit = source.edit();
    let opaque = OpaqueStarMath::new(StarMathVersion::V6, "y + 1").expect("opaque StarMath");
    opaque_edit.set_starmath(&opaque).expect("opaque edit");
    token_edit.join(&opaque_edit).expect("annotation join");
    let commit = token_edit.commit().expect("resource-preserving commit");
    let patch = commit.patch().clone();
    let target = commit.into_formula();
    let reopened = Formula::from_bytes(target.to_bytes()).expect("changed package reopen");
    assert_eq!(reopened.starmath().expect("StarMath").opaque(), &opaque);

    let package = OwnedPackage::from_bytes(reopened.to_bytes()).expect("raw changed package");
    assert_eq!(
        package.get_file(RESOURCE_PATH).expect("preserved resource"),
        RESOURCE_BYTES
    );
    assert_eq!(
        patch.inverse().apply(&target).expect("inverse").as_bytes(),
        source.as_bytes()
    );

    for case in 0..96_usize {
        let mut candidate = raw.clone();
        let index = usize::try_from(deterministic_state(case, 37)).unwrap_or(0) % candidate.len();
        candidate[index] ^= 1_u8 << (case % 8);
        let first = Formula::from_bytes(candidate.clone()).is_ok();
        let second = Formula::from_bytes(candidate.clone()).is_ok();
        assert_eq!(first, second, "raw resource package case {case}");
        if let Ok(accepted) = Formula::from_bytes(candidate) {
            assert!(Formula::from_bytes(accepted.to_bytes()).is_ok());
        }
    }
}

#[test]
fn deterministic_merge_transfer_and_history_properties_reopen_and_reverse() {
    for case in 0..32_usize {
        let base = history_base();
        let insertion_index = case % 7;
        let target_index = (case.wrapping_mul(5).wrapping_add(1)) % 6;

        let mut left = base.edit();
        left.insert_child(
            &NodePath::new([0]),
            insertion_index,
            authoring::operator("+"),
        )
        .expect("left insertion");
        let mut right = base.edit();
        right
            .set_text(&NodePath::new([0, target_index]), &format!("r{case}"))
            .expect("right presentation edit");
        right
            .set_text(&NodePath::new([1, target_index]), &case.to_string())
            .expect("right content edit");

        let transfer = left.plan_join(&right).expect("transfer plan");
        assert!(transfer.is_complete());
        let expected = target_index + usize::from(insertion_index <= target_index);
        assert_eq!(transfer.operations()[0].path().indices(), &[0, expected]);
        assert_eq!(
            transfer.operations()[1].path().indices(),
            &[1, target_index]
        );
        left.join(&right).expect("joined edits");
        let first_commit = left.commit().expect("first commit");
        let first_patch = first_commit.patch().clone();
        let first = first_commit.into_formula();

        let mut annotation_edit = first.edit();
        let opaque = OpaqueStarMath::new(StarMathVersion::V6, format!("case_{case}"))
            .expect("opaque history value");
        annotation_edit
            .set_starmath(&opaque)
            .expect("history StarMath edit");
        let second_commit = annotation_edit.commit().expect("second commit");
        let second_patch = second_commit.patch().clone();
        let target = second_commit.into_formula();

        let mut history = History::new();
        history
            .push(first_patch.clone())
            .expect("first history link");
        history
            .push(second_patch.clone())
            .expect("second history link");
        let durable = history.to_bytes().expect("history bytes");
        let decoded = History::from_bytes(&durable).expect("history reopen");
        let applied = decoded.apply(&base).expect("history apply");
        assert_eq!(applied.as_bytes(), target.as_bytes());
        assert!(Formula::from_bytes(applied.to_bytes()).is_ok());

        let restored_first = second_patch
            .inverse()
            .apply(&target)
            .expect("second inverse");
        let restored_base = first_patch
            .inverse()
            .apply(&restored_first)
            .expect("first inverse");
        assert_eq!(restored_base.as_bytes(), base.as_bytes());

        let mut corrupted = durable;
        let index = usize::try_from(deterministic_state(case, 91)).unwrap_or(0) % corrupted.len();
        corrupted[index] ^= 0x5a;
        let first_result = History::from_bytes(&corrupted).is_ok();
        let second_result = History::from_bytes(&corrupted).is_ok();
        assert_eq!(first_result, second_result, "history case {case}");
        if let Ok(candidate) = History::from_bytes(&corrupted) {
            let _classified = candidate.apply(&base).is_ok();
        }

        let target_path = NodePath::new([0, target_index]);
        let mut deletion = base.edit();
        deletion.remove(&target_path).expect("dependency removal");
        let before_root = deletion.root().clone();
        let before_changes = deletion.changes().to_vec();
        let mut stale_touch = base.edit();
        stale_touch
            .set_text(&target_path, "stale")
            .expect("dependent touch");
        let conflict = deletion.plan_join(&stale_touch).expect("conflict plan");
        assert_eq!(
            conflict.conflicts()[0].kind(),
            DependencyConflictKind::RemovedDependency
        );
        assert!(deletion.join(&stale_touch).is_err());
        assert_eq!(deletion.root(), &before_root);
        assert_eq!(deletion.changes(), before_changes);
    }
}

fn complete_presentation_corpus() -> Vec<Element> {
    let variants = [
        Variant::Normal,
        Variant::Bold,
        Variant::Italic,
        Variant::BoldItalic,
        Variant::DoubleStruck,
        Variant::BoldFraktur,
        Variant::Script,
        Variant::BoldScript,
        Variant::Fraktur,
        Variant::SansSerif,
        Variant::BoldSansSerif,
        Variant::SansSerifItalic,
        Variant::SansSerifBoldItalic,
        Variant::Monospace,
    ];
    let mut corpus: Vec<_> = variants
        .into_iter()
        .map(|variant| authoring::identifier_with_variant("x", variant))
        .collect();
    corpus.extend([
        authoring::identifier("x"),
        authoring::number("1"),
        authoring::operator("+"),
        authoring::literal_text("text"),
        authoring::string_literal("value", "[", "]"),
        authoring::row(vec![authoring::identifier("x"), authoring::number("1")]),
        authoring::fraction(authoring::number("1"), authoring::number("2")),
        authoring::square_root(authoring::identifier("x")),
        authoring::root(authoring::identifier("x"), authoring::number("3")),
        authoring::subscript(authoring::identifier("x"), authoring::number("1")),
        authoring::superscript(authoring::identifier("x"), authoring::number("2")),
        authoring::sub_superscript(
            authoring::identifier("x"),
            authoring::number("1"),
            authoring::number("2"),
        ),
        authoring::under(authoring::operator("sum"), authoring::number("0")),
        authoring::over(authoring::identifier("x"), authoring::operator("bar")),
        authoring::under_over(
            authoring::operator("sum"),
            authoring::number("0"),
            authoring::identifier("n"),
        ),
        authoring::fenced(
            vec![authoring::identifier("x"), authoring::identifier("y")],
            "(",
            ")",
            ",",
        ),
        authoring::table(vec![
            vec![authoring::identifier("a"), authoring::identifier("b")],
            vec![authoring::identifier("c"), authoring::identifier("d")],
        ]),
    ]);
    let opaque = OpaqueStarMath::new(StarMathVersion::V6, "x + 1").expect("opaque corpus value");
    corpus.push(authoring::semantics(
        authoring::identifier("x"),
        Some(&opaque),
    ));
    corpus.push(authoring::semantics(authoring::number("2"), None));
    corpus
}

fn complete_content_corpus() -> Vec<Element> {
    let integer = || content::number("2", NumberType::Integer).expect("integer");
    let condition = || {
        content::qualifier(
            Qualifier::Condition,
            content::relation(ContentSymbol::LT, vec![content::identifier("x"), integer()])
                .expect("condition relation"),
        )
        .expect("condition")
    };
    let domain = || {
        content::qualifier(
            Qualifier::DomainOfApplication,
            content::symbol(ContentSymbol::INTEGERS),
        )
        .expect("domain")
    };
    let bound = || content::bound_variable(content::identifier("i"), None).expect("bound");

    let mut corpus = vec![
        content::declaration(content::identifier("f"), None).expect("declaration"),
        content::declaration(
            content::identifier("g"),
            Some(content::function(content::identifier("x")).expect("function")),
        )
        .expect("defined declaration"),
        content::function(content::identifier("x")).expect("function"),
        content::identifier("x"),
        content::symbol_token("vendor:function"),
        content::number("1.5", NumberType::Real).expect("real"),
        content::number("2", NumberType::Integer).expect("integer"),
        content::number("pi", NumberType::Constant).expect("constant"),
        content::number_pair("1", "3", NumberType::Rational).expect("rational"),
        content::number_pair("1", "2", NumberType::ComplexCartesian).expect("cartesian"),
        content::number_pair("2", "3.14", NumberType::ComplexPolar).expect("polar"),
        content::apply(ContentSymbol::ABS, vec![content::identifier("x")]).expect("unary"),
        content::apply(
            ContentSymbol::DIVIDE,
            vec![content::identifier("x"), integer()],
        )
        .expect("binary"),
        content::apply(
            ContentSymbol::PLUS,
            vec![content::identifier("x"), integer(), integer()],
        )
        .expect("nary"),
        content::relation(ContentSymbol::LT, vec![content::identifier("x"), integer()])
            .expect("relation"),
        content::lambda(vec![bound()], Some(domain()), content::identifier("i")).expect("lambda"),
        content::lambda(
            vec![
                content::bound_variable(content::identifier("j"), Some(integer()))
                    .expect("graded bound"),
            ],
            None,
            content::identifier("j"),
        )
        .expect("graded lambda"),
        content::list(vec![integer(), content::identifier("x")]).expect("enumerated list"),
        content::list(vec![bound(), domain(), content::identifier("i")]).expect("generated list"),
        content::set(vec![integer(), content::identifier("x")]).expect("enumerated set"),
        content::set(vec![bound(), condition(), content::identifier("i")]).expect("generated set"),
        content::vector(vec![integer(), content::identifier("x")]).expect("enumerated vector"),
        content::vector(vec![bound(), domain(), content::identifier("i")])
            .expect("generated vector"),
        content::matrix(vec![
            content::matrix_row(vec![integer(), integer()]).expect("row one"),
            content::matrix_row(vec![integer(), integer()]).expect("row two"),
        ])
        .expect("enumerated matrix"),
        content::matrix(vec![bound(), domain(), content::identifier("i")])
            .expect("generated matrix"),
        content::piecewise(
            vec![(integer(), content::symbol(ContentSymbol::TRUE))],
            Some(content::number("0", NumberType::Integer).expect("zero")),
        )
        .expect("piecewise"),
        content::piecewise(
            vec![(integer(), content::symbol(ContentSymbol::FALSE))],
            None,
        )
        .expect("piecewise without otherwise"),
    ];
    for closure in [
        Closure::Open,
        Closure::Closed,
        Closure::OpenClosed,
        Closure::ClosedOpen,
    ] {
        corpus.push(content::interval(integer(), integer(), closure).expect("interval"));
    }
    corpus.push(
        content::apply(
            ContentSymbol::ROOT,
            vec![
                content::qualifier(Qualifier::Degree, integer()).expect("degree"),
                content::identifier("x"),
            ],
        )
        .expect("qualified root"),
    );
    corpus.push(
        content::apply(
            ContentSymbol::LOG,
            vec![
                content::qualifier(Qualifier::LogBase, integer()).expect("log base"),
                content::identifier("x"),
            ],
        )
        .expect("qualified log"),
    );
    corpus.push(
        content::apply(
            ContentSymbol::SUM,
            vec![
                bound(),
                content::qualifier(Qualifier::LowLimit, integer()).expect("low limit"),
                content::qualifier(Qualifier::UpLimit, integer()).expect("up limit"),
                content::identifier("i"),
            ],
        )
        .expect("bounded sum"),
    );
    corpus.push(
        content::apply(
            ContentSymbol::MOMENT,
            vec![
                content::qualifier(Qualifier::MomentAbout, integer()).expect("moment about"),
                content::identifier("x"),
            ],
        )
        .expect("qualified moment"),
    );
    corpus.extend(ContentSymbol::ALL.iter().copied().map(content::symbol));
    corpus
}

fn history_base() -> Formula {
    let presentation = authoring::row(
        (0..6)
            .map(|index| authoring::identifier(&format!("p{index}")))
            .collect(),
    );
    let content = content::vector(
        (0..6)
            .map(|index| {
                content::number(&index.to_string(), NumberType::Integer).expect("history number")
            })
            .collect(),
    )
    .expect("history vector");
    let opaque = OpaqueStarMath::new(StarMathVersion::V5, "s").expect("history StarMath");
    let semantics = authoring::semantics(authoring::identifier("s"), Some(&opaque));
    let mut root = Element::new("math").expect("history root");
    root.push_child(presentation);
    root.push_child(content);
    root.push_child(semantics);
    Formula::create(codec::serialize(&root)).expect("history base")
}

fn deterministic_state(document: usize, mutation: usize) -> u32 {
    let document = u32::try_from(document).unwrap_or(u32::MAX);
    let mutation = u32::try_from(mutation).unwrap_or(u32::MAX);
    document
        .wrapping_add(1)
        .wrapping_mul(0x9e37_79b9)
        .rotate_left(mutation % 31)
        ^ mutation.wrapping_mul(0x85eb_ca6b)
}

fn raw_formula_package(content: &str, resource: Option<(&str, &[u8])>) -> Vec<u8> {
    let resource_entry = resource.map_or_else(String::new, |(path, _bytes)| {
        format!(
            r#"<manifest:file-entry manifest:full-path="{path}" manifest:media-type="application/octet-stream"/>"#
        )
    });
    let manifest = format!(
        r#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="{ODF_FORMULA}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>{resource_entry}</manifest:manifest>"#
    );
    let mut archive = StreamingArchiveWriter::new();
    archive
        .write_stored("mimetype", ODF_FORMULA.as_bytes())
        .expect("mimetype");
    archive
        .write_deflated("content.xml", content.as_bytes())
        .expect("raw content");
    if let Some((path, bytes)) = resource {
        archive.write_deflated(path, bytes).expect("raw resource");
    }
    archive
        .write_deflated("META-INF/manifest.xml", manifest.as_bytes())
        .expect("manifest");
    archive.finish_to_bytes().expect("raw package")
}
