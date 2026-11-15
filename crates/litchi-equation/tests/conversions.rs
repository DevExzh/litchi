//! The conversions Litchi's document crates rely on.

use litchi_equation::{
    Error, Format, LimitKind, Limits, MtefProduct, OmmlHost, latex_to_omml, mathml_to_latex,
    mtef_to_latex, omml_to_latex, read_mtef,
};

const MATH_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";

#[test]
fn omml_fragments_without_a_namespace_declaration_convert() {
    let latex = omml_to_latex(
        "<m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e>\
         <m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>",
        &Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(latex, "x^{2}");
}

#[test]
fn display_paragraphs_convert() {
    let omml = format!(
        r#"<m:oMathPara xmlns:m="{MATH_NS}"><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></m:oMathPara>"#
    );
    assert_eq!(omml_to_latex(&omml, &Limits::DEFAULT).unwrap(), "y");
}

#[test]
fn latex_becomes_a_namespaced_omath_element() {
    let omml = latex_to_omml("x^2", OmmlHost::Bare, &Limits::DEFAULT).unwrap();
    assert!(omml.losses().is_empty(), "{:?}", omml.losses());
    let omml = omml.into_xml();
    assert!(omml.starts_with("<m:oMath "), "{omml}");
    assert!(omml.contains(MATH_NS), "{omml}");
    assert!(omml.ends_with("</m:oMath>"), "{omml}");
    assert!(omml.contains('2'), "lost the exponent: {omml}");
}

#[test]
fn run_formatting_follows_the_host_dialect() {
    let word = latex_to_omml(
        r"\textcolor{red}{x}",
        OmmlHost::WordprocessingMl,
        &Limits::DEFAULT,
    )
    .unwrap();
    assert!(word.losses().is_empty(), "{:?}", word.losses());
    let word = word.into_xml();
    assert!(word.contains(r#"<w:color w:val="FF0000"/>"#), "{word}");
    let drawing = latex_to_omml(
        r"\textcolor{blue}{y}",
        OmmlHost::DrawingMl,
        &Limits::DEFAULT,
    )
    .unwrap()
    .into_xml();
    assert!(
        drawing.contains(r#"<a:srgbClr val="0000FF"/>"#),
        "{drawing}"
    );
}

#[test]
fn formatting_the_host_cannot_express_is_reported_as_a_loss() {
    let bare = latex_to_omml(r"\textcolor{red}{x}", OmmlHost::Bare, &Limits::DEFAULT).unwrap();
    assert!(!bare.xml().contains("FF0000"), "{}", bare.xml());
    assert_eq!(bare.losses().len(), 1, "{:?}", bare.losses());
}

#[test]
fn latex_survives_a_round_trip_through_omml() {
    let omml = latex_to_omml(r"\frac{a}{b}", OmmlHost::Bare, &Limits::DEFAULT).unwrap();
    assert_eq!(
        omml_to_latex(omml.xml(), &Limits::DEFAULT).unwrap(),
        r"\frac{a}{b}"
    );
}

#[test]
fn malformed_input_is_a_typed_error_not_a_panic() {
    assert!(matches!(
        latex_to_omml(r"\frac{a", OmmlHost::Bare, &Limits::DEFAULT),
        Err(Error::Malformed {
            format: Format::Latex,
            ..
        })
    ));
    assert!(matches!(
        omml_to_latex("<m:oMath><m:r>", &Limits::DEFAULT),
        Err(Error::Malformed {
            format: Format::Omml,
            ..
        })
    ));
    assert!(matches!(
        mtef_to_latex(&[0xFF; 12], &Limits::DEFAULT),
        Err(Error::Malformed {
            format: Format::Mtef,
            ..
        })
    ));
}

#[test]
fn input_beyond_the_byte_limit_is_refused_before_parsing() {
    let omml = "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>";
    let exact = Limits::new(omml.len(), 1_000, 32).unwrap();
    assert!(omml_to_latex(omml, &exact).is_ok());
    let short = Limits::new(omml.len() - 1, 1_000, 32).unwrap();
    assert!(matches!(
        omml_to_latex(omml, &short),
        Err(Error::InputTooLarge { size, limit }) if size == omml.len() && limit == omml.len() - 1
    ));
}

#[test]
fn nesting_beyond_the_depth_limit_is_refused() {
    let limits = Limits::new(4096, 1_000, 1).unwrap();
    assert!(matches!(
        latex_to_omml(r"\frac{\frac{a}{b}}{c}", OmmlHost::Bare, &limits),
        Err(Error::LimitExceeded {
            kind: LimitKind::Depth,
            limit: 1,
        })
    ));
}

#[test]
fn limits_reject_values_out_of_range() {
    assert!(matches!(Limits::new(0, 1, 1), Err(Error::InvalidLimit(_))));
    assert!(matches!(Limits::new(1, 0, 1), Err(Error::InvalidLimit(_))));
    assert!(matches!(Limits::new(1, 1, 0), Err(Error::InvalidLimit(_))));
    assert!(matches!(
        Limits::new(1, 1, Limits::MAX_DEPTH + 1),
        Err(Error::InvalidLimit(_))
    ));
    assert_eq!(Limits::default(), Limits::DEFAULT);
    assert_eq!(Limits::DEFAULT.max_depth(), Limits::MAX_DEPTH);
}

/// An Equation Editor 3 equation holding the digit 2, behind the 28-byte
/// `EQNOLEFILEHDR` of an `Equation Native` stream.
fn equation_native_v3() -> Vec<u8> {
    let mtef = [
        3, 1, 1, 3, 0, 0x0A, 0x01, 0x02, 0x88, 0x32, 0x00, 0x00, 0x00,
    ];
    let mut stream = vec![0u8; 28];
    stream[0] = 28;
    stream[2..6].copy_from_slice(&0x0002_0000_u32.to_le_bytes());
    stream[6..8].copy_from_slice(&0xC1C6_u16.to_le_bytes());
    stream[8..12].copy_from_slice(&u32::try_from(mtef.len()).unwrap().to_le_bytes());
    stream.extend_from_slice(&mtef);
    stream
}

#[test]
fn equation_native_streams_report_latex_and_their_kind() {
    let equation = read_mtef(&equation_native_v3(), &Limits::DEFAULT).unwrap();
    assert_eq!(equation.latex(), "2");
    let info = equation.info().unwrap();
    assert_eq!(info.version, 3);
    assert_eq!(info.product, MtefProduct::EquationEditor);
    assert_eq!(info.ole_header.unwrap().object_size, 13);
}

#[test]
fn mtef_5_equations_convert() {
    // MTEF 5 written by MathType 6: the equation `x`.
    let mut mtef = vec![5, 1, 0, 6, 0];
    mtef.extend_from_slice(b"DSMT6\0");
    mtef.extend_from_slice(&[
        0x01, 0x0A, 0x01, 0x00, 0x02, 0x00, 0x83, b'x', 0x00, 0x00, 0x00,
    ]);
    let equation = read_mtef(&mtef, &Limits::DEFAULT).unwrap();
    assert_eq!(equation.latex(), "x");
    let info = equation.info().unwrap();
    assert_eq!(info.version, 5);
    assert_eq!(info.application_key.as_deref(), Some("DSMT6"));
    assert!(info.ole_header.is_none());
}

#[test]
fn mathml_converts() {
    let mathml = r#"<math xmlns="http://www.w3.org/1998/Math/MathML"><msup><mi>x</mi><mn>2</mn></msup></math>"#;
    assert_eq!(mathml_to_latex(mathml, &Limits::DEFAULT).unwrap(), "x^{2}");
}
