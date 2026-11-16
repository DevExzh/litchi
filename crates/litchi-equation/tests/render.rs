//! Equation pictures.
#![cfg(feature = "render")]

use litchi_equation::{
    Error, Limits, PngOptions, latex_to_png, omml_size_pt, omml_to_png, omml_to_unicode_math,
};

const MATH_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";

fn is_png(bytes: &[u8]) -> bool {
    bytes.starts_with(b"\x89PNG\r\n\x1a\n")
}

#[test]
fn latex_renders_as_a_png_sized_by_the_font_size() {
    let small = latex_to_png(
        r"\frac{a}{b}",
        &PngOptions::new(12.0).unwrap(),
        &Limits::DEFAULT,
    )
    .unwrap();
    assert!(is_png(small.bytes()));
    assert!(small.width() > 0 && small.height() > 0, "{small:?}");
    let large = latex_to_png(
        r"\frac{a}{b}",
        &PngOptions::new(24.0).unwrap(),
        &Limits::DEFAULT,
    )
    .unwrap();
    assert!(large.width() > small.width() && large.height() > small.height());
    // A taller fraction than a single letter.
    let (_, fraction_height) = small.size_pt();
    let (_, letter_height) = latex_to_png("x", &PngOptions::new(12.0).unwrap(), &Limits::DEFAULT)
        .unwrap()
        .size_pt();
    assert!(fraction_height > letter_height);
}

#[test]
fn scale_multiplies_pixels_but_not_the_size_in_points() {
    let options = PngOptions::new(18.0).unwrap();
    let one = latex_to_png("x^2+y^2", &options, &Limits::DEFAULT).unwrap();
    let three = latex_to_png(
        "x^2+y^2",
        &options.with_scale(3.0).unwrap(),
        &Limits::DEFAULT,
    )
    .unwrap();
    assert!(three.width() >= one.width() * 3 - 3, "{one:?} {three:?}");
    let (w1, h1) = one.size_pt();
    let (w3, h3) = three.size_pt();
    assert!(
        (w1 - w3).abs() < 2.0 && (h1 - h3).abs() < 2.0,
        "{w1}x{h1} {w3}x{h3}"
    );
}

#[test]
fn powerpoint_omml_with_drawingml_run_properties_renders() {
    let omml = format!(
        r#"<m:oMath xmlns:m="{MATH_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><m:r><a:rPr lang="en-US" i="1"><a:latin typeface="Cambria Math"/></a:rPr><m:t>𝑥</m:t></m:r><m:r><a:rPr lang="en-US" i="1"><a:latin typeface="Cambria Math"/></a:rPr><m:t>=2</m:t></m:r></m:oMath>"#
    );
    let png = omml_to_png(&omml, &PngOptions::DEFAULT, &Limits::DEFAULT).unwrap();
    assert!(is_png(png.bytes()));
}

#[test]
fn pictures_beyond_the_pixel_limit_are_refused_before_drawing() {
    let options = PngOptions::new(400.0)
        .unwrap()
        .with_max_pixels(10_000)
        .unwrap();
    let error = latex_to_png(r"\sum_{k=1}^{n} k^2", &options, &Limits::DEFAULT).unwrap_err();
    assert!(matches!(error, Error::Render(_)), "{error:?}");
}

#[test]
fn hostile_input_is_a_typed_error() {
    let options = PngOptions::DEFAULT;
    assert!(matches!(
        latex_to_png(r"\frac{", &options, &Limits::DEFAULT),
        Err(Error::Malformed { .. })
    ));
    assert!(matches!(
        omml_to_png("<m:oMath><m:r>", &options, &Limits::DEFAULT),
        Err(Error::Malformed { .. })
    ));
    let deep = format!("{}x{}", "{".repeat(5_000), "}".repeat(5_000));
    assert!(latex_to_png(&deep, &options, &Limits::DEFAULT).is_err());
    let limits = Limits::new(8, 1_000, 64).unwrap();
    assert!(matches!(
        latex_to_png("x+y+z+w+v", &options, &limits),
        Err(Error::InputTooLarge { .. })
    ));
}

#[test]
fn options_reject_values_out_of_range() {
    for size in [0.0, -1.0, f64::NAN, f64::INFINITY, 5000.0] {
        assert!(PngOptions::new(size).is_err(), "{size}");
    }
    for scale in [0.0, -2.0, f64::NAN, 17.0] {
        assert!(PngOptions::DEFAULT.with_scale(scale).is_err(), "{scale}");
    }
    assert!(PngOptions::DEFAULT.with_max_pixels(0).is_err());
}

#[test]
fn a_picture_is_placed_on_a_larger_transparent_canvas() {
    let options = PngOptions::new(18.0)
        .unwrap()
        .with_max_pixels(1_000_000)
        .unwrap();
    let png = latex_to_png("x^2", &options, &Limits::DEFAULT).unwrap();
    let placed = png.placed(400, 300, 10.0, 20.0, 1.0).unwrap();
    assert!(is_png(placed.bytes()));
    assert_eq!((placed.width(), placed.height()), (400, 300));
    let shrunk = png.placed(40, 30, 0.0, 0.0, 0.5).unwrap();
    assert_eq!((shrunk.width(), shrunk.height()), (40, 30));
    for zoom in [0.0, 1.5, f64::NAN] {
        assert!(matches!(
            png.placed(10, 10, 0.0, 0.0, zoom),
            Err(Error::InvalidLimit(_))
        ));
    }
    assert!(matches!(
        png.placed(10, 10, -1.0, 0.0, 1.0),
        Err(Error::InvalidLimit(_))
    ));
    assert!(matches!(
        png.placed(0, 10, 0.0, 0.0, 1.0),
        Err(Error::Render(_))
    ));
    assert!(matches!(
        png.placed(2_000, 2_000, 0.0, 0.0, 1.0),
        Err(Error::Render(_))
    ));
}

#[test]
fn the_size_is_measured_without_drawing_and_matches_the_picture() {
    let omml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>"#;
    let options = PngOptions::new(18.0).unwrap().with_scale(2.0).unwrap();
    let (width, height) = omml_size_pt(omml, &options, &Limits::DEFAULT).unwrap();
    let (drawn_width, drawn_height) = omml_to_png(omml, &options, &Limits::DEFAULT)
        .unwrap()
        .size_pt();
    assert!((width - drawn_width).abs() < 2.0, "{width} {drawn_width}");
    assert!(
        (height - drawn_height).abs() < 2.0,
        "{height} {drawn_height}"
    );
    assert_eq!(omml_to_unicode_math(omml, &Limits::DEFAULT).unwrap(), "a/b");
}

#[test]
fn the_depth_below_the_baseline_follows_the_equation() {
    let options = PngOptions::new(12.0).unwrap().with_scale(2.0).unwrap();
    let depth = |latex| {
        latex_to_png(latex, &options, &Limits::DEFAULT)
            .unwrap()
            .depth_pt()
    };
    // Every picture has its 2-pixel bottom margin below the baseline; a
    // descender, and more so a fraction's denominator, add to it.
    let margin = 2.0 * 72.0 / 96.0;
    assert!(
        depth("x") >= margin && depth("x") < margin + 1.0,
        "{}",
        depth("x")
    );
    assert!(depth("y") > depth("x"));
    assert!(depth(r"\frac{a}{b}") > depth("y"));
    // The depth is a fraction of the height, in the same unit.
    let picture = latex_to_png(r"\frac{a}{b}", &options, &Limits::DEFAULT).unwrap();
    assert!(picture.depth_pt() < picture.size_pt().1);
}

#[test]
fn mtef_renders_as_the_picture_of_the_same_latex() {
    use litchi_equation::{MtefTarget, latex_to_equation_native, mtef_to_png};
    let options = PngOptions::new(12.0).unwrap();
    for target in [MtefTarget::MathType, MtefTarget::EquationEditor3] {
        let native = latex_to_equation_native(r"\frac{a}{b}", &target, &Limits::DEFAULT).unwrap();
        let from_mtef = mtef_to_png(native.bytes(), &options, &Limits::DEFAULT).unwrap();
        let from_latex = latex_to_png(r"\frac{a}{b}", &options, &Limits::DEFAULT).unwrap();
        assert!(is_png(from_mtef.bytes()));
        assert_eq!(from_mtef.size_pt(), from_latex.size_pt());
        // Bare MTEF, without the Equation Native header, renders too.
        let bare = mtef_to_png(&native.bytes()[28..], &options, &Limits::DEFAULT).unwrap();
        assert_eq!(bare.size_pt(), from_mtef.size_pt());
    }
    assert!(matches!(
        mtef_to_png(b"not mtef", &options, &Limits::DEFAULT),
        Err(Error::Malformed { .. })
    ));
}
