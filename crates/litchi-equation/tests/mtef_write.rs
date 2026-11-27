//! Writing `Equation Native` streams for OLE equation objects.

use litchi_equation::{
    Error, Limits, MtefProduct, MtefTarget, latex_to_equation_native, omml_to_equation_native,
    read_mtef,
};

#[test]
fn mathtype_target_writes_mtef5_with_the_ole_header() {
    let native =
        latex_to_equation_native(r"\frac{a}{b}", &MtefTarget::MathType, &Limits::DEFAULT).unwrap();
    let bytes = native.bytes();
    assert_eq!(&bytes[..2], &[28, 0], "EQNOLEFILEHDR length");
    assert_eq!(bytes[28], 5, "MTEF version");
    let declared = u32::from_le_bytes(bytes[8..12].try_into().unwrap()) as usize;
    assert_eq!(declared + 28, bytes.len());
    assert_eq!(native.info().version, 5);
    assert_eq!(native.info().product, MtefProduct::MathType);
    assert!(native.info().ole_header.is_some());
    assert!(native.losses().is_empty());
    let read = read_mtef(bytes, &Limits::DEFAULT).unwrap();
    assert_eq!(read.latex(), r"\frac{a}{b}");
}

#[test]
fn equation_editor_target_writes_mtef3() {
    let native =
        latex_to_equation_native("x^2", &MtefTarget::EquationEditor3, &Limits::DEFAULT).unwrap();
    assert_eq!(native.bytes()[28], 3);
    assert_eq!(native.info().version, 3);
    assert_eq!(native.info().product, MtefProduct::EquationEditor);
    assert_eq!(
        read_mtef(native.bytes(), &Limits::DEFAULT).unwrap().latex(),
        "x^{2}"
    );
}

#[test]
fn same_as_keeps_the_kind_of_an_existing_equation() {
    let old =
        latex_to_equation_native("y", &MtefTarget::EquationEditor3, &Limits::DEFAULT).unwrap();
    let target = MtefTarget::SameAs(old.info().clone());
    assert_eq!(target.prog_id(), "Equation.3");
    assert_eq!(target.clsid()[0], 0x02);
    let new = latex_to_equation_native(r"\sqrt{y}", &target, &Limits::DEFAULT).unwrap();
    assert_eq!(new.info().version, 3);
    assert_eq!(new.info().product, MtefProduct::EquationEditor);

    let old = latex_to_equation_native("y", &MtefTarget::MathType, &Limits::DEFAULT).unwrap();
    let target = MtefTarget::SameAs(old.info().clone());
    assert_eq!(target.prog_id(), "Equation.DSMT4");
    let new = latex_to_equation_native("z", &target, &Limits::DEFAULT).unwrap();
    assert_eq!(new.info().version, 5);
    assert_eq!(new.info().application_key, old.info().application_key);
}

#[test]
fn targets_name_their_ole_identity() {
    assert_eq!(MtefTarget::MathType.prog_id(), "Equation.DSMT4");
    assert_eq!(MtefTarget::EquationEditor3.prog_id(), "Equation.3");
    assert_eq!(
        MtefTarget::MathType.clsid(),
        [
            0x03, 0xCE, 0x02, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46
        ]
    );
    assert_eq!(
        MtefTarget::EquationEditor3.user_type(),
        "Microsoft Equation 3.0"
    );
}

#[test]
fn omml_converts_to_equation_native() {
    let omml = "<m:oMath><m:sSup><m:e><m:r><m:t>e</m:t></m:r></m:e>\
                <m:sup><m:r><m:t>x</m:t></m:r></m:sup></m:sSup></m:oMath>";
    let native = omml_to_equation_native(omml, &MtefTarget::MathType, &Limits::DEFAULT).unwrap();
    assert_eq!(
        read_mtef(native.bytes(), &Limits::DEFAULT).unwrap().latex(),
        "e^x"
    );
}

#[test]
fn malformed_and_oversized_input_is_refused() {
    assert!(matches!(
        latex_to_equation_native(r"\frac{", &MtefTarget::MathType, &Limits::DEFAULT),
        Err(Error::Malformed { .. })
    ));
    let limits = Limits::new(4, 1_000, 16).unwrap();
    assert!(matches!(
        latex_to_equation_native("abcdefgh", &MtefTarget::MathType, &limits),
        Err(Error::InputTooLarge { .. })
    ));
    let deep = "{".repeat(5_000) + &"}".repeat(5_000);
    assert!(latex_to_equation_native(&deep, &MtefTarget::MathType, &Limits::DEFAULT).is_err());
}

#[cfg(feature = "render")]
#[test]
fn mtef_renders_as_png() {
    let native = latex_to_equation_native("x+1", &MtefTarget::MathType, &Limits::DEFAULT).unwrap();
    let png = litchi_equation::mtef_to_png(
        native.bytes(),
        &litchi_equation::PngOptions::DEFAULT,
        &Limits::DEFAULT,
    )
    .unwrap();
    assert!(png.bytes().starts_with(b"\x89PNG"));
    assert!(png.width() > 0 && png.height() > 0);
    assert!(
        litchi_equation::mtef_to_png(
            b"not mtef",
            &litchi_equation::PngOptions::DEFAULT,
            &Limits::DEFAULT
        )
        .is_err()
    );
}
