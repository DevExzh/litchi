//! Unit tests for the value coercions. Expected values follow OfficeCLI's
//! source at `dced0d74` (cited per case); no .NET runtime was available to
//! execute it.

mod booleans;
mod colors;
mod compare;
mod lengths;
mod numbers;
mod roundtrip;
mod tables;
mod tokens;

use super::{
    Color, ColorContext, ErrorKind, Length, LengthContext, Scalar, Test, ValueError, parse_bool,
};

/// The error kind of a failed parse, for compact assertions.
fn kind<T: std::fmt::Debug>(result: Result<T, ValueError<'_>>) -> ErrorKind {
    match result {
        Ok(value) => panic!("expected an error, got {value:?}"),
        Err(error) => error.kind(),
    }
}

#[test]
fn authored_literals_stay_out_of_diagnostics_and_debug() {
    for error in [
        parse_bool("PRIVATE").unwrap_err(),
        Color::parse("PRIVATE", ColorContext::Drawing).unwrap_err(),
        Length::parse_in("PRIVATEin", LengthContext::Emu).unwrap_err(),
    ] {
        assert!(!format!("{error:?}").contains("PRIVATE"));
        assert!(!error.to_string().contains("PRIVATE"));
        let unified = error.into_error();
        assert!(!unified.to_string().contains("PRIVATE"));
    }
    let scalar = Scalar::Text("PRIVATE");
    assert!(!format!("{scalar:?}").contains("PRIVATE"));
    assert!(!format!("{:?}", scalar.display(None)).contains("PRIVATE"));
    let test = Test::with_text(litchi_core::selector::CompareOp::Eq, "PRIVATE", None);
    assert!(!format!("{test:?}").contains("PRIVATE"));
}
