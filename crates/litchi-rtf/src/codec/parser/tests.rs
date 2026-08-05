use super::super::error::RtfError;
use super::codec::parser_classification_error;

#[test]
fn parser_classification_failures_use_the_typed_error_channel() {
    assert!(matches!(
        parser_classification_error(),
        RtfError::ParserError(message)
            if message == "RTF parser control classification invariant failed"
    ));
}
