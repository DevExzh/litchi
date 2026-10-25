use crate::Error;
use litchi_core::edit::RefusalKind;

pub(super) fn not_found(operation: &'static str) -> Error {
    Error::refused(
        RefusalKind::NotFound,
        operation,
        "the Custom XML item does not exist",
    )
}

pub(super) fn unsupported_target(operation: &'static str, reason: &'static str) -> Error {
    Error::refused(RefusalKind::UnsupportedTarget, operation, reason)
}

pub(super) fn unsupported_content(operation: &'static str, reason: &'static str) -> Error {
    Error::refused(RefusalKind::UnsupportedContent, operation, reason)
}
