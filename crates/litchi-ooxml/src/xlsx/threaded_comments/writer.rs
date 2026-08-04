//! Compatibility wrappers for the package-neutral threaded-comments writer.

use litchi_core::sheet::Result as SheetResult;

use super::ThreadedComments;
use super::person::PersonList;

/// Write a workbook-level people part.
pub fn write_persons(person_list: &PersonList) -> SheetResult<String> {
    litchi_xlsx::threaded_comments::write_persons(person_list)
}

/// Write one worksheet threaded-comments part.
pub fn write_threaded_comments(comments: &ThreadedComments) -> SheetResult<String> {
    litchi_xlsx::threaded_comments::write_comments(comments)
}
