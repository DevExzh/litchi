//! OPC/package adapter for the canonical XLSX threaded-comments owner.
//!
//! Semantic values and the bounded XML codec live in
//! [`litchi_xlsx::threaded_comments`]. This host module only resolves OPC
//! relationships, applies MCE preprocessing, and performs package CRUD.

mod package;
pub(crate) mod reader;

pub use litchi_xlsx::threaded_comments::{
    Comment, Comments, Graph, Mention, People, Person, SheetPart, WorkbookPart,
};
pub use package::{
    add_threaded_comment, add_threaded_comment_person, add_threaded_comment_reply,
    find_threaded_comment, find_threaded_comment_person, load_threaded_comment_graph,
    remove_threaded_comment, remove_threaded_comment_person, reorder_threaded_comment_persons,
    reorder_threaded_comments, replace_threaded_comment, replace_threaded_comment_person,
    update_threaded_comment, update_threaded_comment_person,
};
