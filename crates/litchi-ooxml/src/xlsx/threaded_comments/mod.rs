//! Threaded comments module for XLSX files.
//!
//! This module provides structures and functions for reading and writing
//! threaded comments (modern Excel comment threads) in XLSX workbooks.
//!
//! Threaded comments are a modern feature introduced in Office 365 that
//! support conversation threads, @mentions, and richer collaboration features.

pub mod package;
pub mod person;
pub mod reader;
pub mod writer;

pub use package::{
    ThreadedCommentGraph, WorkbookPersonPart, WorksheetThreadedCommentPart, add_threaded_comment,
    add_threaded_comment_person, add_threaded_comment_reply, find_threaded_comment,
    find_threaded_comment_person, load_threaded_comment_graph, remove_threaded_comment,
    remove_threaded_comment_person, reorder_threaded_comment_persons, reorder_threaded_comments,
    replace_threaded_comment, replace_threaded_comment_person, update_threaded_comment,
    update_threaded_comment_person, validate_threaded_comment_graph,
};
pub use person::{Mention, Person, PersonList};
pub use reader::{read_persons, read_threaded_comments};
pub use writer::{write_persons, write_threaded_comments};

pub use litchi_xlsx::threaded_comments::Comment as ThreadedComment;
pub use litchi_xlsx::threaded_comments::Comments as ThreadedComments;
