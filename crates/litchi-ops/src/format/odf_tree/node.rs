//! The typed node of an object tree.

use std::borrow::Cow;

use litchi_core::FileFormat;

use super::OdfTree;
use super::children::Children;
use super::tree::Item;
use crate::error::Result;
use crate::node::{
    Children as ViewChildren, Format, NodePath, Preview, Project, Source, Value, ViewOptions,
};

/// One element of an OpenDocument drawing, chart, image, database or
/// formula: a borrowed position in the handle's object tree.
#[derive(Clone, Copy)]
pub struct Node<'f> {
    file: &'f OdfTree,
    index: u32,
}

impl<'f> Node<'f> {
    pub(crate) const fn new(file: &'f OdfTree, index: u32) -> Self {
        Self { file, index }
    }

    pub(crate) const fn index(&self) -> u32 {
        self.index
    }

    pub(crate) fn item(&self) -> &'f Item {
        self.file.tree().item(self.index)
    }

    /// The OfficeCLI type name.
    #[must_use]
    pub fn kind(&self) -> &'static str {
        self.item().type_name()
    }

    /// The format of the node's document.
    #[must_use]
    pub fn file_format(&self) -> FileFormat {
        self.file.format()
    }

    /// The canonical positional path.
    #[must_use]
    pub fn path(&self) -> NodePath<'f> {
        NodePath::new(self.item().path())
    }

    /// The element's text, if it has any.
    ///
    /// # Errors
    ///
    /// None: the text was read at open; the `Result` keeps node reads
    /// uniform across formats.
    pub fn text(&self) -> Result<Option<Cow<'f, str>>> {
        Ok(self.item().text().map(Cow::Borrowed))
    }

    /// The value of property `key`, if set.
    #[must_use]
    pub fn value(&self, key: &str) -> Option<&'f Value<'static>> {
        self.item().value(key)
    }
}

impl std::fmt::Debug for Node<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("odf::Node")
            .field("kind", &self.kind())
            .finish_non_exhaustive()
    }
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        Node::kind(self)
    }

    fn file_format(&self) -> FileFormat {
        Node::file_format(self)
    }

    fn path(&self) -> NodePath<'_> {
        Node::path(self)
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        Node::text(self)
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        Ok(self
            .item()
            .text()
            .map(|text| Preview::Prefix(Cow::Borrowed(text))))
    }

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        Ok(Some(
            u32::try_from(self.item().children().len()).unwrap_or(u32::MAX),
        ))
    }

    fn format(&self) -> Result<Format<'_>> {
        let item = self.item();
        Ok(Format::new(item.props(), Source::Tree(item.values())))
    }

    fn children(&self, _: &ViewOptions) -> ViewChildren<'_, 'f> {
        ViewChildren::odf_tree(Children::filtered(self.file, self.index, u128::MAX))
    }
}

const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Node<'static>>();
    assert!(size_of::<Node<'static>>() <= 16);
};
