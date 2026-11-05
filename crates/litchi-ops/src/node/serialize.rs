//! Streaming serialization of a node view into OfficeCLI's eight-member
//! `DocumentNode` JSON (record 0010 §2.1, §2.7).
//!
//! The serializer writes each node as a struct in OfficeCLI's member order
//! (`path`, `type`, `text`, `preview`, `style`, `childCount`, `format`,
//! `children`). `path` and `preview` go through `collect_str`, so their text
//! is written straight into the sink; `format` is a map streamed from the
//! property table; `children` is a sequence streamed from the child
//! iterator, each child serialized one level down and dropped before the
//! next is read. Nothing is collected: no `DocumentNode`, no map, no `String`
//! per node.
//!
//! Recursion is one level per view depth, which [`MAX_VIEW_DEPTH`] bounds
//! (record 0010 risk 12). A shared counter stops the walk with
//! `output_too_large` once more than `max_nodes` nodes would be written.
//!
//! [`MAX_VIEW_DEPTH`]: super::MAX_VIEW_DEPTH

use std::cell::Cell;
use std::fmt;

use serde::ser::{Serialize, SerializeMap, SerializeSeq, SerializeStruct, Serializer};

use super::Node;
use super::view::{Project, View, ViewOptions};
use crate::error::{Code, Error};

/// The error of [`View::serialize_into`]: either the view failed, with the
/// unified error kept typed, or the serializer did.
///
/// ```
/// use litchi_ops::Code;
/// use litchi_ops::node::{Node, SerializeError, ViewOptions};
///
/// fn to_json(node: &Node<'_>) -> Result<Vec<u8>, String> {
///     let mut sink = Vec::with_capacity(4096);
///     let view = node.view_with(ViewOptions::default().depth(3).max_nodes(1_000));
///     match view.serialize_into(&mut serde_json::Serializer::new(&mut sink)) {
///         Ok(()) => Ok(sink),
///         Err(SerializeError::View(error)) if error.code() == Code::OutputTooLarge => {
///             Err("too many nodes; lower the depth".to_owned())
///         },
///         Err(error) => Err(error.to_string()),
///     }
/// }
/// # let _ = to_json;
/// ```
#[non_exhaustive]
pub enum SerializeError<E> {
    /// A node could not be read, or the view exceeded its `max_nodes` budget
    /// (`output_too_large`).
    View(Error),
    /// The serializer or its sink failed. The typed cause is available by
    /// matching this variant; it is omitted from the diagnostic source chain
    /// because external errors may contain authored content.
    Serializer(E),
}

impl<E: fmt::Display> fmt::Display for SerializeError<E> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::View(error) => fmt::Display::fmt(error, f),
            Self::Serializer(_) => f.write_str("node serializer failed"),
        }
    }
}

impl<E: std::error::Error + 'static> std::error::Error for SerializeError<E> {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::View(error) => Some(error),
            Self::Serializer(_) => None,
        }
    }
}

impl View<'_, '_> {
    /// Serializes the view into `serializer`, keeping the unified error
    /// typed: a node that fails to read, or a walk that would exceed
    /// `max_nodes` (`output_too_large`), is [`SerializeError::View`], and a
    /// failing sink is [`SerializeError::Serializer`].
    ///
    /// The plain `serde::Serialize` implementation writes the same output
    /// and reports a view failure through `serde::ser::Error::custom`.
    ///
    /// # Errors
    ///
    /// As described above. The sink may hold a partial document.
    pub fn serialize_into<S: Serializer>(
        &self,
        serializer: S,
    ) -> Result<S::Ok, SerializeError<S::Error>> {
        let walk = Walk::new(self.options());
        match walk
            .node(self.node(), self.options().depth)
            .serialize(serializer)
        {
            Ok(ok) => Ok(ok),
            Err(error) => Err(match walk.error.take() {
                Some(view) => SerializeError::View(view),
                None => SerializeError::Serializer(error),
            }),
        }
    }
}

/// Streams the view as OfficeCLI's `DocumentNode` JSON object. A view
/// failure (a node that fails to read, or `output_too_large`) becomes
/// `serde::ser::Error::custom` with the unified error's text; use
/// [`View::serialize_into`] to keep it typed.
impl Serialize for View<'_, '_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let walk = Walk::new(self.options());
        walk.node(self.node(), self.options().depth)
            .serialize(serializer)
    }
}

/// State shared by one serialization: the node budget and the first view
/// error, kept typed for [`View::serialize_into`].
struct Walk {
    options: ViewOptions,
    emitted: Cell<u32>,
    error: Cell<Option<Error>>,
}

impl Walk {
    fn new(options: ViewOptions) -> Self {
        Self {
            options,
            emitted: Cell::new(0),
            error: Cell::new(None),
        }
    }

    fn node<'w, 'n, 'f>(&'w self, node: &'n Node<'f>, depth: u32) -> NodeOut<'w, 'n, 'f> {
        NodeOut {
            walk: self,
            node,
            depth,
        }
    }

    /// Counts one more node, or fails with `output_too_large`.
    fn enter<E: serde::ser::Error>(&self) -> Result<(), E> {
        if self.options.depth > super::MAX_VIEW_DEPTH {
            return Err(self.fail(Error::new(
                Code::MaxDepthExceeded,
                "view depth exceeds the supported recursion limit",
            )));
        }
        let emitted = self.emitted.get();
        if emitted >= self.options.max_nodes {
            return Err(self.fail(
                Error::new(
                    Code::OutputTooLarge,
                    format!(
                        "the node view exceeds its budget of {} nodes",
                        self.options.max_nodes
                    ),
                )
                .with_suggestion("lower `depth` or raise `max_nodes` in the view options"),
            ));
        }
        self.emitted.set(emitted + 1);
        Ok(())
    }

    /// Records `error` (the first one wins: later failures are consequences)
    /// and returns the serializer's error carrying its text.
    fn fail<E: serde::ser::Error>(&self, error: Error) -> E {
        let custom = E::custom(&error);
        let first = self.error.take().unwrap_or(error);
        self.error.set(Some(first));
        custom
    }
}

/// One node of the walk, serialized as a `DocumentNode` object.
struct NodeOut<'w, 'n, 'f> {
    walk: &'w Walk,
    node: &'n Node<'f>,
    depth: u32,
}

impl Serialize for NodeOut<'_, '_, '_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let walk = self.walk;
        let node = self.node;
        // This node's options: the view's, with the depth left at this level.
        let options = walk.options.depth(self.depth);
        walk.enter::<S::Error>()?;
        // Read the optional members before writing anything, so a failing
        // scalar read stops before this node starts; format/child failures may
        // leave a partial node, as permitted by serialize_into.
        let text = node.text().map_err(|error| walk.fail::<S::Error>(error))?;
        let preview = node
            .preview()
            .map_err(|error| walk.fail::<S::Error>(error))?;
        let style = node.style().map_err(|error| walk.fail::<S::Error>(error))?;
        let child_count = node
            .child_count(&options)
            .map_err(|error| walk.fail::<S::Error>(error))?;
        let members = 4
            + usize::from(text.is_some())
            + usize::from(preview.is_some())
            + usize::from(style.is_some())
            + usize::from(child_count.is_some());

        let mut out = serializer.serialize_struct("DocumentNode", members)?;
        out.serialize_field("path", &Collect(node.path()))?;
        out.serialize_field("type", node.kind())?;
        match &text {
            Some(text) => out.serialize_field("text", text.as_ref())?,
            None => out.skip_field("text")?,
        }
        match &preview {
            Some(preview) => out.serialize_field("preview", &Collect(preview))?,
            None => out.skip_field("preview")?,
        }
        match &style {
            Some(style) => out.serialize_field("style", style.as_ref())?,
            None => out.skip_field("style")?,
        }
        match child_count {
            Some(count) => out.serialize_field("childCount", &count)?,
            None => out.skip_field("childCount")?,
        }
        out.serialize_field("format", &FormatOut { walk, node })?;
        out.serialize_field(
            "children",
            &ChildrenOut {
                walk,
                node,
                options,
            },
        )?;
        out.end()
    }
}

/// A node's `format` member, streamed as a map in property-table order.
struct FormatOut<'w, 'n, 'f> {
    walk: &'w Walk,
    node: &'n Node<'f>,
}

impl Serialize for FormatOut<'_, '_, '_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let format =
            Project::format(self.node).map_err(|error| self.walk.fail::<S::Error>(error))?;
        let mut map = serializer.serialize_map(None)?;
        for prop in format {
            if matches!(prop.value(), super::Value::Number(value) if !value.is_finite()) {
                return Err(self
                    .walk
                    .fail(Error::new(Code::InvalidValue, "node number must be finite")));
            }
            map.serialize_entry(prop.key(), &prop)?;
        }
        map.end()
    }
}

/// A node's `children` member, streamed one child at a time.
struct ChildrenOut<'w, 'n, 'f> {
    walk: &'w Walk,
    node: &'n Node<'f>,
    /// The parent's options, whose depth is the levels left below it.
    options: ViewOptions,
}

impl Serialize for ChildrenOut<'_, '_, '_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut seq = serializer.serialize_seq(None)?;
        if let Some(depth) = self.options.depth.checked_sub(1) {
            for child in Project::children(self.node, &self.options) {
                let child = child.map_err(|error| self.walk.fail::<S::Error>(error))?;
                seq.serialize_element(&self.walk.node(&child, depth))?;
            }
        }
        seq.end()
    }
}

/// Serializes a `Display` value as a string written straight into the sink.
struct Collect<T>(T);

impl<T: fmt::Display> Serialize for Collect<T> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        serializer.collect_str(&self.0)
    }
}

impl<E> fmt::Debug for SerializeError<E> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::View(_) => f.debug_struct("View").finish_non_exhaustive(),
            Self::Serializer(_) => f.debug_struct("Serializer").finish_non_exhaustive(),
        }
    }
}
