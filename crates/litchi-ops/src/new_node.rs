//! [`NewNode`]: the element created by `File::add` (record 0015 §2.1).

use std::borrow::Cow;
use std::fmt;
use std::ops::Deref;

use litchi_core::selector::{Path, limits};

use crate::error::{Code, Error, Result};
use crate::file::File;
use crate::placement::{IntoPlacement, Placement};
use crate::props::{IntoProps, Key, PropValue, Props};

/// An element for `add`: its type, property assignments and optional
/// placement among the parent's children. A missing placement means the end.
/// The type is resolved through the format's alias table when the operation
/// is staged, just as a path step is.
///
/// ```
/// use litchi_ops::{IntoNode, NewNode, Placement};
///
/// let node = NewNode::new("paragraph")
///     .prop("text", "Summary")
///     .at(Placement::Start)
///     .into_node()?;
/// assert_eq!(node.kind(), "paragraph");
/// assert_eq!(node.props().len(), 1);
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Clone)]
pub struct NewNode<'a> {
    kind: Cow<'a, str>,
    props: Props<'a>,
    placement: Option<Placement<'a>>,
}

impl fmt::Debug for NewNode<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        // The type name and property values may be caller-authored content.
        // Keep the shape useful without copying them into logs.
        formatter
            .debug_struct("NewNode")
            .field("props", &self.props)
            .field("has_placement", &self.placement.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> NewNode<'a> {
    /// An element of type `kind`, with no assignments and end placement.
    #[must_use]
    pub fn new(kind: impl Into<Cow<'a, str>>) -> Self {
        Self {
            kind: kind.into(),
            props: Props::new(),
            placement: None,
        }
    }

    /// Adds one property assignment in insertion order.
    #[must_use]
    pub fn prop(mut self, key: impl Into<Key<'a>>, value: impl Into<PropValue<'a>>) -> Self {
        self.props.push(key, value);
        self
    }

    /// Adds all assignments from an [`IntoProps`] input.
    ///
    /// # Errors
    ///
    /// Returns the input conversion error before the node is used.
    pub fn with_props(mut self, props: impl IntoProps<'a>) -> Result<Self> {
        self.props.extend(props.into_props()?);
        Ok(self)
    }

    /// Places the element at the requested sibling or parent location.
    #[must_use]
    pub fn at(mut self, placement: Placement<'a>) -> Self {
        self.placement = Some(placement);
        self
    }

    /// The element type as supplied by the caller.
    #[must_use]
    pub fn kind(&self) -> &str {
        &self.kind
    }

    /// The ordered property assignments.
    #[must_use]
    pub fn props(&self) -> &Props<'a> {
        &self.props
    }

    /// The requested placement, or `None` for the parent's end.
    #[must_use]
    pub fn placement(&self) -> Option<&Placement<'a>> {
        self.placement.as_ref()
    }
}

/// A value accepted by `File::add` as the element to create.
///
/// The trait keeps borrowed type names and property values borrowed. It also
/// checks a node's optional placement before the leaf handler sees it.
///
/// ```
/// use litchi_ops::{IntoNode, NewNode};
///
/// let bare = "run".into_node()?;
/// let with_props = ("paragraph", [("bold", true)]).into_node()?;
/// assert!(bare.props().is_empty());
/// assert_eq!(with_props.props().len(), 1);
/// let reusable = NewNode::new("table");
/// let _ = (&reusable).into_node()?;
/// # Ok::<(), litchi_ops::Error>(())
/// ```
pub trait IntoNode<'a> {
    /// Converts this input into a checked node description.
    ///
    /// # Errors
    ///
    /// [`Code::MissingArgument`] for an empty type name,
    /// [`Code::InvalidArgument`] for a control character or an oversized type
    /// name, and the conversion errors from [`IntoProps`].
    fn into_node(self) -> Result<NewNode<'a>>;
}

fn checked(node: NewNode<'_>) -> Result<NewNode<'_>> {
    if node.kind.trim().is_empty() {
        return Err(
            Error::new(Code::MissingArgument, "the element to add has no type")
                .with_suggestion("name its type, e.g. paragraph, table, slide or sheet"),
        );
    }
    if node.kind.len() > limits::MAX_INPUT_BYTES {
        return Err(Error::new(
            Code::InvalidArgument,
            "the element type exceeds the input byte limit",
        ));
    }
    if node.kind.chars().any(char::is_control) {
        return Err(Error::new(
            Code::InvalidArgument,
            "the element type contains a control character",
        ));
    }
    if let Some(placement) = node.placement.clone() {
        placement.into_placement()?;
    }
    Ok(node)
}

impl<'a> IntoNode<'a> for NewNode<'a> {
    fn into_node(self) -> Result<NewNode<'a>> {
        checked(self)
    }
}

impl<'a> IntoNode<'a> for &NewNode<'a> {
    fn into_node(self) -> Result<NewNode<'a>> {
        checked(self.clone())
    }
}

impl<'a> IntoNode<'a> for &'a str {
    fn into_node(self) -> Result<NewNode<'a>> {
        checked(NewNode::new(self))
    }
}

impl<'a> IntoNode<'a> for &'a String {
    fn into_node(self) -> Result<NewNode<'a>> {
        checked(NewNode::new(self.as_str()))
    }
}

impl<'a> IntoNode<'a> for String {
    fn into_node(self) -> Result<NewNode<'a>> {
        checked(NewNode::new(self))
    }
}

impl<'a, K, P> IntoNode<'a> for (K, P)
where
    K: Into<Cow<'a, str>>,
    P: IntoProps<'a>,
{
    fn into_node(self) -> Result<NewNode<'a>> {
        let (kind, props) = self;
        checked(NewNode::new(kind).with_props(props)?)
    }
}

/// The result of `File::add`: the new file snapshot and the created node's
/// canonical path. It dereferences to the new [`File`] for read chaining.
///
/// ```
/// use litchi_ops::{Added, File};
///
/// fn take_snapshot(added: Added) -> File {
///     added.into_file()
/// }
///
/// fn borrow_snapshot(added: &Added) -> &File {
///     added
/// }
/// # let _ = (take_snapshot, borrow_snapshot);
/// ```
#[derive(Clone)]
#[non_exhaustive]
pub struct Added {
    /// The new file snapshot.
    pub file: File,
    /// The created node's canonical path.
    pub path: Path<'static>,
}

impl fmt::Debug for Added {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Added")
            .field("file", &self.file)
            .field("path_steps", &self.path.steps().len())
            .finish_non_exhaustive()
    }
}

impl Added {
    /// Returns the new file snapshot and drops the path.
    #[must_use]
    pub fn into_file(self) -> File {
        self.file
    }
}

impl Deref for Added {
    type Target = File;

    fn deref(&self) -> &Self::Target {
        &self.file
    }
}

#[cfg(test)]
mod tests {
    use super::{IntoNode, NewNode};
    use crate::Code;
    use crate::path::PathArg;
    use crate::placement::Placement;
    use crate::props::PropValue;

    #[test]
    fn node_inputs_keep_order_and_support_owned_and_borrowed_forms() {
        let node = NewNode::new("paragraph")
            .prop("text", "Hi")
            .with_props([("bold", true)])
            .expect("properties");
        assert_eq!(node.kind(), "paragraph");
        assert_eq!(node.props().len(), 2);
        assert_eq!(node.props().get("text"), Some(&PropValue::from("Hi")));
        assert_eq!("table".into_node().expect("type").kind(), "table");
        let owned = String::from("slide");
        assert_eq!(owned.into_node().expect("owned type").kind(), "slide");
        let borrowed = String::from("sheet");
        assert_eq!(
            (&borrowed).into_node().expect("borrowed type").kind(),
            "sheet"
        );
    }

    #[test]
    fn node_inputs_refuse_empty_control_and_oversized_types() {
        for kind in ["", "  ", "bad\nname"] {
            assert_eq!(
                kind.into_node().expect_err("invalid type").code(),
                if kind.trim().is_empty() {
                    Code::MissingArgument
                } else {
                    Code::InvalidArgument
                }
            );
        }
        let oversized = "x".repeat(litchi_core::selector::limits::MAX_INPUT_BYTES + 1);
        assert_eq!(
            oversized.into_node().expect_err("oversized type").code(),
            Code::InvalidArgument
        );
    }

    #[test]
    fn node_debug_redacts_type_and_property_content() {
        let node = NewNode::new("PRIVATE_KIND").prop("text", "PRIVATE_TEXT");
        let debug = format!("{node:?}");
        assert!(!debug.contains("PRIVATE_KIND"));
        assert!(!debug.contains("PRIVATE_TEXT"));
    }

    #[test]
    fn node_placement_is_validated_during_node_admission() {
        let path = PathArg::parse("/body/p[@text!=PRIVATE]").expect("path");
        let error = NewNode::new("paragraph")
            .at(Placement::Before(path))
            .into_node()
            .expect_err("content placement");
        assert_eq!(error.code(), Code::InvalidPath);
    }
}
