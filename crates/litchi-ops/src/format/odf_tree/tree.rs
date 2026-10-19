//! The object tree of one package: an arena of items in document order,
//! each with its kind, properties aligned with its spec table, text, and
//! the leaf positions writes address it by.

use litchi_core::Position;

use crate::node::spec::{ElementSpec, PropSpec};
use crate::node::{PathStep, Value};
use crate::path::navigate::ElementKind;

/// The index of the root item.
pub(crate) const ROOT: u32 = 0;

/// One element of the tree.
#[derive(Debug)]
pub(crate) struct Item {
    /// The kind's ordinal in its flavor's spec table.
    kind: u8,
    /// The OfficeCLI type name (the spec element's name).
    type_name: &'static str,
    /// The kind's readable properties.
    props: &'static [PropSpec],
    /// Values aligned with `props`.
    values: Box<[Option<Value<'static>>]>,
    /// The item's own text, if it has any.
    text: Option<String>,
    /// The name a bare `[NAME]` step matches.
    label: Option<String>,
    children: Vec<u32>,
    /// The zero-based position among its parent's children of its kind.
    ordinal: u32,
    /// The canonical positional path.
    path: Box<[PathStep<'static>]>,
    /// The leaf positions of the element (page and shape, frame, table and
    /// column, …), meaning per flavor.
    #[cfg(any(feature = "odg", feature = "odc", feature = "odi", feature = "odb"))]
    leaf: [usize; 2],
    /// The leaf's own tree path (MathML), when the flavor needs one.
    #[cfg(feature = "odf-formula")]
    route: Box<[usize]>,
}

impl Item {
    pub(crate) const fn kind(&self) -> u8 {
        self.kind
    }

    pub(crate) const fn type_name(&self) -> &'static str {
        self.type_name
    }

    pub(crate) const fn props(&self) -> &'static [PropSpec] {
        self.props
    }

    pub(crate) fn values(&self) -> &[Option<Value<'static>>] {
        &self.values
    }

    pub(crate) fn text(&self) -> Option<&str> {
        self.text.as_deref()
    }

    pub(crate) fn label(&self) -> Option<&str> {
        self.label.as_deref()
    }

    pub(crate) fn children(&self) -> &[u32] {
        &self.children
    }

    /// The zero-based position among its parent's children of its kind.
    pub(crate) const fn ordinal(&self) -> u32 {
        self.ordinal
    }

    pub(crate) fn path(&self) -> &[PathStep<'static>] {
        &self.path
    }

    #[cfg(any(feature = "odg", feature = "odc", feature = "odi", feature = "odb"))]
    pub(crate) const fn leaf(&self) -> [usize; 2] {
        self.leaf
    }

    #[cfg(feature = "odf-formula")]
    pub(crate) fn route(&self) -> &[usize] {
        &self.route
    }

    /// The value of the property `key`, if set.
    pub(crate) fn value(&self, key: &str) -> Option<&Value<'static>> {
        let index = self.props.iter().position(|prop| prop.name == key)?;
        self.values.get(index)?.as_ref()
    }
}

/// The items of one package, the root first.
#[derive(Debug)]
pub(crate) struct Tree {
    items: Vec<Item>,
}

impl Tree {
    /// A tree of only an empty root item.
    pub(crate) fn root_only() -> Self {
        Self {
            items: vec![Item {
                kind: 0,
                type_name: "",
                props: &[],
                values: Box::new([]),
                text: None,
                label: None,
                children: Vec::new(),
                ordinal: 0,
                path: Box::new([]),
                #[cfg(any(feature = "odg", feature = "odc", feature = "odi", feature = "odb"))]
                leaf: [0, 0],
                #[cfg(feature = "odf-formula")]
                route: Box::new([]),
            }],
        }
    }

    pub(crate) fn items(&self) -> &[Item] {
        &self.items
    }

    pub(crate) fn len(&self) -> usize {
        self.items.len()
    }

    pub(crate) fn item(&self, index: u32) -> &Item {
        &self.items[index as usize]
    }
}

/// Builds a [`Tree`] in document order.
pub(crate) struct TreeBuilder<K: 'static> {
    items: Vec<Item>,
    /// Children pushed so far per (parent, kind ordinal) (record 0340: a
    /// push counts in constant time).
    counts: std::collections::HashMap<(u32, u8), usize>,
    elements: &'static [ElementSpec<K>],
    singletons: &'static [K],
}

/// One item under construction, returned by [`TreeBuilder::push`].
pub(crate) struct Draft<'b, K: 'static> {
    builder: &'b mut TreeBuilder<K>,
    index: u32,
}

impl<K: ElementKind> TreeBuilder<K> {
    /// A tree whose root has the spec table's root kind.
    pub(crate) fn new(elements: &'static [ElementSpec<K>], singletons: &'static [K]) -> Self {
        let root = &elements[usize::from(K::ROOT.ordinal())];
        let item = Item {
            kind: K::ROOT.ordinal(),
            type_name: root.name,
            props: root.props,
            values: vec![None; root.props.len()].into_boxed_slice(),
            text: None,
            label: None,
            children: Vec::new(),
            ordinal: 0,
            path: Box::new([]),
            #[cfg(any(feature = "odg", feature = "odc", feature = "odi", feature = "odb"))]
            leaf: [0, 0],
            #[cfg(feature = "odf-formula")]
            route: Box::new([]),
        };
        Self {
            items: vec![item],
            counts: std::collections::HashMap::new(),
            elements,
            singletons,
        }
    }

    /// The root item, for its properties.
    pub(crate) fn root(&mut self) -> Draft<'_, K> {
        Draft {
            builder: self,
            index: ROOT,
        }
    }

    /// Append an item of `kind` under `parent`.
    pub(crate) fn push(&mut self, parent: u32, kind: K) -> Draft<'_, K> {
        let spec = &self.elements[usize::from(kind.ordinal())];
        let count = self.counts.entry((parent, kind.ordinal())).or_insert(0);
        let ordinal = *count;
        *count = count.saturating_add(1);
        let step = if self.singletons.contains(&kind) {
            PathStep::new(kind.name())
        } else {
            PathStep::at(kind.name(), Position::new(ordinal))
        };
        let mut path = Vec::with_capacity(self.items[parent as usize].path.len() + 1);
        path.extend_from_slice(&self.items[parent as usize].path);
        path.push(step);
        let index = u32::try_from(self.items.len()).unwrap_or(u32::MAX);
        self.items.push(Item {
            kind: kind.ordinal(),
            type_name: spec.name,
            props: spec.props,
            values: vec![None; spec.props.len()].into_boxed_slice(),
            text: None,
            label: None,
            children: Vec::new(),
            ordinal: u32::try_from(ordinal).unwrap_or(u32::MAX),
            path: path.into_boxed_slice(),
            #[cfg(any(feature = "odg", feature = "odc", feature = "odi", feature = "odb"))]
            leaf: [0, 0],
            #[cfg(feature = "odf-formula")]
            route: Box::new([]),
        });
        self.items[parent as usize].children.push(index);
        Draft {
            builder: self,
            index,
        }
    }

    pub(crate) fn build(self) -> Tree {
        Tree { items: self.items }
    }
}

impl<K> Draft<'_, K> {
    fn item(&mut self) -> &mut Item {
        &mut self.builder.items[self.index as usize]
    }

    /// Set the property `key` when `value` is present.
    ///
    /// # Panics
    ///
    /// When the kind declares no `key` (a programming error caught by the
    /// flavor tests).
    pub(crate) fn prop(mut self, key: &str, value: Option<Value<'static>>) -> Self {
        if let Some(value) = value {
            let item = self.item();
            let index = item
                .props
                .iter()
                .position(|prop| prop.name == key)
                .unwrap_or_else(|| panic!("`{}` declares no `{key}` property", item.type_name));
            item.values[index] = Some(value);
        }
        self
    }

    /// Set the item's text.
    #[cfg(any(
        feature = "odg",
        feature = "odc",
        feature = "odb",
        feature = "odf-formula"
    ))]
    pub(crate) fn text(mut self, text: impl Into<String>) -> Self {
        self.item().text = Some(text.into());
        self
    }

    /// Set the name a bare `[NAME]` step matches.
    #[cfg(any(feature = "odg", feature = "odc", feature = "odi", feature = "odb"))]
    pub(crate) fn label(mut self, label: Option<&str>) -> Self {
        self.item().label = label.map(str::to_owned);
        self
    }

    /// Record the leaf positions of the element.
    #[cfg(any(feature = "odg", feature = "odc", feature = "odi", feature = "odb"))]
    pub(crate) fn leaf(mut self, first: usize, second: usize) -> Self {
        self.item().leaf = [first, second];
        self
    }

    /// Record the leaf's own tree path of the element.
    #[cfg(feature = "odf-formula")]
    pub(crate) fn route(mut self, route: &[usize]) -> Self {
        self.item().route = route.into();
        self
    }

    /// Finish the item, returning its index.
    pub(crate) const fn done(self) -> u32 {
        self.index
    }
}

/// A text property value.
pub(crate) fn text(value: Option<&str>) -> Option<Value<'static>> {
    value.map(|value| Value::Text(value.to_owned().into()))
}

/// A count property value.
pub(crate) fn count(value: usize) -> Option<Value<'static>> {
    Some(Value::Int(i64::try_from(value).unwrap_or(i64::MAX)))
}

/// An ODF length attribute as a unified length; an unparsable value is
/// absent.
#[cfg(any(feature = "odg", feature = "odc", feature = "odi"))]
pub(crate) fn length(value: Option<&str>) -> Option<Value<'static>> {
    crate::value::Length::parse(value?.trim())
        .ok()
        .map(Value::Length)
}

/// An ODF length attribute value for `length`: the canonical unified
/// spelling, or centimetres with six decimals when no two-decimal unit is
/// exact (the spelling of the ODP adapter, record 0278).
#[cfg(any(feature = "odg", feature = "odc", feature = "odi"))]
pub(crate) fn odf_length(length: crate::value::Length) -> String {
    let display = length.display(crate::value::LengthStyle::Auto).to_string();
    if !display.ends_with("emu") {
        return display;
    }
    let micro = i128::from(length.emu()) * 1_000_000 / 360_000;
    let negative = micro < 0;
    let micro = micro.unsigned_abs();
    let mut text = format!(
        "{}{}.{:06}",
        if negative { "-" } else { "" },
        micro / 1_000_000,
        micro % 1_000_000
    );
    while text.ends_with('0') {
        text.pop();
    }
    if text.ends_with('.') {
        text.pop();
    }
    text.push_str("cm");
    text
}
