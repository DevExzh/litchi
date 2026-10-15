//! The object tree of one iWork package: an arena of items in document
//! order, each with its kind, properties aligned with its spec table, text,
//! and the leaf positions writes address it by.
//!
//! The shape follows the ODF object-tree adapter (`format::odf_tree`,
//! record 0282); the iWork adapter keeps its own copy so the two families
//! evolve independently (record 0296).

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
    /// The name a bare `[NAME]` step (or a `[TOKEN]` step) matches.
    label: Option<String>,
    children: Vec<u32>,
    /// The canonical positional path.
    path: Box<[PathStep<'static>]>,
    /// The leaf positions of the element, meaning per flavor: section and
    /// paragraph (Pages), slide and text role (Keynote), sheet, table, row
    /// and column (Numbers).
    leaf: [u32; 4],
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

    pub(crate) fn path(&self) -> &[PathStep<'static>] {
        &self.path
    }

    pub(crate) const fn leaf(&self) -> [u32; 4] {
        self.leaf
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
    /// A tree of only an empty root item, standing in for a tree whose
    /// build failed.
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
                path: Box::new([]),
                leaf: [0; 4],
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
            path: Box::new([]),
            leaf: [0; 4],
        };
        Self {
            items: vec![item],
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

    /// An item already built, for more properties.
    #[cfg(feature = "pages")]
    pub(crate) fn draft(&mut self, index: u32) -> Draft<'_, K> {
        Draft {
            builder: self,
            index,
        }
    }

    /// Append an item of `kind` under `parent`. A cell's path step is its
    /// A1 address (`step`); every other kind's is its positional step.
    pub(crate) fn push(&mut self, parent: u32, kind: K) -> Draft<'_, K> {
        let ordinal = self.items[parent as usize]
            .children
            .iter()
            .filter(|child| self.items[**child as usize].kind == kind.ordinal())
            .count();
        let step = if self.singletons.contains(&kind) {
            PathStep::new(kind.name())
        } else {
            PathStep::at(kind.name(), Position::new(ordinal))
        };
        self.push_step(parent, kind, step)
    }

    /// Append an item of `kind` under `parent` whose path step is `step`.
    pub(crate) fn push_step(
        &mut self,
        parent: u32,
        kind: K,
        step: PathStep<'static>,
    ) -> Draft<'_, K> {
        let spec = &self.elements[usize::from(kind.ordinal())];
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
            path: path.into_boxed_slice(),
            leaf: [0; 4],
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
    pub(crate) fn text(mut self, text: impl Into<String>) -> Self {
        self.item().text = Some(text.into());
        self
    }

    /// Set the name a bare `[NAME]` step matches.
    pub(crate) fn label(mut self, label: Option<&str>) -> Self {
        self.item().label = label.map(str::to_owned);
        self
    }

    /// Record the leaf positions of the element.
    pub(crate) fn leaf(mut self, leaf: [usize; 4]) -> Self {
        self.item().leaf = leaf.map(|value| u32::try_from(value).unwrap_or(u32::MAX));
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
