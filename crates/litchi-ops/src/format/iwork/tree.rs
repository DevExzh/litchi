//! The object tree of one iWork package: an arena of items in document
//! order, each with its kind, properties aligned with its spec table, text,
//! and the leaf positions writes address it by.
//!
//! The shape follows the ODF object-tree adapter (`format::odf_tree`,
//! record 0282); the iWork adapter keeps its own copy so the two families
//! evolve independently (record 0296).

use litchi_core::Position;
use std::sync::OnceLock;

use crate::node::spec::{ElementSpec, PropSpec};
use crate::node::{PathStep, Value};
use crate::path::navigate::ElementKind;

/// The index of the root item (of a part, its anchor: the item standing in
/// for the eager item whose children the part holds).
pub(crate) const ROOT: u32 = 0;

/// Namespace bit for the private chart-only path cache. It is deliberately
/// outside the normal part arena, whose parts are assigned from one.
pub(crate) const CHART_PART: u32 = u32::MAX;

/// Where an item lives: its part (0 for the tree built on the first read,
/// `k` for the `k`-th lazily built part) and its index there.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct At {
    pub(crate) part: u32,
    pub(crate) index: u32,
}

impl At {
    /// The root of the tree.
    pub(crate) const ROOT: Self = Self::new(0, ROOT);

    pub(crate) const fn new(part: u32, index: u32) -> Self {
        Self { part, index }
    }

    pub(crate) const fn chart(index: u32) -> Self {
        Self::new(CHART_PART, index)
    }

    pub(crate) const fn is_chart(self) -> bool {
        self.part == CHART_PART
    }
}

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
    /// Values whose focused leaf read is deferred until the node view asks
    /// for them (currently chart font-size).  A failed focused read is not
    /// cached: the caller receives its typed leaf error and no refusal is
    /// flattened into an absent property.
    lazy_values: OnceLock<Box<[Option<Value<'static>>]>>,
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
    /// The item's position among its parent's children of its kind.
    ordinal: u32,
    /// The lazily built part holding the item's children (record 0305), or
    /// 0 when they are in its own part.
    part: u32,
}

impl Item {
    /// An item of `spec` at `path` outside any builder: a lazily
    /// materialized element (a Keynote slide, a Numbers cell, record 0305).
    #[cfg(any(feature = "keynote", feature = "numbers"))]
    pub(crate) fn detached<K: ElementKind>(
        spec: &'static ElementSpec<K>,
        path: Box<[PathStep<'static>]>,
        ordinal: u32,
    ) -> Self {
        Self {
            kind: spec.kind.ordinal(),
            type_name: spec.name,
            props: spec.props,
            values: vec![None; spec.props.len()].into_boxed_slice(),
            lazy_values: OnceLock::new(),
            text: None,
            label: None,
            children: Vec::new(),
            path,
            leaf: [0; 4],
            ordinal,
            part: 0,
        }
    }

    /// Set the item's properties, text and leaf positions.
    #[cfg(any(feature = "keynote", feature = "numbers"))]
    pub(crate) fn draft<K>(&mut self) -> Draft<'_, K> {
        Draft {
            item: self,
            index: 0,
            kind: std::marker::PhantomData,
        }
    }

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

    pub(crate) fn lazy_values(&self) -> &OnceLock<Box<[Option<Value<'static>>]>> {
        &self.lazy_values
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

    pub(crate) const fn ordinal(&self) -> u32 {
        self.ordinal
    }

    /// The lazily built part holding the item's children; 0 when they are
    /// in its own part.
    pub(crate) const fn part(&self) -> u32 {
        self.part
    }

    /// The value of the property `key`, if set.
    pub(crate) fn value(&self, key: &str) -> Option<&Value<'static>> {
        let index = self.props.iter().position(|prop| prop.name == key)?;
        self.values.get(index)?.as_ref()
    }
}

/// The items of one package, the root first; or of one lazily built part,
/// its anchor first.
#[derive(Debug)]
pub(crate) struct Tree {
    items: Vec<Item>,
    /// The first item pushed after [`TreeBuilder::defer`]: a part's
    /// deferred items (Keynote character runs), which the document-order
    /// views list after every part's other items.
    #[cfg_attr(
        not(feature = "keynote"),
        allow(dead_code, reason = "only Keynote parts defer items")
    )]
    deferred: u32,
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
                lazy_values: OnceLock::new(),
                text: None,
                label: None,
                children: Vec::new(),
                path: Box::new([]),
                leaf: [0; 4],
                ordinal: 0,
                part: 0,
            }],
            deferred: 1,
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

    /// The index of the first deferred item ([`TreeBuilder::defer`]).
    #[cfg(feature = "keynote")]
    pub(crate) const fn deferred(&self) -> u32 {
        self.deferred
    }
}

/// Builds a [`Tree`] in document order.
pub(crate) struct TreeBuilder<K: 'static> {
    items: Vec<Item>,
    elements: &'static [ElementSpec<K>],
    singletons: &'static [K],
    /// The children pushed so far, by parent and kind: a push's positional
    /// step is its ordinal among them (counting would make building
    /// quadratic in a large parent's child count).
    ordinals: std::collections::HashMap<(u32, u8), usize>,
    /// Where the deferred items start, once [`Self::defer`] was called.
    deferred: Option<u32>,
}

/// One item under construction, returned by [`TreeBuilder::push`].
pub(crate) struct Draft<'b, K: 'static> {
    item: &'b mut Item,
    index: u32,
    kind: std::marker::PhantomData<fn() -> K>,
}

impl<K: ElementKind> TreeBuilder<K> {
    /// A tree whose root has the spec table's root kind.
    pub(crate) fn new(elements: &'static [ElementSpec<K>], singletons: &'static [K]) -> Self {
        Self::under(elements, singletons, K::ROOT, &[])
    }

    /// A lazily built part holding the children of an item of `kind` at
    /// `path`: its anchor (index [`ROOT`]) stands in for that item, so the
    /// children pushed under it take their paths and ordinals from it.
    pub(crate) fn under(
        elements: &'static [ElementSpec<K>],
        singletons: &'static [K],
        kind: K,
        path: &[PathStep<'static>],
    ) -> Self {
        let spec = &elements[usize::from(kind.ordinal())];
        let item = Item {
            kind: kind.ordinal(),
            type_name: spec.name,
            props: spec.props,
            values: vec![None; spec.props.len()].into_boxed_slice(),
            lazy_values: OnceLock::new(),
            text: None,
            label: None,
            children: Vec::new(),
            path: path.into(),
            leaf: [0; 4],
            ordinal: 0,
            part: 0,
        };
        Self {
            items: vec![item],
            elements,
            singletons,
            ordinals: std::collections::HashMap::new(),
            deferred: None,
        }
    }

    /// The contents of a table whose holding tree lists `stand_in`: the
    /// anchor (index [`ROOT`]) is a copy of the stand-in, whose properties
    /// the table's reader completes, and the rows and cells pushed under it
    /// take their paths and ordinals from it (record 0312).
    #[cfg(any(feature = "keynote", feature = "pages"))]
    pub(crate) fn from_stand_in(
        elements: &'static [ElementSpec<K>],
        singletons: &'static [K],
        stand_in: &Item,
    ) -> Self {
        let item = Item {
            kind: stand_in.kind,
            type_name: stand_in.type_name,
            props: stand_in.props,
            values: stand_in.values.clone(),
            lazy_values: OnceLock::new(),
            text: stand_in.text.clone(),
            label: stand_in.label.clone(),
            children: Vec::new(),
            path: stand_in.path.clone(),
            leaf: stand_in.leaf,
            ordinal: stand_in.ordinal,
            part: 0,
        };
        Self {
            items: vec![item],
            elements,
            singletons,
            ordinals: std::collections::HashMap::new(),
            deferred: None,
        }
    }

    /// How many items the tree holds so far.
    #[cfg(feature = "keynote")]
    pub(crate) fn len(&self) -> usize {
        self.items.len()
    }

    /// The root item, for its properties.
    pub(crate) fn root(&mut self) -> Draft<'_, K> {
        self.draft_at(ROOT)
    }

    fn draft_at(&mut self, index: u32) -> Draft<'_, K> {
        Draft {
            item: &mut self.items[index as usize],
            index,
            kind: std::marker::PhantomData,
        }
    }

    /// Mark the items pushed from now on as deferred ([`Tree::deferred`]).
    #[cfg(feature = "keynote")]
    pub(crate) fn defer(&mut self) {
        self.deferred = Some(u32::try_from(self.items.len()).unwrap_or(u32::MAX));
    }

    /// The canonical path of an item already built.
    #[cfg(feature = "numbers")]
    pub(crate) fn path_of(&self, index: u32) -> Box<[PathStep<'static>]> {
        self.items[index as usize].path.clone()
    }

    /// The text of an item already built.
    #[cfg(feature = "keynote")]
    pub(crate) fn item_text(&self, index: u32) -> &str {
        self.items[index as usize]
            .text
            .as_deref()
            .unwrap_or_default()
    }

    /// An item already built, for more properties.
    #[cfg(feature = "pages")]
    pub(crate) fn draft(&mut self, index: u32) -> Draft<'_, K> {
        self.draft_at(index)
    }

    /// Append an item of `kind` under `parent`. A cell's path step is its
    /// A1 address (`step`); every other kind's is its positional step.
    pub(crate) fn push(&mut self, parent: u32, kind: K) -> Draft<'_, K> {
        let ordinal = self
            .ordinals
            .get(&(parent, kind.ordinal()))
            .copied()
            .unwrap_or(0);
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
            lazy_values: OnceLock::new(),
            text: None,
            label: None,
            children: Vec::new(),
            path: path.into_boxed_slice(),
            leaf: [0; 4],
            ordinal: 0,
            part: 0,
        });
        self.items[parent as usize].children.push(index);
        let ordinal = self.ordinals.entry((parent, kind.ordinal())).or_insert(0);
        self.items[index as usize].ordinal = u32::try_from(*ordinal).unwrap_or(u32::MAX);
        *ordinal += 1;
        self.draft_at(index)
    }

    pub(crate) fn build(self) -> Tree {
        let length = u32::try_from(self.items.len()).unwrap_or(u32::MAX);
        Tree {
            items: self.items,
            deferred: self.deferred.unwrap_or(length),
        }
    }
}

impl<K> Draft<'_, K> {
    fn item(&mut self) -> &mut Item {
        &mut *self.item
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

    /// Hold the item's children in lazily built part `part`.
    #[cfg(any(feature = "keynote", feature = "numbers", feature = "pages"))]
    pub(crate) fn part(mut self, part: u32) -> Self {
        self.item().part = part;
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
