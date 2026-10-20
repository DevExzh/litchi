//! Node views as Python dicts (record 0010 §2.1, §8.2).
//!
//! `File.get`, `File.find` and `File.query` return the generic node view of
//! each node they reach: OfficeCLI's `DocumentNode` shape, as a `dict` with
//! the members `path`, `type`, `text`, `preview`, `style`, `childCount`,
//! `format` and `children`, in that order. `text`, `preview`, `style` and
//! `childCount` are left out when the node has none, as the JSON form omits
//! them, so a dict equals `json.loads` of the node view's JSON.
//!
//! A call runs in two phases:
//!
//! 1. **GIL released.** Resolve the path or run the query, list the children
//!    of every node down to the requested depth (level by level), and read
//!    every member of every node through the node view's accessors. The
//!    values stay borrowed from the nodes where the leaf allows (`Cow`), so
//!    nothing is copied and no JSON is written.
//! 2. **GIL held.** Build the dicts bottom-up from what phase 1 read: only
//!    Python object construction happens here. Paths, previews, lengths and
//!    colours print through `Display` into one reused buffer, and the
//!    strings a result repeats (format keys, element types, tokens such as
//!    `"center"`) are created once per call and shared, as `json.loads`
//!    shares repeated keys.
//!
//! One budget, `max_nodes`, bounds the nodes a whole call returns, matches
//! and descendants together: one node past it is `output_too_large`, as
//! for one node view in Rust (ADR 0006 bounded output).

use std::borrow::Cow;
use std::collections::HashMap;
use std::fmt::{self, Write as _};
use std::hash::{BuildHasherDefault, Hasher};
use std::ops::Range;

use litchi::node::Node;
use litchi::node::{Preview, Prop, Value, ViewOptions};
use litchi::query::Diagnostic;
use pyo3::intern;
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDict, PyFloat, PyList, PyString};

use crate::errors::Failure;

/// The node view of the node at `path`.
pub(crate) fn get<'py>(
    py: Python<'py>,
    file: &litchi::File,
    path: &str,
    options: ViewOptions,
) -> PyResult<Bound<'py, PyDict>> {
    let mut views = collect(py, options, || {
        file.get(path)
            .map(|node| std::iter::once(Ok(node)))
            .map_err(|error| Failure::read(&error))
    })?;
    match views.pop() {
        Some(view) if views.is_empty() => Ok(view),
        _ => Err(Failure::new(
            litchi::Code::InternalError,
            "a path resolved to other than one node",
        )
        .into()),
    }
}

/// The node view of the node at `path`, or `None` when nothing is there.
pub(crate) fn find<'py>(
    py: Python<'py>,
    file: &litchi::File,
    path: &str,
    options: ViewOptions,
) -> PyResult<Option<Bound<'py, PyDict>>> {
    let mut views = collect(py, options, || {
        file.find(path)
            .map(|node| node.into_iter().map(Ok))
            .map_err(|error| Failure::read(&error))
    })?;
    Ok(views.pop())
}

/// The node views of the first `limit` matches of `selector`, in document
/// order. The query's compile-time diagnostics (an unknown key, with its
/// did-you-mean) are issued as `UserWarning`s, as OfficeCLI reports them
/// as warnings.
pub(crate) fn query<'py>(
    py: Python<'py>,
    file: &litchi::File,
    selector: &str,
    limit: Option<usize>,
    options: ViewOptions,
) -> PyResult<Bound<'py, PyList>> {
    let mut diagnostics: Vec<Diagnostic> = Vec::new();
    let found = &mut diagnostics;
    let views = collect(py, options, move || {
        let matches = file
            .query(selector)
            .map_err(|error| Failure::read(&error))?;
        found.extend_from_slice(matches.diagnostics());
        let matches = match limit {
            Some(limit) => matches.limit(limit),
            None => matches,
        };
        Ok(matches.map(|node| node.map_err(|error| Failure::read(&error))))
    })?;
    if !diagnostics.is_empty() {
        let category = py.get_type::<pyo3::exceptions::PyUserWarning>();
        for diagnostic in &diagnostics {
            let message = std::ffi::CString::new(diagnostic.to_string().replace('\0', " "))?;
            PyErr::warn(py, &category, &message, 1)?;
        }
    }
    PyList::new(py, views)
}

/// Lists the nodes `roots` returns and their descendants, reads their view
/// members with the GIL released, then builds their dicts.
///
/// `roots` runs with the GIL released; its nodes borrow `'f`, the file the
/// caller holds.
fn collect<'py, 'f, F, I>(
    py: Python<'py>,
    options: ViewOptions,
    roots: F,
) -> PyResult<Vec<Bound<'py, PyDict>>>
where
    F: FnOnce() -> Result<I, Failure> + Send,
    I: Iterator<Item = Result<Node<'f>, Failure>>,
{
    let mut levels: Vec<Level<'f>> = Vec::new();
    let slot = &mut levels;
    let projection = py.detach(move || -> Result<Projection<'_, 'f>, Failure> {
        let mut emitted = 0;
        let mut nodes = Vec::new();
        for node in roots()? {
            let node = node?;
            spend(&mut emitted, options.max_nodes)?;
            nodes.push(node);
        }
        slot.push(Level::new(nodes));
        expand(slot, options, &mut emitted)?;
        // The levels are complete: from here on they are only read.
        let levels: &[Level<'f>] = slot;
        project(levels, options)
    })?;
    projection.build(py)
}

/// One level of a result: level 0 holds the nodes the call reached, level
/// `k + 1` the children of level `k`'s nodes, in document order.
struct Level<'f> {
    nodes: Vec<Node<'f>>,
    /// For each node, the end of its children in the next level: they are
    /// `next[ends[i - 1]..ends[i]]`. Empty on the last level.
    ends: Vec<usize>,
}

impl<'f> Level<'f> {
    fn new(nodes: Vec<Node<'f>>) -> Self {
        Self {
            nodes,
            ends: Vec::new(),
        }
    }
}

/// Counts one more node against the call's budget of `max_nodes`.
fn spend(emitted: &mut u32, max_nodes: u32) -> Result<(), Failure> {
    if *emitted >= max_nodes {
        return Err(Failure::new(
            litchi::Code::OutputTooLarge,
            format!("the node views exceed their budget of {max_nodes} nodes"),
        )
        .with_suggestion("lower depth, pass a limit to query, or raise max_nodes"));
    }
    *emitted += 1;
    Ok(())
}

/// Lists the children of every level, down to `options.depth` levels below
/// level 0. A node `k` levels down is viewed with `depth - k` levels left,
/// as the Rust serializer views it.
fn expand<'f>(
    levels: &mut Vec<Level<'f>>,
    options: ViewOptions,
    emitted: &mut u32,
) -> Result<(), Failure> {
    for level in 0..options.depth as usize {
        let Some(parents) = levels.get(level) else {
            break;
        };
        let remaining = options.depth - level as u32;
        let viewed = options.depth(remaining);
        let mut next = Vec::new();
        let mut ends = Vec::with_capacity(parents.nodes.len());
        for node in &parents.nodes {
            for child in node.view_with(viewed).children() {
                let child = child.map_err(|error| Failure::read(&error))?;
                spend(emitted, options.max_nodes)?;
                next.push(child);
            }
            ends.push(next.len());
        }
        levels[level].ends = ends;
        if next.is_empty() {
            break;
        }
        levels.push(Level::new(next));
    }
    Ok(())
}

/// What phase 1 read of one node: every member but `path`, `type` and
/// `children`, which the dict builder takes from the node and the next
/// level.
struct Record<'n> {
    text: Option<Cow<'n, str>>,
    preview: Option<Preview<'n>>,
    style: Option<Cow<'n, str>>,
    child_count: Option<u32>,
    /// The node's `format` entries, as a range of [`Projection::props`].
    format: Range<usize>,
}

/// Every node of a result and its members, read with the GIL released.
struct Projection<'n, 'f> {
    levels: &'n [Level<'f>],
    /// One record per node, level by level.
    records: Vec<Vec<Record<'n>>>,
    /// Every node's `format` entries, in node order.
    props: Vec<Prop<'n>>,
}

/// Reads the members of every node, in the order the Rust serializer reads
/// them: `text`, `preview`, `style`, `childCount`, then `format`.
fn project<'n, 'f>(
    levels: &'n [Level<'f>],
    options: ViewOptions,
) -> Result<Projection<'n, 'f>, Failure> {
    let read = |error: litchi::Error| Failure::read(&error);
    let mut records = Vec::with_capacity(levels.len());
    let mut props = Vec::new();
    for (index, level) in levels.iter().enumerate() {
        let remaining = options
            .depth
            .saturating_sub(u32::try_from(index).unwrap_or(u32::MAX));
        let viewed = options.depth(remaining);
        let mut level_records = Vec::with_capacity(level.nodes.len());
        for node in &level.nodes {
            let view = node.view_with(viewed);
            let text = view.text().map_err(read)?;
            let preview = view.preview().map_err(read)?;
            let style = view.style().map_err(read)?;
            let child_count = view.child_count().map_err(read)?;
            let start = props.len();
            props.extend(view.format().map_err(read)?);
            level_records.push(Record {
                text,
                preview,
                style,
                child_count,
                format: start..props.len(),
            });
        }
        records.push(level_records);
    }
    Ok(Projection {
        levels,
        records,
        props,
    })
}

impl Projection<'_, '_> {
    /// The dicts of level 0, built bottom-up so that each node's `children`
    /// list is ready when its dict is built.
    fn build<'py>(&self, py: Python<'py>) -> PyResult<Vec<Bound<'py, PyDict>>> {
        let mut out = Out {
            py,
            scratch: String::new(),
            statics: HashMap::default(),
        };
        let mut below: Vec<Bound<'py, PyDict>> = Vec::new();
        for (level, records) in self.levels.iter().zip(&self.records).rev() {
            let mut dicts = Vec::with_capacity(level.nodes.len());
            let mut start = 0;
            for (index, (node, record)) in level.nodes.iter().zip(records).enumerate() {
                let end = level.ends.get(index).copied().unwrap_or(start);
                let children = PyList::new(py, below.get(start..end).unwrap_or_default())?;
                start = end;
                dicts.push(self.dict(&mut out, node, record, children)?);
            }
            below = dicts;
        }
        Ok(below)
    }

    /// One node's dict, members in OfficeCLI's order.
    fn dict<'py>(
        &self,
        out: &mut Out<'py>,
        node: &Node<'_>,
        record: &Record<'_>,
        children: Bound<'py, PyList>,
    ) -> PyResult<Bound<'py, PyDict>> {
        let py = out.py;
        let dict = PyDict::new(py);
        dict.set_item(intern!(py, "path"), out.display(node.path()))?;
        dict.set_item(intern!(py, "type"), out.shared(node.kind()))?;
        if let Some(text) = &record.text {
            dict.set_item(intern!(py, "text"), text.as_ref())?;
        }
        if let Some(preview) = &record.preview {
            dict.set_item(intern!(py, "preview"), out.display(preview))?;
        }
        if let Some(style) = &record.style {
            dict.set_item(intern!(py, "style"), style.as_ref())?;
        }
        if let Some(count) = record.child_count {
            dict.set_item(intern!(py, "childCount"), count)?;
        }
        let format = PyDict::new(py);
        for prop in self.props.get(record.format.clone()).unwrap_or_default() {
            format.set_item(out.shared(prop.key()), out.value(prop)?)?;
        }
        dict.set_item(intern!(py, "format"), format)?;
        dict.set_item(intern!(py, "children"), children)?;
        Ok(dict)
    }
}

/// What building one result's dicts reuses: the `Display` buffer, and one
/// Python string per static string the result repeats.
struct Out<'py> {
    py: Python<'py>,
    scratch: String,
    /// Keyed by address and length, which identify a `&'static str`'s
    /// bytes for the life of the program.
    statics: HashMap<(usize, usize), Bound<'py, PyString>, BuildHasherDefault<AddressHasher>>,
}

impl<'py> Out<'py> {
    /// The shared Python string of a static string: format keys, element
    /// types and tokens, which every node of a kind repeats.
    fn shared(&mut self, text: &'static str) -> Bound<'py, PyString> {
        let py = self.py;
        self.statics
            .entry((text.as_ptr() as usize, text.len()))
            .or_insert_with(|| PyString::new(py, text))
            .clone()
    }

    /// `value`'s text as a `str`, written through the reused buffer.
    fn display(&mut self, value: impl fmt::Display) -> Bound<'py, PyString> {
        self.scratch.clear();
        // Writing into a `String` cannot fail.
        let _ = write!(self.scratch, "{value}");
        PyString::new(self.py, &self.scratch)
    }

    /// A `format` value with the JSON type the node view writes: `bool`,
    /// `int`, `float`, `None` for a number that is not finite, and `str` for
    /// text, tokens, lengths and colours (as the key's value type prints
    /// them).
    fn value(&mut self, prop: &Prop<'_>) -> PyResult<Bound<'py, PyAny>> {
        let py = self.py;
        Ok(match prop.value() {
            Value::Bool(value) => PyBool::new(py, *value).to_owned().into_any(),
            Value::Int(value) => value.into_pyobject(py)?.into_any(),
            Value::Number(value) => match integral(*value) {
                Some(value) => value.into_pyobject(py)?.into_any(),
                None if value.is_finite() => PyFloat::new(py, *value).into_any(),
                None => py.None().into_bound(py),
            },
            Value::Text(text) => PyString::new(py, text).into_any(),
            Value::Token(token) => self.shared(token).into_any(),
            // A list (the metadata keywords, record 0354): a list of str,
            // as the JSON form writes an array.
            Value::List(items) => PyList::new(py, items.iter())?.into_any(),
            // Lengths, colours and any later variant: the text the JSON form
            // writes, the value printed in its key's style.
            _ => self.display(prop).into_any(),
        })
    }
}

/// A hasher for `(address, length)` keys: one multiply and rotate per word
/// (the Fx hash), since the keys are few and not chosen by an attacker.
#[derive(Default)]
struct AddressHasher(u64);

impl Hasher for AddressHasher {
    fn finish(&self) -> u64 {
        self.0
    }

    fn write(&mut self, bytes: &[u8]) {
        for &byte in bytes {
            self.write_u64(u64::from(byte));
        }
    }

    fn write_u64(&mut self, word: u64) {
        self.0 = (self.0.rotate_left(5) ^ word).wrapping_mul(0x51_7c_c1_b7_27_22_0a_95);
    }

    fn write_usize(&mut self, word: usize) {
        self.write_u64(word as u64);
    }
}

/// The integer an integral finite number equals, when it is exactly
/// representable (|value| ≤ 2^53): the node view writes such a number
/// without a fractional part (`1`, not `1.0`), as OfficeCLI does.
fn integral(value: f64) -> Option<i64> {
    const LIMIT: f64 = 9_007_199_254_740_992.0; // 2^53
    // Integral and within ±2^53, so the cast is exact.
    (value.is_finite() && value.fract() == 0.0 && value.abs() <= LIMIT).then_some(value as i64)
}

#[cfg(test)]
mod tests {
    use super::integral;

    #[test]
    fn integral_numbers_are_ints_as_in_the_json_form() {
        assert_eq!(integral(1.0), Some(1));
        assert_eq!(integral(-0.0), Some(0));
        assert_eq!(
            integral(9_007_199_254_740_992.0),
            Some(9_007_199_254_740_992)
        );
        assert_eq!(integral(1.5), None);
        assert_eq!(integral(1e300), None);
        assert_eq!(integral(f64::NAN), None);
        assert_eq!(integral(f64::INFINITY), None);
    }
}
