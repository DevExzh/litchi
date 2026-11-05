//! The unified write surface (design §3 and §8, goal M3): the pieces the
//! write verbs of [`crate::file::File`] share, the [`Commit`] result of
//! `edit`, `batch` and `apply`, and the [`Edit`] transaction.
//!
//! Every verb converts its Python arguments to owned Rust values once, with
//! the GIL held, then runs the Rust verb with the GIL released. A verb never
//! changes the `File` it is called on: it returns a new one.
//!
//! # Transactions
//!
//! A Rust `litchi::transaction::Edit` lives only inside the closure
//! `litchi::File::edit` calls, so a Python object cannot hold it across
//! calls without `unsafe`. [`Edit`] therefore stages the operations
//! `set`, `remove`, `move` and `swap` and runs the staged run in one
//! `litchi::File::edit` call when it must: before `add` (which returns the
//! new element's path), before a read (`get`, `find`, `query`), and at
//! `commit`. Consecutive staged sets share the Rust transaction's native
//! coalescing, as in Rust, where reads and structural commands also close
//! a run. Every run works on an immutable snapshot, so the transaction is
//! atomic: nothing is visible until `commit`, and the `File` the
//! transaction started from never changes. The first failure poisons it:
//! its exception names the zero-based operation (`item`), and every later
//! call raises it again.

use std::borrow::Cow;
use std::sync::{Mutex, PoisonError};

use litchi::common::Position;
use litchi::{BatchItem, BatchMode, ItemResult, NewNode, Placement, Props};
use pyo3::prelude::*;
use pyo3::pybacked::{PyBackedBytes, PyBackedStr};
use pyo3::types::{PyBytes, PyDict, PyList, PyString};

use crate::errors::Failure;
use crate::file::{File, view_options};
use crate::{nodes, props, pyde};

/// Registers `Commit`, `Patch` and `Edit`.
pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Commit>()?;
    m.add_class::<Patch>()?;
    m.add_class::<Edit>()?;
    Ok(())
}

/// Where `add` inserts an element or `move` puts one: OfficeCLI's `to`,
/// `index`, `after` and `before`, owned.
#[derive(Debug, Default)]
pub(crate) struct Where {
    pub(crate) to: Option<String>,
    pub(crate) index: Option<i64>,
    pub(crate) after: Option<String>,
    pub(crate) before: Option<String>,
}

impl Where {
    /// The placement, by the batch executor's rules: at most one of
    /// `index`, `after` and `before`; `to` (moves only) names the new
    /// parent, with an optional `index` in it; nothing means the end.
    pub(crate) fn placement(&self) -> Result<Placement<'_>, Failure> {
        let anchors = usize::from(self.index.is_some())
            + usize::from(self.after.is_some())
            + usize::from(self.before.is_some());
        if anchors > 1 {
            return Err(Failure::argument(
                "only one of index, after or before is allowed",
            ));
        }
        let position = self
            .index
            .map(|index| {
                usize::try_from(index).map(Position::new).map_err(|_| {
                    Failure::argument(format!("index must not be negative, got {index}"))
                })
            })
            .transpose()?;
        let read = |error: litchi::Error| Failure::read(&error);
        if let Some(parent) = &self.to {
            if self.after.is_some() || self.before.is_some() {
                return Err(Failure::new(
                    litchi::Code::UnsupportedCapability,
                    "a move to another parent cannot also name after or before",
                )
                .with_suggestion("pass to with an index, or after/before without to"));
            }
            return match position {
                Some(position) => Placement::inside_at(parent.as_str(), position).map_err(read),
                None => Placement::inside(parent.as_str()).map_err(read),
            };
        }
        if let Some(position) = position {
            return Ok(Placement::Position(position));
        }
        if let Some(path) = &self.after {
            return Placement::after(path.as_str()).map_err(read);
        }
        if let Some(path) = &self.before {
            return Placement::before(path.as_str()).map_err(read);
        }
        Ok(Placement::End)
    }
}

/// The new element of `add`: its type, properties and placement.
pub(crate) fn new_node<'a>(
    kind: &'a str,
    props: &'a Props<'static>,
    place: &'a Where,
) -> Result<NewNode<'a>, Failure> {
    let placement = place.placement()?;
    NewNode::new(kind)
        .with_props(props)
        .map(|node| node.at(placement))
        .map_err(|error| Failure::read(&error))
}

/// `add`'s new element as owned Python arguments.
pub(crate) struct AddArgs {
    pub(crate) parent: String,
    pub(crate) kind: String,
    pub(crate) props: Props<'static>,
    pub(crate) place: Where,
}

// ---------------------------------------------------------------------------
// Batches
// ---------------------------------------------------------------------------

/// Batch items as `File.batch` received them: dicts already converted, or
/// OfficeCLI JSON text, decoded with the GIL released.
pub(crate) enum Items {
    Converted(Vec<BatchItem<'static>>),
    Json(Json),
}

/// JSON text kept as Python owns it, without a copy.
pub(crate) enum Json {
    Text(PyBackedStr),
    Bytes(PyBackedBytes),
}

impl Json {
    fn bytes(&self) -> &[u8] {
        match self {
            Self::Text(text) => text.as_bytes(),
            Self::Bytes(bytes) => bytes,
        }
    }
}

impl Items {
    /// The items of `value`: a list or tuple of dicts, or JSON text (`str`
    /// or `bytes`) holding an array of items, as OfficeCLI's `batch` reads.
    pub(crate) fn extract(value: &Bound<'_, PyAny>) -> PyResult<Self> {
        if let Ok(text) = value.cast::<PyString>() {
            return Ok(Self::Json(Json::Text(text.clone().try_into()?)));
        }
        if let Ok(bytes) = value.cast::<PyBytes>() {
            return Ok(Self::Json(Json::Bytes(bytes.clone().into())));
        }
        if !(value.is_instance_of::<PyList>() || value.is_instance_of::<pyo3::types::PyTuple>()) {
            return Err(Failure::argument(
                "batch items must be a list of dicts, or JSON text holding an array",
            )
            .into());
        }
        let mut items = Vec::with_capacity(value.len().unwrap_or(0));
        for (index, item) in value.try_iter()?.enumerate() {
            items.push(item_of(&item?).map_err(|failure| failure.with_item(index))?);
        }
        Ok(Self::Converted(items))
    }

    /// Run the batch on `file`; call with the GIL released.
    pub(crate) fn run(
        self,
        file: &litchi::File,
        mode: BatchMode,
    ) -> Result<litchi::BatchCommit, Failure> {
        let read = |error: litchi::Error| Failure::read(&error);
        match self {
            Self::Converted(items) => file.batch(items, mode).map_err(read),
            Self::Json(json) => {
                let items = litchi::batch::from_bytes(
                    json.bytes(),
                    &litchi::batch::SourceLimits::default(),
                )
                .map_err(read)?;
                file.batch(items, mode).map_err(read)
            },
        }
    }
}

/// One batch item from a dict, through `BatchItem`'s own wire rules.
pub(crate) fn item_of(value: &Bound<'_, PyAny>) -> Result<BatchItem<'static>, Failure> {
    if !value.is_instance_of::<PyDict>() {
        return Err(Failure::new(
            litchi::Code::CorruptBatch,
            "a batch item must be a dict",
        ));
    }
    serde::Deserialize::deserialize(pyde::Value(value))
        .map_err(|error: pyde::Error| Failure::new(litchi::Code::CorruptBatch, error.0))
}

/// The batch mode a Python caller names.
pub(crate) fn mode(name: &str) -> Result<BatchMode, Failure> {
    if name.eq_ignore_ascii_case("atomic") {
        Ok(BatchMode::Atomic)
    } else if name.eq_ignore_ascii_case("best_effort") || name.eq_ignore_ascii_case("besteffort") {
        Ok(BatchMode::BestEffort)
    } else {
        Err(Failure::argument(format!("unknown batch mode {name:?}"))
            .with_suggestion("use \"atomic\" or \"best_effort\""))
    }
}

/// One dict per item result, in input order: OfficeCLI's `BatchResult`
/// members `index` and `success`, then `changed`, for `add` the new
/// element's `path`, and for a `set` with `find` the match count `matched`;
/// a failed item has `code`, `error` (the message) and `suggestion`
/// instead.
pub(crate) fn results<'py>(
    py: Python<'py>,
    results: &[ItemResult],
) -> PyResult<Bound<'py, PyList>> {
    let list = PyList::empty(py);
    for (index, result) in results.iter().enumerate() {
        let dict = PyDict::new(py);
        dict.set_item(pyo3::intern!(py, "index"), index)?;
        match result {
            ItemResult::Applied(applied) => {
                dict.set_item(pyo3::intern!(py, "success"), true)?;
                dict.set_item(pyo3::intern!(py, "changed"), applied.changed)?;
                if let Some(path) = applied.path.as_deref() {
                    dict.set_item(pyo3::intern!(py, "path"), path.to_string())?;
                }
                if let Some(matched) = applied.matched() {
                    dict.set_item(pyo3::intern!(py, "matched"), matched)?;
                }
            },
            ItemResult::Failed(error) => {
                let failure = Failure::read(error);
                dict.set_item(pyo3::intern!(py, "success"), false)?;
                dict.set_item(pyo3::intern!(py, "code"), failure.code())?;
                dict.set_item(pyo3::intern!(py, "error"), failure.message())?;
                dict.set_item(pyo3::intern!(py, "suggestion"), failure.suggestion())?;
            },
            #[allow(unreachable_patterns, reason = "ItemResult is non-exhaustive")]
            _ => {
                dict.set_item(pyo3::intern!(py, "success"), false)?;
                dict.set_item(pyo3::intern!(py, "code"), "internal_error")?;
                dict.set_item(pyo3::intern!(py, "error"), "an unknown item result")?;
                dict.set_item(pyo3::intern!(py, "suggestion"), py.None())?;
            },
        }
        list.append(dict)?;
    }
    Ok(list)
}

// ---------------------------------------------------------------------------
// Commit
// ---------------------------------------------------------------------------

/// The result of ``File.edit``, ``File.batch``, ``File.apply`` and
/// ``Edit.commit``: the new File and what the change did.
///
/// ``file`` is the new File; the File the change started from is
/// unchanged. ``changed`` says whether the document changed (False for a
/// change that wrote what was already there). ``operations`` counts the
/// successful operations. ``results`` lists one dict per batch item (None
/// for ``edit``), and ``value`` is what ``edit``'s callable returned (None
/// otherwise). ``patch`` is the change's reversible Patch.
#[pyclass(frozen, module = "litchi_py")]
pub struct Commit {
    file: Py<File>,
    patch: Option<Py<Patch>>,
    changed: bool,
    operations: usize,
    results: Option<Py<PyList>>,
    value: Py<PyAny>,
}

impl Commit {
    /// A Commit; `patch` is None for the raw-layer verbs, which are not
    /// transaction steps.
    pub(crate) fn new(
        py: Python<'_>,
        file: litchi::File,
        patch: Option<Patch>,
        changed: bool,
        operations: usize,
        results: Option<Bound<'_, PyList>>,
        value: Option<Py<PyAny>>,
    ) -> PyResult<Self> {
        Ok(Self {
            file: Py::new(py, File::new(file))?,
            patch: patch.map(|patch| Py::new(py, patch)).transpose()?,
            changed,
            operations,
            results: results.map(Bound::unbind),
            value: value.unwrap_or_else(|| py.None()),
        })
    }
}

#[pymethods]
impl Commit {
    /// The new File; the File the change started from is unchanged.
    #[getter]
    fn file(&self, py: Python<'_>) -> Py<File> {
        self.file.clone_ref(py)
    }

    /// Whether the document changed.
    #[getter]
    fn changed(&self) -> bool {
        self.changed
    }

    /// How many operations succeeded.
    #[getter]
    fn operations(&self) -> usize {
        self.operations
    }

    /// One dict per batch item, in input order; None for ``edit``.
    #[getter]
    fn results(&self, py: Python<'_>) -> Option<Py<PyList>> {
        self.results.as_ref().map(|results| results.clone_ref(py))
    }

    /// What ``edit``'s callable returned; None otherwise.
    #[getter]
    fn value(&self, py: Python<'_>) -> Py<PyAny> {
        self.value.clone_ref(py)
    }

    /// The change's reversible Patch; None for ``File.add_part`` and
    /// ``File.raw_set``, which are not transaction steps.
    #[getter]
    fn patch(&self, py: Python<'_>) -> Option<Py<Patch>> {
        self.patch.as_ref().map(|patch| patch.clone_ref(py))
    }

    fn __repr__(&self) -> String {
        format!(
            "<Commit: changed={}, operations={}>",
            if self.changed { "True" } else { "False" },
            self.operations
        )
    }
}

// ---------------------------------------------------------------------------
// Patch
// ---------------------------------------------------------------------------

/// The reversible record of one change: ``Commit.patch``.
///
/// ``apply(file)`` replays the change on the File it started from (or the
/// same File object again) and returns the changed File. ``inverse()`` is
/// the Patch that undoes it: applied to the changed File it gives the
/// original back. A same-format independently opened copy is accepted after
/// its published base passes the bounded content proof; changed content,
/// protected state and stale sources remain typed failures. A Patch lives
/// only in this process; it is not a file format.
///
/// Example:
///     >>> commit = doc.edit(lambda tx: tx.set("/body/p[1]", {"bold": True}))
///     >>> undone = commit.patch.inverse().apply(commit.file)
///     >>> assert undone.to_bytes() == doc.to_bytes()
#[pyclass(frozen, module = "litchi_py")]
pub struct Patch {
    /// The Rust patches, in application order. An ``Edit`` runs one Rust
    /// transaction per staged run, and each run's patch starts from the
    /// previous one's result.
    steps: Vec<litchi::transaction::Patch>,
}

impl Patch {
    /// The Patch of one Rust transaction.
    pub(crate) fn one(patch: litchi::transaction::Patch) -> Self {
        Self { steps: vec![patch] }
    }
}

#[pymethods]
impl Patch {
    /// Whether the change kept the document as it was.
    fn is_noop(&self) -> bool {
        self.steps.iter().all(litchi::transaction::Patch::is_noop)
    }

    /// The Patch that undoes this one.
    fn inverse(&self) -> Self {
        Self {
            steps: self
                .steps
                .iter()
                .rev()
                .map(litchi::transaction::Patch::inverse)
                .collect(),
        }
    }

    /// Replay the change on ``file``; return the changed File.
    ///
    /// The GIL is released while the patch applies.
    ///
    /// Args:
    ///     file: The original File, a clone, or a same-format equal-content
    ///         File reopened from the published base
    ///
    /// Returns:
    ///     The changed File; ``file`` is unchanged
    ///
    /// Raises:
    ///     LitchiError: A typed patch-conflict, protection or source failure
    fn apply(&self, py: Python<'_>, file: PyRef<'_, File>) -> PyResult<File> {
        let mut current = file.handle()?;
        let steps = &self.steps;
        let result = py.detach(|| {
            for step in steps {
                current = step
                    .apply(&current)
                    .map_err(|error| Failure::read(&error))?;
            }
            Ok::<_, Failure>(current)
        })?;
        Ok(File::new(result))
    }

    fn __repr__(&self) -> String {
        format!(
            "<Patch: noop={}>",
            if self.is_noop() { "True" } else { "False" }
        )
    }
}

// ---------------------------------------------------------------------------
// Edit
// ---------------------------------------------------------------------------

/// A staged operation of an [`Edit`].
enum Op {
    Set {
        target: String,
        props: Props<'static>,
    },
    Remove {
        target: String,
    },
    Move {
        path: String,
        place: Where,
    },
    Swap {
        first: String,
        second: String,
    },
}

/// Where an [`Edit`] is in its life.
enum Phase {
    Open,
    Failed(Failure),
    Discarded,
    /// Run and closed; its `Commit` is being built.
    Committing,
    Committed(Py<Commit>),
}

/// The state of an [`Edit`], behind its lock.
struct State {
    /// The snapshot every staged run so far produced.
    current: litchi::File,
    /// Operations staged since the last run.
    staged: Vec<Op>,
    /// The operation index of `staged[0]`.
    first: usize,
    /// Whether a run changed the document.
    changed: bool,
    /// The patch of every run so far, in order.
    patches: Vec<litchi::transaction::Patch>,
    /// Whether a Rust transaction ran yet: the first run happens even with
    /// nothing staged, so the file's admission (locked, signed) is checked
    /// as Rust's `File::edit` checks it.
    ran: bool,
    phase: Phase,
}

impl State {
    /// Fails unless the transaction is open.
    fn check(&self) -> Result<(), Failure> {
        match &self.phase {
            Phase::Open => Ok(()),
            Phase::Failed(failure) => Err(failure.clone()),
            Phase::Discarded => Err(Failure::new(
                litchi::Code::InvalidArgument,
                "the edit was discarded",
            )
            .with_suggestion("start a new one with File.edit()")),
            Phase::Committing | Phase::Committed(_) => Err(Failure::new(
                litchi::Code::InvalidArgument,
                "the edit is already committed",
            )
            .with_suggestion("start a new one with File.edit()")),
        }
    }

    /// Stages `op`.
    fn stage(&mut self, op: Op) -> Result<(), Failure> {
        self.check()?;
        self.staged.push(op);
        Ok(())
    }

    /// Runs the staged operations, then `add` if given, in one Rust
    /// transaction on the current snapshot; returns `add`'s path. A failure
    /// poisons the transaction.
    fn run(&mut self, add: Option<&AddArgs>) -> Result<Option<String>, Failure> {
        self.check()?;
        if self.ran && self.staged.is_empty() && add.is_none() {
            return Ok(None);
        }
        let staged = std::mem::take(&mut self.staged);
        let result = self.current.edit(|edit| {
            for op in &staged {
                match op {
                    Op::Set { target, props } => edit.set(target.as_str(), props)?,
                    Op::Remove { target } => edit.remove(target.as_str())?,
                    Op::Move { path, place } => {
                        let placement = place.placement().map_err(Failure::into_error)?;
                        edit.move_to(path.as_str(), placement)?;
                    },
                    Op::Swap { first, second } => edit.swap(first.as_str(), second.as_str())?,
                }
            }
            match add {
                Some(add) => {
                    let node =
                        new_node(&add.kind, &add.props, &add.place).map_err(Failure::into_error)?;
                    let path = edit.add(add.parent.as_str(), node)?;
                    Ok(Some(path.to_string()))
                },
                None => Ok(None),
            }
        });
        let count = staged.len() + usize::from(add.is_some());
        self.ran = true;
        match result {
            Ok(commit) => {
                self.changed |= commit.diagnostics().changed();
                let (file, patch, value) = commit.into_parts();
                self.current = file;
                self.patches.push(patch);
                self.first += count;
                Ok(value)
            },
            Err(error) => {
                let failure = Failure::read(&error);
                let failure = match error.item() {
                    Some(index) => failure.with_item(self.first + index),
                    None => failure,
                };
                self.phase = Phase::Failed(failure.clone());
                Err(failure)
            },
        }
    }
}

/// A transaction: operations that apply together or not at all.
///
/// ``File.edit()`` returns one. Use it as a context manager, which commits
/// when the block ends normally and discards the operations when it raises,
/// or call ``commit()`` yourself. ``File.edit(fn)`` calls ``fn`` with one
/// and commits it.
///
/// ``set``, ``remove``, ``move`` and ``swap`` are staged and run together:
/// consecutive property sets share one native edit, as in Rust. ``add`` and
/// the reads (``get``, ``find``, ``query``) first run what is staged, so
/// ``add`` can return the new element's path and the reads see every
/// operation before them. An error of a staged operation is raised by the
/// call that runs it (``add``, a read or ``commit``); its ``item`` names the
/// operation (zero-based, in call order). The first failure ends the
/// transaction: every later call raises it again, and the File it started
/// from is unchanged, as it always is.
///
/// An Edit can be shared between threads; calls on it run one at a time.
///
/// Example:
///     >>> with doc.edit() as tx:
///     ...     path = tx.add("/body", "paragraph", {"text": "Hello"})
///     ...     tx.set(path, {"align": "center"})
///     >>> doc = tx.commit().file
#[pyclass(frozen, module = "litchi_py")]
pub struct Edit {
    state: Mutex<State>,
}

impl Edit {
    pub(crate) fn new(file: litchi::File) -> Self {
        Self {
            state: Mutex::new(State {
                current: file,
                staged: Vec::new(),
                first: 0,
                changed: false,
                patches: Vec::new(),
                ran: false,
                phase: Phase::Open,
            }),
        }
    }

    /// Runs `f` on the state with the GIL released, so a thread waiting for
    /// the lock never holds the GIL.
    fn with<T: Send>(
        &self,
        py: Python<'_>,
        f: impl FnOnce(&mut State) -> Result<T, Failure> + Send,
    ) -> Result<T, Failure> {
        py.detach(|| f(&mut self.state.lock().unwrap_or_else(PoisonError::into_inner)))
    }

    /// The current snapshot after running what is staged, for a read.
    fn snapshot(&self, py: Python<'_>) -> Result<litchi::File, Failure> {
        self.with(py, |state| {
            state.run(None)?;
            Ok(state.current.clone())
        })
    }

    /// Commits the transaction with `value` as the commit's value.
    pub(crate) fn commit_with(
        &self,
        py: Python<'_>,
        value: Option<Py<PyAny>>,
    ) -> PyResult<Py<Commit>> {
        if let Some(commit) = self.committed(py) {
            return Ok(commit);
        }
        let (file, patch, changed, operations) = self.with(py, |state| {
            state.run(None)?;
            state.phase = Phase::Committing;
            let patch = Patch {
                steps: std::mem::take(&mut state.patches),
            };
            Ok((state.current.clone(), patch, state.changed, state.first))
        })?;
        let commit = Py::new(
            py,
            Commit::new(py, file, Some(patch), changed, operations, None, value)?,
        )?;
        // Holding the GIL while locking is safe: no thread holds the lock
        // while it waits for the GIL (`with` locks with the GIL released).
        let mut state = self.state.lock().unwrap_or_else(PoisonError::into_inner);
        state.phase = Phase::Committed(commit.clone_ref(py));
        Ok(commit)
    }

    /// The commit of a committed transaction.
    fn committed(&self, py: Python<'_>) -> Option<Py<Commit>> {
        let state = self.state.lock().unwrap_or_else(PoisonError::into_inner);
        match &state.phase {
            Phase::Committed(commit) => Some(commit.clone_ref(py)),
            _ => None,
        }
    }

    /// Discards the transaction unless it is already committed or failed.
    pub(crate) fn discard(&self, py: Python<'_>) {
        let _ = self.with(py, |state| {
            if matches!(state.phase, Phase::Open) {
                state.staged.clear();
                state.phase = Phase::Discarded;
            }
            Ok(())
        });
    }
}

#[pymethods]
impl Edit {
    /// Stage setting ``props`` on the element(s) ``target`` names.
    ///
    /// Args:
    ///     target: An OfficeCLI path, or a selector scoped to a path
    ///     props: The properties, as a dict ({"font.bold": True}) or a
    ///         list of "key=value" strings; None clears where allowed
    ///
    /// Raises:
    ///     PropertyError: (a ValueError) If props is not a dict or list
    ///     LitchiError: If the transaction already failed, was committed or
    ///         was discarded
    fn set(&self, py: Python<'_>, target: String, props: Bound<'_, PyAny>) -> PyResult<()> {
        let props = props::props(Some(&props))?;
        Ok(self.with(py, move |state| state.stage(Op::Set { target, props }))?)
    }

    /// Run what is staged, then add an element; return its path.
    ///
    /// Args:
    ///     parent: The path of the new element's parent
    ///     type: The element type: "paragraph", "run", "sheet", "slide", ...
    ///     props: The new element's properties
    ///     index: Insert at this zero-based position among the children
    ///     after: Insert right after the element at this path
    ///     before: Insert right before the element at this path
    ///
    /// Returns:
    ///     The new element's canonical path
    #[pyo3(signature = (parent, r#type, props = None, *, index = None, after = None, before = None))]
    #[allow(
        clippy::too_many_arguments,
        reason = "the keyword arguments mirror OfficeCLI's add options"
    )]
    fn add(
        &self,
        py: Python<'_>,
        parent: String,
        r#type: String,
        props: Option<Bound<'_, PyAny>>,
        index: Option<i64>,
        after: Option<String>,
        before: Option<String>,
    ) -> PyResult<String> {
        let add = AddArgs {
            parent,
            kind: r#type,
            props: props::props(props.as_ref())?,
            place: Where {
                to: None,
                index,
                after,
                before,
            },
        };
        new_node(&add.kind, &add.props, &add.place)?;
        let path = self.with(py, move |state| state.run(Some(&add)))?;
        path.ok_or_else(|| Failure::new(litchi::Code::InternalError, "add returned no path").into())
    }

    /// Stage removing the element(s) ``target`` names.
    fn remove(&self, py: Python<'_>, target: String) -> PyResult<()> {
        Ok(self.with(py, move |state| state.stage(Op::Remove { target }))?)
    }

    /// Stage moving the element at ``path``.
    ///
    /// Args:
    ///     path: The element to move
    ///     to: The new parent; None keeps the parent
    ///     index: The zero-based position to move to
    ///     after: Move right after the element at this path
    ///     before: Move right before the element at this path
    #[pyo3(name = "move", signature = (path, to = None, *, index = None, after = None, before = None))]
    #[allow(
        clippy::too_many_arguments,
        reason = "the keyword arguments mirror OfficeCLI's move options"
    )]
    fn move_(
        &self,
        py: Python<'_>,
        path: String,
        to: Option<String>,
        index: Option<i64>,
        after: Option<String>,
        before: Option<String>,
    ) -> PyResult<()> {
        let place = Where {
            to,
            index,
            after,
            before,
        };
        // Refuse a malformed placement now, not when the run executes it.
        place.placement()?;
        Ok(self.with(py, move |state| state.stage(Op::Move { path, place }))?)
    }

    /// Stage swapping the elements at ``path`` and ``path2``.
    fn swap(&self, py: Python<'_>, path: String, path2: String) -> PyResult<()> {
        Ok(self.with(py, move |state| {
            state.stage(Op::Swap {
                first: path,
                second: path2,
            })
        })?)
    }

    /// The node view at ``path`` in the transaction's current state, as
    /// ``File.get``.
    #[pyo3(signature = (path, depth = 1, *, max_nodes = 100_000, dense = None))]
    fn get<'py>(
        &self,
        py: Python<'py>,
        path: PyBackedStr,
        depth: i64,
        max_nodes: i64,
        dense: Option<i64>,
    ) -> PyResult<Bound<'py, PyDict>> {
        let options = view_options(depth, max_nodes, dense)?;
        let file = self.snapshot(py)?;
        nodes::get(py, &file, &path, options)
    }

    /// As ``get``, or None when nothing is there.
    #[pyo3(signature = (path, depth = 1, *, max_nodes = 100_000, dense = None))]
    fn find<'py>(
        &self,
        py: Python<'py>,
        path: PyBackedStr,
        depth: i64,
        max_nodes: i64,
        dense: Option<i64>,
    ) -> PyResult<Option<Bound<'py, PyDict>>> {
        let options = view_options(depth, max_nodes, dense)?;
        let file = self.snapshot(py)?;
        nodes::find(py, &file, &path, options)
    }

    /// The node views ``selector`` matches in the transaction's current
    /// state, as ``File.query``.
    #[pyo3(signature = (selector, limit = None, depth = 0, *, max_nodes = 100_000, dense = None))]
    fn query<'py>(
        &self,
        py: Python<'py>,
        selector: PyBackedStr,
        limit: Option<i64>,
        depth: i64,
        max_nodes: i64,
        dense: Option<i64>,
    ) -> PyResult<Bound<'py, PyList>> {
        let options = view_options(depth, max_nodes, dense)?;
        let limit = crate::file::limit_of(limit)?;
        let file = self.snapshot(py)?;
        nodes::query(py, &file, &selector, limit, options)
    }

    /// Run what is staged and finish the transaction.
    ///
    /// Returns:
    ///     A Commit whose ``file`` holds every operation; calling commit
    ///     again returns the same Commit
    ///
    /// Raises:
    ///     LitchiError: The first failed operation's error, with its
    ///         ``item``; the transaction is then over
    fn commit(&self, py: Python<'_>) -> PyResult<Py<Commit>> {
        self.commit_with(py, None)
    }

    fn __enter__(slf: Bound<'_, Self>) -> PyResult<Bound<'_, Self>> {
        let py = slf.py();
        slf.get().with(py, |state| state.check())?;
        Ok(slf)
    }

    fn __exit__(
        &self,
        py: Python<'_>,
        exc_type: Option<Bound<'_, PyAny>>,
        exc_value: Option<Bound<'_, PyAny>>,
        traceback: Option<Bound<'_, PyAny>>,
    ) -> PyResult<bool> {
        let _ = (exc_value, traceback);
        if exc_type.is_some() {
            self.discard(py);
        } else {
            self.commit_with(py, None)?;
        }
        Ok(false)
    }

    fn __repr__(&self) -> Cow<'static, str> {
        let state = self.state.lock().unwrap_or_else(PoisonError::into_inner);
        match state.phase {
            Phase::Open => format!("<Edit: {} staged>", state.staged.len()).into(),
            Phase::Failed(_) => "<Edit: failed>".into(),
            Phase::Discarded => "<Edit: discarded>".into(),
            Phase::Committing | Phase::Committed(_) => "<Edit: committed>".into(),
        }
    }
}
