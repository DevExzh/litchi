//! The unified read surface (record 0010 §8.2, M1 step R9):
//! `litchi_py.open` and `open_bytes` return a [`File`], an immutable handle
//! over `litchi::File` with `get`, `find`, `query`, `view`,
//! `capabilities`, `text` and `metadata`.
//!
//! Every read releases the GIL for the Rust work (opening, resolving,
//! walking, reading the node members, rendering a view) and holds it only
//! to build the Python result. Node views are built as dicts straight from
//! the node view's accessors ([`crate::nodes`]); view JSON, the capability
//! report and the metadata through a `serde` serializer that builds Python
//! objects ([`crate::pyser`]). Neither writes JSON text.
//!
//! The handle is thread-safe: `litchi::File` is `Send + Sync` and cheap to
//! clone (one reference count), so each call clones it out of a lock and
//! reads without holding the lock. `close()` drops the handle's own
//! reference; a call already running on another thread keeps reading its
//! clone.

use std::borrow::Cow;
use std::path::PathBuf;
use std::sync::{PoisonError, RwLock};

use litchi::node::{DEFAULT_MAX_NODES, ViewOptions};
use litchi::view::Mode;
use litchi::{Overwrite, SaveOptions, SaveTo};
use pyo3::prelude::*;
use pyo3::pybacked::PyBackedStr;
use pyo3::types::{PyBytes, PyDict, PyList, PyModule, PyString};

use crate::common::FileFormat;
use crate::errors::Failure;
use crate::write::{self, Commit, Edit, Items, Where};
use crate::{m4, merge, nodes, props, pyser, refresh};

/// Registers `File`, `open` and `open_bytes`.
pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<File>()?;
    m.add_function(wrap_pyfunction!(open, m)?)?;
    m.add_function(wrap_pyfunction!(open_bytes, m)?)?;
    m.add_function(wrap_pyfunction!(create, m)?)?;
    m.add_function(wrap_pyfunction!(schema, m)?)?;
    Ok(())
}

/// The default node budget of `get`, `find` and `query`: the signatures
/// spell it as the literal `100_000`, so `help()` and `inspect` show it.
const _: () = assert!(DEFAULT_MAX_NODES == 100_000);

/// Open the file at `path` through the unified API.
///
/// The format is detected from the file's content, not from its name, in
/// one pass. Any format `litchi::File` holds opens: .docx, .xlsx, .pptx,
/// .xlsb, .doc, .xls, .ppt, .rtf, .odt, .ods and .odp. An encrypted file
/// opens locked (its reads raise ProtectedError; ``capabilities()`` says
/// so), and a signed one opens readable. The GIL is released while the file
/// is opened.
///
/// Args:
///     path: Path to the file (str or os.PathLike)
///
/// Returns:
///     A File
///
/// Raises:
///     IoError: (an OSError) If the file cannot be read
///     UnsupportedError: (a ValueError) If the file is not an Office file
///         (an Apple iWork package directory opens as its single-file form,
///         and save() writes it back as a directory)
///     CorruptFileError: (a ValueError) If the file is malformed
#[pyfunction]
fn open(py: Python<'_>, path: PathBuf) -> PyResult<File> {
    let file = py.detach(|| litchi::open(&path).map_err(|error| Failure::read(&error)))?;
    Ok(File::new(file))
}

/// Open an in-memory file through the unified API.
///
/// As ``open``, from the file's bytes, which are copied once. The GIL is
/// released while the bytes are copied and opened.
///
/// Args:
///     data: The file's content, as bytes
///
/// Returns:
///     A File
///
/// Raises:
///     UnsupportedError: (a ValueError) If the bytes are not an Office
///         file, or are in a format File does not hold
///     CorruptFileError: (a ValueError) If the file is malformed
#[pyfunction]
fn open_bytes(py: Python<'_>, data: &[u8]) -> PyResult<File> {
    let file = py.detach(|| litchi::open_bytes(data).map_err(|error| Failure::read(&error)))?;
    Ok(File::new(file))
}

/// Create a new, empty document of ``format`` through the unified API.
///
/// The new File has no path: write it with ``save_as`` or ``to_bytes``.
/// .docx, .xlsx, .pptx, .xlsb, .doc, .xls, .ppt, .rtf and .odt can be created.
/// The GIL is released while the document is built.
///
/// Args:
///     format: A FileFormat, or its name ("docx", "Xlsx", ".pptx")
///
/// Returns:
///     A File
///
/// Raises:
///     PropertyError: (a ValueError) If format names no format
///         ("invalid_argument")
///     UnsupportedError: (a ValueError) If Litchi cannot create the format
#[pyfunction]
fn create(py: Python<'_>, format: Bound<'_, PyAny>) -> PyResult<File> {
    let format = FileFormat::extract(&format)?.to_core();
    let file = py.detach(|| litchi::create(format).map_err(|error| Failure::read(&error)))?;
    Ok(File::new(file))
}

/// The property registry of ``format``: what ``set``, ``add`` and ``get``
/// accept for each element, as data.
///
/// Each element is in OfficeCLI's help-schema shape. Each property has its
/// ``type``, the enumeration ``values`` and ``valueAliases``, the ``unit``
/// a bare length number is read in, its ``aliases``, ``constraint`` and
/// ``clearable`` flag, the ``add``/``set``/``get`` flags, and a ``writeGap``
/// with its follow-up when it cannot be written yet. ``propertyGaps`` lists
/// an element's keys Litchi does not answer yet; ``gaps`` the format's other
/// gaps. A format without a reader in this build has no elements and a
/// ``pending`` entry. The result equals the Rust ``litchi_ops::schema``
/// JSON.
///
/// Args:
///     format: A FileFormat, or its name ("docx", "Xlsx", ".pptx")
///
/// Returns:
///     A dict
///
/// Raises:
///     PropertyError: (a ValueError) If format names no format
///         ("invalid_argument")
#[pyfunction]
fn schema<'py>(py: Python<'py>, format: Bound<'py, PyAny>) -> PyResult<Bound<'py, PyAny>> {
    let format = FileFormat::extract(&format)?.to_core();
    pyser::to_py(py, &litchi::schema(format))
}

/// An open document, workbook or presentation: the unified API.
///
/// Returned by ``litchi_py.open`` and ``litchi_py.open_bytes``. A File is
/// an immutable handle: reads never change it, and it can be shared between
/// threads. Every read releases the GIL while Litchi works and holds it
/// only to build the result.
///
/// Reads address the document with OfficeCLI's paths (``/body/p[3]``,
/// ``/Sheet1/B2``, ``Sheet1!B2``) and selectors (``paragraph[style=Heading1]``,
/// ``row[Amount>5000]``), and return OfficeCLI's node view as dicts.
/// ``capabilities()`` lists what the format answers; a format whose unified
/// reader has not landed yet refuses every read with UnsupportedError
/// (``unsupported_capability``).
///
/// A File is a context manager: leaving the ``with`` block closes it. After
/// ``close()``, every read raises IoError (``file_required``).
///
/// # Examples
///
/// ```python
/// import litchi_py
///
/// with litchi_py.open("report.docx") as doc:
///     paragraph = doc.get("/body/p[1]")          # dict: path, type, text, format, ...
///     headings = doc.query("paragraph[style=Heading1]", limit=10)
///     print(doc.view("outline"))
/// ```
#[pyclass(frozen, module = "litchi_py")]
pub struct File {
    /// The open handle; `None` after `close()`.
    handle: RwLock<Option<litchi::File>>,
    /// The detected format, kept after `close()`.
    format: Option<litchi::FileFormat>,
}

impl File {
    pub(crate) fn new(file: litchi::File) -> Self {
        Self {
            format: file.format(),
            handle: RwLock::new(Some(file)),
        }
    }

    /// A clone of the open handle (one reference count), so the read runs
    /// without holding the lock; the closed-file failure after `close()`.
    pub(crate) fn handle(&self) -> Result<litchi::File, Failure> {
        self.handle
            .read()
            .unwrap_or_else(PoisonError::into_inner)
            .clone()
            .ok_or_else(Failure::closed)
    }
}

#[pymethods]
impl File {
    /// The detected format, or None for an encrypted OOXML package whose
    /// format is hidden until it is unlocked. Still available after
    /// ``close()``.
    #[getter]
    fn format(&self) -> PyResult<Option<FileFormat>> {
        self.format
            .map(FileFormat::from_core)
            .transpose()
            .map_err(PyErr::from)
    }

    /// Whether ``close()`` has been called.
    #[getter]
    fn closed(&self) -> bool {
        self.handle
            .read()
            .unwrap_or_else(PoisonError::into_inner)
            .is_none()
    }

    /// The node view of the element at ``path``.
    ///
    /// The view is OfficeCLI's node shape, as a dict: ``path`` (canonical),
    /// ``type``, ``text``, ``preview``, ``style`` and ``childCount`` (each
    /// left out when the node has none), ``format`` (the node's own
    /// properties, with JSON types) and ``children`` (child views, ``depth``
    /// levels down). Only the parts on the path are read.
    ///
    /// Args:
    ///     path: An OfficeCLI path, such as "/body/p[3]", "/Sheet1/B2" or
    ///         "Sheet1!B2"
    ///     depth: Child levels to list below the node: 0 lists none; values
    ///         above 64 count as 64
    ///     max_nodes: The most nodes the result may hold, children included
    ///     dense: For a cell range, list empty cells too, up to this many
    ///
    /// Returns:
    ///     The node view
    ///
    /// Raises:
    ///     PathError: (a LookupError) If the path is malformed
    ///         ("invalid_path"), nothing is there ("path_not_found"), or the
    ///         format has no such element ("unsupported_path")
    ///     UnsupportedError: (a ValueError) If Litchi cannot read the
    ///         element or the format yet ("unsupported_capability")
    ///     ProtectedError: (a ValueError) If the file is encrypted
    ///     PropertyError: (a ValueError) If an argument is negative
    ///     LitchiError: If the view exceeds max_nodes ("output_too_large")
    ///     IoError: (an OSError) If the file is closed or changed on disk
    #[pyo3(signature = (path, depth = 1, *, max_nodes = 100_000, dense = None))]
    fn get<'py>(
        &self,
        py: Python<'py>,
        path: PyBackedStr,
        depth: i64,
        max_nodes: i64,
        dense: Option<i64>,
    ) -> PyResult<Bound<'py, PyDict>> {
        let file = self.handle()?;
        let options = view_options(depth, max_nodes, dense)?;
        nodes::get(py, &file, &path, options)
    }

    /// The node view of the element at ``path``, or None when nothing is
    /// there.
    ///
    /// As ``get``, except that a path that resolves to nothing returns None
    /// instead of raising PathError ("path_not_found"). A malformed path
    /// still raises.
    ///
    /// Args:
    ///     path: An OfficeCLI path
    ///     depth: Child levels to list below the node
    ///     max_nodes: The most nodes the result may hold, children included
    ///     dense: For a cell range, list empty cells too, up to this many
    ///
    /// Returns:
    ///     The node view, or None
    #[pyo3(signature = (path, depth = 1, *, max_nodes = 100_000, dense = None))]
    fn find<'py>(
        &self,
        py: Python<'py>,
        path: PyBackedStr,
        depth: i64,
        max_nodes: i64,
        dense: Option<i64>,
    ) -> PyResult<Option<Bound<'py, PyDict>>> {
        let file = self.handle()?;
        let options = view_options(depth, max_nodes, dense)?;
        nodes::find(py, &file, &path, options)
    }

    /// The node views of the elements ``selector`` matches, in document
    /// order.
    ///
    /// The selector is compiled once and the document is walked in one
    /// pass, which stops after ``limit`` matches. Each match is a node view,
    /// as ``get`` returns, with ``depth`` levels of children (none by
    /// default). An unknown property key in the selector is not an error:
    /// it matches nothing and issues a UserWarning that suggests the nearest
    /// key.
    ///
    /// Args:
    ///     selector: An OfficeCLI selector, such as "paragraph",
    ///         "cell[type=Number]" or "Sheet1!row[Amount>5000]"
    ///     limit: The most matches to return; None returns every match
    ///     depth: Child levels to list below each match
    ///     max_nodes: The most nodes the result may hold, matches and their
    ///         children together
    ///     dense: For cell ranges, list empty cells too, up to this many
    ///
    /// Returns:
    ///     A list of node views
    ///
    /// Raises:
    ///     PathError: (a LookupError) If the selector is malformed
    ///         ("invalid_selector") or its scope path does not resolve
    ///     UnsupportedError: (a ValueError) If Litchi cannot query the
    ///         format yet ("unsupported_capability")
    ///     ProtectedError: (a ValueError) If the file is encrypted
    ///     PropertyError: (a ValueError) If an argument is negative
    ///     LitchiError: If more than max_nodes nodes match
    ///         ("output_too_large"); pass a limit
    ///     IoError: (an OSError) If the file is closed or changed on disk
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
        let file = self.handle()?;
        let options = view_options(depth, max_nodes, dense)?;
        let limit = limit_of(limit)?;
        nodes::query(py, &file, &selector, limit, options)
    }

    /// One of OfficeCLI's semantic views of the whole document.
    ///
    /// The modes are "text" (the text, one line per element, with its
    /// path), "annotated" (text with formatting annotations), "outline"
    /// (the structure, with counts) and "stats" (counts and distributions),
    /// or their one-letter aliases. The other OfficeCLI modes ("issues",
    /// "forms", "html", "svg", "screenshot", "pdf") raise UnsupportedError.
    ///
    /// Args:
    ///     mode: The view mode
    ///     json: Return OfficeCLI's JSON form as Python objects instead of
    ///         text
    ///
    /// Returns:
    ///     The view as text, or as a dict when json is True
    ///
    /// Raises:
    ///     PropertyError: (a ValueError) If the mode is unknown
    ///         ("invalid_argument")
    ///     UnsupportedError: (a ValueError) If the format or the mode is not
    ///         supported ("unsupported_capability")
    ///     ProtectedError: (a ValueError) If the file is encrypted
    ///     IoError: (an OSError) If the file is closed or changed on disk
    #[pyo3(signature = (mode, *, json = false))]
    fn view<'py>(
        &self,
        py: Python<'py>,
        mode: PyBackedStr,
        json: bool,
    ) -> PyResult<Bound<'py, PyAny>> {
        let file = self.handle()?;
        let read = |error: litchi::Error| Failure::read(&error);
        let mode = Mode::parse(&mode).ok_or_else(|| {
            let suggestion = match nearest_mode(&mode) {
                Some(name) => format!("did you mean '{name}'?").into(),
                None => Cow::Borrowed("use text, annotated, outline or stats (or t, a, o, s)"),
            };
            Failure::argument("unknown view mode").with_suggestion(suggestion)
        })?;
        if json {
            let view = py.detach(|| file.view_json(mode).map_err(read))?;
            pyser::to_py(py, &view)
        } else {
            let mut text = String::new();
            py.detach(|| file.view_to(mode, &mut text).map_err(read))?;
            Ok(PyString::new(py, &text).into_any())
        }
    }

    /// What this file can do through the unified API: OfficeCLI's help
    /// schema, as a dict.
    ///
    /// The report names the file's ``state`` ("open", "signed" or "locked"),
    /// whether each verb and view mode is supported (or the code its
    /// refusal raises, with the follow-up that adds it), every element the
    /// format's reader knows with its paths and properties, and every listed
    /// gap.
    ///
    /// Returns:
    ///     The capability report
    ///
    /// Raises:
    ///     IoError: (an OSError) If the file is closed
    fn capabilities<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let file = self.handle()?;
        // A view over static tables and the handle's state: building it
        // reads nothing, so only the Python objects cost.
        pyser::to_py(py, &file.capabilities())
    }

    /// All text of the document, in reading order.
    ///
    /// Returns:
    ///     The text
    ///
    /// Raises:
    ///     ProtectedError: (a ValueError) If the file is encrypted
    ///     IoError: (an OSError) If the file is closed or changed on disk
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        let file = self.handle()?;
        py.detach(|| file.text().map_err(|error| Failure::read(&error)))
            .map_err(PyErr::from)
    }

    /// The document properties (title, author, dates, counts, ...), as a
    /// dict with the keys that are set; dates are ISO 8601 text.
    ///
    /// Returns:
    ///     The properties, or None when the format reports none (.pptx and
    ///     .ppt without core properties)
    ///
    /// Raises:
    ///     ProtectedError: (a ValueError) If the file is encrypted
    ///     IoError: (an OSError) If the file is closed or changed on disk
    fn metadata<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyAny>>> {
        let file = self.handle()?;
        let metadata = py.detach(|| file.metadata().map_err(|error| Failure::read(&error)))?;
        metadata
            .map(|metadata| pyser::to_py(py, &metadata))
            .transpose()
    }

    /// Set properties on the element(s) ``target`` names; return the
    /// changed File.
    ///
    /// The File this is called on is unchanged. Properties use OfficeCLI's
    /// names, and dotted keys are fine in a dict. Values keep their Python
    /// types: ``None`` clears a property where it can be cleared, ``bool``,
    /// ``int`` and ``float`` stay typed, and text is coerced through the
    /// key's value type ("12pt", "#1F4E79", "accent1+lumMod75"). The GIL is
    /// released while the change is made.
    ///
    /// Args:
    ///     target: An OfficeCLI path ("/body/p[3]/r[1]", "/Sheet1/B2"), or
    ///         a selector scoped to a path
    ///     props: A dict ({"font.bold": True, "fill": "1F4E79"}) or a list
    ///         of "key=value" strings
    ///
    /// Returns:
    ///     The new File
    ///
    /// Raises:
    ///     PathError: (a LookupError) If the target is malformed or
    ///         resolves to nothing
    ///     PropertyError: (a ValueError) If a property is unknown (with a
    ///         did-you-mean suggestion) or a value is invalid
    ///     UnsupportedError: (a ValueError) If the format cannot set this
    ///         property or element yet ("unsupported_capability")
    ///     ProtectedError: (a ValueError) If the file is encrypted or signed
    ///         (call ``unsign()`` first)
    ///     IoError: (an OSError) If the file is closed or changed on disk
    fn set(&self, py: Python<'_>, target: PyBackedStr, props: Bound<'_, PyAny>) -> PyResult<File> {
        let file = self.handle()?;
        let props = props::props(Some(&props))?;
        let next = py.detach(|| {
            file.set(&*target, &props)
                .map_err(|error| Failure::read(&error))
        })?;
        Ok(Self::new(next))
    }

    /// Add an element under ``parent``; return the new File and the new
    /// element's path.
    ///
    /// Without ``index``, ``after`` or ``before`` (at most one), the element
    /// is added after the parent's last child.
    ///
    /// Args:
    ///     parent: The path of the new element's parent ("/body", "/",
    ///         "/slide[1]")
    ///     type: The element type: "paragraph", "run", "sheet", "slide",
    ///         "shape", ...
    ///     props: The new element's properties, as for ``set``
    ///     index: Insert at this zero-based position among the children
    ///     after: Insert right after the element at this path
    ///     before: Insert right before the element at this path
    ///
    /// Returns:
    ///     A tuple (new File, the new element's canonical path)
    ///
    /// Raises:
    ///     As ``set``
    #[pyo3(signature = (parent, r#type, props = None, *, index = None, after = None, before = None))]
    #[allow(
        clippy::too_many_arguments,
        reason = "the keyword arguments mirror OfficeCLI's add options"
    )]
    fn add(
        &self,
        py: Python<'_>,
        parent: PyBackedStr,
        r#type: PyBackedStr,
        props: Option<Bound<'_, PyAny>>,
        index: Option<i64>,
        after: Option<String>,
        before: Option<String>,
    ) -> PyResult<(File, String)> {
        let file = self.handle()?;
        let props = props::props(props.as_ref())?;
        let place = Where {
            to: None,
            index,
            after,
            before,
        };
        let added = py.detach(|| {
            let node = write::new_node(&r#type, &props, &place)?;
            file.add(&*parent, node)
                .map_err(|error| Failure::read(&error))
        })?;
        let path = added.path.to_string();
        Ok((Self::new(added.file), path))
    }

    /// Remove the element(s) ``target`` names; return the new File.
    ///
    /// Args:
    ///     target: An OfficeCLI path, or a selector scoped to a path
    ///
    /// Returns:
    ///     The new File
    ///
    /// Raises:
    ///     As ``set``
    fn remove(&self, py: Python<'_>, target: PyBackedStr) -> PyResult<File> {
        let file = self.handle()?;
        let next = py.detach(|| file.remove(&*target).map_err(|error| Failure::read(&error)))?;
        Ok(Self::new(next))
    }

    /// Move the element at ``path``; return the new File.
    ///
    /// Without ``to``, the element stays under its parent and ``index``,
    /// ``after`` or ``before`` (at most one) says where it goes; without any
    /// of them it goes last. With ``to``, it moves under that parent, last
    /// or at ``index``. Positions count after the element is taken out.
    ///
    /// Args:
    ///     path: The element to move
    ///     to: The new parent's path
    ///     index: The zero-based position to move to
    ///     after: Move right after the element at this path
    ///     before: Move right before the element at this path
    ///
    /// Returns:
    ///     The new File
    ///
    /// Raises:
    ///     As ``set``
    #[pyo3(name = "move", signature = (path, to = None, *, index = None, after = None, before = None))]
    #[allow(
        clippy::too_many_arguments,
        reason = "the keyword arguments mirror OfficeCLI's move options"
    )]
    fn move_(
        &self,
        py: Python<'_>,
        path: PyBackedStr,
        to: Option<String>,
        index: Option<i64>,
        after: Option<String>,
        before: Option<String>,
    ) -> PyResult<File> {
        let file = self.handle()?;
        let place = Where {
            to,
            index,
            after,
            before,
        };
        let next = py.detach(|| {
            let placement = place.placement()?;
            file.move_to(&*path, placement)
                .map_err(|error| Failure::read(&error))
        })?;
        Ok(Self::new(next))
    }

    /// Swap the elements at ``path`` and ``path2``; return the new File.
    ///
    /// Raises:
    ///     As ``set``
    fn swap(&self, py: Python<'_>, path: PyBackedStr, path2: PyBackedStr) -> PyResult<File> {
        let file = self.handle()?;
        let next = py.detach(|| {
            file.swap(&*path, &*path2)
                .map_err(|error| Failure::read(&error))
        })?;
        Ok(Self::new(next))
    }

    /// Start a transaction, or run ``fn`` in one.
    ///
    /// Without ``fn``, return an Edit: use it as a context manager (it
    /// commits when the block ends normally) or call its ``commit()``.
    /// With ``fn``, call ``fn(edit)`` and commit: the operations apply
    /// together or not at all, and the returned Commit's ``value`` is what
    /// ``fn`` returned. If ``fn`` raises, nothing is committed and the
    /// exception propagates.
    ///
    /// Args:
    ///     fn: A callable that takes the Edit
    ///
    /// Returns:
    ///     An Edit without ``fn``; the Commit with it
    ///
    /// Raises:
    ///     The first failed operation's error (its ``item`` names the
    ///     operation), or what ``fn`` raised
    #[pyo3(signature = (r#fn = None))]
    fn edit(&self, py: Python<'_>, r#fn: Option<Bound<'_, PyAny>>) -> PyResult<Py<PyAny>> {
        let file = self.handle()?;
        let edit = Bound::new(py, Edit::new(file))?;
        let Some(callable) = r#fn else {
            return Ok(edit.into_any().unbind());
        };
        match callable.call1((edit.clone(),)) {
            Ok(value) => Ok(edit.get().commit_with(py, Some(value.unbind()))?.into_any()),
            Err(error) => {
                edit.get().discard(py);
                Err(error)
            },
        }
    }

    /// Run OfficeCLI batch items; return the Commit.
    ///
    /// Each item is a dict in OfficeCLI's ``BatchItem`` shape: ``command``
    /// (or ``op``) and the fields it takes (``path``, ``path2``,
    /// ``parent``, ``type``, ``to``, ``index``, ``after``, ``before``,
    /// ``props``, ...). Keys match ASCII case-insensitively; an unknown or
    /// repeated key, or a value of the wrong type, is ``corrupt_batch``.
    /// ``items`` can also be JSON text (``str`` or ``bytes``) holding the
    /// array, as OfficeCLI's ``batch`` reads it. The items are converted to
    /// Rust once and run with the GIL released.
    ///
    /// In "atomic" mode (the default) the first failure raises, with the
    /// item's ``index`` as ``item``, and nothing changes. In "best_effort"
    /// mode a failed item is skipped and reported in ``results``.
    ///
    /// Args:
    ///     items: A list of item dicts, or JSON text
    ///     mode: "atomic" or "best_effort"
    ///
    /// Returns:
    ///     A Commit: ``file``, ``changed``, ``operations`` and ``results``,
    ///     one dict per item (``index``, ``success``, ``changed``, and
    ///     ``path`` for an add; ``code``, ``error`` and ``suggestion`` for a
    ///     failed item)
    ///
    /// Raises:
    ///     PropertyError: (a ValueError) If an item is malformed
    ///         ("corrupt_batch") or mode is unknown
    ///     LitchiError: The first failed item's error in atomic mode, and in
    ///         either mode a protection, I/O or resource-limit failure
    #[pyo3(signature = (items, *, mode = "atomic"))]
    fn batch(&self, py: Python<'_>, items: Bound<'_, PyAny>, mode: &str) -> PyResult<Commit> {
        let file = self.handle()?;
        let mode = write::mode(mode)?;
        let items = Items::extract(&items)?;
        let commit = py.detach(|| items.run(&file, mode))?;
        let results = write::results(py, commit.value())?;
        let diagnostics = *commit.diagnostics();
        let (file, patch, _) = commit.into_parts();
        Commit::new(
            py,
            file,
            Some(write::Patch::one(patch)),
            diagnostics.changed(),
            diagnostics.operations(),
            Some(results),
            None,
        )
    }

    /// Run one OfficeCLI batch item; return the Commit.
    ///
    /// As ``batch`` with one item in atomic mode.
    ///
    /// Args:
    ///     item: An item dict, such as {"command": "set", "path":
    ///         "/body/p[1]", "props": {"bold": True}}
    ///
    /// Returns:
    ///     A Commit whose ``results`` holds the item's result
    fn apply(&self, py: Python<'_>, item: Bound<'_, PyAny>) -> PyResult<Commit> {
        let file = self.handle()?;
        let item = write::item_of(&item).map_err(|failure| failure.with_item(0))?;
        let commit = py.detach(|| file.apply(item).map_err(|error| Failure::read(&error)))?;
        let diagnostics = *commit.diagnostics();
        let (file, patch, applied) = commit.into_parts();
        let result = litchi::ItemResult::Applied(applied);
        let results = write::results(py, std::slice::from_ref(&result))?;
        Commit::new(
            py,
            file,
            Some(write::Patch::one(patch)),
            diagnostics.changed(),
            diagnostics.operations(),
            Some(results),
            None,
        )
    }

    /// Import CSV or TSV text into a worksheet; return the Commit.
    ///
    /// ``target`` is a sheet ("/Sheet1", origin A1) or an origin cell
    /// ("/Sheet1/B2") of an XLSX, XLSB, XLS or ODS workbook. Every field is
    /// written in one atomic commit; empty input changes nothing. Numbers,
    /// ISO dates and TRUE/FALSE are inferred unless ``infer`` is False;
    /// fields starting with "=" become formulas only with ``formulas``. The
    /// GIL is released while the data is parsed and written.
    ///
    /// Args:
    ///     target: The sheet or origin cell path
    ///     data: The text, as bytes (any supported encoding) or str
    ///     delimiter: One ASCII character ("\\t" or "tab" for a tab); by
    ///         default the data's "sep=" line, else ","
    ///     quote: The quote character, or None to read fields verbatim
    ///     header: Whether the first record is a header, written as text
    ///     encoding: "auto" (byte-order mark, else UTF-8), "utf-8",
    ///         "utf-16le" or "utf-16be"
    ///     infer: Whether numbers, dates and booleans are inferred
    ///     formulas: Whether "=" fields become formulas
    ///     max_cells: The largest number of fields (default 65,536)
    ///     max_bytes: The largest source and staged payload (default 16 MiB)
    ///
    /// Returns:
    ///     A Commit whose ``value`` is the summary dict: ``sheet``, ``row``,
    ///     ``column`` (zero-based origin), ``rows``, ``columns``, ``cells``,
    ///     ``texts``, ``numbers``, ``booleans``, ``dates``,
    ///     ``dates_as_text``, ``formulas`` and ``empty``
    ///
    /// Raises:
    ///     PathError: (a LookupError) If the target does not resolve
    ///     PropertyError: (a ValueError) If the text is malformed, an option
    ///         is invalid or a field is too long
    ///     UnsupportedError: (a ValueError) If the format has no worksheets
    ///     LitchiError: If a limit is exceeded ("limit_exceeded")
    #[pyo3(signature = (
        target,
        data,
        *,
        delimiter = None,
        quote = Some("\""),
        header = false,
        encoding = "auto",
        infer = true,
        formulas = false,
        max_cells = None,
        max_bytes = None
    ))]
    #[allow(clippy::too_many_arguments, reason = "the Python keyword arguments")]
    fn import_csv(
        &self,
        py: Python<'_>,
        target: PyBackedStr,
        data: Bound<'_, PyAny>,
        delimiter: Option<&str>,
        quote: Option<&str>,
        header: bool,
        encoding: &str,
        infer: bool,
        formulas: bool,
        max_cells: Option<usize>,
        max_bytes: Option<usize>,
    ) -> PyResult<Commit> {
        let file = self.handle()?;
        let data = m4::source(&data)?;
        let (options, limits) = m4::options(
            delimiter, quote, header, encoding, infer, formulas, max_cells, max_bytes,
        )?;
        let commit = py.detach(|| {
            file.import_with_options(
                &*target,
                &data,
                &options,
                litchi::EditOptions::new().semantic_limits(limits),
            )
            .map_err(|error| Failure::read(&error))
        })?;
        m4::import_commit(py, commit)
    }

    /// Fill this template's ``{{key}}`` placeholders (and, for DOCX, its
    /// Word ``MERGEFIELD`` fields) from one data record; return the Commit.
    ///
    /// The template is unchanged, so one opened template fills any number
    /// of records. DOCX, PPTX, XLSX, XLSB, DOC, PPT, XLS, RTF, ODT, OTH,
    /// ODM, ODS and ODP templates. Placeholders split across runs
    /// keep each run's formatting. The GIL is released while merging.
    ///
    /// Args:
    ///     data: One record: a mapping (values rendered as JSON renders
    ///         them; nested members also as "a.b" and "a[0]" keys), or a
    ///         JSON object as str or bytes
    ///     missing: "keep" (the default) leaves a placeholder whose key the
    ///         record lacks; "empty" removes it
    ///     fields: Whether Word merge fields are filled
    ///     max_replacements: The largest number of replacements (default
    ///         65,536)
    ///     max_bytes: The largest replacement payload in bytes (default
    ///         16 MiB)
    ///
    /// Returns:
    ///     A Commit whose ``value`` is the summary dict: ``replacements``,
    ///     ``fields``, ``skipped_fields``, ``keys`` (the record keys used)
    ///     and ``unresolved`` (the keys the record lacks)
    ///
    /// Raises:
    ///     PropertyError: (a ValueError) If the data is not one JSON
    ///         object, an option is invalid, or a value holds a character
    ///         the document cannot store
    ///     UnsupportedError: (a ValueError) If the format has no merge
    ///         adapter (the ODF object formats)
    ///     LitchiError: If a limit is exceeded ("limit_exceeded")
    #[pyo3(signature = (
        data,
        *,
        missing = "keep",
        fields = true,
        max_replacements = None,
        max_bytes = None
    ))]
    fn merge(
        &self,
        py: Python<'_>,
        data: Bound<'_, PyAny>,
        missing: &str,
        fields: bool,
        max_replacements: Option<usize>,
        max_bytes: Option<usize>,
    ) -> PyResult<Commit> {
        let file = self.handle()?;
        let source = merge::json_text(&data)?;
        let (options, limits) = merge::options(missing, fields, max_replacements, max_bytes)?;
        let commit = py.detach(|| {
            let record = merge::record(&source, &limits)?;
            file.merge_with_options(
                &record,
                &options,
                litchi::EditOptions::new().semantic_limits(limits),
            )
            .map_err(|error| Failure::read(&error))
        })?;
        merge::commit(py, commit)
    }

    /// Recompute the document's derived field results; return the Commit.
    ///
    /// DOCX: tables of contents (from the headings), SEQ, REF, DATE/TIME,
    /// document-property and count fields; DOC: the same fields but REF;
    /// ODT, OTH, ODM: date, time and document-property fields; PPTX and
    /// ODP: slide-number and date fields; workbooks (XLSX, XLSB, XLS):
    /// formula results (written back as cached values only with
    /// ``write_cache``). Page numbers need a layout
    /// engine Litchi does not have: PAGE-like fields keep their cached
    /// result, and a table of contents keeps each entry's cached page
    /// number. The file is unchanged; a document whose results are current
    /// gives an unchanged Commit. The GIL is released while refreshing.
    ///
    /// Args:
    ///     clock: The date and time date fields show: "YYYY-MM-DD[THH:MM[:SS]]",
    ///         a datetime.datetime (its wall time) or a datetime.date. Without
    ///         it date fields keep their cached result; the system clock is
    ///         never read
    ///     write_cache: Write recalculated XLSX formula results back
    ///     page_numbers: "keep" (the default) keeps cached page numbers;
    ///         "require" refuses ("refresh_failed") a table of contents whose
    ///         page numbers are unknown
    ///     mark_dirty: Mark a table of contents with unknown page numbers for
    ///         its word processor to update
    ///     max_changes: The largest number of changed results (default 65,536)
    ///     max_bytes: The largest changed payload in bytes (default 16 MiB)
    ///
    /// Returns:
    ///     A Commit whose ``value`` is the summary dict: ``refreshed``,
    ///     ``bookmarks``, ``missing_page_numbers`` and ``fields`` (each
    ///     ``type``, ``path``, ``status``, ``reason``)
    ///
    /// Raises:
    ///     PropertyError: (a ValueError) If the clock or an option is invalid
    ///     UnsupportedError: (a ValueError) If the format has no refresh
    ///     LitchiError: If page numbers are required but unknown
    ///         ("refresh_failed") or a limit is exceeded ("limit_exceeded")
    #[pyo3(signature = (
        *,
        clock = None,
        write_cache = false,
        page_numbers = "keep",
        mark_dirty = true,
        max_changes = None,
        max_bytes = None
    ))]
    fn refresh(
        &self,
        py: Python<'_>,
        clock: Option<Bound<'_, PyAny>>,
        write_cache: bool,
        page_numbers: &str,
        mark_dirty: bool,
        max_changes: Option<usize>,
        max_bytes: Option<usize>,
    ) -> PyResult<Commit> {
        let file = self.handle()?;
        let clock = clock.as_ref().map(refresh::clock_text).transpose()?;
        let (options, limits) = refresh::options(
            clock.as_deref(),
            write_cache,
            page_numbers,
            mark_dirty,
            max_changes,
            max_bytes,
        )?;
        let commit = py.detach(|| {
            file.refresh_with_options(&options, litchi::EditOptions::new().semantic_limits(limits))
                .map_err(|error| Failure::read(&error))
        })?;
        refresh::commit(py, commit)
    }

    /// Evaluate one workbook cell.
    ///
    /// A value cell gives its value; a formula gives its cached result, or
    /// the computed one without a cache or with ``recalculate``. Evaluation
    /// reads only: the file is unchanged. The GIL is released while
    /// evaluating.
    ///
    /// Args:
    ///     path: The cell ("/Sheet1/B2")
    ///     recalculate: Compute the formula even when a cache exists
    ///
    /// Returns:
    ///     {"value": None | bool | float | str, "origin": "literal" |
    ///     "cached" | "computed"}; an error value is its spelling
    ///     ("#DIV/0!")
    ///
    /// Raises:
    ///     UnsupportedError: (a ValueError) If the formula uses syntax or a
    ///         function the engine does not evaluate, or calls NOW/TODAY/RAND
    ///         without a cached value
    ///     PropertyError: (a ValueError) If the path is not a cell
    ///     LitchiError: If a limit is exceeded ("limit_exceeded")
    #[pyo3(signature = (path, *, recalculate = false))]
    fn evaluate<'py>(
        &self,
        py: Python<'py>,
        path: PyBackedStr,
        recalculate: bool,
    ) -> PyResult<Bound<'py, PyDict>> {
        let file = self.handle()?;
        let options = m4::calc_options(recalculate);
        let evaluated = py.detach(|| {
            file.evaluate(&*path, &options)
                .map_err(|error| Failure::read(&error))
        })?;
        m4::evaluated(py, &evaluated)
    }

    /// Recalculate every formula of the workbook and compare each result
    /// with its cached value. The file is unchanged; the GIL is released
    /// while evaluating.
    ///
    /// Returns:
    ///     A dict with the counts ``formulas``, ``match``, ``mismatch``,
    ///     ``uncached``, ``unsupported`` and ``volatile``, and ``cells``: one
    ///     dict per formula (``sheet``, ``row``, ``column`` zero-based,
    ///     ``formula``, ``cached_value``, ``computed_value``, ``status``)
    fn recalculate<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let file = self.handle()?;
        let options = m4::calc_options(true);
        let report = py.detach(|| {
            file.recalculate(&options)
                .map_err(|error| Failure::read(&error))
        })?;
        m4::report(py, &report)
    }

    /// Save the File atomically, in its own format; return the bytes
    /// written.
    ///
    /// Without ``path``, the File is written back to the path it was opened
    /// from; with one, this is ``save_as``. The file is written to a
    /// temporary sibling and renamed over the destination, so a failed save
    /// leaves the destination as it was. The GIL is released while the file
    /// is written.
    ///
    /// Args:
    ///     path: Where to save; None saves to the path the File was opened
    ///         from
    ///     overwrite: Whether an existing file at the destination may be
    ///         replaced
    ///
    /// Returns:
    ///     The number of bytes written
    ///
    /// Raises:
    ///     IoError: (an OSError) If the File has no path ("file_required";
    ///         it was created or opened from bytes), the destination exists
    ///         and overwrite is False ("file_exists"), the source changed, or
    ///         writing fails
    ///     ProtectedError: (a ValueError) If the File is encrypted and
    ///         locked, or saving would remove its encryption
    ///     UnsupportedError: (a ValueError) If the format cannot be saved, or
    ///         path names another format's extension
    #[pyo3(signature = (path = None, *, overwrite = true))]
    fn save(&self, py: Python<'_>, path: Option<PathBuf>, overwrite: bool) -> PyResult<u64> {
        let file = self.handle()?;
        let options = SaveOptions::new().overwrite(if overwrite {
            Overwrite::Replace
        } else {
            Overwrite::Refuse
        });
        let saved = py.detach(|| {
            let to = match &path {
                Some(path) => SaveTo::Path(path),
                None => SaveTo::Origin,
            };
            file.save_with(to, &options)
                .map_err(|error| Failure::read(&error))
        })?;
        Ok(saved.bytes)
    }

    /// Save the File atomically to ``path``, in its own format; return the
    /// bytes written. The File keeps the path it was opened from.
    ///
    /// Args:
    ///     path: Where to save
    ///     overwrite: Whether an existing file at path may be replaced
    ///
    /// Returns:
    ///     The number of bytes written
    ///
    /// Raises:
    ///     As ``save``
    #[pyo3(signature = (path, *, overwrite = true))]
    fn save_as(&self, py: Python<'_>, path: PathBuf, overwrite: bool) -> PyResult<u64> {
        self.save(py, Some(path), overwrite)
    }

    /// The File's bytes, in its own format.
    ///
    /// The document is written with the GIL released, then copied once into
    /// the returned ``bytes``.
    ///
    /// Returns:
    ///     The file's content
    ///
    /// Raises:
    ///     As ``save``, without the path errors
    fn to_bytes<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyBytes>> {
        let file = self.handle()?;
        let bytes = py.detach(|| file.to_bytes().map_err(|error| Failure::read(&error)))?;
        Ok(PyBytes::new(py, &bytes))
    }

    /// The document as an OfficeCLI v2 batch: a list of item dicts that
    /// ``batch`` replays onto ``create(format)`` to rebuild it (record 0336).
    ///
    /// The first item is ``{"command": "meta", "dumpVersion": 2}``. Content
    /// the write verbs cannot rebuild (pictures, charts, fields, comments,
    /// resources other than styles) is skipped; with ``strict`` it raises
    /// instead. The GIL is released while the document is walked.
    ///
    /// Args:
    ///     strict: Raise UnsupportedError ("unsupported_element") when the
    ///         dump would skip anything
    ///
    /// Returns:
    ///     The items, as dicts
    ///
    /// Raises:
    ///     UnsupportedError: (a ValueError) If the format cannot be dumped
    ///         ("unsupported_capability"), or with ``strict`` when content
    ///         was skipped
    ///     ProtectedError: (a ValueError) If the file is encrypted
    ///     IoError: (an OSError) If the file is closed or changed on disk
    #[pyo3(signature = (*, strict = false))]
    fn dump<'py>(&self, py: Python<'py>, strict: bool) -> PyResult<Bound<'py, PyAny>> {
        let file = self.handle()?;
        let dump = py.detach(|| file.dump().map_err(|error| Failure::read(&error)))?;
        if let Some(warning) = dump.warnings().first().filter(|_| strict) {
            return Err(Failure::read(warning).into());
        }
        pyser::to_py(py, &dump)
    }

    /// The package parts: a list of ``{"name", "contentType"}`` dicts in
    /// name order (design §12, record 0336).
    ///
    /// Raises:
    ///     UnsupportedError: (a ValueError) If the format has no package
    ///         parts (legacy binary, RTF, flat ODF, iWork)
    ///     IoError: (an OSError) If the file is closed
    fn parts<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let file = self.handle()?;
        let parts = py.detach(|| file.parts().map_err(|error| Failure::read(&error)))?;
        pyser::to_py(py, &parts)
    }

    /// One package part's bytes, as the File would publish them (OfficeCLI
    /// ``raw``; record 0336).
    ///
    /// Args:
    ///     part: A part name ("/word/document.xml") or alias ("/document",
    ///         "/styles", "/slide[1]", "/Sheet1", "/Sheet1/chart[1]",
    ///         "/content")
    ///
    /// Returns:
    ///     The part's bytes (decode XML parts with ``.decode()``)
    ///
    /// Raises:
    ///     PathError: (a LookupError) If there is no such part
    ///     UnsupportedError: (a ValueError) As ``parts``
    fn raw<'py>(&self, py: Python<'py>, part: PyBackedStr) -> PyResult<Bound<'py, PyBytes>> {
        let file = self.handle()?;
        let part = py.detach(|| file.raw(&part).map_err(|error| Failure::read(&error)))?;
        Ok(PyBytes::new(py, part.bytes()))
    }

    /// Add a package part with its content type and a relationship from
    /// ``parent``; return a Commit whose ``value`` is ``{"part",
    /// "relationshipId"}`` (OfficeCLI ``add-part``; record 0336).
    ///
    /// ``type`` is "chart", "header" or "footer" for an empty part of that
    /// kind; otherwise pass ``content_type``, ``relationship``, ``data``
    /// and ``name``. ``parent`` "/" relates the part from the package.
    ///
    /// Raises:
    ///     PathError: (a LookupError) If the parent does not resolve
    ///     PropertyError: (a ValueError) If arguments are missing or the
    ///         type is unknown
    ///     UnsupportedError: (a ValueError) Outside DOCX, XLSX, PPTX and XLSB
    ///     ProtectedError: (a ValueError) If the File is signed or encrypted
    #[pyo3(signature = (parent, r#type = None, *, content_type = None, relationship = None, data = None, name = None))]
    #[allow(clippy::too_many_arguments, reason = "the Python keyword arguments")]
    fn add_part(
        &self,
        py: Python<'_>,
        parent: PyBackedStr,
        r#type: Option<PyBackedStr>,
        content_type: Option<String>,
        relationship: Option<String>,
        data: Option<Vec<u8>>,
        name: Option<String>,
    ) -> PyResult<Commit> {
        let file = self.handle()?;
        let part = crate::raw::new_part(r#type.as_deref(), content_type, relationship, data, name)?;
        let added = py.detach(|| {
            file.add_part(&parent, part)
                .map_err(|error| Failure::read(&error))
        })?;
        let value = PyDict::new(py);
        value.set_item("part", &added.part)?;
        value.set_item("relationshipId", &added.relationship_id)?;
        Commit::new(
            py,
            added.file,
            None,
            true,
            1,
            None,
            Some(value.into_any().unbind()),
        )
    }

    /// Edit one part's XML at every element ``xpath`` matches; return a
    /// Commit whose ``value`` is ``{"matched": n}`` (OfficeCLI ``raw-set``;
    /// record 0336).
    ///
    /// The XPath is a location path of "/" and "//" steps with prefixed
    /// names or "*" and [n], [last()] or [@name='value'] predicates.
    /// ``action`` is append, prepend, insertbefore, insertafter, replace,
    /// remove or setattr (``xml`` is then "name=value"). The edited part
    /// passes the package's publication audit before the new File opens.
    ///
    /// Raises:
    ///     PathError: (a LookupError) If the part does not resolve, the
    ///         XPath is malformed ("invalid_xpath") or matches nothing
    ///         ("zero_matches")
    ///     PropertyError: (a ValueError) If the action or fragment is invalid
    ///     UnsupportedError: (a ValueError) Outside DOCX, XLSX, PPTX and XLSB
    ///     ProtectedError: (a ValueError) If the File is signed or encrypted
    #[pyo3(signature = (part, xpath, action, xml = None))]
    fn raw_set(
        &self,
        py: Python<'_>,
        part: PyBackedStr,
        xpath: PyBackedStr,
        action: PyBackedStr,
        xml: Option<PyBackedStr>,
    ) -> PyResult<Commit> {
        let file = self.handle()?;
        let action: litchi::raw::Action = action.parse().map_err(|error| Failure::read(&error))?;
        let edited = py.detach(|| {
            file.raw_set(&part, &xpath, action, xml.as_deref())
                .map_err(|error| Failure::read(&error))
        })?;
        let value = PyDict::new(py);
        value.set_item("matched", edited.matched)?;
        Commit::new(
            py,
            edited.file,
            None,
            true,
            1,
            None,
            Some(value.into_any().unbind()),
        )
    }

    /// Check the package structure read-only (OfficeCLI ``validate``;
    /// record 0336): a dict with ``format``, ``valid`` (no error or fatal
    /// finding), ``complete``, ``checks`` and ``issues`` (each with
    /// ``code``, ``severity``, ``message`` and ``locations``).
    ///
    /// Raises:
    ///     UnsupportedError: (a ValueError) For iWork packages
    ///     ProtectedError: (a ValueError) If the file is encrypted
    ///     IoError: (an OSError) If the file is closed
    fn validate<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let file = self.handle()?;
        let validation = py.detach(|| file.validate().map_err(|error| Failure::read(&error)))?;
        pyser::to_py(py, &validation)
    }

    /// A File that may edit this signed document.
    ///
    /// Editing a signed document invalidates its signature, so the write
    /// verbs refuse it (ProtectedError). The returned File reads the same
    /// content; its first write removes the signature and publishes the
    /// document without it. This File stays signed. A File that is not
    /// signed is returned as it is.
    ///
    /// Returns:
    ///     The File that may be edited
    ///
    /// Raises:
    ///     IoError: (an OSError) If the file is closed
    fn unsign(&self) -> PyResult<File> {
        let file = self.handle()?;
        Ok(Self::new(match file {
            litchi::File::Signed(signed) => signed.unsign(),
            other => other,
        }))
    }

    /// Whether the File is signed: its write verbs refuse until
    /// ``unsign()``. False after ``close()``.
    #[getter]
    fn signed(&self) -> bool {
        self.handle
            .read()
            .unwrap_or_else(PoisonError::into_inner)
            .as_ref()
            .is_some_and(litchi::File::is_signed)
    }

    /// Whether the File is encrypted and locked: it reads and writes
    /// nothing. False after ``close()``.
    #[getter]
    fn locked(&self) -> bool {
        matches!(
            self.handle
                .read()
                .unwrap_or_else(PoisonError::into_inner)
                .as_ref(),
            Some(litchi::File::Locked(_))
        )
    }

    /// Close the file: drop this handle's reference to the document. Reads
    /// raise IoError ("file_required") afterwards; closing again does
    /// nothing. A read already running on another thread finishes first.
    fn close(&self, py: Python<'_>) {
        let file = self
            .handle
            .write()
            .unwrap_or_else(PoisonError::into_inner)
            .take();
        if let Some(file) = file {
            // Freeing a document's caches can take a while.
            py.detach(move || drop(file));
        }
    }

    fn __enter__(slf: Bound<'_, Self>) -> PyResult<Bound<'_, Self>> {
        slf.get().handle()?;
        Ok(slf)
    }

    fn __exit__(
        &self,
        py: Python<'_>,
        exc_type: Option<Bound<'_, PyAny>>,
        exc_value: Option<Bound<'_, PyAny>>,
        traceback: Option<Bound<'_, PyAny>>,
    ) -> bool {
        let _ = (exc_type, exc_value, traceback);
        self.close(py);
        false
    }

    fn __repr__(&self) -> String {
        let format = match self.format {
            Some(format) => format!("{format:?}"),
            None => "unknown format".to_owned(),
        };
        let state = match self
            .handle
            .read()
            .unwrap_or_else(PoisonError::into_inner)
            .as_ref()
        {
            None => ", closed",
            Some(litchi::File::Locked(_)) => ", locked",
            Some(file) if file.is_signed() => ", signed",
            Some(_) => "",
        };
        format!("<File: {format}{state}>")
    }
}

/// The node view options of `get`, `find` and `query`.
pub(crate) fn view_options(
    depth: i64,
    max_nodes: i64,
    dense: Option<i64>,
) -> Result<ViewOptions, Failure> {
    let mut options = ViewOptions::default()
        .depth(count("depth", depth)?)
        .max_nodes(count("max_nodes", max_nodes)?);
    if let Some(cells) = dense {
        options = options.dense(count("dense", cells)?);
    }
    Ok(options)
}

/// The view mode nearest to an unknown `name`, within two edits; only on
/// the error path.
fn nearest_mode(name: &str) -> Option<&'static str> {
    let name = name.to_ascii_lowercase();
    Mode::ALL
        .iter()
        .map(|mode| mode.name())
        .map(|candidate| (distance(&name, candidate), candidate))
        .filter(|(distance, _)| *distance <= 2)
        .min_by_key(|(distance, _)| *distance)
        .map(|(_, candidate)| candidate)
}

/// The Levenshtein distance of two short ASCII names.
fn distance(a: &str, b: &str) -> usize {
    let b = b.as_bytes();
    let mut row: Vec<usize> = (0..=b.len()).collect();
    for (i, &left) in a.as_bytes().iter().enumerate() {
        let mut diagonal = row[0];
        row[0] = i + 1;
        for (j, &right) in b.iter().enumerate() {
            let above = row[j + 1];
            row[j + 1] = (above + 1)
                .min(row[j] + 1)
                .min(diagonal + usize::from(left != right));
            diagonal = above;
        }
    }
    row[b.len()]
}

/// `query`'s `limit` argument: negative is `invalid_argument`.
pub(crate) fn limit_of(limit: Option<i64>) -> Result<Option<usize>, Failure> {
    limit
        .map(|limit| {
            usize::try_from(limit)
                .map_err(|_| Failure::argument(format!("limit must not be negative, got {limit}")))
        })
        .transpose()
}

/// A count argument: negative is `invalid_argument`; beyond `u32::MAX`
/// counts as `u32::MAX`, which no document reaches.
fn count(name: &str, value: i64) -> Result<u32, Failure> {
    if value < 0 {
        return Err(Failure::argument(format!(
            "{name} must not be negative, got {value}"
        )));
    }
    Ok(u32::try_from(value).unwrap_or(u32::MAX))
}
