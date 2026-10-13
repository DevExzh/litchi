//! The command-line grammar.
//!
//! Verbs, argument order and option names follow OfficeCLI
//! (OfficeCLI `src/officecli/CommandBuilder*.cs`): the file comes
//! first, then the path, then options. Litchi-specific spellings are noted on
//! each option.

use std::path::PathBuf;

use clap::{Args, Parser, Subcommand};

/// Open, read, query and edit Office documents through Litchi's unified
/// operations layer.
///
/// Writes save in place atomically unless `--out` names another file. Exit
/// statuses: 0 success, 1 other failure, 2 usage, 3 path, 4 property or
/// argument, 5 unsupported, 6 I/O, 7 corrupt input, 8 protected, 9 a
/// best-effort batch with failed items.
#[derive(Debug, Parser)]
#[command(name = "litchi", version, about, long_about = None)]
pub(crate) struct Cli {
    /// Print JSON: the node view for reads, the batch result for writes.
    #[arg(long, global = true)]
    json: bool,
    #[command(subcommand)]
    pub(crate) command: Command,
}

impl Cli {
    /// Whether `--json` was given.
    pub(crate) const fn json(&self) -> bool {
        self.json
    }
}

/// The verbs.
#[derive(Debug, Subcommand)]
pub(crate) enum Command {
    /// Create a blank document; the type comes from `--type` or the
    /// extension.
    Create(Create),
    /// Print the node at a path.
    Get(Get),
    /// Print every node a selector matches.
    Query(Query),
    /// Render a document view: text, annotated, outline, stats or issues.
    View(View),
    /// Set properties on the node at a path, or on every match of a scoped
    /// selector.
    Set(Set),
    /// Add a new element under a parent.
    Add(Add),
    /// Remove the node at a path.
    Remove(Remove),
    /// Move a node to another position or parent.
    Move(Move),
    /// Swap two nodes.
    Swap(Swap),
    /// Run a JSON array of OfficeCLI batch items.
    Batch(Batch),
    /// Print what a file or a format supports.
    Capabilities(Capabilities),
    /// Print a format's property registry: every element's properties with
    /// type, unit, values, aliases, constraint and get/set/add flags.
    Schema(Schema),
    /// Import CSV/TSV data into a worksheet (record 0266).
    Import(crate::m4::Import),
    /// Evaluate one cell: its value, cached formula result or computed
    /// result (record 0266).
    Evaluate(crate::m4::Evaluate),
    /// Recalculate every formula and compare it with its cached value
    /// (record 0266).
    Recalculate(crate::m4::Recalculate),
}

/// Where a write publishes its result.
#[derive(Debug, Args)]
pub(crate) struct Output {
    /// Save to this path instead of replacing the input file.
    #[arg(long, value_name = "FILE")]
    pub(crate) out: Option<PathBuf>,
}

/// Placement options shared by `add` and `move`.
#[derive(Debug, Args)]
pub(crate) struct Position {
    /// Zero-based position among the parent's children (OfficeCLI
    /// `--index`; `--at` is accepted as the same option).
    #[arg(long, visible_alias = "at", value_name = "N", conflicts_with_all = ["after", "before"])]
    pub(crate) index: Option<i32>,
    /// Insert immediately after the element at this path.
    #[arg(long, value_name = "PATH", conflicts_with = "before")]
    pub(crate) after: Option<String>,
    /// Insert immediately before the element at this path.
    #[arg(long, value_name = "PATH")]
    pub(crate) before: Option<String>,
}

#[derive(Debug, Args)]
pub(crate) struct Create {
    /// The file to create.
    pub(crate) file: PathBuf,
    /// The document type (docx, xlsx, pptx, ...); defaults to the extension.
    #[arg(long = "type", value_name = "TYPE")]
    pub(crate) kind: Option<String>,
    /// Replace an existing file instead of refusing it.
    #[arg(long)]
    pub(crate) force: bool,
}

#[derive(Debug, Args)]
pub(crate) struct Get {
    /// The document.
    pub(crate) file: PathBuf,
    /// The node path, for example `/body/p[1]`; `/` is the root.
    #[arg(default_value = "/")]
    pub(crate) path: String,
    /// Child levels to list below the node (0 lists none).
    #[arg(long, default_value_t = 1)]
    pub(crate) depth: u32,
}

#[derive(Debug, Args)]
pub(crate) struct Query {
    /// The document.
    pub(crate) file: PathBuf,
    /// A CSS-like selector, for example `p[style=Heading1]`.
    pub(crate) selector: String,
    /// Stop after this many matches.
    #[arg(long, value_name = "N")]
    pub(crate) limit: Option<usize>,
    /// Child levels to list below each match (0 lists none).
    #[arg(long, default_value_t = 1)]
    pub(crate) depth: u32,
}

#[derive(Debug, Args)]
pub(crate) struct View {
    /// The document.
    pub(crate) file: PathBuf,
    /// The view mode, positionally as OfficeCLI takes it.
    #[arg(value_name = "MODE", conflicts_with = "mode")]
    pub(crate) positional_mode: Option<String>,
    /// The view mode: text, annotated, outline, stats or issues (default
    /// text).
    #[arg(long, value_name = "MODE")]
    pub(crate) mode: Option<String>,
}

#[derive(Debug, Args)]
pub(crate) struct Set {
    /// The document.
    pub(crate) file: PathBuf,
    /// The node path or scoped selector.
    pub(crate) path: String,
    /// A property assignment `key=value`; repeat for several.
    #[arg(long = "prop", value_name = "KEY=VALUE", required = true)]
    pub(crate) props: Vec<String>,
    #[command(flatten)]
    pub(crate) output: Output,
}

#[derive(Debug, Args)]
pub(crate) struct Add {
    /// The document.
    pub(crate) file: PathBuf,
    /// The parent path, for example `/body` or `/Sheet1`.
    pub(crate) parent: String,
    /// The element type, for example paragraph, run, sheet, slide, shape.
    #[arg(long = "type", value_name = "TYPE")]
    pub(crate) kind: String,
    /// A property assignment `key=value`; repeat for several.
    #[arg(long = "prop", value_name = "KEY=VALUE")]
    pub(crate) props: Vec<String>,
    #[command(flatten)]
    pub(crate) position: Position,
    #[command(flatten)]
    pub(crate) output: Output,
}

#[derive(Debug, Args)]
pub(crate) struct Remove {
    /// The document.
    pub(crate) file: PathBuf,
    /// The node path or scoped selector.
    pub(crate) path: String,
    #[command(flatten)]
    pub(crate) output: Output,
}

#[derive(Debug, Args)]
pub(crate) struct Move {
    /// The document.
    pub(crate) file: PathBuf,
    /// The node to move.
    pub(crate) path: String,
    /// The new parent; without it the node moves among its siblings.
    #[arg(long, value_name = "PATH")]
    pub(crate) to: Option<String>,
    #[command(flatten)]
    pub(crate) position: Position,
    #[command(flatten)]
    pub(crate) output: Output,
}

#[derive(Debug, Args)]
pub(crate) struct Swap {
    /// The document.
    pub(crate) file: PathBuf,
    /// The first node.
    pub(crate) path1: String,
    /// The second node.
    pub(crate) path2: String,
    #[command(flatten)]
    pub(crate) output: Output,
}

#[derive(Debug, Args)]
pub(crate) struct Batch {
    /// The document.
    pub(crate) file: PathBuf,
    /// A JSON file holding the items; `-` reads standard input.
    #[arg(
        long,
        value_name = "FILE",
        conflicts_with = "commands",
        required_unless_present = "commands"
    )]
    pub(crate) input: Option<PathBuf>,
    /// The items as inline JSON.
    #[arg(long, value_name = "JSON")]
    pub(crate) commands: Option<String>,
    /// Discard every change when any item fails (the default).
    #[arg(long, conflicts_with = "best_effort")]
    pub(crate) atomic: bool,
    /// Apply the items that succeed, report the ones that fail, and exit 9
    /// when any failed. Never the default.
    #[arg(long)]
    pub(crate) best_effort: bool,
    #[command(flatten)]
    pub(crate) output: Output,
}

#[derive(Debug, Args)]
pub(crate) struct Capabilities {
    /// A document whose capabilities (including its open state) to print.
    #[arg(conflicts_with = "format", required_unless_present = "format")]
    pub(crate) file: Option<PathBuf>,
    /// A format whose static capabilities to print, for example docx.
    #[arg(long, value_name = "TYPE")]
    pub(crate) format: Option<String>,
}

#[derive(Debug, Args)]
pub(crate) struct Schema {
    /// The format whose registry to print, for example docx.
    #[arg(value_name = "TYPE")]
    pub(crate) format: String,
}

#[cfg(test)]
mod tests {
    use clap::CommandFactory as _;

    #[test]
    fn grammar_is_consistent() {
        super::Cli::command().debug_assert();
    }
}
