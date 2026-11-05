//! `litchi`: the command line of Litchi's unified operations layer.
//!
//! The `litchi` binary is a thin shell over `litchi-ops`' public API: it
//! parses arguments, calls one operation, saves the result and prints it. It
//! adds no document behaviour of its own (unified-ops record 0256). Verbs and
//! options follow OfficeCLI's command shapes, so OfficeCLI example scripts
//! replay with `officecli` replaced by `litchi`.
//!
//! | Verb | `litchi-ops` call |
//! |---|---|
//! | `create <file> [--type T] [--force]` | `create`, then `File::save_with` (refusing an existing file unless `--force`) |
//! | `get <file> [<path>] [--depth N]` | `open`, `File::get`, `Node::view_with` |
//! | `query <file> <selector> [--limit N] [--depth N]` | `File::query`, `Matches::limit` |
//! | `view <file> [<mode> \| --mode M]` | `File::view` or `File::view_json` |
//! | `set <file> <path> --prop k=v... [--find T] [--replace T] [--unsupported P]` | `File::apply_with` with a `set` item (`--find T` is `--prop find=T`, `--replace T` is `--prop replace=T`) |
//! | `add <file> <parent> (--type T \| --from P) [--prop k=v...] [--index N \| --after P \| --before P] [--unsupported P]` | `File::apply_with` with an `add` item |
//! | `remove <file> <path> [--prop k=v...]` | `File::apply` with a `remove` item |
//! | `move <file> <path> [--to P] [--index N \| --after P \| --before P]` | `File::apply` with a `move` item |
//! | `swap <file> <path1> <path2>` | `File::apply` with a `swap` item |
//! | `batch <file> (--input F \| --commands JSON) [--atomic \| --best-effort] [--unsupported P]` | `batch::from_bytes` with its required `SourceLimits`, then `File::batch_with` |
//! | `capabilities (<file> \| --format T)` | `File::capabilities` or `Capabilities::of` |
//! | `schema <type>` | `litchi_ops::schema` |
//! | `merge <template> <data> --out <file> [--format json\|csv\|tsv] [--missing keep\|empty] [--no-fields] [--force]`, or OfficeCLI's `merge <template> <output> --data <json\|file>` | `merge::Data::from_json`/`from_csv`, then `File::merge` and `File::save_with` per record (refusing existing outputs unless `--force`); several records need `{n}` in the output name |
//! | `refresh <file> [--clock DATETIME] [--write-cache] [--require-page-numbers] [--no-dirty] [--out F]` | `File::refresh_with_options`, then `File::save` (in place, as OfficeCLI) or `File::save_as`; nothing is written when every result is current |
//! | `dump <file> [/] [--format batch] [--out F]` | `File::dump`; the JSON array on stdout (or `--out`), skipped content as `warning[...]` lines on stderr (record 0336) |
//! | `raw <file> [<part>] [--out F]` | `File::parts` without a part, else `File::raw`: the XML on stdout, or the bytes to `--out` |
//! | `raw-set <file> <part> --xpath X --action A [--xml X] [--out F]` | `File::raw_set`, then a save in place (or to `--out`) |
//! | `add-part <file> <parent> --type T [--name P] [--content-type C --relationship R --data F] [--out F]` | `File::add_part`; prints OfficeCLI's `Created T part: relId=... path=...` |
//! | `validate <file>` | `File::validate`; exits 1 when the report has an error, and prints the errors on stderr as OfficeCLI does |
//!
//! `--at` is accepted as another name of `--index`.
//!
//! # Writes
//!
//! Single-command writes go through the command channel (`File::apply`), the
//! same admission a batch item gets. A write saves the file in place with
//! `File::save`, the leaf's atomic sibling-temp replacement, unless `--out`
//! names another destination (`File::save_as`). A write that changes
//! nothing leaves the file untouched. A batch is atomic unless
//! `--best-effort` is given; OfficeCLI's "apply the rest and report" mode is
//! that flag, never the default.
//!
//! A property the target does not support refuses the whole command
//! (`unsupported_property`, exit 4) unless `--unsupported warn` is given on
//! `set`, `add` or `batch`. Then, as OfficeCLI does by default, the other
//! properties apply atomically, each skipped property is printed on stderr
//! as `warning[unsupported_property]: ...` with its `property` line (and in
//! the item's `warnings` with `--json`), and the run exits 9. Invalid values
//! and every other failure still refuse the command (`batch::Unsupported`,
//! record 0269). A command that still fails after skipping prints the
//! skipped properties as `warning[...]` blocks before its `error[...]`
//! (with `--json`, the error's `warnings`), and exits with the failure's own
//! status: the failure is that of the command without them
//! (`litchi_ops::Error::warnings`, record 0271).
//!
//! # Output
//!
//! Text output is for people. `--json` output is for programs:
//!
//! - `get`: the `litchi-ops` node view (OfficeCLI's `DocumentNode`), exactly
//!   what serializing `Node::view_with` produces.
//! - `query`: OfficeCLI's `NodesResult`, `{"matches": N, "results": [...]}`,
//!   each result the node view.
//! - `view`: `File::view_json`.
//! - `capabilities`: the serialized `Capabilities`.
//! - `schema`: the serialized `Schema`, OfficeCLI's help-schema shape per
//!   element with the registry's units, aliases and constraints.
//! - `create`: `{"path", "type", "bytes"}`.
//! - writes and `batch`: OfficeCLI's batch rendering, `{"results": [...],
//!   "summary": {...}}`. Each result has `index` and `success`; a success
//!   adds `changed` and, for `add`, the created `path`; a failure adds
//!   `code`, `error` (the message), `suggestion`, `property` and the failed
//!   input `item`; under `--unsupported warn` a success, or a failure after
//!   skipping, adds `warnings`, one `ErrorResult` per skipped property. The summary counts `total`, `executed`, `succeeded`,
//!   `failed` and `skipped`, and adds `changed`, `saved` (the destination)
//!   and `atomicRolledBack`.
//!
//! # Errors and exit statuses
//!
//! A failure prints `error[<code>]: <message>` and any `item`, `property` and
//! `suggestion` lines on stderr; with `--json` it prints OfficeCLI's
//! `ErrorResult` object, `{"error": message, "code": code, "suggestion": ...}`,
//! on stderr instead, so stdout carries only results. The exit status names
//! the family of `litchi_ops::Code`, the same families as the Python
//! exception classes:
//!
//! | Status | Meaning | Codes |
//! |---|---|---|
//! | 0 | success | |
//! | 1 | other failure | `internal_error`, `cancelled`, `limit_exceeded`, `output_too_large`, `max_depth_exceeded`, `validation_error`, `merge_not_exact`, `refresh_failed`, and any code added later |
//! | 2 | usage: the command line itself is malformed | (argument parser) |
//! | 3 | path | `not_found`, `path_not_found`, `zero_matches`, `out_of_range`, `range_target_not_found`, `invalid_path`, `invalid_selector`, `invalid_xpath`, `unsupported_path`, `bare_selector_rejected` |
//! | 4 | property or argument | `invalid_value`, `unsupported_property`, `missing_property`, `duplicate_name`, `invalid_argument`, `missing_argument`, `invalid_json`, `corrupt_batch` |
//! | 5 | unsupported | `unsupported_format`, `unsupported_capability`, `unsupported_element`, `unsupported_type`, `unsupported_command` |
//! | 6 | I/O | `file_not_found`, `file_locked`, `file_exists`, `file_required`, `io_error` |
//! | 7 | corrupt input | `corrupt_file`, `invalid_input` |
//! | 8 | protected | `document_protected` |
//! | 9 | partial: a `--best-effort` batch saved its successful items, and at least one item failed; or a `--unsupported warn` write saved the properties it could apply and skipped at least one | (per-item codes and warnings in the result) |
//!
//! A closed stdout pipe (`litchi get ... | head`) ends the run quietly.

#![forbid(unsafe_code)]

mod cli;
mod exit;
mod formats;
mod m4;
mod merge;
mod raw;
mod refresh;
mod render;
mod run;

use std::io::Write as _;
use std::process::ExitCode;

use clap::Parser as _;

/// Runs the `litchi` command line on the process arguments and returns its
/// exit status. The `litchi` binary is this function.
///
/// ```no_run
/// fn main() -> std::process::ExitCode {
///     litchi_cli::main()
/// }
/// ```
#[must_use]
pub fn main() -> ExitCode {
    let cli = match cli::Cli::try_parse() {
        Ok(cli) => cli,
        Err(error) => {
            // Help and version requests are successes; everything else is a
            // usage error (exit::USAGE).
            let code = if error.use_stderr() {
                exit::USAGE
            } else {
                exit::OK
            };
            let _ = error.print();
            return ExitCode::from(code);
        },
    };
    let json = cli.json();
    let mut stdout = std::io::stdout().lock();
    let status = match run::run(cli, &mut stdout) {
        Ok(status) => status,
        Err(_) if render::broken_pipe() => return ExitCode::from(exit::OK),
        Err(error) => {
            let _ = stdout.flush();
            render::error(&error, json);
            exit::of(error.code())
        },
    };
    if let Err(error) = stdout.flush() {
        if error.kind() != std::io::ErrorKind::BrokenPipe {
            render::error(&render::io(&error), json);
            return ExitCode::from(exit::of(litchi_ops::Code::IoError));
        }
    }
    ExitCode::from(status)
}
