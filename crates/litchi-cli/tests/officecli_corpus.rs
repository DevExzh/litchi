//! The OfficeCLI example-corpus harness (unified-ops record 0265).
//!
//! Replays every OfficeCLI example script vendored under
//! `test-data/officecli/examples/` (record 0257; provenance in
//! `test-data/officecli/PROVENANCE.md`) against the `litchi` binary, one
//! scratch directory per script, and records the outcome of every
//! `officecli` call.
//!
//! How a script is replayed:
//!
//! - bash runs a copy of the script whose first line (the shebang) is
//!   replaced by a `DEBUG` trap exporting the current line number, so line
//!   numbers stay those of the vendored file; `set -e` lines become `set +e`,
//!   so one refused line does not end the replay;
//! - `tests/officecli_corpus/officecli` comes first on `PATH` as `officecli`:
//!   it translates the call (resident `open`/`close` become no-ops, `batch`
//!   reads stdin, a bare slide gets a layout; record 0265 §3), passes
//!   `--unsupported warn` to `set`, `add` and `batch` as OfficeCLI applies
//!   the supported properties by default (record 0269), runs `litchi` and
//!   records the call; a call that skipped properties is `warned`, and one
//!   that still failed after skipping records its code with detail
//!   `<property>;skipped:<properties>` (record 0271);
//! - `python3` and `python` run the real interpreter with `-I -S`, so the
//!   scripts' helper snippets see only the standard library and every
//!   machine replays them the same way (no Pillow, no numpy);
//! - after the script, each file it created is read back with
//!   `view <file> text` and `get <file> / --depth 2` (line 0).
//!
//! The outcomes are compared with `tests/officecli_corpus/outcomes.tsv`. The
//! test fails when any call panics, dies by a signal, times out or returns
//! `internal_error`, or when a call that passed in the recorded outcomes no
//! longer passes, or a script family passes fewer calls. Other differences
//! (newly passing lines, changed refusal codes) are printed and do not fail
//! the test; re-record them with
//! `LITCHI_BLESS=1 cargo test -p litchi-cli --test officecli_corpus`, then
//! regenerate `docs/unified-ops/officecli-corpus-deltas.md` with
//! `python3 tools/officecli_corpus.py`.
//!
//! The harness needs `bash` and `python3`, and runs on Unix only.

#![cfg(unix)]

use std::collections::{BTreeMap, BTreeSet};
use std::fmt::Write as _;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::Mutex;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::{Duration, Instant};

const LITCHI: &str = env!("CARGO_BIN_EXE_litchi");

/// The longest one script may run before the replay counts as hung.
const SCRIPT_TIMEOUT: Duration = Duration::from_secs(900);

fn manifest_dir() -> &'static Path {
    Path::new(env!("CARGO_MANIFEST_DIR"))
}

fn examples_dir() -> PathBuf {
    manifest_dir().join("../../test-data/officecli/examples")
}

fn harness_dir() -> PathBuf {
    manifest_dir().join("tests/officecli_corpus")
}

/// One recorded `officecli` call.
#[derive(Clone, Debug, PartialEq, Eq)]
struct Call {
    line: u32,
    verb: String,
    format: String,
    kind: String,
    result: String,
    detail: String,
    target: String,
}

impl Call {
    fn passed(&self) -> bool {
        self.result == "ok"
    }

    /// A panic, a signal, a timeout or `internal_error`: never acceptable.
    fn fatal(&self) -> bool {
        self.result == "panic"
            || self.result == "internal_error"
            || self.result == "timeout"
            || self.result.starts_with("signal")
    }

    fn tsv(&self) -> String {
        [
            self.line.to_string().as_str(),
            &self.verb,
            &self.format,
            &self.kind,
            &self.result,
            &self.detail,
            &self.target,
        ]
        .join("\t")
    }
}

/// Replaces tabs and newlines, bounds the length to 48 characters and spells
/// empty as `-`.
fn field(text: &str) -> String {
    field_within(text, 48)
}

/// [`field`] with a caller's bound: a skipped-property list is recorded in
/// full (record 0339), so the corpus report can classify every name.
fn field_within(text: &str, bound: usize) -> String {
    let mut clean: String = text
        .chars()
        .map(|c| if c.is_control() { ' ' } else { c })
        .take(bound)
        .collect();
    clean = clean.trim().to_owned();
    if clean.is_empty() {
        "-".to_owned()
    } else {
        clean
    }
}

/// Every vendored script, as a `/`-separated path relative to the examples
/// directory, sorted.
fn scripts() -> Vec<String> {
    fn walk(dir: &Path, prefix: &str, out: &mut Vec<String>) {
        let mut entries: Vec<_> = fs::read_dir(dir)
            .unwrap_or_else(|error| panic!("reading {}: {error}", dir.display()))
            .map(|entry| entry.expect("directory entry"))
            .collect();
        entries.sort_by_key(fs::DirEntry::file_name);
        for entry in entries {
            let name = entry.file_name().to_string_lossy().into_owned();
            let relative = if prefix.is_empty() {
                name.clone()
            } else {
                format!("{prefix}/{name}")
            };
            let kind = entry.file_type().expect("file type");
            if kind.is_dir() {
                walk(&entry.path(), &relative, out);
            } else if name.ends_with(".sh") {
                out.push(relative);
            }
        }
    }
    let mut out = Vec::new();
    walk(&examples_dir(), "", &mut out);
    out.sort();
    out
}

/// The family a script is counted under: its directory, with the template
/// styles folded into `ppt/templates`.
fn family(script: &str) -> String {
    let directory = script.rsplit_once('/').map_or("", |(dir, _)| dir);
    if directory.starts_with("ppt/templates") {
        "ppt/templates".to_owned()
    } else {
        directory.to_owned()
    }
}

/// The script as replayed: same line numbers, a `DEBUG` trap on line 1 and
/// no `set -e`.
fn instrument(source: &str) -> String {
    const TRAP: &str = "trap 'export OFFICECLI_LINE=$LINENO' DEBUG";
    let mut out = String::with_capacity(source.len() + TRAP.len() + 4);
    for (index, line) in source.split_inclusive('\n').enumerate() {
        let (body, end) = line
            .strip_suffix('\n')
            .map_or((line, ""), |body| (body, "\n"));
        if index == 0 {
            if body.starts_with("#!") {
                out.push_str(TRAP);
            } else {
                out.push_str(TRAP);
                out.push_str("; ");
                out.push_str(body);
            }
        } else if sets_errexit(body) {
            out.push_str("set +e");
        } else {
            out.push_str(body);
        }
        out.push_str(end);
    }
    out
}

/// Whether the line is a bare `set` turning on `-e` (`set -e`,
/// `set -euo pipefail`, ...).
fn sets_errexit(line: &str) -> bool {
    let code = line.split('#').next().unwrap_or("");
    let mut words = code.split_whitespace();
    if words.next() != Some("set") {
        return false;
    }
    let mut errexit = false;
    for word in words {
        if word == "pipefail" || word == "-o" {
            continue;
        }
        let Some(flags) = word.strip_prefix('-') else {
            return false;
        };
        if !flags.chars().all(|c| c.is_ascii_alphabetic()) {
            return false;
        }
        errexit |= flags.contains('e');
    }
    errexit
}

fn find_on_path(program: &str) -> Option<PathBuf> {
    let path = std::env::var_os("PATH")?;
    std::env::split_paths(&path)
        .map(|dir| dir.join(program))
        .find(|candidate| candidate.is_file())
}

fn write_executable(path: &Path, text: &str) {
    use std::os::unix::fs::PermissionsExt as _;
    fs::write(path, text).unwrap_or_else(|error| panic!("writing {}: {error}", path.display()));
    fs::set_permissions(path, fs::Permissions::from_mode(0o755)).expect("chmod");
}

/// The shared `bin` directory: the `officecli` stand-in and the isolated
/// Python interpreters.
fn install_bin(scratch: &Path) -> PathBuf {
    let bin = scratch.join("bin");
    fs::create_dir_all(&bin).expect("bin directory");
    let shim = fs::read_to_string(harness_dir().join("officecli")).expect("the officecli shim");
    write_executable(&bin.join("officecli"), &shim);
    let python = find_on_path("python3")
        .expect("the corpus harness needs python3 on PATH (the example scripts call it)");
    let wrapper = format!(
        "#!/bin/sh\nexec '{}' -I -S \"$@\"\n",
        python.display().to_string().replace('\'', r"'\''")
    );
    write_executable(&bin.join("python3"), &wrapper);
    write_executable(&bin.join("python"), &wrapper);
    bin
}

struct Replay {
    calls: Vec<Call>,
    /// Set when the script itself did not finish in time.
    timed_out: bool,
}

/// Replays one script in `scratch/<n>` and returns its calls.
fn replay(script: &str, scratch: &Path, bin: &Path) -> Replay {
    let source_path = examples_dir().join(script);
    let source_dir = source_path.parent().expect("script directory");
    let work = scratch.join("work");
    let log = scratch.join("log");
    let tmp = scratch.join("tmp");
    let home = scratch.join("home");
    for dir in [&work, &log, &tmp, &home] {
        fs::create_dir_all(dir).expect("scratch directory");
    }
    // The script's sibling inputs (pictures, thumbnails), never other
    // scripts.
    for entry in fs::read_dir(source_dir).expect("script directory") {
        let entry = entry.expect("directory entry");
        let name = entry.file_name();
        if entry.file_type().expect("file type").is_file()
            && !name.to_string_lossy().ends_with(".sh")
        {
            fs::copy(entry.path(), work.join(&name)).expect("copy script input");
        }
    }
    let name = source_path.file_name().expect("script name");
    let source = fs::read_to_string(&source_path).expect("script text");
    fs::write(work.join(name), instrument(&source)).expect("instrumented script");

    let path = std::env::var_os("PATH").unwrap_or_default();
    let mut dirs = vec![bin.to_path_buf()];
    dirs.extend(std::env::split_paths(&path));
    let path = std::env::join_paths(dirs).expect("PATH");
    let environment = |command: &mut Command| {
        command
            .current_dir(&work)
            .env("PATH", &path)
            .env("TMPDIR", &tmp)
            .env("HOME", &home)
            .env("LITCHI_CORPUS_BIN", LITCHI)
            .env("LITCHI_CORPUS_LOG", &log)
            .env_remove("OFFICECLI_LINE")
            .env_remove("BASH_ENV")
            .env_remove("ENV")
            .stdin(Stdio::null());
    };

    let mut command = Command::new("bash");
    command.arg(name);
    environment(&mut command);
    command
        .stdout(fs::File::create(scratch.join("stdout")).expect("stdout file"))
        .stderr(fs::File::create(scratch.join("stderr")).expect("stderr file"));
    let mut child = command.spawn().expect("bash runs");
    let deadline = Instant::now() + SCRIPT_TIMEOUT;
    let timed_out = loop {
        if child.try_wait().expect("waiting for bash").is_some() {
            break false;
        }
        if Instant::now() >= deadline {
            let _ = child.kill();
            let _ = child.wait();
            break true;
        }
        std::thread::sleep(Duration::from_millis(20));
    };

    // Read back every file the script created.
    let created: BTreeSet<PathBuf> = records(&log)
        .into_iter()
        .filter(|record| record.argv.first().map(String::as_str) == Some("create"))
        .filter(|record| record.status == "0")
        .filter_map(|record| record.argv.get(1).map(|file| work.join(file)))
        .filter(|file| file.is_file())
        .collect();
    for file in &created {
        for args in [
            vec!["view", "", "text"],
            vec!["get", "", "/", "--depth", "2"],
        ] {
            let mut command = Command::new(bin.join("officecli"));
            for arg in args {
                if arg.is_empty() {
                    command.arg(file);
                } else {
                    command.arg(arg);
                }
            }
            environment(&mut command);
            command
                .env("OFFICECLI_LINE", "0")
                .stdout(Stdio::null())
                .stderr(Stdio::null());
            let _ = command.status().expect("the officecli shim runs");
        }
    }

    let calls = records(&log).iter().map(Record::call).collect();
    Replay { calls, timed_out }
}

/// One call as the shim recorded it.
struct Record {
    line: u32,
    argv: Vec<String>,
    status: String,
    stderr: String,
    translated: Option<String>,
}

fn records(log: &Path) -> Vec<Record> {
    let mut names: Vec<PathBuf> = match fs::read_dir(log) {
        Ok(entries) => entries
            .map(|entry| entry.expect("log entry").path())
            .filter(|path| path.is_dir())
            .collect(),
        Err(_) => return Vec::new(),
    };
    names.sort();
    names
        .into_iter()
        .map(|dir| {
            let read = |name: &str| {
                fs::read(dir.join(name))
                    .map(|bytes| String::from_utf8_lossy(&bytes).into_owned())
                    .ok()
            };
            Record {
                line: read("line")
                    .and_then(|line| line.trim().parse().ok())
                    .unwrap_or(0),
                argv: read("argv")
                    .unwrap_or_default()
                    .split_terminator('\0')
                    .map(str::to_owned)
                    .collect(),
                // A missing status means litchi never returned.
                status: read("status").unwrap_or_else(|| "lost".to_owned()),
                stderr: read("stderr").unwrap_or_default(),
                translated: read("translated"),
            }
        })
        .collect()
}

/// The text between the first `start` and the next quote after it.
fn quoted_after<'a>(text: &'a str, start: &str) -> Option<&'a str> {
    let rest = &text[text.find(start)? + start.len()..];
    Some(&rest[..rest.find('\'')?])
}

/// `"key": "value"` in a JSON error object, or `key: value` in the text one.
fn error_field(stderr: &str, key: &str) -> Option<String> {
    if stderr.trim_start().starts_with('{') {
        let value: serde_json::Value = serde_json::from_str(stderr.trim()).ok()?;
        return value.get(key)?.as_str().map(str::to_owned);
    }
    // The error's own lines follow its `error[...]` line; `warning[...]`
    // blocks of skipped properties may precede it (record 0271).
    let prefix = format!("  {key}: ");
    let error = stderr
        .find("error[")
        .map_or(stderr, |start| &stderr[start..]);
    error
        .lines()
        .skip(1)
        .take_while(|line| line.starts_with("  "))
        .find_map(|line| line.strip_prefix(&prefix))
        .map(str::to_owned)
}

/// Whether stderr holds only `warning[...]` reports: a `--unsupported warn`
/// write that applied and skipped properties.
fn warned(stderr: &str) -> bool {
    stderr.contains("warning[") && !stderr.contains("error[")
}

/// The properties a `--unsupported warn` write skipped, in report order,
/// each once.
fn skipped(stderr: &str) -> Vec<String> {
    let mut properties: Vec<String> = Vec::new();
    let mut in_warning = false;
    for line in stderr.lines() {
        if line.starts_with("warning[") {
            in_warning = true;
        } else if !line.starts_with("  ") {
            in_warning = false;
        } else if let Some(property) = line.strip_prefix("  property: ") {
            if in_warning && !properties.iter().any(|known| known == property) {
                properties.push(property.to_owned());
            }
        }
    }
    properties
}

/// The skipped properties, in full (record 0339): comma-separated, in
/// report order, with each run of consecutive names that share a first
/// dotted segment written once as `<head>.{<rest>,...}`
/// (`title.font,title.size,fill` is `title.{font,size},fill`).
fn skipped_detail(properties: &[String]) -> String {
    let mut detail = String::new();
    let mut index = 0;
    while index < properties.len() {
        let property = &properties[index];
        let head = property.split_once('.').map(|(head, _)| head);
        let run = head.map_or(1, |head| {
            properties[index..]
                .iter()
                .take_while(|name| name.split_once('.').is_some_and(|(other, _)| other == head))
                .count()
        });
        if !detail.is_empty() {
            detail.push(',');
        }
        match head {
            Some(head) if run > 1 => {
                detail.push_str(head);
                detail.push_str(".{");
                for (n, name) in properties[index..index + run].iter().enumerate() {
                    if n > 0 {
                        detail.push(',');
                    }
                    detail.push_str(&name[head.len() + 1..]);
                }
                detail.push('}');
            },
            _ => detail.push_str(property),
        }
        index += run;
    }
    detail
}

/// The detail of a call that failed after skipping properties: the error's
/// own property, `;skipped:` and the skipped properties in full.
fn failed_after_skipping_detail(property: &str, properties: &[String]) -> String {
    format!("{property};skipped:{}", skipped_detail(properties))
}

fn error_code(stderr: &str) -> Option<String> {
    if stderr.trim_start().starts_with('{') {
        return error_field(stderr, "code");
    }
    let rest = &stderr[stderr.find("error[")? + "error[".len()..];
    Some(rest[..rest.find(']')?].to_owned())
}

impl Record {
    fn call(&self) -> Call {
        let verb = self.argv.first().cloned().unwrap_or_default();
        // Every verb takes the document first; its extension names the
        // format.
        let format = self
            .argv
            .get(1)
            .and_then(|file| Path::new(file).extension())
            .map(|extension| extension.to_string_lossy().to_ascii_lowercase())
            .unwrap_or_default();
        let kind = self
            .argv
            .iter()
            .position(|arg| arg == "--type")
            .and_then(|index| self.argv.get(index + 1))
            .cloned()
            .or_else(|| {
                self.argv
                    .iter()
                    .find_map(|arg| arg.strip_prefix("--type=").map(str::to_owned))
            })
            .or_else(|| {
                // `add --from` copies an element instead of naming a type.
                self.argv
                    .iter()
                    .any(|arg| arg == "--from" || arg.starts_with("--from="))
                    .then(|| "(from)".to_owned())
            })
            .unwrap_or_default();
        // The element word of the addressed node: the last step of a path
        // (`/` for the root), or the leading word of a query selector.
        let target = match verb.as_str() {
            "set" | "get" | "remove" | "move" | "swap" => self
                .argv
                .get(2)
                .map(|path| {
                    if path == "/" {
                        return "/".to_owned();
                    }
                    let step = path.rsplit('/').next().unwrap_or("");
                    step.split('[').next().unwrap_or("").to_owned()
                })
                .unwrap_or_default(),
            "query" => self
                .argv
                .get(2)
                .map(|selector| {
                    selector
                        .chars()
                        .take_while(|c| c.is_ascii_alphanumeric())
                        .collect()
                })
                .unwrap_or_default(),
            _ => String::new(),
        };
        let status = self.status.trim();
        let (result, detail) = if status == "skip" {
            ("skip".to_owned(), String::new())
        } else if status == "0" {
            let detail = self
                .translated
                .as_deref()
                .map(|id| format!("translated:{}", id.trim()))
                .unwrap_or_default();
            ("ok".to_owned(), detail)
        } else if self.stderr.contains("panicked at") {
            ("panic".to_owned(), String::new())
        } else if status == "9" && warned(&self.stderr) {
            // `--unsupported warn` applied the call without the properties
            // it lists.
            ("warned".to_owned(), skipped_detail(&skipped(&self.stderr)))
        } else if let Some(code) = error_code(&self.stderr) {
            let property = error_field(&self.stderr, "property").unwrap_or_default();
            let skipped = skipped(&self.stderr);
            if skipped.is_empty() {
                (code, property)
            } else {
                // `--unsupported warn` skipped properties, then the rest of
                // the call failed (record 0271).
                (code, failed_after_skipping_detail(&property, &skipped))
            }
        } else if status == "2" {
            let token = quoted_after(&self.stderr, "unexpected argument '")
                .or_else(|| quoted_after(&self.stderr, "unrecognized subcommand '"))
                .map(str::to_owned)
                .or_else(|| {
                    self.stderr
                        .lines()
                        .find(|line| line.starts_with("error:"))
                        .map(|line| line.trim_start_matches("error:").trim().to_owned())
                })
                .unwrap_or_default();
            ("usage".to_owned(), token)
        } else {
            match status.parse::<i32>() {
                Ok(code) if code > 128 => (format!("signal{}", code - 128), String::new()),
                Ok(code) => (format!("exit{code}"), String::new()),
                Err(_) => ("lost".to_owned(), String::new()),
            }
        };
        // A skipped-property list is recorded in full (record 0339).
        let skipped_list = result == "warned" || detail.contains(";skipped:");
        let detail = if skipped_list {
            field_within(&detail, usize::MAX)
        } else {
            field(&detail)
        };
        Call {
            line: self.line,
            verb: field(&verb),
            format: field(&format),
            kind: field(&kind),
            result: field(&result),
            detail,
            target: field(&target),
        }
    }
}

const HEADER: &str = "\
# OfficeCLI example-corpus outcomes (unified-ops record 0265).
# Generated by `LITCHI_BLESS=1 cargo test -p litchi-cli --test officecli_corpus`;
# read by `python3 tools/officecli_corpus.py`. Do not edit by hand.
# `@ <script>` starts a script; each row is one `officecli` call in replay order:
# line (0 = harness read-back), verb, format (the document's extension), --type,
# result (ok, skip, usage, warned, error code),
# detail (property, the skipped properties of a warned call,
# `<property>;skipped:<properties>` for a call that failed after skipping,
# offending argument or translated:<delta>), last path step.
# Skipped properties are listed in full, in report order; a run of names
# sharing a first dotted segment is written `<head>.{<rest>,...}`.
";

fn render(outcomes: &BTreeMap<String, Vec<Call>>) -> String {
    let mut text = HEADER.to_owned();
    for (script, calls) in outcomes {
        let _ = writeln!(text, "@ {script}");
        for call in calls {
            text.push_str(&call.tsv());
            text.push('\n');
        }
    }
    text
}

fn parse(text: &str) -> BTreeMap<String, Vec<Call>> {
    let mut outcomes = BTreeMap::new();
    let mut current: Option<String> = None;
    for line in text.lines() {
        if line.starts_with('#') || line.is_empty() {
            continue;
        }
        if let Some(script) = line.strip_prefix("@ ") {
            current = Some(script.to_owned());
            outcomes.entry(script.to_owned()).or_insert_with(Vec::new);
            continue;
        }
        let fields: Vec<&str> = line.split('\t').collect();
        assert_eq!(fields.len(), 7, "malformed outcome row: {line}");
        let script = current.as_ref().expect("a row before any script");
        outcomes.get_mut(script).expect("script").push(Call {
            line: fields[0].parse().expect("line number"),
            verb: fields[1].to_owned(),
            format: fields[2].to_owned(),
            kind: fields[3].to_owned(),
            result: fields[4].to_owned(),
            detail: fields[5].to_owned(),
            target: fields[6].to_owned(),
        });
    }
    outcomes
}

/// Calls keyed by (line, verb, occurrence of that line and verb).
fn keyed(calls: &[Call]) -> BTreeMap<(u32, String, usize), &Call> {
    let mut seen: BTreeMap<(u32, &str), usize> = BTreeMap::new();
    calls
        .iter()
        .map(|call| {
            let occurrence = seen.entry((call.line, call.verb.as_str())).or_insert(0);
            *occurrence += 1;
            ((call.line, call.verb.clone(), *occurrence), call)
        })
        .collect()
}

#[test]
fn officecli_example_corpus() {
    let scripts = scripts();
    assert!(
        scripts.len() >= 90,
        "expected the vendored OfficeCLI example scripts, found {}",
        scripts.len()
    );
    let scratch = tempfile::tempdir().expect("scratch directory");
    let bin = install_bin(scratch.path());

    let next = AtomicUsize::new(0);
    let results: Mutex<BTreeMap<String, Replay>> = Mutex::new(BTreeMap::new());
    let workers = std::thread::available_parallelism()
        .map_or(2, usize::from)
        .clamp(1, 6);
    std::thread::scope(|scope| {
        for _ in 0..workers {
            scope.spawn(|| {
                loop {
                    let index = next.fetch_add(1, Ordering::Relaxed);
                    let Some(script) = scripts.get(index) else {
                        break;
                    };
                    let dir = scratch.path().join(format!("s{index:03}"));
                    let replay = replay(script, &dir, &bin);
                    // Free the scratch space as soon as the script is done.
                    let _ = fs::remove_dir_all(&dir);
                    results
                        .lock()
                        .expect("results lock")
                        .insert(script.clone(), replay);
                }
            });
        }
    });
    let results = results.into_inner().expect("results lock");

    let mut fatal = Vec::new();
    let mut outcomes = BTreeMap::new();
    for (script, replay) in results {
        if replay.timed_out {
            fatal.push(format!("{script}: the replay did not finish in time"));
        }
        for call in &replay.calls {
            if call.fatal() {
                fatal.push(format!(
                    "{script}:{} {} -> {} {}",
                    call.line, call.verb, call.result, call.detail
                ));
            }
        }
        outcomes.insert(script, replay.calls);
    }
    let fresh = render(&outcomes);
    let fresh_path = Path::new(env!("CARGO_TARGET_TMPDIR")).join("officecli-corpus-outcomes.tsv");
    let _ = fs::write(&fresh_path, &fresh);

    assert!(
        fatal.is_empty(),
        "calls that panicked, died, hung or returned internal_error:\n{}",
        fatal.join("\n")
    );

    let recorded_path = harness_dir().join("outcomes.tsv");
    if std::env::var_os("LITCHI_BLESS").is_some() {
        fs::write(&recorded_path, &fresh).expect("writing the recorded outcomes");
        return;
    }
    let recorded_text = fs::read_to_string(&recorded_path).unwrap_or_else(|error| {
        panic!(
            "reading {} ({error}); record it with LITCHI_BLESS=1",
            recorded_path.display()
        )
    });
    let recorded = parse(&recorded_text);

    let mut regressions = Vec::new();
    let mut drift = Vec::new();
    let mut family_passed: BTreeMap<String, (usize, usize)> = BTreeMap::new();
    for script in recorded
        .keys()
        .chain(outcomes.keys())
        .collect::<BTreeSet<_>>()
    {
        let before = recorded.get(script).map(Vec::as_slice).unwrap_or_default();
        let after = outcomes.get(script).map(Vec::as_slice).unwrap_or_default();
        let counts = family_passed.entry(family(script)).or_default();
        counts.0 += before.iter().filter(|call| call.passed()).count();
        counts.1 += after.iter().filter(|call| call.passed()).count();
        let before_keyed = keyed(before);
        let after_keyed = keyed(after);
        for (key, old) in &before_keyed {
            match after_keyed.get(key) {
                Some(new) if new == old => {},
                Some(new) if old.passed() && !new.passed() => regressions.push(format!(
                    "{script}:{} {}: passed, now {} {}",
                    key.0, key.1, new.result, new.detail
                )),
                Some(new) => drift.push(format!(
                    "{script}:{} {}: {} {} -> {} {}",
                    key.0, key.1, old.result, old.detail, new.result, new.detail
                )),
                None if old.passed() => regressions.push(format!(
                    "{script}:{} {}: passed, now not replayed",
                    key.0, key.1
                )),
                None => drift.push(format!("{script}:{} {}: no longer replayed", key.0, key.1)),
            }
        }
        for key in after_keyed.keys() {
            if !before_keyed.contains_key(key) {
                drift.push(format!("{script}:{} {}: newly replayed", key.0, key.1));
            }
        }
    }
    for (family, (before, after)) in &family_passed {
        if after < before {
            regressions.push(format!(
                "family {family}: {before} calls passed, now {after}"
            ));
        }
    }
    if !drift.is_empty() {
        eprintln!(
            "{} corpus outcomes differ from {} without regressing; re-record with \
             LITCHI_BLESS=1 and regenerate the delta table (python3 tools/officecli_corpus.py):",
            drift.len(),
            recorded_path.display()
        );
        for line in drift.iter().take(50) {
            eprintln!("  {line}");
        }
    }
    assert!(
        regressions.is_empty(),
        "OfficeCLI corpus regressions (fresh outcomes in {}):\n{}",
        fresh_path.display(),
        regressions.join("\n")
    );
}

#[test]
fn instrumentation_keeps_line_numbers() {
    let source = "#!/bin/bash\nset -euo pipefail\necho hi\n  set -e # strict\nset -x\n";
    let replayed = instrument(source);
    assert_eq!(source.lines().count(), replayed.lines().count());
    let lines: Vec<&str> = replayed.lines().collect();
    assert!(lines[0].starts_with("trap "));
    assert_eq!(lines[1], "set +e");
    assert_eq!(lines[3], "set +e");
    assert_eq!(lines[4], "set -x");
    assert!(
        instrument("echo x\n").starts_with("trap 'export OFFICECLI_LINE=$LINENO' DEBUG; echo x")
    );
}

#[test]
fn warned_calls_list_their_skipped_properties() {
    let stderr = "warning[unsupported_property]: the element has no such property\n  \
                  property: fill\n  suggestion: did you mean 'x'?\n\
                  warning[unsupported_property]: m\n  property: geometry\nset: ok\n";
    assert!(warned(stderr));
    assert_eq!(skipped(stderr), ["fill", "geometry"]);
    assert!(!warned("warning[x]: y\nerror[invalid_value]: z\n"));
    // Every name is kept (record 0339); runs sharing a head are grouped.
    let long: Vec<String> = (0..20).map(|n| format!("property{n}")).collect();
    let detail = skipped_detail(&long);
    assert_eq!(detail.split(',').count(), 20, "{detail}");
    assert!(detail.ends_with(",property19"), "{detail}");
    let names: Vec<String> = [
        "title.font",
        "title.size",
        "fill",
        "title.bold",
        "plotArea.x",
    ]
    .map(str::to_owned)
    .to_vec();
    assert_eq!(
        skipped_detail(&names),
        "title.{font,size},fill,title.bold,plotArea.x"
    );
    assert_eq!(
        skipped_detail(&["a.b.c".to_owned(), "a.d".to_owned()]),
        "a.{b.c,d}"
    );
}

#[test]
fn calls_that_failed_after_skipping_name_the_skipped_properties() {
    let stderr = "warning[unsupported_property]: the element has no such property\n  \
                  item: 0\n  property: series1.name\n\
                  warning[unsupported_property]: m\n  property: series1.values\n\
                  error[missing_property]: a chart needs data\n  item: 0\n  \
                  property: data\n  suggestion: pass data=Name:1,2,3\n";
    assert!(!warned(stderr));
    assert_eq!(error_code(stderr).as_deref(), Some("missing_property"));
    // The error's own property, not the first warning's.
    assert_eq!(error_field(stderr, "property").as_deref(), Some("data"));
    assert_eq!(skipped(stderr), ["series1.name", "series1.values"]);
    assert_eq!(
        failed_after_skipping_detail("data", &skipped(stderr)),
        "data;skipped:series1.{name,values}"
    );
    let long: Vec<String> = (0..20).map(|n| format!("series{n}.color")).collect();
    let detail = failed_after_skipping_detail("data", &long);
    assert!(
        detail.starts_with("data;skipped:series0.color,series1.color,"),
        "{detail}"
    );
    assert!(detail.ends_with(",series19.color"), "{detail}");
}

#[test]
fn outcomes_round_trip() {
    let call = Call {
        line: 7,
        verb: "add".to_owned(),
        format: "docx".to_owned(),
        kind: "paragraph".to_owned(),
        result: "unsupported_property".to_owned(),
        detail: "font.eastAsia".to_owned(),
        target: "-".to_owned(),
    };
    let mut outcomes = BTreeMap::new();
    outcomes.insert("word/x.sh".to_owned(), vec![call]);
    assert_eq!(parse(&render(&outcomes)), outcomes);
}
