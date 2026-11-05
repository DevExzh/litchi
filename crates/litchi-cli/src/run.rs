//! One function per verb, each a direct call into `litchi-ops`.

use std::io::{Read as _, Write};
use std::path::{Path, PathBuf};

use litchi_ops::batch::{BatchItem, ItemResult, Options, SourceLimits, Unsupported};
use litchi_ops::capabilities::Capabilities;
use litchi_ops::node::ViewOptions;
use litchi_ops::view::Mode;
use litchi_ops::{BatchMode, Code, Error, File, Overwrite, Props, Result, SaveOptions, SaveTo};

use crate::cli::{self, Cli, Command};
use crate::render::{self, ItemJson, Report, Summary};
use crate::{exit, formats};

/// Runs the parsed command line, writing its result to `out`, and returns
/// the exit status of a run that did not fail outright.
pub(crate) fn run(cli: Cli, out: &mut impl Write) -> Result<u8> {
    let json = cli.json();
    match cli.command {
        Command::Create(args) => create(&args, json, out),
        Command::Get(args) => get(&args, json, out),
        Command::Query(args) => query(&args, json, out),
        Command::View(args) => view(&args, json, out),
        Command::Capabilities(args) => capabilities(&args, json, out),
        Command::Schema(args) => schema(&args, json, out),
        Command::Set(args) => {
            // OfficeCLI's `--find T` and `--replace T` are `--prop find=T`
            // and `--prop replace=T`; they lead the assignments.
            let assignments: Vec<String> = [("find", &args.find), ("replace", &args.replace)]
                .into_iter()
                .filter_map(|(key, value)| value.as_ref().map(|value| format!("{key}={value}")))
                .chain(args.props.iter().cloned())
                .collect();
            let mut item = BatchItem::new("set");
            item.path = Some(args.path.as_str().into());
            item.props = Some(props(&assignments)?);
            let options = options(&args.policy);
            write(&args.file, item, &options, &args.output, "set", json, out)
        },
        Command::Add(args) => {
            let mut item = BatchItem::new("add");
            item.parent = Some(args.parent.as_str().into());
            item.kind = args.kind.as_deref().map(Into::into);
            item.from = args.from.as_deref().map(Into::into);
            if !args.props.is_empty() {
                item.props = Some(props(&args.props)?);
            }
            position(&mut item, &args.position);
            let options = options(&args.policy);
            write(&args.file, item, &options, &args.output, "add", json, out)
        },
        Command::Remove(args) => {
            let mut item = BatchItem::new("remove");
            item.path = Some(args.path.as_str().into());
            if !args.props.is_empty() {
                item.props = Some(props(&args.props)?);
            }
            write(
                &args.file,
                item,
                &Options::new(),
                &args.output,
                "remove",
                json,
                out,
            )
        },
        Command::Move(args) => {
            let mut item = BatchItem::new("move");
            item.path = Some(args.path.as_str().into());
            item.to = args.to.as_deref().map(Into::into);
            position(&mut item, &args.position);
            write(
                &args.file,
                item,
                &Options::new(),
                &args.output,
                "move",
                json,
                out,
            )
        },
        Command::Swap(args) => {
            let mut item = BatchItem::new("swap");
            item.path = Some(args.path1.as_str().into());
            item.path2 = Some(args.path2.as_str().into());
            write(
                &args.file,
                item,
                &Options::new(),
                &args.output,
                "swap",
                json,
                out,
            )
        },
        Command::Batch(args) => batch(&args, json, out),
        Command::Import(args) => crate::m4::import(&args, json, out),
        Command::Evaluate(args) => crate::m4::evaluate(&args, json, out),
        Command::Recalculate(args) => crate::m4::recalculate(&args, json, out),
        Command::Merge(args) => crate::merge::merge(&args, json, out),
        Command::Refresh(args) => crate::refresh::refresh(&args, json, out),
        Command::Dump(args) => crate::raw::dump(&args, json, out),
        Command::Raw(args) => crate::raw::raw(&args, json, out),
        Command::RawSet(args) => crate::raw::raw_set(&args, json, out),
        Command::AddPart(args) => crate::raw::add_part(&args, json, out),
        Command::Validate(args) => crate::raw::validate(&args, json, out),
    }
}

/// `--prop key=value` assignments, parsed by `litchi-ops`.
fn props(assignments: &[String]) -> Result<Props<'_>> {
    Props::from_props(assignments)
}

/// The command-channel options a `--unsupported` policy selects.
fn options(policy: &cli::Policy) -> Options {
    Options::new().unsupported(match policy.unsupported {
        cli::UnsupportedArg::Refuse => Unsupported::Refuse,
        cli::UnsupportedArg::Warn => Unsupported::Warn,
    })
}

fn position<'a>(item: &mut BatchItem<'a>, position: &'a cli::Position) {
    item.index = position.index;
    item.after = position.after.as_deref().map(Into::into);
    item.before = position.before.as_deref().map(Into::into);
}

fn create(args: &cli::Create, json: bool, out: &mut impl Write) -> Result<u8> {
    let format = match &args.kind {
        Some(kind) => formats::parse(kind)?,
        None => formats::of_path(&args.file).ok_or_else(|| {
            Error::new(
                Code::MissingArgument,
                "the document type is neither given nor implied by the extension",
            )
            .with_suggestion("pass --type, for example --type docx")
        })?,
    };
    // As OfficeCLI does, a path without an extension gains the type's.
    let mut path = args.file.clone();
    if path.extension().is_none() {
        path.set_extension(formats::name(format));
    }
    let overwrite = if args.force {
        Overwrite::Replace
    } else {
        Overwrite::Refuse
    };
    let saved = litchi_ops::create(format)?.save_with(
        SaveTo::Path(&path),
        &SaveOptions::new().overwrite(overwrite),
    )?;
    if json {
        render::json(
            out,
            &serde_json::json!({
                "path": saved.path.display().to_string(),
                "type": formats::name(format),
                "bytes": saved.bytes,
            }),
        )?;
    } else {
        writeln!(out, "created {}", saved.path.display()).map_err(|error| render::io(&error))?;
    }
    Ok(exit::OK)
}

fn get(args: &cli::Get, json: bool, out: &mut impl Write) -> Result<u8> {
    let file = litchi_ops::open(&args.file)?;
    let node = file.get(args.path.as_str())?;
    let view = node.view_with(ViewOptions::default().depth(args.depth));
    if json {
        render::view_json(out, &view)?;
    } else {
        render::node_text(out, &view, 0)?;
    }
    Ok(exit::OK)
}

fn query(args: &cli::Query, json: bool, out: &mut impl Write) -> Result<u8> {
    let file = litchi_ops::open(&args.file)?;
    let mut matches = file.query(args.selector.as_str())?;
    if let Some(limit) = args.limit {
        matches = matches.limit(limit);
    }
    let options = ViewOptions::default().depth(args.depth);
    let nodes = matches.collect::<Result<Vec<_>>>()?;
    if json {
        render::nodes_json(out, &nodes, options)?;
    } else {
        for node in &nodes {
            render::node_text(out, &node.view_with(options), 0)?;
        }
        writeln!(out, "{} matches", nodes.len()).map_err(|error| render::io(&error))?;
    }
    Ok(exit::OK)
}

fn view(args: &cli::View, json: bool, out: &mut impl Write) -> Result<u8> {
    let name = args
        .positional_mode
        .as_deref()
        .or(args.mode.as_deref())
        .unwrap_or("text");
    let mode = Mode::parse(name).ok_or_else(|| {
        Error::new(Code::InvalidArgument, format!("unknown view mode `{name}`"))
            .with_suggestion("use text, annotated, outline, stats or issues")
    })?;
    let file = litchi_ops::open(&args.file)?;
    if json {
        render::json(out, &file.view_json(mode)?)?;
    } else {
        let text = file.view(mode)?;
        out.write_all(text.as_bytes())
            .map_err(|error| render::io(&error))?;
        if !text.is_empty() && !text.ends_with('\n') {
            writeln!(out).map_err(|error| render::io(&error))?;
        }
    }
    Ok(exit::OK)
}

fn capabilities(args: &cli::Capabilities, json: bool, out: &mut impl Write) -> Result<u8> {
    let capabilities = match (&args.file, &args.format) {
        (Some(path), _) => litchi_ops::open(path)?.capabilities(),
        (None, Some(format)) => Capabilities::of(formats::parse(format)?),
        (None, None) => {
            return Err(Error::new(
                Code::MissingArgument,
                "capabilities needs a file or --format",
            ));
        },
    };
    if json {
        render::json(out, &capabilities)?;
        return Ok(exit::OK);
    }
    let format = capabilities.format().map_or("unknown", formats::name);
    let mut text = format!(
        "format: {format}\nstate: {}\nverbs:\n",
        capabilities.state().as_str()
    );
    for verb in litchi_ops::capabilities::Verb::ALL {
        text.push_str(&support_line(verb.as_str(), capabilities.verb(verb)));
    }
    text.push_str("views:\n");
    for &mode in Mode::ALL {
        text.push_str(&support_line(mode.name(), capabilities.view(mode)));
    }
    text.push_str("elements:");
    for element in capabilities.elements() {
        text.push(' ');
        text.push_str(element.name());
    }
    text.push('\n');
    out.write_all(text.as_bytes())
        .map_err(|error| render::io(&error))?;
    Ok(exit::OK)
}

fn schema(args: &cli::Schema, json: bool, out: &mut impl Write) -> Result<u8> {
    let schema = litchi_ops::schema(formats::parse(&args.format)?);
    if json {
        render::json(out, &schema)?;
        return Ok(exit::OK);
    }
    let mut text = format!("format: {}\n", formats::name(schema.format()));
    if !schema.registered() {
        text.push_str("registered: no\n");
    }
    for element in schema.elements() {
        text.push_str(element.name());
        if let Some(gap) = element.gap() {
            text.push_str(&format!(" (gap: {})", gap.follow_up()));
        }
        text.push('\n');
        for property in element.properties() {
            let flags: String = [
                ('a', property.add()),
                ('s', property.set()),
                ('g', property.get()),
            ]
            .iter()
            .map(|&(letter, on)| if on { letter } else { '-' })
            .collect();
            text.push_str(&format!(
                "  {} {} {flags}",
                property.name(),
                property.value_type().schema_type()
            ));
            let aliases: Vec<_> = property.aliases().collect();
            if !aliases.is_empty() {
                text.push_str(&format!(" aliases={}", aliases.join(",")));
            }
            if let Some(gap) = property.write_gap() {
                text.push_str(&format!(" write-gap={}", gap.follow_up()));
            }
            text.push('\n');
        }
    }
    out.write_all(text.as_bytes())
        .map_err(|error| render::io(&error))?;
    Ok(exit::OK)
}

fn support_line(name: &str, support: litchi_ops::capabilities::Support) -> String {
    use litchi_ops::capabilities::Support;
    match support {
        Support::Supported => format!("  {name}: supported\n"),
        Support::Unsupported { code, reason, .. } => format!("  {name}: {code}: {reason}\n"),
        _ => format!("  {name}: unknown\n"),
    }
}

/// Publishes `file`: to `--out` when given, else in place when it changed.
fn publish(file: &File, changed: bool, output: &cli::Output) -> Result<Option<PathBuf>> {
    match &output.out {
        Some(path) => file.save_as(path).map(|saved| Some(saved.path)),
        None if changed => file.save().map(|saved| Some(saved.path)),
        None => Ok(None),
    }
}

/// A single-command write through the command channel (`File::apply`).
fn write(
    path: &Path,
    item: BatchItem<'_>,
    options: &Options,
    output: &cli::Output,
    verb: &str,
    json: bool,
    out: &mut impl Write,
) -> Result<u8> {
    let file = litchi_ops::open(path)?;
    let commit = file.apply_with(item, options)?;
    let warned = !commit.value().warnings.is_empty();
    if !json {
        render::warnings(&commit.value().warnings);
    }
    let changed = commit.diagnostics().changed();
    let saved = publish(commit.file(), changed, output)?;
    let report = Report {
        results: vec![ItemJson::applied(0, commit.value())],
        summary: Summary {
            total: 1,
            executed: 1,
            succeeded: 1,
            changed,
            saved: saved.map(|path| path.display().to_string()),
            ..Summary::default()
        },
    };
    report.write(out, json, verb)?;
    Ok(if warned { exit::PARTIAL } else { exit::OK })
}

/// Reads the batch payload from `--commands`, `--input FILE` or stdin
/// (`--input -`), bounded by the decoder's input limit.
fn payload(args: &cli::Batch, limits: &SourceLimits) -> Result<Vec<u8>> {
    if let Some(commands) = &args.commands {
        return Ok(commands.as_bytes().to_vec());
    }
    let Some(input) = &args.input else {
        return Err(Error::new(
            Code::MissingArgument,
            "batch needs --input or --commands",
        ));
    };
    // One byte over the limit lets the decoder report `limit_exceeded`.
    let bound = u64::try_from(limits.max_input_bytes())
        .unwrap_or(u64::MAX)
        .saturating_add(1);
    let mut bytes = Vec::new();
    let result = if input.as_os_str() == "-" {
        std::io::stdin().lock().take(bound).read_to_end(&mut bytes)
    } else {
        std::fs::File::open(input).and_then(|file| file.take(bound).read_to_end(&mut bytes))
    };
    result.map_err(|error| {
        let code = if error.kind() == std::io::ErrorKind::NotFound {
            Code::FileNotFound
        } else {
            Code::IoError
        };
        Error::new(code, format!("reading the batch input failed: {error}"))
    })?;
    Ok(bytes)
}

fn batch(args: &cli::Batch, json: bool, out: &mut impl Write) -> Result<u8> {
    let limits = SourceLimits::default();
    let bytes = payload(args, &limits)?;
    let items = litchi_ops::batch::from_bytes(&bytes, &limits)?;
    let mode = if args.best_effort {
        BatchMode::BestEffort
    } else {
        BatchMode::Atomic
    };
    let options = options(&args.policy).mode(mode);
    let file = litchi_ops::open(&args.file)?;
    let total = items.len();
    let commit = match file.batch_with(&items, &options) {
        Ok(commit) => commit,
        Err(error) => {
            // An atomic failure discards every item; report which one
            // failed, then fail with its code.
            if let Some(index) = error.item() {
                let report = Report {
                    results: vec![ItemJson::failed(index, &error, items.get(index))],
                    summary: Summary {
                        total,
                        executed: index.saturating_add(1).min(total),
                        succeeded: index.min(total),
                        failed: 1,
                        skipped: total.saturating_sub(index.saturating_add(1)),
                        atomic_rolled_back: true,
                        ..Summary::default()
                    },
                };
                report.write(out, json, "batch")?;
            }
            return Err(error);
        },
    };
    let mut warned = false;
    for result in commit.value() {
        if let ItemResult::Applied(applied) = result {
            warned |= !applied.warnings.is_empty();
            if !json {
                render::warnings(&applied.warnings);
            }
        }
    }
    let changed = commit.diagnostics().changed();
    let saved = publish(commit.file(), changed, &args.output)?;
    let results: Vec<ItemJson<'_>> = commit
        .value()
        .iter()
        .enumerate()
        .map(|(index, result)| ItemJson::of(index, result, items.get(index)))
        .collect();
    let failed = results.iter().filter(|item| !item.success()).count();
    let report = Report {
        summary: Summary {
            total,
            executed: results.len(),
            succeeded: results.len() - failed,
            failed,
            skipped: total.saturating_sub(results.len()),
            changed,
            saved: saved.map(|path| path.display().to_string()),
            ..Summary::default()
        },
        results,
    };
    report.write(out, json, "batch")?;
    Ok(if failed == 0 && !warned {
        exit::OK
    } else {
        exit::PARTIAL
    })
}
