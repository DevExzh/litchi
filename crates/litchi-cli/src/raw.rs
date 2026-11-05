//! `dump`, `raw`, `raw-set`, `add-part` and `validate` (record 0336): one
//! `litchi-ops` call each.

use std::io::Write;
use std::path::Path;

use litchi_ops::raw::{Action, NewPart};
use litchi_ops::{Code, Error, File, Overwrite, Result, SaveOptions, SaveTo};

use crate::{cli, exit, render};

pub(crate) fn dump(args: &cli::Dump, json: bool, out: &mut impl Write) -> Result<u8> {
    if !args.format.eq_ignore_ascii_case("batch") {
        return Err(Error::new(Code::InvalidArgument, "unknown dump format")
            .with_suggestion("the only dump format is batch"));
    }
    if args.path != "/" {
        return Err(
            Error::new(Code::UnsupportedPath, "dump covers the whole document only")
                .with_suggestion(
                    "dump `/` (the default); subtree dumps are a record 0336 follow-up",
                ),
        );
    }
    let file = litchi_ops::open(&args.file)?;
    let dump = file.dump()?;
    // Warnings go to stderr, so a dump piped into `batch --input -` stays
    // a JSON array.
    render::warnings(dump.warnings());
    let mut text = serde_json::to_string(&dump)
        .map_err(|error| Error::new(Code::InternalError, format!("dump serialization: {error}")))?;
    text.push('\n');
    match args.out.as_deref().filter(|path| path.as_os_str() != "-") {
        Some(path) => {
            std::fs::write(path, text.as_bytes()).map_err(|error| render::io(&error))?;
            if json {
                render::json(
                    out,
                    &serde_json::json!({
                        "outputFile": path.display().to_string(),
                        "itemCount": dump.items().len(),
                        "warnings": dump.warnings().len(),
                    }),
                )?;
            } else {
                writeln!(out, "{}", path.display()).map_err(|error| render::io(&error))?;
            }
        },
        None => out
            .write_all(text.as_bytes())
            .map_err(|error| render::io(&error))?,
    }
    Ok(exit::OK)
}

pub(crate) fn raw(args: &cli::Raw, json: bool, out: &mut impl Write) -> Result<u8> {
    let file = litchi_ops::open(&args.file)?;
    let Some(part) = &args.part else {
        let parts = file.parts()?;
        if json {
            render::json(out, &parts)?;
        } else {
            for part in &parts {
                writeln!(out, "{}\t{}", part.name(), part.content_type())
                    .map_err(|error| render::io(&error))?;
            }
        }
        return Ok(exit::OK);
    };
    let part = file.raw(part)?;
    if let Some(path) = &args.out {
        std::fs::write(path, part.bytes()).map_err(|error| render::io(&error))?;
        if json {
            render::json(
                out,
                &serde_json::json!({
                    "part": part.name(),
                    "contentType": part.content_type(),
                    "bytes": part.bytes().len(),
                    "outputFile": path.display().to_string(),
                }),
            )?;
        } else {
            writeln!(out, "{}", path.display()).map_err(|error| render::io(&error))?;
        }
        return Ok(exit::OK);
    }
    let text = part.text().map_err(|error| {
        error.with_suggestion("the part is binary; write it to a file with --out")
    })?;
    if json {
        render::json(
            out,
            &serde_json::json!({
                "part": part.name(),
                "contentType": part.content_type(),
                "xml": text,
            }),
        )?;
    } else {
        out.write_all(text.as_bytes())
            .map_err(|error| render::io(&error))?;
        if !text.ends_with('\n') {
            writeln!(out).map_err(|error| render::io(&error))?;
        }
    }
    Ok(exit::OK)
}

pub(crate) fn raw_set(args: &cli::RawSet, json: bool, out: &mut impl Write) -> Result<u8> {
    let action: Action = args.action.parse()?;
    let file = litchi_ops::open(&args.file)?;
    let edited = file.raw_set(&args.part, &args.xpath, action, args.xml.as_deref())?;
    let saved = save(&edited.file, &args.file, &args.output)?;
    if json {
        render::json(
            out,
            &serde_json::json!({
                "action": action.as_str(),
                "matched": edited.matched,
                "saved": saved,
            }),
        )?;
    } else {
        writeln!(
            out,
            "raw-set applied: {} at {} ({} matched)",
            action.as_str(),
            args.xpath,
            edited.matched
        )
        .map_err(|error| render::io(&error))?;
    }
    Ok(exit::OK)
}

pub(crate) fn add_part(args: &cli::AddPart, json: bool, out: &mut impl Write) -> Result<u8> {
    let part = if args.kind.eq_ignore_ascii_case("custom") {
        let missing = |what: &'static str| {
            Error::new(Code::MissingArgument, what).with_suggestion(
                "a custom part needs --name, --content-type, --relationship and --data",
            )
        };
        let data = args
            .data
            .as_ref()
            .ok_or_else(|| missing("a custom part needs --data"))?;
        let bytes = std::fs::read(data).map_err(|error| {
            let code = if error.kind() == std::io::ErrorKind::NotFound {
                Code::FileNotFound
            } else {
                Code::IoError
            };
            Error::new(code, format!("reading the part data failed: {error}"))
        })?;
        NewPart::new(
            args.content_type
                .clone()
                .ok_or_else(|| missing("a custom part needs --content-type"))?,
            args.relationship
                .clone()
                .ok_or_else(|| missing("a custom part needs --relationship"))?,
            bytes,
        )
        .named(
            args.name
                .clone()
                .ok_or_else(|| missing("a custom part needs --name"))?,
        )
    } else {
        let part = NewPart::of_type(&args.kind)?;
        match &args.name {
            Some(name) => part.named(name.clone()),
            None => part,
        }
    };
    let file = litchi_ops::open(&args.file)?;
    let added = file.add_part(&args.parent, part)?;
    let saved = save(&added.file, &args.file, &args.output)?;
    if json {
        render::json(
            out,
            &serde_json::json!({
                "part": added.part,
                "relationshipId": added.relationship_id,
                "saved": saved,
            }),
        )?;
    } else {
        // OfficeCLI's line, which scripts grep for `relId=`.
        writeln!(
            out,
            "Created {} part: relId={} path={}",
            args.kind, added.relationship_id, added.part
        )
        .map_err(|error| render::io(&error))?;
    }
    Ok(exit::OK)
}

pub(crate) fn validate(args: &cli::Validate, json: bool, out: &mut impl Write) -> Result<u8> {
    let file = litchi_ops::open(&args.file)?;
    let validation = file.validate()?;
    if json {
        render::json(out, &validation)?;
    } else if validation.is_valid() {
        writeln!(out, "Validation passed: no errors found.").map_err(|error| render::io(&error))?;
    } else {
        // As OfficeCLI: the failure summary on stderr.
        let errors: Vec<_> = validation.errors().collect();
        let mut stderr = std::io::stderr().lock();
        let _ = writeln!(stderr, "Found {} validation error(s):", errors.len());
        for issue in errors {
            let _ = writeln!(stderr, "  [{}] {}", issue.code(), issue.message());
            for location in issue.locations() {
                if let Some(part) = location.part() {
                    let _ = writeln!(stderr, "    Part: {part}");
                }
                if let Some(path) = location.path() {
                    let _ = writeln!(stderr, "    Path: {path}");
                }
                if let Some(offset) = location.byte_offset() {
                    let _ = writeln!(stderr, "    Offset: {offset}");
                }
            }
        }
    }
    Ok(if validation.is_valid() {
        exit::OK
    } else {
        exit::OTHER
    })
}

/// Saves a raw write's new handle to `--out`, or over the input file. The
/// new handle was opened from published bytes, so it has no origin path of
/// its own.
fn save(file: &File, input: &Path, output: &cli::Output) -> Result<String> {
    let destination = output.out.as_deref().unwrap_or(input);
    let saved = file.save_with(
        SaveTo::Path(destination),
        &SaveOptions::new().overwrite(Overwrite::Replace),
    )?;
    Ok(saved.path.display().to_string())
}
