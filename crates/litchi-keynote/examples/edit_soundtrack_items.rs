//! Edit the semantic audio collection of an existing Keynote presentation.
//!
//! This example deliberately uses only the focused `litchi-keynote` surface:
//! item positions are semantic [`litchi_core::Position`] values and the input
//! is validated by [`litchi_keynote::soundtrack::items::AudioSource`].  Native
//! data identifiers, PackageMetadata records, IWA components, and raw wire
//! messages are not part of the command-line contract.
//!
//! ```text
//! edit_soundtrack_items <input.key> <output.key> <operation> [arguments]
//!
//! operations:
//!   list
//!   add <filename> <audio>
//!   insert <position> <filename> <audio>
//!   replace <position> <filename> <audio>
//!   remove <position>
//!
//! optional:
//!   --inverse <path>
//! ```

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports semantic item diagnostics"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::fs;
use std::path::PathBuf;

use litchi_keynote::soundtrack::items::AudioSource;
use litchi_keynote::{Package, Position};

const USAGE: &str = "usage: edit_soundtrack_items <input.key> <output.key> \
                     <list|add|insert|replace|remove> [position] \
                     [filename audio] [--inverse PATH]";

#[derive(Debug)]
enum Operation {
    List,
    Add {
        filename: String,
        audio: PathBuf,
    },
    Insert {
        position: Position,
        filename: String,
        audio: PathBuf,
    },
    Replace {
        position: Position,
        filename: String,
        audio: PathBuf,
    },
    Remove {
        position: Position,
    },
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let operation_name = required_text(&mut arguments, "missing operation")?;
    let operation = parse_operation(&operation_name, &mut arguments)?;
    let inverse = parse_inverse(&mut arguments)?;

    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }
    if inverse
        .as_deref()
        .is_some_and(|path| path == input || path == output)
    {
        return Err(invalid_input(
            "inverse path must differ from input and output paths",
        ));
    }

    let package = Package::open(&input)?;
    if matches!(&operation, Operation::List) {
        print_items(&package)?;
        ensure_no_trailing_arguments(&mut arguments)?;
        return Ok(());
    }

    let mut edit = package.edit_soundtrack_items()?;
    stage(&mut edit, operation)?;
    let commit = edit.commit()?;

    if let Some(path) = inverse {
        let restored = commit
            .package()
            .apply_soundtrack_items(&commit.patch().inverse())?;
        restored.package().save(path)?;
    }
    commit.package().save(output)?;

    println!(
        "soundtrack items: changed={}, touched_components={}, full_reparse_performed={}",
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().full_reparse_performed(),
    );
    print_items(commit.package())?;
    Ok(())
}

fn stage(
    edit: &mut litchi_keynote::soundtrack::items::Edit<'_>,
    operation: Operation,
) -> Result<(), Box<dyn Error>> {
    match operation {
        Operation::List => unreachable!("list is handled before staging"),
        Operation::Add { filename, audio } => {
            edit.add(AudioSource::new(filename, fs::read(audio)?)?)?;
        },
        Operation::Insert {
            position,
            filename,
            audio,
        } => {
            edit.insert(position, AudioSource::new(filename, fs::read(audio)?)?)?;
        },
        Operation::Replace {
            position,
            filename,
            audio,
        } => {
            edit.replace(position, AudioSource::new(filename, fs::read(audio)?)?)?;
        },
        Operation::Remove { position } => {
            edit.remove(position)?;
        },
    }
    Ok(())
}

fn print_items(package: &Package) -> Result<(), Box<dyn Error>> {
    match package.soundtrack_items()? {
        None => println!("soundtrack: absent"),
        Some(items) => {
            println!("soundtrack items: {}", items.len());
            for item in items.iter() {
                println!(
                    "  position={} filename={:?} bytes={}",
                    item.position().get(),
                    item.filename(),
                    item.byte_length(),
                );
            }
        },
    }
    Ok(())
}

fn parse_operation(
    name: &str,
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    match name {
        "list" => Ok(Operation::List),
        "add" => Ok(Operation::Add {
            filename: required_text(arguments, "missing semantic filename")?,
            audio: PathBuf::from(required_argument(arguments, "missing audio path")?),
        }),
        "insert" => Ok(Operation::Insert {
            position: parse_position(required_text(arguments, "missing position")?)?,
            filename: required_text(arguments, "missing semantic filename")?,
            audio: PathBuf::from(required_argument(arguments, "missing audio path")?),
        }),
        "replace" => Ok(Operation::Replace {
            position: parse_position(required_text(arguments, "missing position")?)?,
            filename: required_text(arguments, "missing semantic filename")?,
            audio: PathBuf::from(required_argument(arguments, "missing audio path")?),
        }),
        "remove" => Ok(Operation::Remove {
            position: parse_position(required_text(arguments, "missing position")?)?,
        }),
        _ => Err(invalid_input(
            "operation must be list, add, insert, replace, or remove",
        )),
    }
}

fn parse_inverse(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Option<PathBuf>, Box<dyn Error>> {
    let Some(flag) = arguments.next() else {
        return Ok(None);
    };
    if flag != OsStr::new("--inverse") {
        return Err(invalid_input(
            "unexpected trailing argument; expected --inverse PATH",
        ));
    }
    Ok(Some(PathBuf::from(required_argument(
        arguments,
        "missing inverse path",
    )?)))
}

fn ensure_no_trailing_arguments(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<(), Box<dyn Error>> {
    if arguments.next().is_some() {
        return Err(invalid_input("list does not accept trailing arguments"));
    }
    Ok(())
}

fn parse_position(value: String) -> Result<Position, Box<dyn Error>> {
    let index = value
        .parse::<usize>()
        .map_err(|_| invalid_input("position must be a non-negative integer"))?;
    Ok(Position::new(index))
}

fn required_argument(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<OsString, Box<dyn Error>> {
    arguments.next().ok_or_else(|| invalid_input(message))
}

fn required_text(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<String, Box<dyn Error>> {
    required_argument(arguments, message)?
        .into_string()
        .map_err(|_| invalid_input("arguments must be valid UTF-8"))
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(std::io::Error::new(
        std::io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
