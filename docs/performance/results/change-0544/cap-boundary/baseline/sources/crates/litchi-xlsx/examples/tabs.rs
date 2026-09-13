use std::env;
use std::error::Error;
use std::io;

use litchi_xlsx::Workbook;

const USAGE: &str = "usage: tabs <input.xlsx> <output.xlsx> <sheet-name> <rename|show|hide|very-hide|activate|show-activate|before|after|to> [new-name|anchor|position]";

fn main() -> Result<(), Box<dyn Error>> {
    let mut args = env::args().skip(1);
    let input = args.next().ok_or(USAGE)?;
    let output = args.next().ok_or(USAGE)?;
    let name = args.next().ok_or(USAGE)?;
    let operation = args.next().ok_or(USAGE)?;
    let target = args.next();

    let source = Workbook::open(&input)?;
    let mut edit = source.edit()?;
    let mut reported_name = name.as_str();
    match operation.as_str() {
        "rename" => {
            let new_name = target.as_deref().ok_or(USAGE)?;
            edit.tab(name.as_str())?
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::NotFound,
                        format!("workbook has no sheet named '{name}'"),
                    )
                })?
                .rename(new_name)?;
            reported_name = new_name;
        },
        "before" | "after" => {
            let anchor = target.as_deref().ok_or(USAGE)?;
            let moved = if operation == "before" {
                edit.move_before(name.as_str(), anchor)?
            } else {
                edit.move_after(name.as_str(), anchor)?
            };
            moved.ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::NotFound,
                    format!("workbook has no tab matching '{name}' or '{anchor}'"),
                )
            })?;
        },
        "to" => {
            let position = target
                .as_deref()
                .ok_or(USAGE)?
                .parse::<usize>()
                .map_err(|_| {
                    io::Error::new(io::ErrorKind::InvalidInput, "position must be an integer")
                })?;
            edit.move_to(name.as_str(), position)?.ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::NotFound,
                    format!("workbook has no tab '{name}' or position {position}"),
                )
            })?;
        },
        "show" | "hide" | "very-hide" | "activate" | "show-activate" => {
            let mut tab = edit.tab(name.as_str())?.ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::NotFound,
                    format!("workbook has no sheet named '{name}'"),
                )
            })?;
            match operation.as_str() {
                "show" => {
                    tab.show();
                },
                "hide" => {
                    tab.hide();
                },
                "very-hide" => {
                    tab.very_hide();
                },
                "activate" => {
                    tab.activate();
                },
                "show-activate" => {
                    tab.show().activate();
                },
                other => {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        format!("unknown tab operation '{other}'"),
                    )
                    .into());
                },
            }
        },
        other => {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("unknown tab operation '{other}'"),
            )
            .into());
        },
    }

    let committed = edit.commit()?;
    let tab = committed
        .workbook()
        .sheet(reported_name)?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "edited tab disappeared"))?;
    committed.workbook().save(&output)?;
    println!(
        "saved {} semantic change(s); {} is {:?}, active={}, position={}",
        committed.patch().len(),
        tab.name(),
        tab.visibility(),
        tab.is_active(),
        tab.position()
    );
    Ok(())
}
