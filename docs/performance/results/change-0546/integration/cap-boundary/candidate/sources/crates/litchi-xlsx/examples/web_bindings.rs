use std::error::Error;
use std::ffi::OsString;
use std::io;

use litchi_ooxml_common::web::{
    AddIn, Binding as AddInBinding, Conformance, Pane, Panes, Reference, Store,
};
use litchi_xlsx::Workbook;
use litchi_xlsx::web::Binding as RangeBinding;

const INSTANCE_ID: &str = "litchi-office-roundtrip";
const APP_REF: &str = "sales-range";

fn main() -> Result<(), Box<dyn Error>> {
    let mut args = std::env::args_os().skip(1);
    let first = args.next().ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            "usage: web_bindings OUTPUT.xlsx | web_bindings --check INPUT.xlsx",
        )
    })?;
    if first.to_str() == Some("--check") {
        let path = args
            .next()
            .ok_or_else(|| invalid_input("--check requires an input workbook"))?;
        if args.next().is_some() {
            return Err(invalid_input("unexpected argument after input workbook").into());
        }
        return check(path);
    }
    if args.next().is_some() {
        return Err(invalid_input("unexpected argument after output workbook").into());
    }

    create(first)
}

fn create(path: OsString) -> Result<(), Box<dyn Error>> {
    let workbook = Workbook::new()?;

    let reference = Reference::file("Example3", "15.0", r"C:\Example")?;
    let add_in = AddIn::new(INSTANCE_ID, reference)?
        .bind(AddInBinding::new("Matrix1", "matrix", APP_REF)?)?;
    let mut panes = Panes::new();
    panes.push(Pane::new(add_in).show(false))?;

    let mut package = workbook.edit()?;
    package.put_task_panes(panes, Conformance::Transitional)?;

    let mut range = workbook.edit()?;
    range
        .sheet("Sheet1")?
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "Sheet1 is missing"))?
        .set("A1", "Region")?
        .set("B1", "Revenue")?
        .set("A2", "North")?
        .set("B2", 42_i32)?
        .bind(RangeBinding::new(APP_REF, "Sheet1!$A$1:$B$2")?)?;

    package.join(range)?;
    package.commit()?.workbook().save(path)?;
    Ok(())
}

fn check(path: OsString) -> Result<(), Box<dyn Error>> {
    let workbook = Workbook::open(path)?;
    let panes = workbook
        .task_panes()?
        .ok_or_else(|| invalid_data("task panes are missing"))?;
    let pane = panes
        .get(INSTANCE_ID)
        .ok_or_else(|| invalid_data("expected add-in instance is missing"))?;
    let reference = pane.add_in().reference();
    let location = reference
        .location_name()
        .ok_or_else(|| invalid_data("file-system add-in location is missing"))?;
    if reference.store() != Store::FileSystem || location != r"C:\Example" {
        return Err(invalid_data("file-system add-in location changed").into());
    }
    let package_binding = pane
        .add_in()
        .binding("Matrix1")
        .ok_or_else(|| invalid_data("package binding is missing"))?;
    if package_binding.app_ref() != APP_REF {
        return Err(invalid_data("package binding appRef changed").into());
    }

    let sheet = workbook
        .sheet("Sheet1")?
        .ok_or_else(|| invalid_data("Sheet1 is missing"))?;
    let range = sheet
        .web_bindings()?
        .get(APP_REF)
        .ok_or_else(|| invalid_data("worksheet binding is missing"))?;
    if range.formula() != "Sheet1!$A$1:$B$2" {
        return Err(invalid_data("worksheet binding formula changed").into());
    }

    println!(
        "{} -> {} ({})",
        package_binding.app_ref(),
        range.formula(),
        location
    );
    Ok(())
}

fn invalid_input(message: &'static str) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidInput, message)
}

fn invalid_data(message: &'static str) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidData, message)
}
