//! Python property arguments as `litchi::Props` (design §8: "props passed
//! as plain dicts; dotted keys are legal in dicts").
//!
//! `File.set`, `File.add` and the `Edit` verbs take their properties as a
//! `dict` (`{"font.bold": True, "fill": "1F4E79"}`) or as OfficeCLI's string
//! list (`["font.bold=true", "fill=1F4E79"]`). The conversion keeps each
//! value's type, as `litchi::PropValue` does: `None` clears the property,
//! `bool`, `int` and `float` stay typed, and `str` is coerced through the
//! key's value type when the verb stages it ("12pt", "#1F4E79", "true").
//! The registry, not the bindings, decides what a key accepts, so an
//! unknown key or an invalid value raises the same `PropertyError` with
//! the same did-you-mean suggestion as Rust.
//!
//! Every value is copied into an owned `Props<'static>` while the GIL is
//! held, once, so the verb can then run with the GIL released.

use litchi::{PropValue, Props};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDict, PyFloat, PyInt, PyList, PyString, PyTuple};

use crate::errors::Failure;

/// The properties of `value`: a `dict`, OfficeCLI's list of `"key=value"`
/// strings, or `None` for no properties.
pub(crate) fn props(value: Option<&Bound<'_, PyAny>>) -> PyResult<Props<'static>> {
    let Some(value) = value.filter(|value| !value.is_none()) else {
        return Ok(Props::new());
    };
    if let Ok(dict) = value.cast::<PyDict>() {
        let mut props = Props::with_capacity(dict.len());
        for (key, value) in dict.iter() {
            let Ok(key) = key.cast::<PyString>() else {
                return Err(Failure::new(
                    litchi::Code::InvalidArgument,
                    "property names must be str",
                )
                .into());
            };
            let key = key.to_cow()?.into_owned();
            let value = prop_value(&key, &value)?;
            props.push(key, value);
        }
        return Ok(props);
    }
    if value.is_instance_of::<PyList>() || value.is_instance_of::<PyTuple>() {
        let mut entries = Vec::new();
        for entry in value.try_iter()? {
            let entry = entry?;
            let Ok(entry) = entry.cast::<PyString>() else {
                return Err(Failure::new(
                    litchi::Code::InvalidArgument,
                    "a property list holds \"key=value\" strings",
                )
                .into());
            };
            entries.push(entry.to_cow()?.into_owned());
        }
        return Props::from_props(&entries)
            .map(Props::into_owned)
            .map_err(|error| Failure::read(&error).into());
    }
    Err(Failure::new(
        litchi::Code::InvalidArgument,
        "props must be a dict, a list of \"key=value\" strings, or None",
    )
    .into())
}

/// One property value, typed as `PropValue` types it.
fn prop_value(key: &str, value: &Bound<'_, PyAny>) -> PyResult<PropValue<'static>> {
    if value.is_none() {
        return Ok(PropValue::Null);
    }
    // `bool` is a subclass of `int`: test it first.
    if let Ok(flag) = value.cast::<PyBool>() {
        return Ok(PropValue::Bool(flag.is_true()));
    }
    if value.is_instance_of::<PyInt>() {
        if let Ok(number) = value.extract::<i64>() {
            return Ok(PropValue::Int(number));
        }
        if let Ok(number) = value.extract::<u64>() {
            return Ok(PropValue::from(number));
        }
        if let Ok(number) = value.extract::<i128>() {
            return Ok(PropValue::from(number));
        }
        if let Ok(number) = value.extract::<u128>() {
            return Ok(PropValue::from(number));
        }
        return Err(invalid(key, "the integer is out of range").into());
    }
    if let Ok(number) = value.cast::<PyFloat>() {
        return Ok(PropValue::Number(number.value()));
    }
    if let Ok(text) = value.cast::<PyString>() {
        return Ok(PropValue::from(text.to_cow()?.into_owned()));
    }
    // A list or tuple of str: a list value (the metadata keywords, record
    // 0354).
    if value.is_instance_of::<PyList>() || value.is_instance_of::<PyTuple>() {
        let mut items = Vec::new();
        for item in value.try_iter()? {
            let item = item?;
            let Ok(text) = item.cast::<PyString>() else {
                return Err(invalid(key, "a list value holds only str").into());
            };
            items.push(text.to_cow()?.into_owned());
        }
        return Ok(PropValue::from(items));
    }
    Err(invalid(
        key,
        "a property value must be None, bool, int, float, str or a list of str",
    )
    .into())
}

/// An `invalid_value` failure naming the property (never its value, which is
/// authored content).
fn invalid(key: &str, reason: &str) -> Failure {
    Failure::new(
        litchi::Code::InvalidValue,
        format!("property \"{key}\": {reason}"),
    )
}
