//! A `serde` serializer that builds Python objects: the JSON value model,
//! written as `dict`, `list`, `str`, `int`, `float`, `bool` and `None`.
//!
//! Litchi's view JSON (`File::view_json`), capability report
//! (`Capabilities`) and document properties (`Metadata`) are `Serialize`
//! values. Serializing them through this module gives the Python objects
//! `json.loads` would make of their JSON, without writing JSON text: no
//! string round trip (goal, "PERFORMANCE BAR", Python).
//!
//! The mapping follows `serde_json`, so a result equals `json.loads` of the
//! Rust JSON: a `None` or unit is `None`; a non-finite float is `None`; a
//! unit variant is its name; a newtype, tuple or struct variant is a
//! one-entry `dict` keyed by the variant name; bytes are a `list` of `int`;
//! a map key is a `str`, an integer key its decimal text.
//!
//! Python objects need the GIL, so a caller computes the value with the GIL
//! released and serializes it here afterwards.

use std::fmt;

use pyo3::IntoPyObjectExt;
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDict, PyFloat, PyList, PyString};
use serde::ser::{self, Serialize};

use crate::errors::Failure;

/// The Python object of `value`.
pub(crate) fn to_py<'py, T: Serialize + ?Sized>(
    py: Python<'py>,
    value: &T,
) -> PyResult<Bound<'py, PyAny>> {
    value.serialize(ToPy { py }).map_err(|error| error.0)
}

/// A failure while building objects: a Python error (memory, a raising
/// `__hash__`), or a `custom` error a `Serialize` implementation reported.
pub(crate) struct Error(PyErr);

impl fmt::Debug for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::Debug::fmt(&self.0, f)
    }
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::Display::fmt(&self.0, f)
    }
}

impl std::error::Error for Error {}

impl ser::Error for Error {
    /// A value that cannot be serialized: a defect of the Rust value, so
    /// an `internal_error`.
    fn custom<T: fmt::Display>(message: T) -> Self {
        Self(Failure::new(litchi::Code::InternalError, message.to_string()).into())
    }
}

impl From<PyErr> for Error {
    fn from(error: PyErr) -> Self {
        Self(error)
    }
}

type Result<T> = std::result::Result<T, Error>;

/// The serializer: one Python object per value.
#[derive(Clone, Copy)]
struct ToPy<'py> {
    py: Python<'py>,
}

impl<'py> ToPy<'py> {
    fn int(self, value: impl IntoPyObject<'py>) -> Result<Bound<'py, PyAny>> {
        Ok(value.into_bound_py_any(self.py)?)
    }

    /// A one-entry `{variant: value}` dict, `serde_json`'s externally
    /// tagged form.
    fn tagged(self, variant: &str, value: Bound<'py, PyAny>) -> Result<Bound<'py, PyAny>> {
        let dict = PyDict::new(self.py);
        dict.set_item(variant, value)?;
        Ok(dict.into_any())
    }
}

impl<'py> ser::Serializer for ToPy<'py> {
    type Ok = Bound<'py, PyAny>;
    type Error = Error;
    type SerializeSeq = Seq<'py>;
    type SerializeTuple = Seq<'py>;
    type SerializeTupleStruct = Seq<'py>;
    type SerializeTupleVariant = Seq<'py>;
    type SerializeMap = Map<'py>;
    type SerializeStruct = Map<'py>;
    type SerializeStructVariant = Map<'py>;

    fn serialize_bool(self, value: bool) -> Result<Self::Ok> {
        Ok(PyBool::new(self.py, value).to_owned().into_any())
    }

    fn serialize_i8(self, value: i8) -> Result<Self::Ok> {
        self.int(value)
    }

    fn serialize_i16(self, value: i16) -> Result<Self::Ok> {
        self.int(value)
    }

    fn serialize_i32(self, value: i32) -> Result<Self::Ok> {
        self.int(value)
    }

    fn serialize_i64(self, value: i64) -> Result<Self::Ok> {
        self.int(value)
    }

    fn serialize_i128(self, value: i128) -> Result<Self::Ok> {
        self.int(value)
    }

    fn serialize_u8(self, value: u8) -> Result<Self::Ok> {
        self.int(value)
    }

    fn serialize_u16(self, value: u16) -> Result<Self::Ok> {
        self.int(value)
    }

    fn serialize_u32(self, value: u32) -> Result<Self::Ok> {
        self.int(value)
    }

    fn serialize_u64(self, value: u64) -> Result<Self::Ok> {
        self.int(value)
    }

    fn serialize_u128(self, value: u128) -> Result<Self::Ok> {
        self.int(value)
    }

    fn serialize_f32(self, value: f32) -> Result<Self::Ok> {
        self.serialize_f64(f64::from(value))
    }

    /// A float; `None` when it is not finite, as JSON has no such number.
    fn serialize_f64(self, value: f64) -> Result<Self::Ok> {
        if value.is_finite() {
            Ok(PyFloat::new(self.py, value).into_any())
        } else {
            self.serialize_unit()
        }
    }

    fn serialize_char(self, value: char) -> Result<Self::Ok> {
        self.serialize_str(value.encode_utf8(&mut [0; 4]))
    }

    fn serialize_str(self, value: &str) -> Result<Self::Ok> {
        Ok(PyString::new(self.py, value).into_any())
    }

    /// Bytes as a list of numbers, as JSON writes them.
    fn serialize_bytes(self, value: &[u8]) -> Result<Self::Ok> {
        Ok(PyList::new(self.py, value.iter().copied())?.into_any())
    }

    fn serialize_none(self) -> Result<Self::Ok> {
        self.serialize_unit()
    }

    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<Self::Ok> {
        value.serialize(self)
    }

    fn serialize_unit(self) -> Result<Self::Ok> {
        Ok(self.py.None().into_bound(self.py))
    }

    fn serialize_unit_struct(self, _name: &'static str) -> Result<Self::Ok> {
        self.serialize_unit()
    }

    fn serialize_unit_variant(
        self,
        _name: &'static str,
        _index: u32,
        variant: &'static str,
    ) -> Result<Self::Ok> {
        self.serialize_str(variant)
    }

    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _name: &'static str,
        value: &T,
    ) -> Result<Self::Ok> {
        value.serialize(self)
    }

    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _name: &'static str,
        _index: u32,
        variant: &'static str,
        value: &T,
    ) -> Result<Self::Ok> {
        let value = value.serialize(self)?;
        self.tagged(variant, value)
    }

    fn serialize_seq(self, _len: Option<usize>) -> Result<Self::SerializeSeq> {
        Ok(Seq::new(self, None))
    }

    fn serialize_tuple(self, len: usize) -> Result<Self::SerializeTuple> {
        self.serialize_seq(Some(len))
    }

    fn serialize_tuple_struct(
        self,
        _name: &'static str,
        len: usize,
    ) -> Result<Self::SerializeTupleStruct> {
        self.serialize_seq(Some(len))
    }

    fn serialize_tuple_variant(
        self,
        _name: &'static str,
        _index: u32,
        variant: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeTupleVariant> {
        Ok(Seq::new(self, Some(variant)))
    }

    fn serialize_map(self, _len: Option<usize>) -> Result<Self::SerializeMap> {
        Ok(Map::new(self, None))
    }

    fn serialize_struct(self, _name: &'static str, _len: usize) -> Result<Self::SerializeStruct> {
        Ok(Map::new(self, None))
    }

    fn serialize_struct_variant(
        self,
        _name: &'static str,
        _index: u32,
        variant: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeStructVariant> {
        Ok(Map::new(self, Some(variant)))
    }

    fn collect_str<T: fmt::Display + ?Sized>(self, value: &T) -> Result<Self::Ok> {
        self.serialize_str(&value.to_string())
    }
}

/// A `list` being filled; `variant` wraps it as `{variant: [...]}`.
pub(crate) struct Seq<'py> {
    to: ToPy<'py>,
    list: Bound<'py, PyList>,
    variant: Option<&'static str>,
}

impl<'py> Seq<'py> {
    fn new(to: ToPy<'py>, variant: Option<&'static str>) -> Self {
        Self {
            to,
            list: PyList::empty(to.py),
            variant,
        }
    }

    fn push<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<()> {
        self.list.append(value.serialize(self.to)?)?;
        Ok(())
    }

    fn finish(self) -> Result<Bound<'py, PyAny>> {
        let list = self.list.into_any();
        match self.variant {
            Some(variant) => self.to.tagged(variant, list),
            None => Ok(list),
        }
    }
}

impl<'py> ser::SerializeSeq for Seq<'py> {
    type Ok = Bound<'py, PyAny>;
    type Error = Error;

    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<()> {
        self.push(value)
    }

    fn end(self) -> Result<Self::Ok> {
        self.finish()
    }
}

impl<'py> ser::SerializeTuple for Seq<'py> {
    type Ok = Bound<'py, PyAny>;
    type Error = Error;

    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<()> {
        self.push(value)
    }

    fn end(self) -> Result<Self::Ok> {
        self.finish()
    }
}

impl<'py> ser::SerializeTupleStruct for Seq<'py> {
    type Ok = Bound<'py, PyAny>;
    type Error = Error;

    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<()> {
        self.push(value)
    }

    fn end(self) -> Result<Self::Ok> {
        self.finish()
    }
}

impl<'py> ser::SerializeTupleVariant for Seq<'py> {
    type Ok = Bound<'py, PyAny>;
    type Error = Error;

    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<()> {
        self.push(value)
    }

    fn end(self) -> Result<Self::Ok> {
        self.finish()
    }
}

/// A `dict` being filled; `variant` wraps it as `{variant: {...}}`.
pub(crate) struct Map<'py> {
    to: ToPy<'py>,
    dict: Bound<'py, PyDict>,
    /// The key of `serialize_key`, until its value arrives.
    key: Option<Bound<'py, PyString>>,
    variant: Option<&'static str>,
}

impl<'py> Map<'py> {
    fn new(to: ToPy<'py>, variant: Option<&'static str>) -> Self {
        Self {
            to,
            dict: PyDict::new(to.py),
            key: None,
            variant,
        }
    }

    fn finish(self) -> Result<Bound<'py, PyAny>> {
        let dict = self.dict.into_any();
        match self.variant {
            Some(variant) => self.to.tagged(variant, dict),
            None => Ok(dict),
        }
    }
}

impl<'py> ser::SerializeMap for Map<'py> {
    type Ok = Bound<'py, PyAny>;
    type Error = Error;

    fn serialize_key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<()> {
        self.key = Some(key.serialize(Key { py: self.to.py })?);
        Ok(())
    }

    fn serialize_value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<()> {
        let key = self
            .key
            .take()
            .ok_or_else(|| <Error as ser::Error>::custom("a map value arrived without its key"))?;
        self.dict.set_item(key, value.serialize(self.to)?)?;
        Ok(())
    }

    fn end(self) -> Result<Self::Ok> {
        self.finish()
    }
}

impl<'py> ser::SerializeStruct for Map<'py> {
    type Ok = Bound<'py, PyAny>;
    type Error = Error;

    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<()> {
        self.dict.set_item(key, value.serialize(self.to)?)?;
        Ok(())
    }

    fn end(self) -> Result<Self::Ok> {
        self.finish()
    }
}

impl<'py> ser::SerializeStructVariant for Map<'py> {
    type Ok = Bound<'py, PyAny>;
    type Error = Error;

    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<()> {
        self.dict.set_item(key, value.serialize(self.to)?)?;
        Ok(())
    }

    fn end(self) -> Result<Self::Ok> {
        self.finish()
    }
}

/// The serializer of a map key: a `str`, as JSON object keys are. Integer
/// keys become their decimal text, as `serde_json` writes them; any other
/// key is refused.
struct Key<'py> {
    py: Python<'py>,
}

impl<'py> Key<'py> {
    fn text(self, text: &str) -> Result<Bound<'py, PyString>> {
        Ok(PyString::new(self.py, text))
    }

    fn number(self, value: impl fmt::Display) -> Result<Bound<'py, PyString>> {
        self.text(&value.to_string())
    }

    fn refuse<T>(what: &str) -> Result<T> {
        Err(<Error as ser::Error>::custom(format!(
            "a map key must be a string, not {what}"
        )))
    }
}

impl<'py> ser::Serializer for Key<'py> {
    type Ok = Bound<'py, PyString>;
    type Error = Error;
    type SerializeSeq = ser::Impossible<Self::Ok, Error>;
    type SerializeTuple = ser::Impossible<Self::Ok, Error>;
    type SerializeTupleStruct = ser::Impossible<Self::Ok, Error>;
    type SerializeTupleVariant = ser::Impossible<Self::Ok, Error>;
    type SerializeMap = ser::Impossible<Self::Ok, Error>;
    type SerializeStruct = ser::Impossible<Self::Ok, Error>;
    type SerializeStructVariant = ser::Impossible<Self::Ok, Error>;

    fn serialize_bool(self, _value: bool) -> Result<Self::Ok> {
        Self::refuse("a boolean")
    }

    fn serialize_i8(self, value: i8) -> Result<Self::Ok> {
        self.number(value)
    }

    fn serialize_i16(self, value: i16) -> Result<Self::Ok> {
        self.number(value)
    }

    fn serialize_i32(self, value: i32) -> Result<Self::Ok> {
        self.number(value)
    }

    fn serialize_i64(self, value: i64) -> Result<Self::Ok> {
        self.number(value)
    }

    fn serialize_i128(self, value: i128) -> Result<Self::Ok> {
        self.number(value)
    }

    fn serialize_u8(self, value: u8) -> Result<Self::Ok> {
        self.number(value)
    }

    fn serialize_u16(self, value: u16) -> Result<Self::Ok> {
        self.number(value)
    }

    fn serialize_u32(self, value: u32) -> Result<Self::Ok> {
        self.number(value)
    }

    fn serialize_u64(self, value: u64) -> Result<Self::Ok> {
        self.number(value)
    }

    fn serialize_u128(self, value: u128) -> Result<Self::Ok> {
        self.number(value)
    }

    fn serialize_f32(self, _value: f32) -> Result<Self::Ok> {
        Self::refuse("a float")
    }

    fn serialize_f64(self, _value: f64) -> Result<Self::Ok> {
        Self::refuse("a float")
    }

    fn serialize_char(self, value: char) -> Result<Self::Ok> {
        self.text(value.encode_utf8(&mut [0; 4]))
    }

    fn serialize_str(self, value: &str) -> Result<Self::Ok> {
        self.text(value)
    }

    fn serialize_bytes(self, _value: &[u8]) -> Result<Self::Ok> {
        Self::refuse("bytes")
    }

    fn serialize_none(self) -> Result<Self::Ok> {
        Self::refuse("None")
    }

    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<Self::Ok> {
        value.serialize(self)
    }

    fn serialize_unit(self) -> Result<Self::Ok> {
        Self::refuse("a unit")
    }

    fn serialize_unit_struct(self, _name: &'static str) -> Result<Self::Ok> {
        Self::refuse("a unit")
    }

    fn serialize_unit_variant(
        self,
        _name: &'static str,
        _index: u32,
        variant: &'static str,
    ) -> Result<Self::Ok> {
        self.text(variant)
    }

    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _name: &'static str,
        value: &T,
    ) -> Result<Self::Ok> {
        value.serialize(self)
    }

    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _name: &'static str,
        _index: u32,
        _variant: &'static str,
        _value: &T,
    ) -> Result<Self::Ok> {
        Self::refuse("an enum variant with data")
    }

    fn serialize_seq(self, _len: Option<usize>) -> Result<Self::SerializeSeq> {
        Self::refuse("a sequence")
    }

    fn serialize_tuple(self, _len: usize) -> Result<Self::SerializeTuple> {
        Self::refuse("a tuple")
    }

    fn serialize_tuple_struct(
        self,
        _name: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeTupleStruct> {
        Self::refuse("a tuple")
    }

    fn serialize_tuple_variant(
        self,
        _name: &'static str,
        _index: u32,
        _variant: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeTupleVariant> {
        Self::refuse("a tuple")
    }

    fn serialize_map(self, _len: Option<usize>) -> Result<Self::SerializeMap> {
        Self::refuse("a map")
    }

    fn serialize_struct(self, _name: &'static str, _len: usize) -> Result<Self::SerializeStruct> {
        Self::refuse("a struct")
    }

    fn serialize_struct_variant(
        self,
        _name: &'static str,
        _index: u32,
        _variant: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeStructVariant> {
        Self::refuse("a struct")
    }

    fn collect_str<T: fmt::Display + ?Sized>(self, value: &T) -> Result<Self::Ok> {
        self.text(&value.to_string())
    }
}
