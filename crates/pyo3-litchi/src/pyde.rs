//! A `serde` deserializer that reads Python objects: the JSON value model
//! (`dict`, `list`/`tuple`, `str`, `int`, `float`, `bool`, `None`), the
//! inverse of [`crate::pyser`].
//!
//! `File.batch` and `File.apply` take OfficeCLI batch items as dicts.
//! Deserializing them through `litchi::BatchItem`'s own `Deserialize`
//! implementation applies exactly the wire rules the JSON decoder applies
//! (ASCII case-insensitive keys, `op` as an alias of `command`, unknown and
//! repeated keys refused, `None` as an absent field, signed 32-bit numbers,
//! `props` as an object or a `"key=value"` list, typed property values), and
//! converts every item to Rust once, without writing JSON text (goal,
//! "PERFORMANCE BAR", Python: "Batch payloads are converted to Rust once").
//!
//! Python strings are copied as they are read, so the items own their text
//! and can be executed with the GIL released.

use std::fmt;

use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDict, PyDictMethods, PyFloat, PyInt, PyList, PyString, PyTuple};
use serde::de::{self, DeserializeSeed, Deserializer, MapAccess, SeqAccess, Visitor};

/// A deserialization failure: the message of a `custom` error, or a Python
/// error met while reading an object.
#[derive(Debug)]
pub(crate) struct Error(pub(crate) String);

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::error::Error for Error {}

impl de::Error for Error {
    fn custom<T: fmt::Display>(message: T) -> Self {
        Self(message.to_string())
    }
}

impl From<PyErr> for Error {
    fn from(error: PyErr) -> Self {
        Self(error.to_string())
    }
}

type Result<T> = std::result::Result<T, Error>;

/// One Python object, read as a JSON value.
pub(crate) struct Value<'a, 'py>(pub(crate) &'a Bound<'py, PyAny>);

impl<'de> Deserializer<'de> for Value<'_, '_> {
    type Error = Error;

    fn deserialize_any<V: Visitor<'de>>(self, visitor: V) -> Result<V::Value> {
        let object = self.0;
        if object.is_none() {
            return visitor.visit_unit();
        }
        // `bool` is a subclass of `int`: test it first.
        if let Ok(flag) = object.cast::<PyBool>() {
            return visitor.visit_bool(flag.is_true());
        }
        if object.is_instance_of::<PyInt>() {
            if let Ok(number) = object.extract::<i64>() {
                return visitor.visit_i64(number);
            }
            if let Ok(number) = object.extract::<u64>() {
                return visitor.visit_u64(number);
            }
            if let Ok(number) = object.extract::<i128>() {
                return visitor.visit_i128(number);
            }
            if let Ok(number) = object.extract::<u128>() {
                return visitor.visit_u128(number);
            }
            return Err(de::Error::custom("integer out of range"));
        }
        if let Ok(number) = object.cast::<PyFloat>() {
            return visitor.visit_f64(number.value());
        }
        if let Ok(text) = object.cast::<PyString>() {
            return visitor.visit_str(&text.to_cow()?);
        }
        if let Ok(dict) = object.cast::<PyDict>() {
            let entries: Vec<_> = dict.iter().collect();
            return visitor.visit_map(Map {
                entries: entries.into_iter(),
                value: None,
            });
        }
        if let Ok(list) = object.cast::<PyList>() {
            let items: Vec<_> = list.iter().collect();
            return visitor.visit_seq(Seq(items.into_iter()));
        }
        if let Ok(tuple) = object.cast::<PyTuple>() {
            let items: Vec<_> = tuple.iter().collect();
            return visitor.visit_seq(Seq(items.into_iter()));
        }
        Err(de::Error::custom(
            "a batch value must be a dict, list, tuple, str, int, float, bool or None",
        ))
    }

    fn deserialize_option<V: Visitor<'de>>(self, visitor: V) -> Result<V::Value> {
        if self.0.is_none() {
            visitor.visit_none()
        } else {
            visitor.visit_some(self)
        }
    }

    serde::forward_to_deserialize_any! {
        bool i8 i16 i32 i64 i128 u8 u16 u32 u64 u128 f32 f64 char str string
        bytes byte_buf unit unit_struct newtype_struct seq tuple
        tuple_struct map struct enum identifier ignored_any
    }
}

/// A `dict`'s entries, in insertion order.
struct Map<'py> {
    entries: std::vec::IntoIter<(Bound<'py, PyAny>, Bound<'py, PyAny>)>,
    value: Option<Bound<'py, PyAny>>,
}

impl<'de> MapAccess<'de> for Map<'_> {
    type Error = Error;

    fn next_key_seed<K: DeserializeSeed<'de>>(&mut self, seed: K) -> Result<Option<K::Value>> {
        let Some((key, value)) = self.entries.next() else {
            return Ok(None);
        };
        self.value = Some(value);
        seed.deserialize(Value(&key)).map(Some)
    }

    fn next_value_seed<V: DeserializeSeed<'de>>(&mut self, seed: V) -> Result<V::Value> {
        let value = self
            .value
            .take()
            .ok_or_else(|| Error("a map value was read before its key".to_owned()))?;
        seed.deserialize(Value(&value))
    }

    fn size_hint(&self) -> Option<usize> {
        Some(self.entries.len())
    }
}

/// A `list`'s or `tuple`'s items.
struct Seq<'py>(std::vec::IntoIter<Bound<'py, PyAny>>);

impl<'de> SeqAccess<'de> for Seq<'_> {
    type Error = Error;

    fn next_element_seed<T: DeserializeSeed<'de>>(&mut self, seed: T) -> Result<Option<T::Value>> {
        match self.0.next() {
            Some(item) => seed.deserialize(Value(&item)).map(Some),
            None => Ok(None),
        }
    }

    fn size_hint(&self) -> Option<usize> {
        Some(self.0.len())
    }
}
