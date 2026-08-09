//! `ActiveX` descriptor and nested property-bag models.

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Persistence {
    PropertyBag,
    Stream,
    StreamInit,
    Storage,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Picture {
    pub relationship_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Font {
    pub persistence: Option<Persistence>,
    pub relationship_id: Option<String>,
    pub properties: Vec<Property>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PropertyObject {
    Font(Font),
    Picture(Picture),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Property {
    pub name: String,
    pub value: Option<String>,
    pub object: Option<PropertyObject>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Descriptor {
    pub class_id: String,
    pub license: Option<String>,
    pub persistence: Persistence,
    pub relationship_id: Option<String>,
    pub properties: Vec<Property>,
}
