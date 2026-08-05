use litchi_opc::PackURI;

/// ActiveX persistence mode declared by an `ax:ocx` descriptor.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Persistence {
    PropertyBag,
    Stream,
    StreamInit,
    Storage,
    Unknown,
}

/// Inert slide control metadata and its resolved descriptor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Control {
    pub(crate) slide_index: usize,
    pub(crate) index: usize,
    pub(crate) shape_id: Option<String>,
    pub(crate) name: Option<String>,
    pub(crate) show_as_icon: Option<bool>,
    pub(crate) image_width: Option<u32>,
    pub(crate) image_height: Option<u32>,
    pub(crate) relationship_id: Option<String>,
    pub(crate) descriptor: Option<Descriptor>,
}

impl Control {
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    pub fn index(&self) -> usize {
        self.index
    }

    pub fn shape_id(&self) -> Option<&str> {
        self.shape_id.as_deref()
    }

    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    pub fn show_as_icon(&self) -> Option<bool> {
        self.show_as_icon
    }

    pub fn image_width(&self) -> Option<u32> {
        self.image_width
    }

    pub fn image_height(&self) -> Option<u32> {
        self.image_height
    }

    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }

    pub fn descriptor(&self) -> Option<&Descriptor> {
        self.descriptor.as_ref()
    }
}

/// Inert `ax:ocx` descriptor metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Descriptor {
    pub(crate) part_name: PackURI,
    pub(crate) class_id: String,
    pub(crate) license: Option<String>,
    pub(crate) persistence: Persistence,
    pub(crate) binary: Option<Binary>,
}

impl Descriptor {
    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }

    pub fn class_id(&self) -> &str {
        &self.class_id
    }

    pub fn license(&self) -> Option<&str> {
        self.license.as_deref()
    }

    pub fn persistence(&self) -> Persistence {
        self.persistence
    }

    pub fn binary(&self) -> Option<&Binary> {
        self.binary.as_ref()
    }
}

/// Inert metadata for an ActiveX binary state part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Binary {
    pub(crate) relationship_id: String,
    pub(crate) part_name: PackURI,
    pub(crate) byte_length: usize,
}

impl Binary {
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }

    pub fn byte_length(&self) -> usize {
        self.byte_length
    }
}
