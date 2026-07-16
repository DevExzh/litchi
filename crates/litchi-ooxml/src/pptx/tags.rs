//! Typed PresentationML programmable tag-list parts.
//!
//! Tag values are exposed as inert strings. Embedded XML, paths, or commands
//! contained in a value are never parsed or executed.

use crate::common::mce::process_ooxml;
use crate::error::{OoxmlError, Result};
use litchi_opc::{OpcPackage, Part};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashSet;

const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const PML_TEXT: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_TEXT: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const TAG_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags";
const STRICT_TAG_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/tags";
const TAG_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml";
const MAX_PART_BYTES: usize = 8 * 1024 * 1024;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_TAGS: usize = 16_384;
const MAX_TAG_PARTS: usize = 1_024;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TagListConformance { Transitional, Strict }
impl TagListConformance { fn namespace(self)->&'static str{match self{Self::Transitional=>PML_TEXT,Self::Strict=>STRICT_TEXT}} }

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TagExtensionAttribute { qualified_name:String, value:String }
impl TagExtensionAttribute { pub fn qualified_name(&self)->&str{&self.qualified_name} pub fn value(&self)->&str{&self.value} }

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProgrammableTag { name:String, value:String, extension_attributes:Vec<TagExtensionAttribute> }
impl ProgrammableTag { pub fn name(&self)->&str{&self.name} pub fn value(&self)->&str{&self.value} pub fn extension_attributes(&self)->&[TagExtensionAttribute]{&self.extension_attributes} }

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TagList { tags:Vec<ProgrammableTag>, namespaces:Vec<TagExtensionAttribute>, extension_attributes:Vec<TagExtensionAttribute> }
impl TagList {
    pub fn tags(&self)->&[ProgrammableTag]{&self.tags}
    pub fn extension_attributes(&self)->&[TagExtensionAttribute]{&self.extension_attributes}
    pub fn to_xml(&self,conformance:TagListConformance)->Result<Vec<u8>>{write_tag_list(self,conformance)}
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideTagList { relationship_id:String, part_name:String, tag_list:TagList }
impl SlideTagList {
    pub fn relationship_id(&self)->&str{&self.relationship_id}
    pub fn part_name(&self)->&str{&self.part_name}
    pub fn tag_list(&self)->&TagList{&self.tag_list}
}

pub fn is_tag_relationship(value:&str)->bool{matches!(value,TAG_REL|STRICT_TAG_REL)}

pub fn parse_tag_list(xml:&[u8])->Result<TagList>{
    if xml.len()>MAX_PART_BYTES{return Err(invalid("tag-list part is too large"))}
    let xml=process_ooxml(xml)?;
    if xml.len()>MAX_PART_BYTES{return Err(invalid("MCE-expanded tag-list part is too large"))}
    let mut reader=NsReader::from_reader(xml.as_ref());reader.config_mut().trim_text(false);
    let mut buffer=Vec::new();let mut depth=0usize;let mut root=false;let mut closed=false;
    let mut open_tag:Option<(usize,ProgrammableTag)>=None;let mut tags=Vec::new();
    let mut namespaces=Vec::new();let mut root_extensions=Vec::new();
    loop{let(namespace,event)=reader.read_resolved_event_into(&mut buffer).map_err(xml_error)?;match event{
        Event::Start(e)=>{let name=e.local_name();if !root&&depth==0&&pml(&namespace)&&name.as_ref()==b"tagLst"{root=true;let parsed=parse_attributes(&e,&[],reader.decoder())?;namespaces=parsed.namespaces;root_extensions=parsed.extensions;depth=1;}else if root&&!closed&&depth==1&&open_tag.is_none()&&pml(&namespace)&&name.as_ref()==b"tag"{if tags.len()==MAX_TAGS{return Err(invalid("tag list has too many tags"))}let tag=parse_tag(&e,reader.decoder())?;depth+=1;open_tag=Some((depth,tag));}else{return Err(invalid(format!("unexpected tag-list element '{}'",String::from_utf8_lossy(name.as_ref()))))}},
        Event::Empty(e)=>{let name=e.local_name();if !root&&depth==0&&pml(&namespace)&&name.as_ref()==b"tagLst"{let parsed=parse_attributes(&e,&[],reader.decoder())?;namespaces=parsed.namespaces;root_extensions=parsed.extensions;root=true;closed=true;}else if root&&!closed&&depth==1&&open_tag.is_none()&&pml(&namespace)&&name.as_ref()==b"tag"{if tags.len()==MAX_TAGS{return Err(invalid("tag list has too many tags"))}tags.push(parse_tag(&e,reader.decoder())?);}else{return Err(invalid(format!("unexpected tag-list element '{}'",String::from_utf8_lossy(name.as_ref()))))}},
        Event::End(e)=>{let name=e.local_name();if open_tag.as_ref().is_some_and(|(level,_)|*level==depth)&&pml(&namespace)&&name.as_ref()==b"tag"{let(_,tag)=open_tag.take().unwrap();tags.push(tag);depth-=1;}else if root&&!closed&&depth==1&&pml(&namespace)&&name.as_ref()==b"tagLst"{closed=true;depth=0;}else{return Err(invalid("unexpected tag-list end element"))}},
        Event::Text(e)=>{let value=e.decode().map_err(xml_error)?;let value=quick_xml::escape::unescape(&value).map_err(xml_error)?;if !value.trim().is_empty(){return Err(invalid("tag elements cannot contain text"))}},
        Event::CData(e)=>if !e.decode().map_err(xml_error)?.trim().is_empty(){return Err(invalid("tag elements cannot contain CDATA"))},
        Event::DocType(_)|Event::PI(_)=>return Err(invalid("DTD and processing instructions are rejected")),
        Event::Decl(_)|Event::Comment(_)|Event::GeneralRef(_)=>{},Event::Eof=>break}
        buffer.clear();}
    if !root||!closed||depth!=0||open_tag.is_some(){return Err(invalid("unterminated tag-list part"))}
    Ok(TagList{tags,namespaces,extension_attributes:root_extensions})
}

struct ParsedAttributes{values:Vec<(String,String)>,namespaces:Vec<TagExtensionAttribute>,extensions:Vec<TagExtensionAttribute>}
fn parse_attributes(e:&BytesStart<'_>,known:&[&str],decoder:Decoder)->Result<ParsedAttributes>{let mut values=Vec::new();let mut namespaces=Vec::new();let mut extensions=Vec::new();let mut seen=HashSet::new();for a in e.attributes().with_checks(true){let a=a.map_err(xml_error)?;let name=std::str::from_utf8(a.key.as_ref()).map_err(xml_error)?.to_string();if !seen.insert(name.clone()){return Err(invalid("duplicate XML attribute"))}let value=a.decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0,decoder).map_err(xml_error)?.into_owned();if value.len()>MAX_TEXT_BYTES{return Err(invalid("tag attribute is too large"))}if name=="xmlns"||name.starts_with("xmlns:"){if !matches!(value.as_str(),PML_TEXT|STRICT_TEXT|"http://schemas.openxmlformats.org/drawingml/2006/main"|"http://schemas.openxmlformats.org/officeDocument/2006/relationships"|"http://purl.oclc.org/ooxml/officeDocument/relationships"|"http://schemas.openxmlformats.org/markup-compatibility/2006"){namespaces.push(TagExtensionAttribute{qualified_name:name,value});}}else if !name.contains(':')&&known.contains(&name.as_str()){values.push((name,value));}else if name.contains(':'){extensions.push(TagExtensionAttribute{qualified_name:name,value});}else{return Err(invalid(format!("unexpected tag attribute '{name}'")))}}Ok(ParsedAttributes{values,namespaces,extensions})}
fn parse_tag(e:&BytesStart<'_>,decoder:Decoder)->Result<ProgrammableTag>{let p=parse_attributes(e,&["name","val"],decoder)?;let get=|name:&str|p.values.iter().find(|(k,_)|k==name).map(|(_,v)|v.clone()).ok_or_else(||invalid(format!("tag is missing '{name}'")));let name=get("name")?;let value=get("val")?;let mut extension_attributes=p.namespaces;extension_attributes.extend(p.extensions);Ok(ProgrammableTag{name,value,extension_attributes})}

pub fn write_tag_list(value:&TagList,conformance:TagListConformance)->Result<Vec<u8>>{let mut out=br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec();out.extend_from_slice(b"<p:tagLst xmlns:p=\"");escape(&mut out,conformance.namespace());out.push(b'\"');for a in &value.namespaces{write_preserved(&mut out,a)?}for a in &value.extension_attributes{write_preserved(&mut out,a)?}if value.tags.is_empty(){out.extend_from_slice(b"/>");return Ok(out)}out.push(b'>');for tag in &value.tags{out.extend_from_slice(b"<p:tag");for a in &tag.extension_attributes{write_preserved(&mut out,a)?}write_attr(&mut out,"name",&tag.name);write_attr(&mut out,"val",&tag.value);out.extend_from_slice(b"/>")}out.extend_from_slice(b"</p:tagLst>");Ok(out)}
fn write_preserved(out:&mut Vec<u8>,a:&TagExtensionAttribute)->Result<()>{if a.qualified_name.is_empty()||a.qualified_name.bytes().any(|b|b.is_ascii_whitespace()||matches!(b,b'<'|b'>'|b'='|b'\''|b'\"')){return Err(invalid("invalid preserved tag attribute name"))}write_attr(out,&a.qualified_name,&a.value);Ok(())}
fn write_attr(out:&mut Vec<u8>,name:&str,value:&str){out.push(b' ');out.extend_from_slice(name.as_bytes());out.extend_from_slice(b"=\"");escape(out,value);out.push(b'\"')}
fn escape(out:&mut Vec<u8>,value:&str){for c in value.chars(){match c{'&'=>out.extend_from_slice(b"&amp;"),'<'=>out.extend_from_slice(b"&lt;"),'"'=>out.extend_from_slice(b"&quot;"),'\t'=>out.extend_from_slice(b"&#x9;"),'\n'=>out.extend_from_slice(b"&#xA;"),'\r'=>out.extend_from_slice(b"&#xD;"),_=>{let mut b=[0;4];out.extend_from_slice(c.encode_utf8(&mut b).as_bytes())}}}}

pub(crate)fn load_slide_tag_lists(slide:&dyn Part,package:&OpcPackage)->Result<Vec<SlideTagList>>{let relationships:Vec<_>=slide.rels().iter().filter(|r|is_tag_relationship(r.reltype())).collect();if relationships.len()>MAX_TAG_PARTS{return Err(invalid("slide has too many tag-list relationships"))}let mut targets=HashSet::new();let mut output=Vec::with_capacity(relationships.len());for relationship in relationships{if relationship.is_external(){return Err(invalid(format!("tag-list relationship '{}' cannot be external",relationship.r_id())))}let target=relationship.target_partname()?;if !targets.insert(target.to_string()){return Err(invalid(format!("duplicate slide tag-list target '{target}'")))}let part=package.get_part(&target)?;if part.content_type()!=TAG_CONTENT_TYPE{return Err(OoxmlError::InvalidContentType{expected:TAG_CONTENT_TYPE.into(),got:part.content_type().into()})}if part.rels().iter().next().is_some(){return Err(invalid(format!("tag-list part '{target}' has unexpected relationships")))}output.push(SlideTagList{relationship_id:relationship.r_id().into(),part_name:target.to_string(),tag_list:parse_tag_list(part.blob())?})}Ok(output)}
fn pml(ns:&ResolveResult<'_>)->bool{matches!(ns,ResolveResult::Bound(v)if v.as_ref()==PML||v.as_ref()==STRICT)}fn xml_error(e:impl std::fmt::Display)->OoxmlError{OoxmlError::Xml(e.to_string())}fn invalid(e:impl Into<String>)->OoxmlError{OoxmlError::InvalidFormat(e.into())}

#[cfg(test)]mod tests{use super::*;
#[test]fn strict_round_trip_and_inert_values(){let xml=br#"<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tag name="PATH" val="C:\Docs\file"/><p:tag name="XML" val="&lt;root command=&quot;none&quot;/&gt;"></p:tag></p:tagLst>"#;let value=parse_tag_list(xml).unwrap();assert_eq!(value.tags()[1].value(),"<root command=\"none\"/>");let strict=value.to_xml(TagListConformance::Strict).unwrap();assert!(std::str::from_utf8(&strict).unwrap().contains(STRICT_TEXT));assert_eq!(parse_tag_list(&strict).unwrap(),value)}
#[test]fn mce_fallback(){let xml=br#"<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><mc:AlternateContent><mc:Choice Requires="x"><x:tag/></mc:Choice><mc:Fallback><p:tag name="fallback" val="1"/></mc:Fallback></mc:AlternateContent></p:tagLst>"#;assert_eq!(parse_tag_list(xml).unwrap().tags()[0].name(),"fallback")}
#[test]fn malformed_and_bounds(){for xml in [r#"<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tag val="x"/></p:tagLst>"#,r#"<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tag name="x"/></p:tagLst>"#,r#"<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tag name="x" val="y"><p:tag name="z" val="q"/></p:tag></p:tagLst>"#,r#"<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:other/></p:tagLst>"#,r#"<!DOCTYPE x><p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,r#"<?bad x?><p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#]{assert!(parse_tag_list(xml.as_bytes()).is_err(),"{xml}")}assert!(parse_tag_list(&vec![b' ';MAX_PART_BYTES+1]).is_err())}
#[test]fn real_libreoffice_package_tags(){let root=std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");let package=crate::pptx::Package::open(root.join("3rdparty/libreoffice-core/sd/qa/unit/data/pptx/tdf103477.pptx")).unwrap();let presentation=package.presentation().unwrap();let slides=presentation.slides().unwrap();let lists=slides[0].tag_lists().unwrap();assert_eq!(lists.len(),7);assert!(lists.iter().all(|v|!v.tag_list().tags().is_empty()));assert!(lists.iter().all(|v|v.part_name().starts_with("/ppt/tags/")))}}
