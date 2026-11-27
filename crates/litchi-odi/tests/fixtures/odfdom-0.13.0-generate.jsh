// Provenance transcript for the checked-in ODFDOM 0.13.0 ODI pair.
// Run with JShell and the pinned ODFDOM runtime dependencies on --class-path.
import org.odftoolkit.odfdom.doc.OdfImageDocument;
import org.w3c.dom.Element;

var original = "/tmp/odfdom-0.13.0-producer-original.odi";
var changed = "/tmp/odfdom-0.13.0-producer-changed.odi";
var draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
var office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
var svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";

var document = OdfImageDocument.newImageDocument();
var content = document.getContentDom();
var imageRoot = document.getContentRoot();
var frame = (Element) imageRoot.getFirstChild();
frame.setAttributeNS(draw, "draw:name", "ODFDOM-0.13.0-Original");
frame.setAttributeNS(svg, "svg:width", "1cm");
frame.setAttributeNS(svg, "svg:height", "1cm");
var image = content.createElementNS(draw, "draw:image");
var binary = content.createElementNS(office, "office:binary-data");
binary.setTextContent("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
image.appendChild(binary);
frame.appendChild(image);
document.save(original);
document.close();

var reopened = OdfImageDocument.loadDocument(original);
var changedFrame = (Element) reopened.getContentRoot().getFirstChild();
changedFrame.setAttributeNS(draw, "draw:name", "ODFDOM-0.13.0-Changed");
reopened.save(changed);
reopened.close();
/exit
