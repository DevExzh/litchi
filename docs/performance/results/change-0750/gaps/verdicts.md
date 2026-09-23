# Every probed input: the source policy before and after, and whether the other auditors changed

Inputs are the probe cases of change 0750. Control characters and noncharacters are shown escaped as `\u{…}`. "unchanged" means `verify`, `verify_authored` and `verify_reader` return the identical verdict text on both legs.

| input | bytes | `verify_source` before | `verify_source` after | other auditors |
| --- | --- | --- | --- | --- |
| gap: ref outside root | `<r/>& ;` | OK | malformed XML at byte 4: character data outside the document element | unchanged |
| gap: empty ref | `<r>&;</r>` | OK | malformed XML at byte 3: malformed reference | unchanged |
| gap: ref with space | `<r>&a b;</r>` | OK | malformed XML at byte 3: malformed reference | unchanged |
| gap: bad hex charref + NUL | `<r>&#xZZ;&#0;</r>` | OK | malformed XML at byte 3: malformed character reference | unchanged |
| gap: bad hex charref | `<r>&#xZZ;</r>` | OK | malformed XML at byte 3: malformed character reference | unchanged |
| gap: NUL charref | `<r>&#0;</r>` | OK | malformed XML at byte 3: character reference to a character XML does not allow | unchanged |
| gap: decl inside root | `<r><?xml version="1.0"?></r>` | OK | malformed XML at byte 3: an XML declaration is allowed only at the start of the document | unchanged |
| gap: ]]> in text | `<r>]]></r>` | OK | malformed XML at byte 3: ']]>' is not allowed in character data | unchanged |
| gap: undeclared element prefix | `<p:r/>` | OK | malformed XML at byte 1: undeclared namespace prefix | unchanged |
| gap: undeclared attribute prefix | `<r p:a="1"/>` | OK | malformed XML at byte 3: undeclared namespace prefix | unchanged |
| C0 in text | `<r>\u{1}</r>` | OK | malformed XML at byte 3: character U+0001 is not allowed in XML | unchanged |
| NUL in text | `<r>\u{0}</r>` | OK | malformed XML at byte 3: character U+0000 is not allowed in XML | unchanged |
| VT in text | `<r>\u{b}</r>` | OK | malformed XML at byte 3: character U+000B is not allowed in XML | unchanged |
| C0 in attr | `<r a="\u{1}"/>` | OK | malformed XML at byte 6: character U+0001 is not allowed in XML | unchanged |
| C0 in comment | `<r><!--\u{1}--></r>` | OK | malformed XML at byte 7: character U+0001 is not allowed in XML | unchanged |
| C0 in PI | `<r><?p \u{1}?></r>` | OK | malformed XML at byte 7: character U+0001 is not allowed in XML | unchanged |
| C0 in CDATA | `<r><![CDATA[\u{1}]]></r>` | OK | malformed XML at byte 12: character U+0001 is not allowed in XML | unchanged |
| C0 outside root | `<r/>\u{1}` | malformed XML at byte 4: character data outside the document element | malformed XML at byte 4: character U+0001 is not allowed in XML | unchanged |
| U+FFFE text | `<r>\u{fffe}</r>` | OK | malformed XML at byte 3: character U+FFFE is not allowed in XML | unchanged |
| U+FFFF text | `<r>\u{ffff}</r>` | OK | malformed XML at byte 3: character U+FFFF is not allowed in XML | unchanged |
| U+FFFF attr | `<r a="\u{ffff}"/>` | OK | malformed XML at byte 6: character U+FFFF is not allowed in XML | unchanged |
| surrogate utf8 | `<r>\xed\xa0\x80</r>` | XML is not UTF-8 at byte 3 | XML is not UTF-8 at byte 3 | unchanged |
| lt in attr | `<r a="<"/>` | OK | malformed XML at byte 6: '<' is not allowed in an attribute value | unchanged |
| lt in attr 2 | `<r a='x<y'/>` | OK | malformed XML at byte 7: '<' is not allowed in an attribute value | unchanged |
| amp alone in attr | `<r a="&"/>` | OK | malformed XML at byte 6: unterminated reference in an attribute value | unchanged |
| amp in attr | `<r a="a&b"/>` | OK | malformed XML at byte 7: unterminated reference in an attribute value | unchanged |
| undeclared entity in attr | `<r a="&bogus;"/>` | OK | malformed XML at byte 6: reference to an undeclared entity | unchanged |
| NUL charref in attr | `<r a="&#0;"/>` | OK | malformed XML at byte 6: character reference to a character XML does not allow | unchanged |
| surrogate charref in attr | `<r a="&#xD800;"/>` | OK | malformed XML at byte 6: character reference to a character XML does not allow | unchanged |
| dup attr after ns | `<r xmlns:a="u" xmlns:b="u" a:x="1" b:x="2"/>` | OK | malformed XML at byte 35: two attributes have the same namespace name and local name | unchanged |
| dup attr exact | `<r a="1" a="2"/>` | malformed XML at byte 0: position 8: duplicated attribute, previous declaration at position 2 | malformed XML at byte 0: position 8: duplicated attribute, previous declaration at position 2 | unchanged |
| xmlns:xml other | `<r xmlns:xml="urn:other"/>` | OK | malformed XML at byte 3: the xml prefix must not be bound to another namespace name | unchanged |
| xmlns:xml correct | `<r xmlns:xml="http://www.w3.org/XML/1998/namespace"/>` | OK | OK | unchanged |
| xmlns:xmlns | `<r xmlns:xmlns="http://www.w3.org/2000/xmlns/"/>` | OK | malformed XML at byte 3: the xmlns prefix must not be declared | unchanged |
| xmlns:xmlns other | `<r xmlns:xmlns="urn:x"/>` | OK | malformed XML at byte 3: the xmlns prefix must not be declared | unchanged |
| p bound to xml ns | `<r xmlns:p="http://www.w3.org/XML/1998/namespace"/>` | OK | malformed XML at byte 3: only a reserved prefix may be bound to a reserved namespace name | unchanged |
| p bound to xmlns ns | `<r xmlns:p="http://www.w3.org/2000/xmlns/"/>` | OK | malformed XML at byte 3: only a reserved prefix may be bound to a reserved namespace name | unchanged |
| default bound to xml ns | `<r xmlns="http://www.w3.org/XML/1998/namespace"/>` | OK | malformed XML at byte 3: the default namespace must not be a reserved namespace name | unchanged |
| default bound to xmlns ns | `<r xmlns="http://www.w3.org/2000/xmlns/"/>` | OK | malformed XML at byte 3: the default namespace must not be a reserved namespace name | unchanged |
| prefix bound to empty | `<r xmlns:p=""/>` | OK | malformed XML at byte 3: a namespace prefix must not be undeclared | unchanged |
| default undeclared (legal) | `<r xmlns=""/>` | OK | OK | unchanged |
| element prefix xmlns | `<xmlns:r/>` | OK | malformed XML at byte 1: an element name must not use the xmlns prefix | unchanged |
| mismatched end | `<r><a></b></r>` | malformed XML at byte 10: ill-formed document: expected \`</a>\`, but \`</b>\` was found | malformed XML at byte 10: ill-formed document: expected \`</a>\`, but \`</b>\` was found | unchanged |
| mismatched case end | `<r></R>` | malformed XML at byte 7: ill-formed document: expected \`</r>\`, but \`</R>\` was found | malformed XML at byte 7: ill-formed document: expected \`</r>\`, but \`</R>\` was found | unchanged |
| comment -- | `<r><!-- a -- b --></r>` | OK | malformed XML at byte 10: '--' is not allowed in a comment | unchanged |
| comment ending - | `<r><!-- a ---></r>` | OK | malformed XML at byte 10: a comment must not end with '-' | unchanged |
| name digit start | `<1r/>` | OK | malformed XML at byte 1: invalid XML name | unchanged |
| attr name digit start | `<r 1a="1"/>` | OK | malformed XML at byte 3: invalid XML name | unchanged |
| name hyphen start | `<-r/>` | OK | malformed XML at byte 1: invalid XML name | unchanged |
| name with lt | `<r><a<b/></r>` | OK | malformed XML at byte 4: invalid XML name | unchanged |
| name with quote | `<r><a"b"/></r>` | OK | malformed XML at byte 4: invalid XML name | unchanged |
| space before name | `<r>< a/></r>` | malformed XML at byte 6: attribute name must be followed by '=' | malformed XML at byte 4: invalid XML name | unchanged |
| empty name with attr | `<r>< a="1"/></r>` | OK | malformed XML at byte 4: invalid XML name | unchanged |
| name trailing colon | `<r:/>` | OK | malformed XML at byte 1: name is not namespace-well-formed: a qualified name has one colon between two names | unchanged |
| name leading colon | `<:r/>` | OK | malformed XML at byte 1: name is not namespace-well-formed: a qualified name has one colon between two names | unchanged |
| name two colons | `<a:b:c xmlns:a="u"/>` | OK | malformed XML at byte 1: name is not namespace-well-formed: a qualified name has one colon between two names | unchanged |
| attr trailing colon | `<r a:="1"/>` | OK | malformed XML at byte 3: name is not namespace-well-formed: a qualified name has one colon between two names | unchanged |
| attr leading colon | `<r :a="1"/>` | OK | malformed XML at byte 3: name is not namespace-well-formed: a qualified name has one colon between two names | unchanged |
| attr two colons | `<r xmlns:a="u" a:b:c="1"/>` | OK | malformed XML at byte 15: name is not namespace-well-formed: a qualified name has one colon between two names | unchanged |
| PI xml-stylesheet (legal) | `<r><?xml-stylesheet href="x"?></r>` | OK | OK | unchanged |
| PI XML reserved | `<r><?XML x?></r>` | OK | malformed XML at byte 5: processing-instruction target 'xml' is reserved | unchanged |
| PI xMl reserved | `<r><?xMl?></r>` | OK | malformed XML at byte 5: processing-instruction target 'xml' is reserved | unchanged |
| PI empty target | `<r><? x?></r>` | OK | malformed XML at byte 5: invalid processing-instruction target | unchanged |
| PI digit target | `<r><?1x?></r>` | OK | malformed XML at byte 5: invalid processing-instruction target | unchanged |
| PI colon target | `<r><?a:b x?></r>` | OK | malformed XML at byte 5: invalid processing-instruction target | unchanged |
| PI empty <??> | `<??><r/>` | OK | malformed XML at byte 2: invalid processing-instruction target | unchanged |
| decl bare | `<?xml?><r/>` | OK | malformed XML at byte 2: XML declaration must declare a version | unchanged |
| decl no version | `<?xml encoding="UTF-8"?><r/>` | OK | malformed XML at byte 6: XML declaration must declare its version first | unchanged |
| decl version 2.0 | `<?xml version="2.0"?><r/>` | OK | malformed XML at byte 6: XML declaration version must be 1.x | unchanged |
| decl standalone maybe | `<?xml version="1.0" standalone="maybe"?><r/>` | OK | malformed XML at byte 20: XML declaration standalone must be 'yes' or 'no' | unchanged |
| decl wrong order | `<?xml standalone="yes" version="1.0"?><r/>` | OK | malformed XML at byte 6: XML declaration must declare its version first | unchanged |
| decl unknown attr | `<?xml version="1.0" foo="bar"?><r/>` | OK | malformed XML at byte 20: unexpected attribute in the XML declaration | unchanged |
| decl latin1 | `<?xml version="1.0" encoding="ISO-8859-1"?><r/>` | OK | malformed XML at byte 20: XML declaration names an encoding other than UTF-8 | unchanged |
| decl utf16 | `<?xml version="1.0" encoding="UTF-16"?><r/>` | OK | malformed XML at byte 20: XML declaration names an encoding other than UTF-8 | unchanged |
| decl utf-8 lower single (legal) | `<?xml version='1.0' encoding='utf-8'?><r/>` | OK | OK | unchanged |
| decl 1.1 | `<?xml version="1.1"?><r/>` | OK | OK | unchanged |
| decl after newline | `\u{a}<?xml version="1.0"?><r/>` | OK | malformed XML at byte 1: an XML declaration is allowed only at the start of the document | unchanged |
| decl after space | ` <?xml version="1.0"?><r/>` | OK | malformed XML at byte 1: an XML declaration is allowed only at the start of the document | unchanged |
| decl after comment | `<!--c--><?xml version="1.0"?><r/>` | OK | malformed XML at byte 8: an XML declaration is allowed only at the start of the document | unchanged |
| decl after root | `<r/><?xml version="1.0"?>` | OK | malformed XML at byte 4: an XML declaration is allowed only at the start of the document | unchanged |
| decl twice | `<?xml version="1.0"?><?xml version="1.0"?><r/>` | OK | malformed XML at byte 21: an XML declaration is allowed only at the start of the document | unchanged |
| decl after BOM (legal) | `\u{feff}<?xml version="1.0"?><r/>` | OK | OK | unchanged |
| charref > max | `<r>&#x110000;</r>` | OK | malformed XML at byte 3: character reference to a character XML does not allow | unchanged |
| charref surrogate | `<r>&#xD800;</r>` | OK | malformed XML at byte 3: character reference to a character XML does not allow | unchanged |
| charref FFFE | `<r>&#xFFFE;</r>` | OK | malformed XML at byte 3: character reference to a character XML does not allow | unchanged |
| charrefs legal ws | `<r>&#9;&#10;&#13;&#x20;</r>` | OK | OK | unchanged |
| charref overflow | `<r>&#99999999999999999999;</r>` | OK | malformed XML at byte 3: character reference to a character XML does not allow | unchanged |
| charref empty hex | `<r>&#x;</r>` | OK | malformed XML at byte 3: malformed character reference | unchanged |
| charref empty | `<r>&#;</r>` | OK | malformed XML at byte 3: malformed character reference | unchanged |
| charref negative | `<r>&#-1;</r>` | OK | malformed XML at byte 3: malformed character reference | unchanged |
| charref uppercase X | `<r>&#X41;</r>` | OK | malformed XML at byte 3: malformed character reference | unchanged |
| entity nbsp | `<r>&nbsp;</r>` | OK | malformed XML at byte 3: reference to an undeclared entity | unchanged |
| entity predefined (legal) | `<r>&amp;&lt;&gt;&apos;&quot;</r>` | OK | OK | unchanged |
| entity AMP | `<r>&AMP;</r>` | OK | malformed XML at byte 3: reference to an undeclared entity | unchanged |
| ref amp outside root | `<r/>&amp;` | malformed XML at byte 4: character data outside the document element | malformed XML at byte 4: character data outside the document element | unchanged |
| charref space before root | `&#32;<r/>` | malformed XML at byte 0: character data outside the document element | malformed XML at byte 0: character data outside the document element | unchanged |
| charref space after root | `<r/>&#32;` | malformed XML at byte 4: character data outside the document element | malformed XML at byte 4: character data outside the document element | unchanged |
| lt in text | `<r>a < b</r>` | malformed XML at byte 11: attribute name must be followed by '=' | malformed XML at byte 6: invalid XML name | unchanged |
| empty element | `<r><></></r>` | OK | malformed XML at byte 4: invalid XML name | unchanged |
| empty end | `<r><a></></r>` | malformed XML at byte 9: ill-formed document: expected \`</a>\`, but \`</>\` was found | malformed XML at byte 9: ill-formed document: expected \`</a>\`, but \`</>\` was found | unchanged |
| doctype lower | `<!doctype r><r/>` | DTD and DOCTYPE are not allowed at byte 0 | DTD and DOCTYPE are not allowed at byte 0 | unchanged |
| ELEMENT decl | `<r><!ELEMENT x ANY></r>` | malformed XML at byte 3: syntax error: unknown or missed symbol in markup | malformed XML at byte 3: syntax error: unknown or missed symbol in markup | unchanged |
| ]]&gt; (legal) | `<r>]]&gt;</r>` | OK | OK | unchanged |
| attr ws values (legal) | `<r a="\u{9}\u{a}"/>` | OK | OK | unchanged |
| non-ascii names (legal) | `<é é="1"/>` | OK | OK | unchanged |
| name start times | `<×/>` | OK | malformed XML at byte 1: invalid XML name | unchanged |
| name middle dot (legal) | `<a·/>` | OK | OK | unchanged |
| name combining start | `<̀a/>` | OK | malformed XML at byte 1: invalid XML name | unchanged |
| unbound in sibling scope | `<r><a xmlns:p="u"/><p:b/></r>` | OK | malformed XML at byte 20: undeclared namespace prefix | unchanged |
| self-declared prefix (legal) | `<p:r xmlns:p="u"/>` | OK | OK | unchanged |
| attr prefix declared later (legal) | `<r p:a="1" xmlns:p="u"/>` | OK | OK | unchanged |
| xml:lang (legal) | `<r xml:lang="en"/>` | OK | OK | unchanged |
| xml:foo (legal) | `<r xml:foo="1"/>` | OK | OK | unchanged |
| xmlns attr prefix on element | `<r xmlns:="u"/>` | OK | malformed XML at byte 3: name is not namespace-well-formed: a qualified name has one colon between two names | unchanged |
| tab before root decl ok? | `<?xml version="1.0"?>\u{9}<r/>` | OK | OK | unchanged |
| decl ws before ?> (legal) | `<?xml version="1.0" ?><r/>` | OK | OK | unchanged |
| decl no space before version | `<?xmlversion="1.0"?><r/>` | OK | malformed XML at byte 2: invalid processing-instruction target | unchanged |
| decl encoding bad name | `<?xml version="1.0" encoding="8bit"?><r/>` | OK | malformed XML at byte 20: invalid encoding name in the XML declaration | unchanged |
| comment <!-->  | `<r><!--></r>` | malformed XML at byte 12: syntax error: comment not closed: \`-->\` not found before end of input | malformed XML at byte 12: syntax error: comment not closed: \`-->\` not found before end of input | unchanged |
| comment <!--->  | `<r><!---></r>` | malformed XML at byte 13: syntax error: comment not closed: \`-->\` not found before end of input | malformed XML at byte 13: syntax error: comment not closed: \`-->\` not found before end of input | unchanged |
| comment empty (legal) | `<r><!----></r>` | OK | OK | unchanged |
| comment <!----->  | `<r><!-----></r>` | OK | malformed XML at byte 7: a comment must not end with '-' | unchanged |
| comment single dash (legal) | `<r><!-- a-b - c --></r>` | OK | OK | unchanged |
| PI no content (legal) | `<r><?pi?></r>` | OK | OK | unchanged |
| PI empty content (legal) | `<r><?pi ?></r>` | OK | OK | unchanged |
| PI xmlfoo target (legal) | `<r><?xmlfoo x?></r>` | OK | OK | unchanged |
| shadowing (legal) | `<r xmlns:p="u"><p:a xmlns:p="v"><p:b/></p:a><p:c/></r>` | OK | OK | unchanged |
| scope popped | `<r><a xmlns:p="u"><p:b/></a><p:c/></r>` | OK | malformed XML at byte 29: undeclared namespace prefix | unchanged |
| empty element scope popped | `<r><a xmlns:p="u"/><p:c/></r>` | OK | malformed XML at byte 20: undeclared namespace prefix | unchanged |
| alias distinct locals (legal) | `<r xmlns:a="u" xmlns:b="u" a:x="1" b:y="2"/>` | OK | OK | unchanged |
| alias three prefixes dup | `<r xmlns:a="u" xmlns:b="v" xmlns:c="u" a:x="1" b:x="2" c:x="3"/>` | OK | malformed XML at byte 55: two attributes have the same namespace name and local name | unchanged |
| xml attr vs other (legal) | `<r xmlns:p="u" xml:lang="en" p:lang="x"/>` | OK | OK | unchanged |
| uri via charref to xml ns | `<r xmlns:p="http://www.w3.org/XML/1998/&#110;amespace"/>` | OK | malformed XML at byte 3: only a reserved prefix may be bound to a reserved namespace name | unchanged |
| uri whitespace (legal) | `<r xmlns:p=" "><p:a/></r>` | OK | OK | unchanged |
| xmlns as local name (legal) | `<r xmlns:p="u" p:xmlns="1"/>` | OK | OK | unchanged |
| xml element prefix (legal) | `<xml:r/>` | OK | OK | unchanged |
| CDATA ]] edge (legal) | `<r><![CDATA[]]]]><![CDATA[>]]></r>` | OK | OK | unchanged |
| text ]] then > via ref (legal) | `<r>]]&gt;]]</r>` | OK | OK | unchanged |
| text ] > (legal) | `<r>] ]> ]]x></r>` | OK | OK | unchanged |
| attr gt and apos (legal) | `<r a="'>" b='">'/>` | OK | OK | unchanged |
| decl encoding UTF8 no dash | `<?xml version="1.0" encoding="UTF8"?><r/>` | OK | malformed XML at byte 20: XML declaration names an encoding other than UTF-8 | unchanged |
| decl version 1. | `<?xml version="1."?><r/>` | OK | malformed XML at byte 6: XML declaration version must be 1.x | unchanged |
| decl version 1.0 sp | `<?xml version="1.0 "?><r/>` | OK | malformed XML at byte 6: XML declaration version must be 1.x | unchanged |
| decl empty encoding | `<?xml version="1.0" encoding=""?><r/>` | OK | malformed XML at byte 20: invalid encoding name in the XML declaration | unchanged |
| decl tab separators (legal) | `<?xml\u{9}version="1.0"\u{9}encoding="UTF-8"\u{9}?><r/>` | OK | OK | unchanged |
| decl eq spaces (legal) | `<?xml version = "1.0" ?><r/>` | OK | OK | unchanged |
| nel and lsep (legal) | `<r>\u{85}\u{2028}</r>` | OK | OK | unchanged |
| DEL (legal) | `<r>\u{7f}</r>` | OK | OK | unchanged |
| private use FFFD (legal) | `<r>�</r>` | OK | OK | unchanged |
| decl standalone no (legal) | `<?xml version="1.0" encoding="UTF-8" standalone="no"?><r/>` | OK | OK | unchanged |
