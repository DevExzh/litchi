## 3.2 Content Parts and Ink

![Sample ink shape](assets/image40.png)

*Figure 31: Sample ink shape*

The following code example represents the Ink content part that describes the example ink shown in the preceding figure. The **brush** definitions specify its size, shape, color, and coordinate space. Traces are gathered into a **traceGroup** hierarchy of **writingRegion**, **paragraph**, **line**, and **inkWord**, with the **inkWord** containing the results of ink recognition that was generated at runtime.

2736. \<?xml version="1.0" encoding="UTF-8" standalone="yes"?\>

      \<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"\>

      \<inkml:definitions\>

      \<inkml:context xml:id="ctx0"\>

      \<inkml:inkSource xml:id="inkSrc0"\>

      \<inkml:traceFormat\>

      \<inkml:channel name="X" type="integer" max="9600" units="cm"/\>

      \<inkml:channel name="Y" type="integer" max="7200" units="cm"/\>

      \<inkml:channel name="F" type="integer" max="256" units="dev"/\>

      \</inkml:traceFormat\>

      \<inkml:channelProperties\>

      \<inkml:channelProperty channel="X" name="resolution" value="367.8161" units="1/cm"/\>

      \<inkml:channelProperty channel="Y" name="resolution" value="440.36697" units="1/cm"/\>

      \<inkml:channelProperty channel="F" name="resolution" value="0" units="1/dev"/\>

      \</inkml:channelProperties\>

      \</inkml:inkSource\>

      \<inkml:timestamp xml:id="ts0" timeString="2010-01-10T13:10:12.822"/\>

      \</inkml:context\>

      \<inkml:brush xml:id="br0"\>

      \<inkml:brushProperty name="width" value="0.06667" units="cm"/\>

      \<inkml:brushProperty name="height" value="0.06667" units="cm"/\>

      \<inkml:brushProperty name="color" value="#FF0000"/\>

      \<inkml:brushProperty name="fitToCurve" value="1"/\>

      \</inkml:brush\>

      \</inkml:definitions\>

      \<inkml:traceGroup\>

      \<inkml:annotationXML\>

      \<emma:emma xmlns:emma="http://www.w3.org/2003/04/emma" version="1.0"\>

      \<emma:interpretation id="{8646EB18-6E67-4FFA-8739-E20C3C1A0F80}" emma:medium="tactile" emma:mode="ink"\>

      \<msink:context xmlns:msink="http://schemas.microsoft.com/ink/2010/main" type="writingRegion" rotatedBoundingBox="4334,2653 13237,2575 13262,5384 4359,5462"/\>

      \</emma:interpretation\>

      \</emma:emma\>

      \</inkml:annotationXML\>

      \<inkml:traceGroup\>

      \<inkml:annotationXML\>

      \<emma:emma xmlns:emma="http://www.w3.org/2003/04/emma" version="1.0"\>

      \<emma:interpretation id="{4A0797F9-1386-486A-B2BA-709708E24147}" emma:medium="tactile" emma:mode="ink"\>

      \<msink:context xmlns:msink="http://schemas.microsoft.com/ink/2010/main" type="paragraph" rotatedBoundingBox="4334,2653 13237,2575 13262,5384 4359,5462" alignmentLevel="1"/\>

      \</emma:interpretation\>

      \</emma:emma\>

      \</inkml:annotationXML\>

      \<inkml:traceGroup\>

      \<inkml:annotationXML\>

      \<emma:emma xmlns:emma="http://www.w3.org/2003/04/emma" version="1.0"\>

      \<emma:interpretation id="{0430FEDC-2E4E-4A32-91D8-87424D412245}" emma:medium="tactile" emma:mode="ink"\>

      \<msink:context xmlns:msink="http://schemas.microsoft.com/ink/2010/main" type="line" rotatedBoundingBox="4334,2653 13237,2575 13262,5384 4359,5462"/\>

      \</emma:interpretation\>

      \</emma:emma\>

      \</inkml:annotationXML\>

      \<inkml:traceGroup\>

      \<inkml:annotationXML\>

      \<emma:emma xmlns:emma="http://www.w3.org/2003/04/emma" version="1.0"\>

      \<emma:interpretation id="{583A8050-3462-4BC6-915B-A797D921D61F}" emma:medium="tactile" emma:mode="ink"\>

      \<msink:context xmlns:msink="http://schemas.microsoft.com/ink/2010/main" type="inkWord" rotatedBoundingBox="4334,2653 13237,2575 13262,5384 4359,5462"/\>

      \</emma:interpretation\>

      \<emma:one-of disjunction-type="recognition" id="oneOf0"\>

      \<emma:interpretation id="interp0" emma:lang="en-US" emma:confidence="1"\>

      \<emma:literal\>hello\</emma:literal\>

      \</emma:interpretation\>

      \<emma:interpretation id="interp1" emma:lang="en-US" emma:confidence="0.5"\>

      \<emma:literal\>hello \</emma:literal\>

      \</emma:interpretation\>

      \<emma:interpretation id="interp2" emma:lang="en-US" emma:confidence="0.5"\>

      \<emma:literal\> hello\</emma:literal\>

      \</emma:interpretation\>

      \<emma:interpretation id="interp3" emma:lang="en-US" emma:confidence="0.5"\>

      \<emma:literal\>he 11 o\</emma:literal\>

      \</emma:interpretation\>

      \<emma:interpretation id="interp4" emma:lang="en-US" emma:confidence="0.5"\>

      \<emma:literal\>he 110\</emma:literal\>

      \</emma:interpretation\>

      \</emma:one-of\>

      \</emma:emma\>

      \</inkml:annotationXML\>

      \<inkml:trace contextRef="#ctx0" brushRef="#br0"\>-2 0 43,'0'0'14,"0"0"-6,0 0 1,8 23 2,-16 12-3,16 7 3,17 18-4,9 28-7,-9 0 4,8 14 3,-8 10 3,25 18-2,-17 18-1,1-18-1,6 18-2,18 24 0,-16 6-1,-9-1 0,-8 1 0,0-19-1,0-17 0,0-5 0,-8-8-1,0-16 0,8-30 0,-17 5 0,-8-34-1,0-12 1,0-25-1,8 13 0,-8-36 0,9-30 0,-1-17 1,-16-6 0,8-24 1,8 6-1,-16 6 0,16-1 0,9-5-1,-1-17 1,9 0 0,16 4-1,17 2 2,17 5 1,8 12-3,9 12 1,-17 18 0,7 5 0,1 12 1,9 24-1,-17 19 0,0 15 0,-10 32 0,2 22 0,-25 13 0,-1 0 0,1-6 0,-9-13 0,9 1-1,-17-6-1,0-12-7,-8 7-10\</inkml:trace\>

      \<inkml:trace contextRef="#ctx0" brushRef="#br0"\>2375 1858 106,'0'0'0,"0"0"4,0 0 4,48 11-1,-15 2 0,26 4-1,-1 7-1,17 5 0,9-11-1,7 5-1,-9-4-1,10 4 0,-17-11-2,0-12 0,0-18-2,-10-11 0,2-13 0,-17-11-1,8 0 1,-25 6 0,1-19 1,-26 14 0,-8-2 1,-25 1-2,0 6 0,-17 5 1,1 18 0,-1-6 0,-8 30 1,8 6-1,-23 24 1,7 6 0,0 16 0,-9 2 0,-17 18 0,1 3 2,1 15 2,-1 10 1,33-5 0,17-12 1,7 0-1,10 0-2,32-1 0,35 1 1,24-17 1,25-1 0,23-12-2,-6-17-1,25-14-3,23-9-4,1-31-6,25-35-9\</inkml:trace\>

      \<inkml:trace contextRef="#ctx0" brushRef="#br0"\>4591 295 72,'0'0'4,"0"0"5,0 0-1,0 0 0,-8 6 0,8-6 1,8 24 1,1 23-1,7 7 0,9 22 0,-16 19 0,24 24 0,0 11-2,1-6-1,16 24-1,0 12-1,8-18-1,-1 12-1,1-13 0,-8 2 0,0-19 0,0-36-1,0 2-2,-8-7-5,0-19-7,-9-4-12\</inkml:trace\>

      \<inkml:trace contextRef="#ctx0" brushRef="#br0"\>5730 212 81,'0'0'1,"0"0"8,0 0 4,9 42-1,7 18-1,1 10 1,0 19 0,-1 30-1,16 16-2,-7 8 0,25-2-1,0 19-2,9-6-1,7 0-1,9-6-1,7 0-1,-7-7-1,-8-28-2,0-12-5,-9-18-6,9-18-14\</inkml:trace\>

      \<inkml:trace contextRef="#ctx0" brushRef="#br0"\>8281 1189 93,'0'0'3,"0"0"5,0 0 2,-25-6-1,0-6-1,25 19-2,-34-7-1,9-19-1,8 31-1,-8-24 0,-41 12-1,16 12 0,-9 12-1,-6 11 1,15 13 1,0 6 0,-8 4 3,-1 8 0,9 11-1,-8 0 0,16 5-1,11 7 0,14 5-1,17 7 0,25-7 0,7-4-1,26 4 0,9 13-2,25-13 0,24-11 0,-1-12 0,2-18 0,0-12 0,-25-11-1,-19-30 0,10-23 0,9-31-1,0-28 0,-17-18 1,-2-13-1,-6-18 1,-34-4 1,-8 5 0,-25 4 2,-16 8 0,-18 24 0,-24-7 1,1 12 0,-18 12 0,0 24-1,0 6-1,-17 12-3,-6 23-5,-36 12-16\</inkml:trace\>

      \</inkml:traceGroup\>

      \</inkml:traceGroup\>

      \</inkml:traceGroup\>

      \</inkml:traceGroup\>

      \</inkml:ink\>

The following example shows a Document Part reference to the previous content part as used in a **PresentationML** slide.

2820. \<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"\>

      \<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"\>

      \<p:contentPart p14:bwMode="auto" r:id="rId2"\>

      \<p14:nvContentPartPr\>

      \<p14:cNvPr id="8" name="Ink 7"/\>

      \<p14:cNvContentPartPr/\>

      \<p14:nvPr/\>

      \</p14:nvContentPartPr\>

      \<p14:xfrm\>

      \<a:off x="1561526" y="971040"/\>

      \<a:ext cx="3210480" cy="1010160"/\>

      \</p14:xfrm\>

      \</p:contentPart\>

      \</mc:Choice\>

      \<mc:Fallback\>

      \<p:pic\>

      \<p:nvPicPr\>

      \<p:cNvPr id="8" name="Ink 7"/\>

      \<p:cNvPicPr/\>

      \<p:nvPr/\>

      \</p:nvPicPr\>

      \<p:blipFill\>

      \<a:blip r:embed="rId3"/\>

      \<a:stretch\>

      \<a:fillRect/\>

      \</a:stretch\>

      \</p:blipFill\>

      \<p:spPr\>

      \<a:xfrm\>

      \<a:off x="1561526" y="971040"/\>

      \<a:ext cx="3210480" cy="1010160"/\>

      \</a:xfrm\>

      \<a:prstGeom prst="rect"\>

      \<a:avLst/\>

      \</a:prstGeom\>

      \</p:spPr\>

      \</p:pic\>

      \</mc:Fallback\>

      \</mc:AlternateContent\>

      ...
