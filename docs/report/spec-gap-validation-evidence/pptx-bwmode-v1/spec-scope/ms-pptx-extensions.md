## 2.2 Extensions

This section specifies the elements from [\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) that are extended by this format. Either the **Ignorable** attribute ([\[ISO/IEC29500-3:2015\]](https://go.microsoft.com/fwlink/?linkid=861154) section 7.2), **AlternateContent** element (\[ISO/IEC29500-3:2015\] section 7.5), or the **extLst** element (\[ISO/IEC29500-1:2016\] section 19.2.1.12) MUST be used to maintain compatibility with \[ISO/IEC29500-1:2016\] implementations.

### 2.2.1 Slide Transition Extensions

The **sld** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.38), the **sldLayout** element (\[ISO/IEC29500-1:2016\] section 19.3.1.39), and the **sldMaster** element (\[ISO/IEC29500-1:2016\] section 19.3.1.42) are extended by the addition of an **AlternateContent** child element ([\[ISO/IEC29500-3:2015\]](https://go.microsoft.com/fwlink/?linkid=861154) section 7.5), whose structure is specified in the following table.

<table>
<colgroup>
<col style="width: 56%" />
<col style="width: 43%" />
</colgroup>
<thead>
<tr>
<th>AlternateContent components</th>
<th>Child element</th>
</tr>
</thead>
<tbody>
<tr>
<td>Choice: http://schemas.microsoft.com/office/powerpoint/2010/main</td>
<td><strong>transition</strong> ([ISO/IEC29500-1:2016] section 19.3.1.50)</td>
</tr>
<tr>
<td><p>Choice:</p>
<p>http://schemas.microsoft.com/office/powerpoint/2012/main</p></td>
<td><strong>transition</strong> ([ISO/IEC29500-1:2016] section 19.3.1.50)</td>
</tr>
<tr>
<td><p>Choice:</p>
<p>http://schemas.microsoft.com/office/powerpoint/2015/09/main</p></td>
<td><strong>transition</strong> ([ISO/IEC29500-1:2016] section 19.3.1.50)</td>
</tr>
<tr>
<td>Fallback</td>
<td><strong>transition</strong> ([ISO/IEC29500-1:2016] section 19.3.1.50)</td>
</tr>
</tbody>
</table>

The **transition** element (\[ISO/IEC29500-1:2016\] section 19.3.1.50) is extended by the addition of the following child elements to the **xsd:choice** content model of the **CT_SlideTransition** complex type (\[ISO/IEC29500-1:2016\] section A.3):

- **vortex** (section 2.3.1.30)

- **switch** (section 2.3.1.29)

- **flip** (section 2.3.1.11)

- **ripple** (section 2.3.1.24)

- **honeycomb** (section 2.3.1.15)

- **prism** (section 2.3.1.22)

- **doors** (section 2.3.1.7)

- **window** (section 2.3.1.33)

- **ferris** (section 2.3.1.9)

- **gallery** (section 2.3.1.13)

- **conveyor** (section 2.3.1.3)

- **pan** (section 2.3.1.21)

- **glitter** (section 2.3.1.14)

- **warp** (section 2.3.1.31)

- **flythrough** (section 2.3.1.12)

- **flash** (section 2.3.1.10)

- **shred** (section 2.3.1.28)

- **reveal** (section 2.3.1.23)

- **wheelReverse** (section 2.3.1.32)

- **morph** (section 2.6.1.1)

- **prstTrans** (section 2.4.1.5)

The **transition** element (\[ISO/IEC29500-1:2016\] section 19.3.1.50) is further extended by the addition of the following attribute to the **CT_SlideTransition** complex type (\[ISO/IEC29500-1:2016\] section A.3): **dur** (section 2.3.2.3).

### 2.2.2 Animation Info Extensions

The **sld** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.38), the **sldLayout** element (\[ISO/IEC29500-1:2016\] section 19.3.1.39), and the **sldMaster** element (\[ISO/IEC29500-1:2016\] section 19.3.1.42) are extended by the addition of an **AlternateContent** child element ([\[ISO/IEC29500-3:2015\]](https://go.microsoft.com/fwlink/?linkid=861154) section 7.5), whose structure is specified in the following table.

| AlternateContent components | Child element |
|----|----|
| Choice: http://schemas.microsoft.com/office/powerpoint/2010/main | **timing** (\[ISO/IEC29500-1:2016\] section 19.3.1.48) |
| Fallback | **timing** (\[ISO/IEC29500-1:2016\] section 19.3.1.48) |

The **tgtEl** descendant element (\[ISO/IEC29500-1:2016\] section 19.5.81) of the **timing** element is extended by the addition of the following child elements to the **xsd:choice** content model of the **CT_TLTimeTargetElement** complex type (\[ISO/IEC29500-1:2016\] section A.3): **bmkTgt** (section 2.3.1.1).

The **cTn** descendant element (\[ISO/IEC29500-1:2016\] section 19.5.33) of the **timing** element is extended by the addition of the following attribute to the **CT_TLCommonTimeNodeData** complex type (\[ISO/IEC29500-1:2016\] section A.3): **presetBounceEnd** (section 2.3.2.4).

The **anim** descendant element (\[ISO/IEC29500-1:2016\] section 19.5.1) of the **timing** element is extended by the addition of the following attribute to the **CT_TLAnimateBehavior** complex type (\[ISO/IEC29500-1:2016\] section A.3): **bounceEnd** (section 2.3.2.1).

The **animMotion** descendant element (\[ISO/IEC29500-1:2016\] section 19.5.4) of the **timing** element is extended by the addition of the following attribute to the **CT_TLAnimateMotionBehavior** complex type (\[ISO/IEC29500-1:2016\] section A.3): **bounceEnd** (section 2.3.2.1).

The **animRot** descendant element (\[ISO/IEC29500-1:2016\] section 19.5.5) of the **timing** element is extended by the addition of the following attribute to the **CT_TLAnimateRotationBehavior** complex type (\[ISO/IEC29500-1:2016\] section A.3): **bounceEnd** (section 2.3.2.1).

The **animScale** descendant element (\[ISO/IEC29500-1:2016\] section 19.5.6) of the **timing** element is extended by the addition of the following attribute to the **CT_TLAnimateScaleBehavior** complex type (\[ISO/IEC29500-1:2016\] section A.3): **bounceEnd** (section 2.3.2.1).

### 2.2.3 Content Part Extensions

The **grpSp** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.22) is extended by the addition of an **AlternateContent** child element ([\[ISO/IEC29500-3:2015\]](https://go.microsoft.com/fwlink/?linkid=861154) section 7.5), whose structure is specified in the following table.

| AlternateContent components | Child element |
|----|----|
| Choice: http://schemas.microsoft.com/office/powerpoint/2010/main | **contentPart** (\[ISO/IEC29500-1:2016\] section 19.3.1.14) |
| Fallback | **sp** (\[ISO/IEC29500-1:2016\] section 19.3.1.43) |

The **contentPart** element (\[ISO/IEC29500-1:2016\] section 19.3.1.14) is extended by the addition of the following child elements to a new **xsd:sequence** content model of the **CT_Rel** complex type (\[ISO/IEC29500-1:2016\] section A.3):

- **nvContentPartPr** (section 2.3.1.20)

- **xfrm** (section 2.3.1.34)

- **extLst** (section 2.3.1.8).

The **contentPart** element (\[ISO/IEC29500-1:2016\] section 19.3.1.14) is further extended by the addition of the following attribute to the **CT_Rel** complex type (\[ISO/IEC29500-1:2016\] section A.3): **bwMode** (section 2.3.2.2).

#### 2.2.3.1 Ink Extensions

The spTree element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.45) and the grpSp element (\[ISO/IEC29500-1:2016\] section 19.3.1.22) are extended by the addition of an AlternateContent child element\<1\> whose structure is specified in the following table.

<table>
<colgroup>
<col style="width: 57%" />
<col style="width: 42%" />
</colgroup>
<thead>
<tr>
<th>AlternateContent components</th>
<th>Child element</th>
</tr>
</thead>
<tbody>
<tr>
<td><p>Choice:</p>
<p>http://schemas.microsoft.com/office/powerpoint/2010/main</p>
<p>http://schemas.microsoft.com/office/powerpoint/2014/inkAction (<a href="%5bMS-ODRAWXML%5d.pdf#Section_06cff208c6e14db7bb68665135e5f0de">[MS-ODRAWXML]</a> section 2.21)</p></td>
<td><p>contentPart ([ISO/IEC29500-1:2016]</p>
<p>section 19.3.1.14)</p></td>
</tr>
<tr>
<td>Fallback</td>
<td><p>pic ([ISO/IEC29500-1:2016] section</p>
<p>19.3.1.37)</p></td>
</tr>
</tbody>
</table>

### 2.2.4 Media Extensions

The **extLst** child element of the **nvPr** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.33) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.

| Ext uri | Child element |
|----|----|
| {DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230} | **media** (section 2.3.1.18) |

The **extLst** child element of the **showPr** element (\[ISO/IEC29500-1:2016\] section 19.2.1.30) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.

| Ext uri | Child element |
|----|----|
| {2FDB2607-1784-4EEB-B798-7EB5836EED8A} | **showMediaCtrls** (section 2.3.1.27) |

#### 2.2.4.1 TracksInfo Extensions

The **extLst** child element of the **media** element (section 2.3.1.18) is extended by the addition of a new **ext** child element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.2.1.11), whose structure is specified in the following table:

| Ext uri | Child element |
|----|----|
| {3AFAAA56-56D3-431D-BCD4-E75A35582382} | **tracksInfo** (section 2.13.1.1) |

### 2.2.5 Section Extensions

The **extLst** child element of the **presentation** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.2.1.26) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table:

| Ext uri | Child element |
|----|----|
| {521415D9-36F7-43E2-AB2F-B90AF26B5E84} | **sectionLst** (section 2.3.1.25) |

### 2.2.6 Slide Show Extensions

The **extLst** child element of the **showPr** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.2.1.30) is extended by the addition of new **ext** child elements (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structures are specified in the following table.

| Ext uri | Child element |
|----|----|
| {F99C55AA-B7CB-42B0-86F8-08522FDF87E8} | **browseMode** (section 2.3.1.2) |
| {EC167BDD-8182-4AB7-AECC-EB403E3ABB37} | **laserClr** (section 2.3.1.16) |

The **extLst** child element of the **sld** element (\[ISO/IEC29500-1:2016\] section 19.3.1.38) is extended by the addition of new **ext** child elements (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structures are specified in the following table.

| Ext uri | Child element |
|----|----|
| {3A86A75C-4F4B-4683-9AE1-C65F6400EC91} | **laserTraceLst** (section 2.3.1.17) |
| {E180D4A7-C9FB-4DFB-919C-405C955672EB} | **showEvtLst** (section 2.3.1.26) |

### 2.2.7 Image Extensions

The **extLst** child element of the **presentationPr** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.2.1.27) is extended by the addition of new **ext** child elements (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structures are specified in the following table.

| Ext uri | Child element |
|----|----|
| {E76CE94A-603C-4142-B9EB-6D1370010A27} | **discardImageEditData** (section 2.3.1.6) |
| {D31A062A-798A-4329-ABDD-BBA856620510} | **defaultImageDpi** (section 2.3.1.5) |

### 2.2.8 Math Extensions

The **extLst** child element of the **presentationPr** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.2.1.27) is extended by the addition of new **ext** child elements (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structures are specified in the following table.

| Ext uri | Child element |
|----|----|
| {4599F94E-CEE6-441E-89CC-EB005ECD8F06} | **a14:m** ([\[MS-ODRAWXML\]](%5bMS-ODRAWXML%5d.pdf#Section_06cff208c6e14db7bb68665135e5f0de) section 2.3.1.11) |

### 2.2.9 Change Tracking Extensions

The **extLst** child element of the **nvPr** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.33) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.

|  Ext uri | Child element |
|----|----|
| {D42A27DB-BD31-4B8C-83A1-F6EECF244321} | **modId** (section 2.3.1.19) |

The **extLst** child element of the **cSld** element (\[ISO/IEC29500-1:2016\] section 19.3.1.16) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.

|  Ext uri | Child element |
|----|----|
| {BB962C8B-B14F-4D97-AF65-F5344CB8AC3E} | **creationId** (section 2.3.1.4) |

### 2.2.10 Comment Extensions

The **extLst** child element of the **cmAuthor** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.4.2) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.\<2\>

|  Ext uri | Child element |
|----|----|
| {19B8F6BF-5375-455C-9EA6-DF929625EA0E} | **presenceInfo** (section 2.4.1.4) |

The **extLst** child element of the **cm** element (\[ISO/IEC29500-1:2016\] section 19.4.1) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.\<3\>

|  Ext uri | Child element |
|----|----|
| {C676402C-5697-4E1C-873F-D02D1690AC5C} | **threadingInfo** (section 2.4.1.7) |

The **extLst** child element of the **sld** element (\[ISO/IEC29500-1:2016\] section 19.3.1.38) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.\<4\>

|  Ext uri | Child element |
|----|----|
| {6950BFC3-D8DA-4A85-94F7-54DA5524770B} | **commentRel** (section 2.16.1.3) |

### 2.2.11 Guide Extensions

The **extLst** child element of the **presentation** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.2.1.26) is extended by the addition of new **ext** child elements (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structures are specified in the following table.\<5\>

| Ext uri | Child element |
|----|----|
| {EFAFB233-063F-42B5-8137-9DF3F51BA10A} | **sldGuideLst** (section 2.4.1.6) |
| {2D200454-40CA-4A62-9FC3-DE9A4176ACB9} | **notesGuideLst** (section 2.4.1.3) |

The **extLst** child element of the **sldLayout** element (\[ISO/IEC29500-1:2016\] section 19.3.1.39) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.\<6\>

| Ext uri                                | Child element                     |
|----------------------------------------|-----------------------------------|
| {DCECCB84-F9BA-43D5-87BE-67443E8EF086} | **sldGuideLst** (section 2.4.1.6) |

The **extLst** child element of the **sldMaster** element (\[ISO/IEC29500-1:2016\] section 19.3.1.42) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.\<7\>

| Ext uri                                | Child element                     |
|----------------------------------------|-----------------------------------|
| {27BBF7A9-308A-43DC-89C8-2F10F3537804} | **sldGuideLst** (section 2.4.1.6) |

The **extLst** child element of the **handoutMaster** element (\[ISO/IEC29500-1:2016\] section 19.3.1.24) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.\<8\>

| Ext uri                                | Child element                     |
|----------------------------------------|-----------------------------------|
| {56416CCD-93CA-4268-BC5B-53C4BB910035} | **sldGuideLst** (section 2.4.1.6) |

The **extLst** child element of the **notesMaster** element (\[ISO/IEC29500-1:2016\] section 19.3.1.27) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.\<9\>

| Ext uri                                | Child element                     |
|----------------------------------------|-----------------------------------|
| {620B2872-D7B9-4A21-9093-7833F8D536E1} | **sldGuideLst** (section 2.4.1.6) |

### 2.2.12 Charting Extensions

The **extLst** child element of the **presentationPr** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.2.1.27) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.\<10\>

| Ext uri | Child element |
|----|----|
| {FD5EFAAD-0ECE-453E-9831-46B23BE46B34} | **chartTrackingRefBased** (section 2.4.1.1) |

### 2.2.13 Office App Extensions

The **spTree** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.45) and the **grpSp** element (\[ISO/IEC29500-1:2016\] section 19.3.1.22) are extended by the addition of an **AlternateContent** child element whose structure is specified in the following table.

<table>
<colgroup>
<col style="width: 67%" />
<col style="width: 32%" />
</colgroup>
<thead>
<tr>
<th>AlternateContent components</th>
<th>Child element</th>
</tr>
</thead>
<tbody>
<tr>
<td><p>Choice:<br />
http://schemas.microsoft.com/office/webextensions/webextension/2010/11</p>
<p>http://schemas.microsoft.com/office/powerpoint/2013/contentapp</p></td>
<td><strong>webextensionref (</strong><a href="%5bMS-OWEXML%5d.pdf#Section_a2cd741a4cca4b1aade4b2c443972afa">[MS-OWEXML]</a> section 2.1.3)</td>
</tr>
<tr>
<td>Fallback</td>
<td><strong>pic</strong> ([ISO/IEC29500-1:2016] section 19.3.1.37)</td>
</tr>
</tbody>
</table>

### 2.2.14 Narration Extensions

The **extLst** child element of the **nvPr** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.33) is extended by the addition of a new **ext** child element\<11\> [(](https://go.microsoft.com/fwlink/?LinkId=325242)\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.

<table style="width:95%;">
<colgroup>
<col style="width: 50%" />
<col style="width: 44%" />
</colgroup>
<thead>
<tr>
<th>Ext uri</th>
<th><blockquote>
<p>Child element</p>
</blockquote></th>
</tr>
</thead>
<tbody>
<tr>
<td>{42D2F446-02D8-4167-A562-619A0277C38B}</td>
<td><blockquote>
<p><strong>isNarration</strong> (section <a href="#Section_4ae507ab2b3d41b094b2f59a1626023e">2.4.1.2</a>)</p>
</blockquote></td>
</tr>
</tbody>
</table>

### 2.2.15 Zoom Extensions

The **spTree** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.45) and the **grpSp** element (\[ISO/IEC29500-1:2016\] section 19.3.1.22) are extended by the addition of an **AlternateContent** child element whose structure is specified in the following tables.

<table style="width:100%;">
<colgroup>
<col style="width: 60%" />
<col style="width: 39%" />
</colgroup>
<thead>
<tr>
<th>AlternateContent components</th>
<th>Child element</th>
</tr>
</thead>
<tbody>
<tr>
<td>Choice:<br />
http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom</td>
<td><strong>sectionZm</strong> (section <a href="#Section_40214a8b104443fca9d53f8635ea5bc2">2.9.1.1</a>)</td>
</tr>
<tr>
<td>Fallback</td>
<td><strong>pic</strong> ([ISO/IEC29500-1:2016] section 19.3.1.37)</td>
</tr>
</tbody>
</table>

<table>
<colgroup>
<col style="width: 58%" />
<col style="width: 41%" />
</colgroup>
<thead>
<tr>
<th>AlternateContent components</th>
<th>Child element</th>
</tr>
</thead>
<tbody>
<tr>
<td>Choice:<br />
http://schemas.microsoft.com/office/powerpoint/2016/slidezoom</td>
<td><strong>sldZm</strong> (section <a href="#Section_623f4c8ecd0a4c188c9df7be9cfbfd72">2.10.1.1</a>)</td>
</tr>
<tr>
<td>Fallback</td>
<td><strong>pic</strong> ([ISO/IEC29500-1:2016] section 19.3.1.37)</td>
</tr>
</tbody>
</table>

<table>
<colgroup>
<col style="width: 62%" />
<col style="width: 37%" />
</colgroup>
<thead>
<tr>
<th>AlternateContent components</th>
<th>Child element</th>
</tr>
</thead>
<tbody>
<tr>
<td>Choice:<br />
http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom</td>
<td><strong>summaryZm</strong> (section <a href="#Section_68d01d06c23447b3b191af013beb1a19">2.11.1.1</a>)</td>
</tr>
<tr>
<td>Fallback</td>
<td><strong>grpSp</strong> ([ISO/IEC29500-1:2016] section 19.3.1.22)</td>
</tr>
</tbody>
</table>

### 2.2.16 View Mode Extensions

The **extLst** child element of the **presentationPr** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.2.1.27) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.\<12\>

| Ext uri | Child element |
|----|----|
| {1BD7E111-0CB8-44D6-8891-C1BB2F81B7CC} | **readonlyRecommended** (section 2.14.1.1) |

### 2.2.17 Design Element Extensions

The **extLst** child element of the **nvPr** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.33) is extended by the addition of a new **ext** child element\<13\> [(](http://go.microsoft.com/fwlink/?LinkId=325242)\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.

<table style="width:95%;">
<colgroup>
<col style="width: 50%" />
<col style="width: 44%" />
</colgroup>
<thead>
<tr>
<th>Ext uri</th>
<th><blockquote>
<p>Child element</p>
</blockquote></th>
</tr>
</thead>
<tbody>
<tr>
<td>{386F3935-93C4-4BCD-93E2-E3B085C9AB24}</td>
<td><blockquote>
<p><strong>designElem</strong> (section <a href="#Section_208645d6756f4479a3d2249ebd379d28">2.5.3.1</a>)</p>
</blockquote></td>
</tr>
</tbody>
</table>

### 2.2.18 Classification Element Extensions

The **extLst** child element of the **nvPr** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.33) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.\<14\>

| Ext uri | Child element |
|----|----|
| {1162E1C5-73C7-4A58-AE30-91384D911F3F} | **classification** (section 2.15.1.1) |

### 2.2.19 Designer Properties Extensions

The **extLst** child element of the **nvPr** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.3.1.33) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.\<15\>

| Ext uri | Child element |
|----|----|
| {E7BDC344-281C-4309-B0C6-D0EE65EED2A8} | **designPr** (section 2.17.1.1) |

### 2.2.20 Designer Tags Extensions

The **extLst** child element of the **sldId** element ([\[ISO/IEC29500-1:2016\]](https://go.microsoft.com/fwlink/?linkid=861065) section 19.2.1.33) is extended by the addition of a new **ext** child element (\[ISO/IEC29500-1:2016\] section 19.2.1.11), whose structure is specified in the following table.\<16\>

| Ext uri | Child element |
|----|----|
| {E3EDB536-0D56-4F60-86BA-61A60CA02DAB} | **designTagLst** (section 2.17.1.2) |
