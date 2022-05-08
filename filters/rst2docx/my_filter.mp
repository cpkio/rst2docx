if _G.FORMAT != 'json'
  print [[
MINI FILTER HELP

This filter helps to convert reStructuredText files to DOCX

Contrary to previous filter version, inserting references now is field-based.

Some elements' data is hashed and wrapped into a bookmark with this hash:

    .. figure: image\source.png
       :name: Figure reference title

       Figure displayed title

To reference this figure, write
  - :link:`Figure reference title`
  - :linknum:`Header reference`
  - :linkpage:`Figure reference title`

:prop: field will insert document property.

:input: will create an input field in resulting DOCX.

Other fields are:
    'area'
    'b' (bold)
    'button'
    'command'
    'field'
    'file'
    'flag'
    'folder'
    'i' (italic)
    'icon` (for inline images of buttons)
    'key'
    'menu'
    'page'
    'parameter'
    'path'
    'screen'
    'section'
    'switch'
    'tab'
    'url' (underlined URL)
    'user'
    'userole'
    'value'
    'window'

! Referencing links of '.. _`some text`:' type (which wrap the next para to
! bookmark) has to be done with :link:`Xsome text` so that hashing would work
! correctly. This is done in walkDiv function.

]]

export *

paraName = "Main"
pictureName = "Picture"
pictureCaptionName = "Picture Caption"
tableCaptionName = "Table Caption 1"
tableRowName = "Table Row"
headingName = "Num Heading"
bulletName = "Unnumbered"

paraAttr = pandoc.Attr("", { "Main" }, {{"custom-style", paraName }})
imageAttr = pandoc.Attr("", { "Picture" }, {{"custom-style", pictureName }})
imageCaptionAttr = pandoc.Attr("", { "Picture Caption" }, {{"custom-style", pictureCaptionName }})
tableCaptionAttr = pandoc.Attr( "", { "Table Caption 1" }, {{"custom-style", tableCaptionName }} )
tableRowAttr = pandoc.Attr( "", { "Table Row" }, {{"custom-style", tableRowName }} )

h1Attr = pandoc.Attr("", { "Num Heading 1" }, {{"custom-style", headingName .. " 1" }})
h2Attr = pandoc.Attr("", { "Num Heading 2" }, {{"custom-style", headingName .. " 2" }})
h3Attr = pandoc.Attr("", { "Num Heading 3" }, {{"custom-style", headingName .. " 3" }})
h4Attr = pandoc.Attr("", { "Num Heading 4" }, {{"custom-style", headingName .. " 4" }})
h5Attr = pandoc.Attr("", { "Num Heading 5" }, {{"custom-style", headingName .. " 5" }})
h6Attr = pandoc.Attr("", { "Num Heading 6" }, {{"custom-style", headingName .. " 6" }})

string.starts = (str, start) ->
  str\sub(1, #start) == start

string.ends = (str, ending) ->
  ending == "" or str\sub(-#ending) == ending

chain = (value, ...) ->
  items = { ... }
  t = value
  for item in *items
    t = item(t)
  return t

shaX = (value) ->
  text = pandoc.utils.stringify value
  protohash = pandoc.sha1 text
  hash = 'Z' .. string.sub(protohash, 2, 20)
  return hash

idfunc = (element) ->
  return element

wrapQuote = (element) ->
  return { pandoc.Str("'"), element, pandoc.Str("'") }

wrapDblQuote = (element) ->
  return { pandoc.Str("\""), element, pandoc.Str("\"") }

wrapRusQuote = (element) ->
  return { pandoc.Str("«"), element, pandoc.Str("»") }

wrapBrackets = (element) ->
  return { pandoc.Str("["), element, pandoc.Str("]") }

wrapAngleBrackets = (element) ->
  return { pandoc.Str("<"), element, pandoc.Str(">") }

-- This function returns the name of reStructuredText interpreted text
-- element. Pandoc treats :role: differently based on whether this role has
-- been previously declared (.. role::) or not. The function basically returns
-- role name as string.
getRole = (element) ->
  switch element.tag
    when 'Code' -- the role has not been declared
      if (element.classes\includes 'interpreted-text', 1) and element.attributes['role']
        element.attributes['role']
    when 'Span' -- the role has been declared
      if (#element.classes > 0) and (element.classes[1])
        element.classes[1]

makeRef = (element) ->
  text = pandoc.utils.stringify element
  protohash = pandoc.sha1 text
  hash = 'Z' .. string.sub(protohash, 2, 20)
  ref = string.format '<w:bookmarkStart w:id="%s" w:name="%s"/><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Рисунок </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="%s"/>', hash, hash, hash
  return pandoc.RawInline 'openxml', ref

-- Transform text to document property field. Used for :prop:`` role
-- processing
makeProperty = (element) ->
  text = pandoc.utils.stringify element
  prop = string.format '<w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">DOCPROPERTY %s</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', text
  return pandoc.RawInline 'openxml', prop

-- Transform text to bookmark reference field; bookmark has to be of same text
-- as reference, or the hashes would not match. Used for :link:`` role
-- processing
makeLink = (element) ->
  inText = pandoc.utils.stringify element
  splitPosition = string.find(inText, ';')
  protohash = nil
  link = nil
  if splitPosition
    protohash = pandoc.sha1 string.sub(inText, 1, splitPosition-1)
    hash = 'Z' .. string.sub(protohash, 2, 20)
    params = string.sub(inText, splitPosition+1, #inText)
    paramsSplitPos = string.find(params, ';')
    param1 = nil
    param2 = nil
    if paramsSplitPos
      param1 = string.sub(params, 1, paramsSplitPos-1)
      param2 = string.sub(params, paramsSplitPos+1, #params)
    else
      param1 = params
      param2 = params
    link = string.format '<w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> = </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText> &lt;&gt; </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> PAGEREF </w:instrText></w:r>
    <w:r><w:instrText xml:space="preserve"> "%s" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText xml:space="preserve"> + </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> REF </w:instrText></w:r>
    <w:r><w:instrText xml:space="preserve"> "%s" \\p </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText xml:space="preserve"> = "выше" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText xml:space="preserve"> &gt;= 1 </w:instrText></w:r>
    <w:r><w:instrText xml:space="preserve"> "см. %s" "%s"</w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>'..'<w:r><w:t xml:space="preserve"> </w:t></w:r>'..'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF "%s" \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash, param2, param1, hash, hash
  else
    protohash = pandoc.sha1 inText
    hash = 'Z' .. string.sub(protohash, 2, 20)
    link = string.format '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF "%s" \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash
  return pandoc.RawInline 'openxml', link

makeLinkNum = (element) ->
  inText = pandoc.utils.stringify element
  splitPosition = string.find(inText, ';')
  protohash = nil
  link = nil
  if splitPosition
    protohash = pandoc.sha1 string.sub(inText, 1, splitPosition-1)
    hash = 'Z' .. string.sub(protohash, 2, 20)
    params = string.sub(inText, splitPosition+1, #inText)
    paramsSplitPos = string.find(params, ';')
    param1 = nil
    param2 = nil
    if paramsSplitPos
      param1 = string.sub(params, 1, paramsSplitPos-1)
      param2 = string.sub(params, paramsSplitPos+1, #params)
    else
      param1 = params
      param2 = params
    link = string.format '<w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> = </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText> &lt;&gt; </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> PAGEREF </w:instrText></w:r>
    <w:r><w:instrText xml:space="preserve"> "%s" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText xml:space="preserve"> + </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> REF </w:instrText></w:r>
    <w:r><w:instrText xml:space="preserve"> "%s" \\p </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText xml:space="preserve"> = "выше" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText xml:space="preserve"> &gt;= 1 </w:instrText></w:r><w:r><w:instrText xml:space="preserve"> "см. %s" "%s"</w:instrText></w:r> <w:r><w:fldChar w:fldCharType="end"/></w:r>'..'<w:r><w:t xml:space="preserve"> </w:t></w:r>'..'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF "%s" \\h \\n</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash, param2, param1, hash, hash
  else
    protohash = pandoc.sha1 inText
    hash = 'Z' .. string.sub(protohash, 2, 20)
    link = string.format '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF "%s" \\h \\n</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash
  return pandoc.RawInline 'openxml', link
makeLinkPage = (element) ->
  inText = pandoc.utils.stringify element
  splitPosition = string.find(inText, ';')
  protohash = nil
  link = nil
  if splitPosition
    protohash = pandoc.sha1 string.sub(inText, 1, splitPosition-1)
    hash = 'Z' .. string.sub(protohash, 2, 20)
    params = string.sub(inText, splitPosition+1, #inText)
    paramsSplitPos = string.find(params, ';')
    param1 = nil
    param2 = nil
    if paramsSplitPos
      param1 = string.sub(params, 1, paramsSplitPos-1)
      param2 = string.sub(params, paramsSplitPos+1, #params)
    else
      param1 = params
      param2 = params
    link = string.format '<w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> = </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText> &lt;&gt; </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> PAGEREF </w:instrText></w:r>
    <w:r><w:instrText xml:space="preserve"> "%s" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText xml:space="preserve"> + </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> REF </w:instrText></w:r>
    <w:r><w:instrText xml:space="preserve"> "%s" \\p </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText xml:space="preserve"> = "выше" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText xml:space="preserve"> &gt;= 1 </w:instrText></w:r>
    <w:r><w:instrText xml:space="preserve"> "см. %s" "%s"</w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>'..'<w:r><w:t xml:space="preserve"> </w:t></w:r>'..'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF "%s" \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'..'<w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">PAGEREF "%s" \\p \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash, param2, param1, hash, hash, hash
  else
    protohash = pandoc.sha1 inText
    hash = 'Z' .. string.sub(protohash, 2, 20)
    link = string.format '<w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">REF "%s" \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">PAGEREF "%s" \\p \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash
  return pandoc.RawInline 'openxml', link


makeLinkNumPage = (element) ->
  inText = pandoc.utils.stringify element
  splitPosition = string.find(inText, ';')
  protohash = nil
  link = nil
  if splitPosition
    protohash = pandoc.sha1 string.sub(inText, 1, splitPosition-1)
    hash = 'Z' .. string.sub(protohash, 2, 20)
    params = string.sub(inText, splitPosition+1, #inText)
    paramsSplitPos = string.find(params, ';')
    param1 = nil
    param2 = nil
    if paramsSplitPos
      param1 = string.sub(params, 1, paramsSplitPos-1)
      param2 = string.sub(params, paramsSplitPos+1, #params)
    else
      param1 = params
      param2 = params
    link = string.format '<w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> = </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText> &lt;&gt; </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> PAGEREF </w:instrText></w:r>
    <w:r><w:instrText xml:space="preserve"> "%s" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText xml:space="preserve"> + </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText> REF </w:instrText></w:r>
    <w:r><w:instrText xml:space="preserve"> "%s" \\p </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText xml:space="preserve"> = "выше" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText xml:space="preserve"> &gt;= 1 </w:instrText></w:r>
    <w:r><w:instrText xml:space="preserve"> "см. %s" "%s"</w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>'..'<w:r><w:t xml:space="preserve"> </w:t></w:r>'..'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF "%s" \\h \\n</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'..'<w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">PAGEREF "%s" \\p \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash, param2, param1, hash, hash, hash
  else
    protohash = pandoc.sha1 inText
    hash = 'Z' .. string.sub(protohash, 2, 20)
    link = string.format '<w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">REF "%s" \\h \\n</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">PAGEREF "%s" \\p \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash
  return pandoc.RawInline 'openxml', link

-- Transform text to input field with same text as default, wrapped into
-- bookmark
makeInputField = (element) ->
  hash = shaX element
  input = string.format '<w:bookmarkStart w:id="%s" w:name="%s"/><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="%s"/><w:enabled/><w:calcOnExit w:val="1"/><w:textInput><w:default w:val="%s"/></w:textInput></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>%s</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="%s"/>', hash, hash, hash, text, text, hash
  return pandoc.RawInline 'openxml', input

-- Wrap element into a Div with default Para styling
makeDiv = (element) ->
  return pandoc.Div makePara(element), paraAttr

-- Transform block element with list of inlines as content into a Para. Para
-- cannot have styling ('custom-style') for Word, so Div wrapping is required
makePara = (element) ->
  return pandoc.Para element.content

-- Transform block element with list of inlines as content into Plain
makePlain = (element) ->
  return pandoc.Plain element.content

-- This function inserts inline role elements wrapped into required symbols
putRole = (element) ->
  role = getRole element
  el = pandoc.utils.stringify element
  switch role
    when 'ref'
      makeRef element
    when 'prop'
      makeProperty element
    when 'link'
      makeLink element
    when 'linknum'
      makeLinkNum element
    when 'linkpage'
      makeLinkPage element
    when 'linknumpage'
      makeLinkNumPage element
    when 'input'
      makeInputField element
    when 'area'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Область" }})
      return wrapDblQuote pandoc.Span(el, roleAttr)
    when 'button'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Кнопка" }})
      return wrapDblQuote pandoc.Span(el, roleAttr)
    when 'command'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Команда" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'field'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Поле" }})
      return wrapDblQuote pandoc.Span(el, roleAttr)
    when 'file'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Файл" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'flag'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Флаг" }})
      return wrapDblQuote pandoc.Span(el, roleAttr)
    when 'folder'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Папка" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'icon'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Иконка" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'key'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Клавиша" }})
      return wrapAngleBrackets pandoc.Span(el, roleAttr)
    when 'menu'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Меню" }})
      return wrapDblQuote pandoc.Span(el, roleAttr)
    when 'page'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Страница" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'parameter'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Параметр" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'path'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Путь" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'screen'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Экран" }})
      return wrapDblQuote pandoc.Span(el, roleAttr)
    when 'section'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Раздел" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'switch'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Переключатель" }})
      return wrapDblQuote pandoc.Span(el, roleAttr)
    when 'tab'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Вкладка" }})
      return wrapDblQuote pandoc.Span(el, roleAttr)
    when 'url'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "URL" }})
      link = pandoc.utils.stringify el
      return pandoc.Span(link, roleAttr)
    when 'user'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Пользователь" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'userole'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Роль" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'value'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Значение" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'window'
      roleAttr = pandoc.Attr("",{  },{{ "custom-style", "Окно" }})
      return wrapDblQuote pandoc.Span(el, roleAttr)
    when 'i'
      roleAttr = pandoc.Attr("",{ },{{ "custom-style", "Курсив" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'b'
      roleAttr = pandoc.Attr("",{ },{{ "custom-style", "Полужирный" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'yellow'
      roleAttr = pandoc.Attr("",{ },{{ "custom-style", "yellow" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'fuchsia'
      roleAttr = pandoc.Attr("",{ },{{ "custom-style", "fuchsia" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'green'
      roleAttr = pandoc.Attr("",{ },{{ "custom-style", "green" }})
      return idfunc pandoc.Span(el, roleAttr)
    when 'red'
      roleAttr = pandoc.Attr("",{ },{{ "custom-style", "red" }})
      return idfunc pandoc.Span(el, roleAttr)
    else
      return pandoc.Span(element)

putHeaders = (element) ->
  content = bookmarkWrapBlock element
  para = makePara content
  switch element.level
    when 1
      return pandoc.Div(para, h1Attr)
    when 2
      return pandoc.Div(para, h2Attr)
    when 3
      return pandoc.Div(para, h3Attr)
    when 4
      return pandoc.Div(para, h4Attr)
    when 5
      return pandoc.Div(para, h5Attr)
    when 6
      return pandoc.Div(para, h6Attr)

putFootnote = (element) ->
  return { pandoc.Str(" ["), element, pandoc.Str("]") }

-- This helper function prints a role name before returning same object to the
-- caller. Used for debugging
printRole = (element) ->
  print 'Found role:', (getRole element)
  return element

-- This helper function prints Div class before returning same object to the
-- caller. Used for debugging
printClasses = (element) ->
  if element.classes[1]
    print 'Found Div with class:', element.classes[1]

-- Wrap any Inline with bookmark wrapper. Bookmarks should not wrap whole text
-- sections
bookmarkWrapInline = (element) ->
  hash = shaX element
  bmkstart = string.format '<w:bookmarkStart w:id="%s" w:name="%s"/>', hash, hash
  bmkend = string.format '<w:bookmarkEnd w:id="%s"/>', hash
  oobmkstart = pandoc.RawInline 'openxml', bmkstart
  oobmkend = pandoc.RawInline 'openxml', bmkend
  return { oobmkstart, element, oobmkend }

bookmarkWrapBlock = (element) ->
  hash = shaX element
  bmkstart = string.format '<w:bookmarkStart w:id="%s" w:name="%s"/>', hash, hash
  bmkend = string.format '<w:bookmarkEnd w:id="%s"/>', hash
  oobmkstart = pandoc.RawInline 'openxml', bmkstart
  oobmkend = pandoc.RawInline 'openxml', bmkend
  table.insert(element.content, 1, oobmkstart)
  table.insert(element.content, oobmkend)
  return element

-- Images processing
walkImage = (element) ->
  if element.classes[1] == 'icon'
    table.insert(element.attributes, { "custom-style", "Иконка" })
    table.insert(element.attributes, { "height", "16pt" })
    -- If inline image is :icon:, set its style and height. This is used only in my Pandoc patch.
  return element

-- Paragraph walker function
walkPara = (element) ->
  if element.content[1].tag == 'Image' and element.content[1].title == 'fig:'
    return {
      nullifyCaption element.content[1]
      getImageCaption element.content[1]
    }
  return pandoc.Div (pandoc.walk_block element, { Span: putRole, Code: putRole }), paraAttr

walkDiv = (element) ->
  if element.identifier
    hash = shaX element.identifier
    element.identifier = hash
  if element.classes[1] == 'note'
    prefix = pandoc.Para pandoc.Str'Примечание'
    sep = { pandoc.Str' – ' }
    data = { prefix, element.content[2] }
    inlines = pandoc.utils.blocks_to_inlines(data, sep)
    return pandoc.Div(pandoc.Para(inlines), pandoc.Attr( "", { "Note" }, {{ "custom-style", "Примечание" }} ))
  if element.classes[1] == 'attention'
    prefix = pandoc.Para pandoc.Str'ВНИМАНИЕ!'
    sep = { pandoc.Str' ' }
    data = { prefix, element.content[2] }
    inlines = pandoc.utils.blocks_to_inlines(data, sep)
    return pandoc.Div(pandoc.Para(inlines), pandoc.Attr( "", { "Note" }, {{ "custom-style", "Внимание" }} ))
  if element.classes[1] == 'tip'
    prefix = pandoc.Para pandoc.Str'\tСовет:'
    sep = { pandoc.Str'\t' }
    data = { prefix, element.content[2] }
    inlines = pandoc.utils.blocks_to_inlines(data, sep)
    return pandoc.Div(pandoc.Para(inlines), pandoc.Attr( "", { "Note" }, {{ "custom-style", "Совет" }} ))
  return element

walkPlain = (element) ->
  return pandoc.walk_block element, { Span: putRole, Code: putRole }

-- Обработка вложенных ненумерованных списков
bulletLevel = 0
walkBullet = (element) ->
  el = {}
  for k,v in pairs element.content -- проходим по списку и получаем списки блоков
    table.insert(el, {}) -- создаём новый список, дублирующий исходный
    if #v > 0 -- если длина полученного списка блоков не нулевая
      for l,w in pairs v -- получаем элементы списка = блоки
        bulletLevel += 1
        if w.tag == 'Plain' or w.tag == 'Para'
          z = pandoc.Div(pandoc.Para(w.content),pandoc.Attr("",{},{{"custom-style", bulletName .. " " .. bulletLevel }}))
          table.insert(el[k], z)
        else
          walkPara = (smth) ->
            listAttr = pandoc.Attr("",{},{{"custom-style", bulletName .. " " .. bulletLevel }})
            return pandoc.Div (pandoc.walk_block smth, { Span: putRole, Code: putRole }), listAttr
          z = pandoc.walk_block(w, { BulletList: walkBullet, OrderedList: walkOrdered, Para: walkPara, Plain: walkPara })
          table.insert(el[k], z)
        bulletLevel -= 1
  element.content = el
  return element

-- Обработка вложенных нумерованных списков
-- Тут проблема: элемент нумерованного списка ПРОПАДАЕТ, если у него есть
-- вложенный отдельный абзац. Нужно переписывать.
walkOrdered = (element) ->
  level = element.start or 1
  listelements = {}
  for k,v in pairs element.content -- получаем списки элементов одного пункта
    for l,w in pairs v -- проходим по остальным элементам пункта
      if l == 1
        code = string.format '<w:r><w:t>%d)</w:t><w:tab/></w:r>',level
        table.insert(v[1].content[1].content, 1, pandoc.RawInline('openxml', code))
        table.insert(listelements, #listelements+1, w)
      if l > 1
        table.insert(listelements, #listelements+1, w)
    level = level + 1
  return pandoc.Div(listelements, paraAttr)

walkOrderedSeq = (element) ->
  level = element.start or 1
  listelements = {}
  wPlain = (el) ->
    if #listelements == 0
      code = string.format '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ NumList \\r %d </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>)</w:t><w:tab/></w:r>',level
      table.insert(el.content, 1, pandoc.RawInline('openxml', code))
    if #listelements > 0
      code = string.format '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ NumList </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>)</w:t><w:tab/></w:r>',level
      table.insert(el.content, 1, pandoc.RawInline('openxml', code))
    table.insert(listelements, #listelements+1, el)
  pandoc.walk_block(element,{ Plain: wPlain })

  return pandoc.Div(listelements, paraAttr)

walkTable = (element) ->
  cap = getTableCaption element
  element.caption.long = {}
  walkPara = (el) ->
    if el.tag == 'Para' or el.tag == 'Plain'
      return pandoc.Div (pandoc.Para el.content),tableRowAttr
    return el
  pandoc.walk_block element.bodies[1],{ Para: walkPara, Plain: walkPara }
  return {
    cap
    element
    pandoc.Div (pandoc.Para (pandoc.Str ''))
  }

getImageCaption = (element) ->
  id = pandoc.utils.stringify element.identifier or 'none'
  protohash = pandoc.sha1 id
  hash = 'Z' .. string.sub(protohash, 2, 20)
  if #element.caption > 0
    capSrc = string.format '<w:r><w:t xml:space="preserve">Рисунок </w:t></w:r><w:bookmarkStart w:id="%s" w:name="%s"/><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Рисунок </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="%s"/><w:r><w:t xml:space="preserve"> – </w:t></w:r>', hash, hash, hash
    rawCap = pandoc.RawInline 'openxml', capSrc
    cap = pandoc.walk_block pandoc.Para(element.caption), { Span: putRole, Code: putRole }
    element.caption = cap.content
    table.insert(element.caption, 1, rawCap)
  if #element.caption == 0 or element.caption == nil
    capSrc = string.format '<w:r><w:t xml:space="preserve">Рисунок </w:t></w:r><w:bookmarkStart w:id="%s" w:name="%s"/><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Рисунок </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="%s"/>', hash, hash, hash
    rawCap = pandoc.RawInline 'openxml', capSrc
    element.caption = {rawCap}
  return pandoc.Div (pandoc.Para element.caption), imageCaptionAttr

getTableCaption = (element) ->
  hash = shaX pandoc.utils.blocks_to_inlines(element.caption.long)
  capSrc = string.format '<w:r><w:t xml:space="preserve">Таблица </w:t></w:r><w:bookmarkStart w:id="%s" w:name="%s"/><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Таблица </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="%s"/><w:r><w:t xml:space="preserve"> –\t</w:t></w:r>', hash, hash, hash
  rawCap = pandoc.RawInline 'openxml', capSrc
  pandoc.walk_block element.caption.long[1],{ Span: putRole, Code: putRole }
  table.insert(element.caption.long[1].content, 1, rawCap)
  return pandoc.Div (element.caption.long), tableCaptionAttr

nullify = (element) ->
  return pandoc.Null

nullifyCaption = (element) ->
  figure = pandoc.Image(pandoc.Str'', element.src)
  figure.identifier = element.identifier
  figure.classes = element.classes
  return pandoc.Div (pandoc.Para figure), imageAttr

Pandoc = (doc) ->
  tree = pandoc.walk_block (pandoc.Div doc.blocks),
    {
      Note: putFootnote,
      Div: walkDiv,
      Header: putHeaders,
      BulletList: walkBullet,
      OrderedList: walkOrdered,
      Para: walkPara,
      Table: walkTable,
      Plain: walkPlain,
      Image: walkImage,
    }
  pandoc.Pandoc tree.content, doc.meta
