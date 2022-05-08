if _G.FORMAT ~= 'json' then
  print([[MINI FILTER HELP

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

]])
end
paraName = "Main"
pictureName = "Picture"
pictureCaptionName = "Picture Caption"
tableCaptionName = "Table Caption 1"
tableRowName = "Table Row"
headingName = "Num Heading"
bulletName = "Unnumbered"
paraAttr = pandoc.Attr("", {
  "Main"
}, {
  {
    "custom-style",
    paraName
  }
})
imageAttr = pandoc.Attr("", {
  "Picture"
}, {
  {
    "custom-style",
    pictureName
  }
})
imageCaptionAttr = pandoc.Attr("", {
  "Picture Caption"
}, {
  {
    "custom-style",
    pictureCaptionName
  }
})
tableCaptionAttr = pandoc.Attr("", {
  "Table Caption 1"
}, {
  {
    "custom-style",
    tableCaptionName
  }
})
tableRowAttr = pandoc.Attr("", {
  "Table Row"
}, {
  {
    "custom-style",
    tableRowName
  }
})
h1Attr = pandoc.Attr("", {
  "Num Heading 1"
}, {
  {
    "custom-style",
    headingName .. " 1"
  }
})
h2Attr = pandoc.Attr("", {
  "Num Heading 2"
}, {
  {
    "custom-style",
    headingName .. " 2"
  }
})
h3Attr = pandoc.Attr("", {
  "Num Heading 3"
}, {
  {
    "custom-style",
    headingName .. " 3"
  }
})
h4Attr = pandoc.Attr("", {
  "Num Heading 4"
}, {
  {
    "custom-style",
    headingName .. " 4"
  }
})
h5Attr = pandoc.Attr("", {
  "Num Heading 5"
}, {
  {
    "custom-style",
    headingName .. " 5"
  }
})
h6Attr = pandoc.Attr("", {
  "Num Heading 6"
}, {
  {
    "custom-style",
    headingName .. " 6"
  }
})
string.starts = function(str, start)
  return str:sub(1, #start) == start
end
string.ends = function(str, ending)
  return ending == "" or str:sub(-#ending) == ending
end
chain = function(value, ...)
  local items = {
    ...
  }
  local t = value
  for _index_0 = 1, #items do
    local item = items[_index_0]
    t = item(t)
  end
  return t
end
shaX = function(value)
  local text = pandoc.utils.stringify(value)
  local protohash = pandoc.sha1(text)
  local hash = 'Z' .. string.sub(protohash, 2, 20)
  return hash
end
idfunc = function(element)
  return element
end
wrapQuote = function(element)
  return {
    pandoc.Str("'"),
    element,
    pandoc.Str("'")
  }
end
wrapDblQuote = function(element)
  return {
    pandoc.Str("\""),
    element,
    pandoc.Str("\"")
  }
end
wrapRusQuote = function(element)
  return {
    pandoc.Str("«"),
    element,
    pandoc.Str("»")
  }
end
wrapBrackets = function(element)
  return {
    pandoc.Str("["),
    element,
    pandoc.Str("]")
  }
end
wrapAngleBrackets = function(element)
  return {
    pandoc.Str("<"),
    element,
    pandoc.Str(">")
  }
end
getRole = function(element)
  local _exp_0 = element.tag
  if 'Code' == _exp_0 then
    if (element.classes:includes('interpreted-text', 1)) and element.attributes['role'] then
      return element.attributes['role']
    end
  elseif 'Span' == _exp_0 then
    if (#element.classes > 0) and (element.classes[1]) then
      return element.classes[1]
    end
  end
end
makeRef = function(element)
  local text = pandoc.utils.stringify(element)
  local protohash = pandoc.sha1(text)
  local hash = 'Z' .. string.sub(protohash, 2, 20)
  local ref = string.format('<w:bookmarkStart w:id="%s" w:name="%s"/><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Рисунок </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="%s"/>', hash, hash, hash)
  return pandoc.RawInline('openxml', ref)
end
makeProperty = function(element)
  local text = pandoc.utils.stringify(element)
  local prop = string.format('<w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">DOCPROPERTY %s</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', text)
  return pandoc.RawInline('openxml', prop)
end
makeLink = function(element)
  local inText = pandoc.utils.stringify(element)
  local splitPosition = string.find(inText, ';')
  local protohash = nil
  local link = nil
  if splitPosition then
    protohash = pandoc.sha1(string.sub(inText, 1, splitPosition - 1))
    local hash = 'Z' .. string.sub(protohash, 2, 20)
    local params = string.sub(inText, splitPosition + 1, #inText)
    local paramsSplitPos = string.find(params, ';')
    local param1 = nil
    local param2 = nil
    if paramsSplitPos then
      param1 = string.sub(params, 1, paramsSplitPos - 1)
      param2 = string.sub(params, paramsSplitPos + 1, #params)
    else
      param1 = params
      param2 = params
    end
    link = string.format('<w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> = </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText> PAGE </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText> &lt;&gt; </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText> PAGEREF </w:instrText></w:r>\n    <w:r><w:instrText xml:space="preserve"> "%s" </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> + </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText> REF </w:instrText></w:r>\n    <w:r><w:instrText xml:space="preserve"> "%s" \\p </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> = "выше" </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> &gt;= 1 </w:instrText></w:r>\n    <w:r><w:instrText xml:space="preserve"> "см. %s" "%s"</w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>' .. '<w:r><w:t xml:space="preserve"> </w:t></w:r>' .. '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF "%s" \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash, param2, param1, hash, hash)
  else
    protohash = pandoc.sha1(inText)
    local hash = 'Z' .. string.sub(protohash, 2, 20)
    link = string.format('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF "%s" \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash)
  end
  return pandoc.RawInline('openxml', link)
end
makeLinkNum = function(element)
  local inText = pandoc.utils.stringify(element)
  local splitPosition = string.find(inText, ';')
  local protohash = nil
  local link = nil
  if splitPosition then
    protohash = pandoc.sha1(string.sub(inText, 1, splitPosition - 1))
    local hash = 'Z' .. string.sub(protohash, 2, 20)
    local params = string.sub(inText, splitPosition + 1, #inText)
    local paramsSplitPos = string.find(params, ';')
    local param1 = nil
    local param2 = nil
    if paramsSplitPos then
      param1 = string.sub(params, 1, paramsSplitPos - 1)
      param2 = string.sub(params, paramsSplitPos + 1, #params)
    else
      param1 = params
      param2 = params
    end
    link = string.format('<w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> = </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText> PAGE </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText> &lt;&gt; </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText> PAGEREF </w:instrText></w:r>\n    <w:r><w:instrText xml:space="preserve"> "%s" </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> + </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText> REF </w:instrText></w:r>\n    <w:r><w:instrText xml:space="preserve"> "%s" \\p </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> = "выше" </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> &gt;= 1 </w:instrText></w:r><w:r><w:instrText xml:space="preserve"> "см. %s" "%s"</w:instrText></w:r> <w:r><w:fldChar w:fldCharType="end"/></w:r>' .. '<w:r><w:t xml:space="preserve"> </w:t></w:r>' .. '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF "%s" \\h \\n</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash, param2, param1, hash, hash)
  else
    protohash = pandoc.sha1(inText)
    local hash = 'Z' .. string.sub(protohash, 2, 20)
    link = string.format('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF "%s" \\h \\n</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash)
  end
  return pandoc.RawInline('openxml', link)
end
makeLinkPage = function(element)
  local inText = pandoc.utils.stringify(element)
  local splitPosition = string.find(inText, ';')
  local protohash = nil
  local link = nil
  if splitPosition then
    protohash = pandoc.sha1(string.sub(inText, 1, splitPosition - 1))
    local hash = 'Z' .. string.sub(protohash, 2, 20)
    local params = string.sub(inText, splitPosition + 1, #inText)
    local paramsSplitPos = string.find(params, ';')
    local param1 = nil
    local param2 = nil
    if paramsSplitPos then
      param1 = string.sub(params, 1, paramsSplitPos - 1)
      param2 = string.sub(params, paramsSplitPos + 1, #params)
    else
      param1 = params
      param2 = params
    end
    link = string.format('<w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> = </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText> PAGE </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText> &lt;&gt; </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText> PAGEREF </w:instrText></w:r>\n    <w:r><w:instrText xml:space="preserve"> "%s" </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> + </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText> REF </w:instrText></w:r>\n    <w:r><w:instrText xml:space="preserve"> "%s" \\p </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> = "выше" </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> &gt;= 1 </w:instrText></w:r>\n    <w:r><w:instrText xml:space="preserve"> "см. %s" "%s"</w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>' .. '<w:r><w:t xml:space="preserve"> </w:t></w:r>' .. '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF "%s" \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' .. '<w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">PAGEREF "%s" \\p \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash, param2, param1, hash, hash, hash)
  else
    protohash = pandoc.sha1(inText)
    local hash = 'Z' .. string.sub(protohash, 2, 20)
    link = string.format('<w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">REF "%s" \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">PAGEREF "%s" \\p \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash)
  end
  return pandoc.RawInline('openxml', link)
end
makeLinkNumPage = function(element)
  local inText = pandoc.utils.stringify(element)
  local splitPosition = string.find(inText, ';')
  local protohash = nil
  local link = nil
  if splitPosition then
    protohash = pandoc.sha1(string.sub(inText, 1, splitPosition - 1))
    local hash = 'Z' .. string.sub(protohash, 2, 20)
    local params = string.sub(inText, splitPosition + 1, #inText)
    local paramsSplitPos = string.find(params, ';')
    local param1 = nil
    local param2 = nil
    if paramsSplitPos then
      param1 = string.sub(params, 1, paramsSplitPos - 1)
      param2 = string.sub(params, paramsSplitPos + 1, #params)
    else
      param1 = params
      param2 = params
    end
    link = string.format('<w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> = </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText> PAGE </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText> &lt;&gt; </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText> PAGEREF </w:instrText></w:r>\n    <w:r><w:instrText xml:space="preserve"> "%s" </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> + </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> COMPARE </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="begin"/></w:r>\n    <w:r><w:instrText> REF </w:instrText></w:r>\n    <w:r><w:instrText xml:space="preserve"> "%s" \\p </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> = "выше" </w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>\n    <w:r><w:instrText xml:space="preserve"> &gt;= 1 </w:instrText></w:r>\n    <w:r><w:instrText xml:space="preserve"> "см. %s" "%s"</w:instrText></w:r>\n    <w:r><w:fldChar w:fldCharType="end"/></w:r>' .. '<w:r><w:t xml:space="preserve"> </w:t></w:r>' .. '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF "%s" \\h \\n</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' .. '<w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">PAGEREF "%s" \\p \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash, param2, param1, hash, hash, hash)
  else
    protohash = pandoc.sha1(inText)
    local hash = 'Z' .. string.sub(protohash, 2, 20)
    link = string.format('<w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">REF "%s" \\h \\n</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"></w:fldChar></w:r><w:r><w:instrText xml:space="preserve">PAGEREF "%s" \\p \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', hash, hash)
  end
  return pandoc.RawInline('openxml', link)
end
makeInputField = function(element)
  local hash = shaX(element)
  local input = string.format('<w:bookmarkStart w:id="%s" w:name="%s"/><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="%s"/><w:enabled/><w:calcOnExit w:val="1"/><w:textInput><w:default w:val="%s"/></w:textInput></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>%s</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="%s"/>', hash, hash, hash, text, text, hash)
  return pandoc.RawInline('openxml', input)
end
makeDiv = function(element)
  return pandoc.Div(makePara(element), paraAttr)
end
makePara = function(element)
  return pandoc.Para(element.content)
end
makePlain = function(element)
  return pandoc.Plain(element.content)
end
putRole = function(element)
  local role = getRole(element)
  local el = pandoc.utils.stringify(element)
  local _exp_0 = role
  if 'ref' == _exp_0 then
    return makeRef(element)
  elseif 'prop' == _exp_0 then
    return makeProperty(element)
  elseif 'link' == _exp_0 then
    return makeLink(element)
  elseif 'linknum' == _exp_0 then
    return makeLinkNum(element)
  elseif 'linkpage' == _exp_0 then
    return makeLinkPage(element)
  elseif 'linknumpage' == _exp_0 then
    return makeLinkNumPage(element)
  elseif 'input' == _exp_0 then
    return makeInputField(element)
  elseif 'area' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Область"
      }
    })
    return wrapDblQuote(pandoc.Span(el, roleAttr))
  elseif 'button' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Кнопка"
      }
    })
    return wrapDblQuote(pandoc.Span(el, roleAttr))
  elseif 'command' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Команда"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'field' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Поле"
      }
    })
    return wrapDblQuote(pandoc.Span(el, roleAttr))
  elseif 'file' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Файл"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'flag' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Флаг"
      }
    })
    return wrapDblQuote(pandoc.Span(el, roleAttr))
  elseif 'folder' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Папка"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'icon' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Иконка"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'key' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Клавиша"
      }
    })
    return wrapAngleBrackets(pandoc.Span(el, roleAttr))
  elseif 'menu' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Меню"
      }
    })
    return wrapDblQuote(pandoc.Span(el, roleAttr))
  elseif 'page' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Страница"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'parameter' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Параметр"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'path' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Путь"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'screen' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Экран"
      }
    })
    return wrapDblQuote(pandoc.Span(el, roleAttr))
  elseif 'section' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Раздел"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'switch' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Переключатель"
      }
    })
    return wrapDblQuote(pandoc.Span(el, roleAttr))
  elseif 'tab' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Вкладка"
      }
    })
    return wrapDblQuote(pandoc.Span(el, roleAttr))
  elseif 'url' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "URL"
      }
    })
    local link = pandoc.utils.stringify(el)
    return pandoc.Span(link, roleAttr)
  elseif 'user' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Пользователь"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'userole' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Роль"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'value' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Значение"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'window' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Окно"
      }
    })
    return wrapDblQuote(pandoc.Span(el, roleAttr))
  elseif 'i' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Курсив"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'b' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "Полужирный"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'yellow' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "yellow"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'fuchsia' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "fuchsia"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'green' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "green"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  elseif 'red' == _exp_0 then
    local roleAttr = pandoc.Attr("", { }, {
      {
        "custom-style",
        "red"
      }
    })
    return idfunc(pandoc.Span(el, roleAttr))
  else
    return pandoc.Span(element)
  end
end
putHeaders = function(element)
  local content = bookmarkWrapBlock(element)
  local para = makePara(content)
  local _exp_0 = element.level
  if 1 == _exp_0 then
    return pandoc.Div(para, h1Attr)
  elseif 2 == _exp_0 then
    return pandoc.Div(para, h2Attr)
  elseif 3 == _exp_0 then
    return pandoc.Div(para, h3Attr)
  elseif 4 == _exp_0 then
    return pandoc.Div(para, h4Attr)
  elseif 5 == _exp_0 then
    return pandoc.Div(para, h5Attr)
  elseif 6 == _exp_0 then
    return pandoc.Div(para, h6Attr)
  end
end
putFootnote = function(element)
  return {
    pandoc.Str(" ["),
    element,
    pandoc.Str("]")
  }
end
printRole = function(element)
  print('Found role:', (getRole(element)))
  return element
end
printClasses = function(element)
  if element.classes[1] then
    return print('Found Div with class:', element.classes[1])
  end
end
bookmarkWrapInline = function(element)
  local hash = shaX(element)
  local bmkstart = string.format('<w:bookmarkStart w:id="%s" w:name="%s"/>', hash, hash)
  local bmkend = string.format('<w:bookmarkEnd w:id="%s"/>', hash)
  local oobmkstart = pandoc.RawInline('openxml', bmkstart)
  local oobmkend = pandoc.RawInline('openxml', bmkend)
  return {
    oobmkstart,
    element,
    oobmkend
  }
end
bookmarkWrapBlock = function(element)
  local hash = shaX(element)
  local bmkstart = string.format('<w:bookmarkStart w:id="%s" w:name="%s"/>', hash, hash)
  local bmkend = string.format('<w:bookmarkEnd w:id="%s"/>', hash)
  local oobmkstart = pandoc.RawInline('openxml', bmkstart)
  local oobmkend = pandoc.RawInline('openxml', bmkend)
  table.insert(element.content, 1, oobmkstart)
  table.insert(element.content, oobmkend)
  return element
end
walkImage = function(element)
  if element.classes[1] == 'icon' then
    table.insert(element.attributes, {
      "custom-style",
      "Иконка"
    })
    table.insert(element.attributes, {
      "height",
      "16pt"
    })
  end
  return element
end
walkPara = function(element)
  if element.content[1].tag == 'Image' and element.content[1].title == 'fig:' then
    return {
      nullifyCaption(element.content[1]),
      getImageCaption(element.content[1])
    }
  end
  return pandoc.Div((pandoc.walk_block(element, {
    Span = putRole,
    Code = putRole
  })), paraAttr)
end
walkDiv = function(element)
  if element.identifier then
    local hash = shaX(element.identifier)
    element.identifier = hash
  end
  if element.classes[1] == 'note' then
    local prefix = pandoc.Para(pandoc.Str('Примечание'))
    local sep = {
      pandoc.Str(' – ')
    }
    local data = {
      prefix,
      element.content[2]
    }
    local inlines = pandoc.utils.blocks_to_inlines(data, sep)
    return pandoc.Div(pandoc.Para(inlines), pandoc.Attr("", {
      "Note"
    }, {
      {
        "custom-style",
        "Примечание"
      }
    }))
  end
  if element.classes[1] == 'attention' then
    local prefix = pandoc.Para(pandoc.Str('ВНИМАНИЕ!'))
    local sep = {
      pandoc.Str(' ')
    }
    local data = {
      prefix,
      element.content[2]
    }
    local inlines = pandoc.utils.blocks_to_inlines(data, sep)
    return pandoc.Div(pandoc.Para(inlines), pandoc.Attr("", {
      "Note"
    }, {
      {
        "custom-style",
        "Внимание"
      }
    }))
  end
  if element.classes[1] == 'tip' then
    local prefix = pandoc.Para(pandoc.Str('\tСовет:'))
    local sep = {
      pandoc.Str('\t')
    }
    local data = {
      prefix,
      element.content[2]
    }
    local inlines = pandoc.utils.blocks_to_inlines(data, sep)
    return pandoc.Div(pandoc.Para(inlines), pandoc.Attr("", {
      "Note"
    }, {
      {
        "custom-style",
        "Совет"
      }
    }))
  end
  return element
end
walkPlain = function(element)
  return pandoc.walk_block(element, {
    Span = putRole,
    Code = putRole
  })
end
bulletLevel = 0
walkBullet = function(element)
  local el = { }
  for k, v in pairs(element.content) do
    table.insert(el, { })
    if #v > 0 then
      for l, w in pairs(v) do
        bulletLevel = bulletLevel + 1
        if w.tag == 'Plain' or w.tag == 'Para' then
          local z = pandoc.Div(pandoc.Para(w.content), pandoc.Attr("", { }, {
            {
              "custom-style",
              bulletName .. " " .. bulletLevel
            }
          }))
          table.insert(el[k], z)
        else
          walkPara = function(smth)
            local listAttr = pandoc.Attr("", { }, {
              {
                "custom-style",
                bulletName .. " " .. bulletLevel
              }
            })
            return pandoc.Div((pandoc.walk_block(smth, {
              Span = putRole,
              Code = putRole
            })), listAttr)
          end
          local z = pandoc.walk_block(w, {
            BulletList = walkBullet,
            OrderedList = walkOrdered,
            Para = walkPara,
            Plain = walkPara
          })
          table.insert(el[k], z)
        end
        bulletLevel = bulletLevel - 1
      end
    end
  end
  element.content = el
  return element
end
walkOrdered = function(element)
  local level = element.start or 1
  local listelements = { }
  for k, v in pairs(element.content) do
    for l, w in pairs(v) do
      if l == 1 then
        local code = string.format('<w:r><w:t>%d)</w:t><w:tab/></w:r>', level)
        table.insert(v[1].content[1].content, 1, pandoc.RawInline('openxml', code))
        table.insert(listelements, #listelements + 1, w)
      end
      if l > 1 then
        table.insert(listelements, #listelements + 1, w)
      end
    end
    level = level + 1
  end
  return pandoc.Div(listelements, paraAttr)
end
walkOrderedSeq = function(element)
  local level = element.start or 1
  local listelements = { }
  local wPlain
  wPlain = function(el)
    if #listelements == 0 then
      local code = string.format('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ NumList \\r %d </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>)</w:t><w:tab/></w:r>', level)
      table.insert(el.content, 1, pandoc.RawInline('openxml', code))
    end
    if #listelements > 0 then
      local code = string.format('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ NumList </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>)</w:t><w:tab/></w:r>', level)
      table.insert(el.content, 1, pandoc.RawInline('openxml', code))
    end
    return table.insert(listelements, #listelements + 1, el)
  end
  pandoc.walk_block(element, {
    Plain = wPlain
  })
  return pandoc.Div(listelements, paraAttr)
end
walkTable = function(element)
  local cap = getTableCaption(element)
  element.caption.long = { }
  walkPara = function(el)
    if el.tag == 'Para' or el.tag == 'Plain' then
      return pandoc.Div((pandoc.Para(el.content)), tableRowAttr)
    end
    return el
  end
  pandoc.walk_block(element.bodies[1], {
    Para = walkPara,
    Plain = walkPara
  })
  return {
    cap,
    element,
    pandoc.Div((pandoc.Para((pandoc.Str('')))))
  }
end
getImageCaption = function(element)
  local id = pandoc.utils.stringify(element.identifier or 'none')
  local protohash = pandoc.sha1(id)
  local hash = 'Z' .. string.sub(protohash, 2, 20)
  if #element.caption > 0 then
    local capSrc = string.format('<w:r><w:t xml:space="preserve">Рисунок </w:t></w:r><w:bookmarkStart w:id="%s" w:name="%s"/><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Рисунок </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="%s"/><w:r><w:t xml:space="preserve"> – </w:t></w:r>', hash, hash, hash)
    local rawCap = pandoc.RawInline('openxml', capSrc)
    local cap = pandoc.walk_block(pandoc.Para(element.caption), {
      Span = putRole,
      Code = putRole
    })
    element.caption = cap.content
    table.insert(element.caption, 1, rawCap)
  end
  if #element.caption == 0 or element.caption == nil then
    local capSrc = string.format('<w:r><w:t xml:space="preserve">Рисунок </w:t></w:r><w:bookmarkStart w:id="%s" w:name="%s"/><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Рисунок </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="%s"/>', hash, hash, hash)
    local rawCap = pandoc.RawInline('openxml', capSrc)
    element.caption = {
      rawCap
    }
  end
  return pandoc.Div((pandoc.Para(element.caption)), imageCaptionAttr)
end
getTableCaption = function(element)
  local hash = shaX(pandoc.utils.blocks_to_inlines(element.caption.long))
  local capSrc = string.format('<w:r><w:t xml:space="preserve">Таблица </w:t></w:r><w:bookmarkStart w:id="%s" w:name="%s"/><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Таблица </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="%s"/><w:r><w:t xml:space="preserve"> –\t</w:t></w:r>', hash, hash, hash)
  local rawCap = pandoc.RawInline('openxml', capSrc)
  pandoc.walk_block(element.caption.long[1], {
    Span = putRole,
    Code = putRole
  })
  table.insert(element.caption.long[1].content, 1, rawCap)
  return pandoc.Div((element.caption.long), tableCaptionAttr)
end
nullify = function(element)
  return pandoc.Null
end
nullifyCaption = function(element)
  local figure = pandoc.Image(pandoc.Str(''), element.src)
  figure.identifier = element.identifier
  figure.classes = element.classes
  return pandoc.Div((pandoc.Para(figure)), imageAttr)
end
Pandoc = function(doc)
  local tree = pandoc.walk_block((pandoc.Div(doc.blocks)), {
    Note = putFootnote,
    Div = walkDiv,
    Header = putHeaders,
    BulletList = walkBullet,
    OrderedList = walkOrdered,
    Para = walkPara,
    Table = walkTable,
    Plain = walkPlain,
    Image = walkImage
  })
  return pandoc.Pandoc(tree.content, doc.meta)
end
