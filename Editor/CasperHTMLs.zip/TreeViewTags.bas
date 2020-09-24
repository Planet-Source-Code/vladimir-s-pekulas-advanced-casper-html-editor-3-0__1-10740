Attribute VB_Name = "TreeViewTags"
Function AddTags()
'-------------------------'
'#    Add basic tags     #'
'-------------------------'
 
 Dim intFileNum As Integer, strFilename As String
 strFilename = App.Path & "\Temp\tags.dat"
 intFileNum = FreeFile
 Open strFilename For Input As #intFileNum
 Do While Not EOF(intFileNum)
  Line Input #intFileNum, SValue
  If Not Trim(SValue) = "" Then fMainForm.TRV.Nodes.Add , , , SValue, 1
 Loop
 Close #intFileNum


'-------------------------'
'# Add property families #'
'-------------------------'
Dim tempNode As Node

'add Anchor
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "ANCHOR", "<a></a>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("ANCHOR", tvwChild, , " href=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("ANCHOR", tvwChild, , " target=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("ANCHOR", tvwChild, , " name=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("ANCHOR", tvwChild, , " title=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("ANCHOR", tvwChild, , " rel=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("ANCHOR", tvwChild, , " rev=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("ANCHOR", tvwChild, , " type=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("ANCHOR", tvwChild, , " charset=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("ANCHOR", tvwChild, , " hreflang=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("ANCHOR", tvwChild, , " media=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("ANCHOR", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("ANCHOR", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("ANCHOR", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 

'add Applet
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Applet", "<applet></applet>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Applet", tvwChild, , " code=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Applet", tvwChild, , " codebase=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Applet", tvwChild, , " name=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Applet", tvwChild, , " title=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Applet", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Applet", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Applet", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Applet", tvwChild, , " archive=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Applet", tvwChild, , " alt=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Applet", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Applet", tvwChild, , " height=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Applet", tvwChild, , " width=" & Chr(34) & Chr(34), 2)



'add Area
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Area", "<area>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Area", tvwChild, , " shape=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Area", tvwChild, , " coords=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Area", tvwChild, , " href=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Area", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Area", tvwChild, , " title=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Area", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Area", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Area", tvwChild, , " target=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Area", tvwChild, , " nohref=" & Chr(34) & Chr(34), 2)



'add Base
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Base", "<base>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Base", tvwChild, , " href=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Base", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Base", tvwChild, , " target=" & Chr(34) & Chr(34), 2)



'add Basefont
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Basefont", "<basefont>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Basefont", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Basefont", tvwChild, , " face=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Basefont", tvwChild, , " size=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Basefont", tvwChild, , " color=" & Chr(34) & Chr(34), 2)



'add bgsound
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Bgsound", "<bgsound>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Bgsound", tvwChild, , " src=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Bgsound", tvwChild, , " loop=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Bgsound", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Bgsound", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Bgsound", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Bgsound", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add Body
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Body", "<body></body>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Body", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Body", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Body", tvwChild, , " title=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Body", tvwChild, , " background=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Body", tvwChild, , " bgcolor=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Body", tvwChild, , " text=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Body", tvwChild, , " link=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Body", tvwChild, , " vlink=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Body", tvwChild, , " alink=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Body", tvwChild, , " leftmargin=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Body", tvwChild, , " topmargin=" & Chr(34) & Chr(34), 2)



'add Col
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Col", "<col>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Col", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Col", tvwChild, , " span=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Col", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Col", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Col", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Col", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add Colgroup
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Colgroup", "<colgroup>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Colgroup", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Colgroup", tvwChild, , " valign=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Colgroup", tvwChild, , " span=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Colgroup", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Colgroup", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Colgroup", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Colgroup", tvwChild, , " title=" & Chr(34) & Chr(34), 2)


'add Div
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "div", "<div></div>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("div", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("div", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("div", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("div", tvwChild, , " title=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("div", tvwChild, , " align=" & Chr(34) & Chr(34), 2)



'add Embed
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Embed", "<embed>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Embed", tvwChild, , " src=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Embed", tvwChild, , " height=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Embed", tvwChild, , " width=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Embed", tvwChild, , " hidden=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Embed", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Embed", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Embed", tvwChild, , " class=" & Chr(34) & Chr(34), 2)



'add Font
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Font", "<font></font>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Font", tvwChild, , " face=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Font", tvwChild, , " size=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Font", tvwChild, , " color=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Font", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Font", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Font", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Font", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add Form
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Form", "<form></form>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Form", tvwChild, , " action=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Form", tvwChild, , " target=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Form", tvwChild, , " method=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Form", tvwChild, , " enctype=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Form", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Form", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Form", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Form", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add Frame
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Frame", "<Frame>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frame", tvwChild, , " src=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frame", tvwChild, , " name=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frame", tvwChild, , " scrolling=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frame", tvwChild, , " marginwidth=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frame", tvwChild, , " framespacing=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frame", tvwChild, , " marginheight=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frame", tvwChild, , " noresize", 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frame", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frame", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frame", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frame", tvwChild, , " title=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frame", tvwChild, , " frameborder=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frame", tvwChild, , " bordercolor=" & Chr(34) & Chr(34), 2)



'add Frameset
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Frameset", "<frameset></frameset>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frameset", tvwChild, , " rows=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frameset", tvwChild, , " cols=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frameset", tvwChild, , " frameborder=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frameset", tvwChild, , " framespacing=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frameset", tvwChild, , " border=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frameset", tvwChild, , " bordercolor=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frameset", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frameset", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frameset", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Frameset", tvwChild, , " title=" & Chr(34) & Chr(34), 2)


'add H1
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "h1", "<h1></h1>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("h1", tvwChild, , " title=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("h1", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("h1", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("h1", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("h1", tvwChild, , " class=" & Chr(34) & Chr(34), 2)



'add H2
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "h2", "<h2></h2>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("h2", tvwChild, , " title=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("h2", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("h2", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("h2", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("h2", tvwChild, , " class=" & Chr(34) & Chr(34), 2)



'add H3
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "h3", "<h3></h3>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("h3", tvwChild, , " title=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("h3", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("h3", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("h3", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("h3", tvwChild, , " class=" & Chr(34) & Chr(34), 2)



'add HR
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "HR", "<hr>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("HR", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("HR", tvwChild, , " size=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("HR", tvwChild, , " color=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("HR", tvwChild, , " width=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("HR", tvwChild, , " noshade", 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("HR", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("HR", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("HR", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("HR", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add Iframe
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Iframe", "<iframe></iframe>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Iframe", tvwChild, , " src=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Iframe", tvwChild, , " name=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Iframe", tvwChild, , " scrolling=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Iframe", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Iframe", tvwChild, , " height=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Iframe", tvwChild, , " width=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Iframe", tvwChild, , " marginwidth=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Iframe", tvwChild, , " marginheight=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Iframe", tvwChild, , " frameborder=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Iframe", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Iframe", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Iframe", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Iframe", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add IMG
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "IMG", "<img>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " src=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " alt=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " border=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " height=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " width=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " hspace=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " vspace=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " ismap=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " usemap=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " dynsrc=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " start=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " loop=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " controls", 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " loopdelay=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " lowsrc=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("IMG", tvwChild, , " class=" & Chr(34) & Chr(34), 2)




'add INPUT
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "INPUT", "<input>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " type=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " name=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " value=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " size=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " maxlength=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " tabindex=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " notab", 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " checked", 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " src=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " border", 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " width", 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " height", 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " vspace", 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " hspace", 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " accept=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("INPUT", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add LINK
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "LINK", "<link>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("LINK", tvwChild, , " rel=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("LINK", tvwChild, , " href=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("LINK", tvwChild, , " type=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("LINK", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add MAP
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "MAP", "<map></map>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("MAP", tvwChild, , " name=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MAP", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MAP", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MAP", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MAP", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add MARAQUEE
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "MARAQUEE", "<marquee></marquee>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " behavior=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " direction=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " bgcolor=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " height=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " width=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " hspace=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " vspace=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " loop=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " scrollamount=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " scrolldelay=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MARAQUEE", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add META
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "META", "<meta>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("META", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("META", tvwChild, , " http-equiv=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("META", tvwChild, , " name=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("META", tvwChild, , " content=" & Chr(34) & Chr(34), 2)


'add MULTICOL
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "MULTICOL", "<multicol>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("MULTICOL", tvwChild, , " cols=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MULTICOL", tvwChild, , " width=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MULTICOL", tvwChild, , " gutter=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MULTICOL", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MULTICOL", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MULTICOL", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("MULTICOL", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add P
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "PARAGRAPH", "<p>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("PARAGRAPH", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("PARAGRAPH", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("PARAGRAPH", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("PARAGRAPH", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("PARAGRAPH", tvwChild, , " title=" & Chr(34) & Chr(34), 2)




'add Param
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "Param", "<param></param>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("Param", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Param", tvwChild, , " name=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Param", tvwChild, , " value=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Param", tvwChild, , " type=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("Param", tvwChild, , " valuetype=" & Chr(34) & Chr(34), 2)




'add SCRIPT
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "SCRIPT", "<script></script>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("SCRIPT", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SCRIPT", tvwChild, , " type=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SCRIPT", tvwChild, , " language=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SCRIPT", tvwChild, , " src=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SCRIPT", tvwChild, , " for=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SCRIPT", tvwChild, , " defer=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SCRIPT", tvwChild, , " runat=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SCRIPT", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SCRIPT", tvwChild, , " charset=" & Chr(34) & Chr(34), 2)



'add SOUND
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "SOUND", "<sound></sound>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("SOUND", tvwChild, , " src=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SOUND", tvwChild, , " loop=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SOUND", tvwChild, , " delay=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SOUND", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SOUND", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SOUND", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SOUND", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add SPAN
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "SPAN", "<span></span>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("SPAN", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SPAN", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SPAN", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SPAN", tvwChild, , " title=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("SPAN", tvwChild, , " align=" & Chr(34) & Chr(34), 2)



'add STYLE
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "STYLE", "<style></style>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("STYLE", tvwChild, , " type=" & Chr(34) & Chr(34), 2)



'add TABLE
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "TABLE", "<table></table>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " cellpadding=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " border=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " valign=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " cellspacing=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " nowrap", 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " background=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " bgcolor=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " bordercolor=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " bordercolorlight=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " bordercolordark=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " cols=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " clear=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " frame=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " rules=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TABLE", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add TD
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "TD", "<td></td>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " bgcolor=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " bordercolor=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " bordercolordark=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " bordercolorlight=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " background=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " width=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " height=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " rowspan=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " colspan=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " valign=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " nowrap=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TD", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add TH
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "TH", "<th></th>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " width=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " height=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " rowspan=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " colspan=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " valign=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " nowrap=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " bgcolor=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " bordercolor=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " bordercolordark=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " bordercolorlight=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " background=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TH", tvwChild, , " title=" & Chr(34) & Chr(34), 2)



'add TR
 Set tempNode = fMainForm.TRV.Nodes.Add(, , "TR", "<tr></tr>", 1)
 Set tempNode = fMainForm.TRV.Nodes.Add("TR", tvwChild, , " bgcolor=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TR", tvwChild, , " bordercolor=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TR", tvwChild, , " bordercolorlight=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TR", tvwChild, , " bordercolordark=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TR", tvwChild, , " align=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TR", tvwChild, , " valign=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TR", tvwChild, , " nowrap=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TR", tvwChild, , " style=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TR", tvwChild, , " id=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TR", tvwChild, , " class=" & Chr(34) & Chr(34), 2)
 Set tempNode = fMainForm.TRV.Nodes.Add("TR", tvwChild, , " title=" & Chr(34) & Chr(34), 2)
 'tempNode.Expanded = False

End Function
