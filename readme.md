#EasyOpenXML - easily parse OOXML document

##Overview
EasyOpenXML is an api to parse word document easily. It use OpenXML api from Microsoft and it is intended to be used in a server environment (so you won't see any interop here)

##Summary
[What the api can do now, and future updates] (#current-status-and-future-updates)
[OOXML Format quick guide] (#ooxml-quick-guide)
[Word content control] (#word-content-control)
[Methods] (#methods)
[Access time] (#access-time-and-optimisation)

##Current status and future updates
This API is written in C#, and first used in an asp.NET project, this is why I don't use interop (interop need office suite installed on the computer, and microsoft doesn't support it on server environment)
For now, the api can only get text and images from word's content control (more on next chapter). Some part of the code used is not from me.
In the future, I'll develop a parser that output a JSON string and do not require content control.
Anyway, feel free to push updates and develop this project.

##OOXML Quick guide
Please note that it's better to understand basics of xml and xml namespace!
Word (and every other office suite's file) are compressed folders.
When you open this folder, you have a bunch of other files and folders, the one interesting is here: yourfile.docx\word
Here you have everything you need to start:
document.xml contain the body part of the document
other files speak for themselves (endnotes, footer, footnotes, etc)

Basically, if you don’t find any documentation about some content you’re trying to extract, just open a blanc word, put only the content you’re trying to extract, and then observe the file document.xml

```xml
<w:p>
</w:p>
```
This is a paragraph, if it’s empty, the tag will be closed
```xml
<w:p/>
```
Paragraph is a basic element of word, visually it’s what you’re inserting when you hit enter on your keyboard.


###Paragraph’s children tags
```xml
<w:p>
    <w:pPr>
        <w:pStyle w:Val=”Title1”/>
    </w:p>
    <!—some other tags-->
</w:p>
```
pStyle is a text style. It appear only if there is applied to the text. Val is indicating which style is applied. Be careful: custom style have also custom Val, and localization for word also change this Val (for example, in French, Val=”Titre1”)

```xml
<w:p>
    <w:r>
        <w:t>My text</w:t>
    </w:r>
</w:p>
```
This is a text tag.

```xml
<w:p>
    <w:r>
        <w:drawing>
            <a:graphic>
                <a:blip r:embed=”rId”>
                </a:blip>
            </a:graphic>
        </w:drawing>
    </w:r>
</w:p>
```
This is an image. Images are in a subfolder of your word file. There is two way to get pictures from your word file: 

direct access to the document folder
yourdocument.docx\word\media

or use the blip. Blip is a relationship between the xml sheet and the image



###Text formatting
```xml
<w:r>
    <w:br/>
</w:r>
```
This is a new line (not a new paragraph). This what appear when you hit " shift + enter " on your keyboard

```xml
<w:r>
    <w:rPr>
        <w:color w:val=’’ FFFFFF’’/>
        <w:highlight w:val=’’red’’/>
    </w:rPr>
  <w:t>white text on red background</w:t>
</w:r>
```
w:color is text color while highlith is background color.

text formatting is not in a different paragraph, but in a different r tag:
ex: my text is *colored*
```xml
<w:p>
  <w:r>
    <w:t>my text is</w:t>
  </w:r>
  <w:r>
    <w:t>colored</w:t>
  <w:r>
<w:p>
```
stay aware that text color is nested in r tag!

###Table
```xml
<w:tbl>
```
This is a table.
It's not a paragraph's child! it's replacing a paragraph tag

###Table's children
```xml
<w:tr>
```
like in html, this is a line


```xml
<w:tbl>
    <w:tr>
        <w:tc>
            <w:p>
                <w:r>
                    <w:t>this is text in line 1 column 1</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:p>
                <w:r>
                    <w:t> this is text in line 1 column 2</w :t>
                </w:r>
            </w:p>
        </w:tc>
    <w:tr>
</w:tbl>
```
Here, instead of html tag td and th, you'll have only tc, for columns. there is no header like html. Header are just text style or text formating.


##Word content controls

##Methods
###GetText

###GetImage

###GetRowString

##Access time and optimisation

