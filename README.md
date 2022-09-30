# Word as native LaTeX editor

* Prepare LaTeX documents for revision by co-authors and supervisors using MS Word (i.e., add pictures of environments, hide LaTeX clutter). The revised .doc (which is in fact a LaTeX file) can be edited further in Word, or copied back to your favourite LaTeX editor (just Ctrl-A, Ctrl-C).
* Use MS Word as basic LaTeX editor (i.e., no file format conversion) 

## Files/Installation

* **Word template with macros:** `LatexEditor.dotm`. Put this template into the start-up folder of MS Word: `C:\Users\<your username>\AppData\Roaming\Microsoft\Word\STARTUP`
* **LaTeX-ribbon-tab:** `Word Customizations.exportedUI`. In Word, import the customization file in to the ribbon: Right-Click on ribbon -> `Customize the Ribbon...`. This is the most straighforward way to install. However, this simple import overwrites other custom ribbons. However, most users have no custom ribbons (ACROBAT and other add-in ribbon-tabs seems to get preserved;-))

![LaTeXribbon](https://user-images.githubusercontent.com/113771815/192064971-68e64966-ec35-418c-a8e7-f0a8dcb03c0c.PNG)

## Features

The VBA-code let you (all functionality through buttons on a new `LaTeX` tab in the Word ribbon):

* Preview LaTeX-environments: `equation`, `align`, `table`, `figure`, etc.
* Format some common Latex-markup to gray or white, and text inside `label`, `(eq)ref`, `cite(tp)` to blue
* Format LaTeX-section markups into Word headings. With this, navigation and folding of parts is possible
* Hide LaTeX-markup (using the hidden text feature of Word)
* Run pdfLaTeX for the whole Word document and display the resulting .pdf
* Delete LaTeX-markup


### Prerequisites

* LaTeX installation (with the usual `latemk` command)
 * Installed LaTeX packages: `standalone`, `preview`, `amsmath`, `amssymb`
* `ghostscript` installation (e.g. from https://www.ghostscript.com/). For previewing, the macro-code first ask you to browse manually for the .exe, or else it searches the usual 32-bit or 64-bit progam folders on the c-drive (i.e. `gswin32c.exe` or `gswin64c.exe`). Info: Ghostcript is used as the .pdf -> .png converter.
* Tick the required References in the VBA-editor of Word (`ALT-F11` -> `References`):
  * VBA Scripting Runtime (or similar named)
  * VBA Regex (or similar named)
  
  Else the code gives errors: Object not found, etc.
* Trust Center of Word: Add location of macro-file to Trusted Location etc.

## Example workflow

* After you drafted your .tex in your preferred LaTeX-editor, open the .tex in Word (MS-DOS import setting e.g. works, or you draft your LaTeX directly in Word
* For a nice viewing:
     * Set paragraph spacing to zero
	 * Align pargraphs to block
	 * LateX-comments must be written with two %% to be seen for formatting (hiding). Hence, comments with single % are comments in Word visible for co-authors
* Run the VBA-macro (accessible through the new LaTeX-Ribbon) to preview: table, equation, figure, align etc.
* Prepare the file for revision in MS Word:  Convert the previews into inline pictures, then hide the markup completely. If some LateX clutter is not recognized by hiding, marke this text and set hidden yourself in Word
* Note: If the pictures (shapes) are converted to inline shapes (shapes is Word jargon) than a control character is added by Word to the text. Hence, delete the shapes first if you want to copy text (e.g. the `RunLaTeX` does this also to run pdftex on the document). 

### Bibliography
There is no preview of bibliography. This can be done in 20 seconds manually. Run the runLaTeX-macro, or use your preferred Latex-editor to produce a .pdf of your work (which has the reference section), then view the .pdf with word, and paste the reference section at the end.

## Example: test.docx

### Original .tex inside Word

![Original](https://user-images.githubusercontent.com/113771815/193218750-05aeb71a-1396-4813-b470-b6da7d0b177e.PNG)



### Prettified .tex inside Word

![Prettified_1](https://user-images.githubusercontent.com/113771815/193218780-43bcc801-f94e-4846-af6f-27233c4735a8.PNG)

![Prettified_2](https://user-images.githubusercontent.com/113771815/193218822-ec1c9d33-f94b-4e68-b4c6-4ecec306df72.PNG)



## Limitations/Bugs

* Tested on MS Windows 10 and Word 2016.
* Currently, you can preview and delete all generated pictures at once only
* Not all markup gets recognized, and labels, refs, and cites get not deleted
* Equation numbering is basic (align is not yet recognized with multiple numbers)
* The snippets when converted to inline-shapes do not wrap properly on a line on their own 
* Only ASCII character files work surely. The conversion is not fully UTF compatible.
* A MsbBox still says that you can browse to `ghostscript` already installed by the common LaTeX distributions `TeXLive` or `MikTeX`, but this (truncated) versions seem not to work





