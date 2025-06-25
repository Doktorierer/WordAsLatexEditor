# Load LaTeX documents into MS Word with rendered math-environments, figures, and tables


> Make a LateX document easily readable in MS Word, without any file-conversion. 



## Purpose

- This Word-plugin prepares a LaTeX documents in MS Word such that revision can be done using MS Word itself
- Optional: Use MS Word as a LaTeX editor


## Features

- LaTeX-environments are added as pictures in MS Word: Figure, table, math environments (such as equation, align etc.), and LaTeX markup can be hidden or greyed out, by making use of the formatting/hiding features of MS Word 
- The prettified .docx (which is in fact a LaTeX file) can be edited further in MS Word, or copied back to your favourite LaTeX editor (just Ctrl-A, Ctrl-C).
 


## Files/Installation

* **MS Word template with macros:** `LatexEditor.dotm`. Put this Word template into the start-up folder of MS Word: `C:\Users\<your username>\AppData\Roaming\Microsoft\Word\STARTUP`
* **LaTeX-ribbon-tab:** `Word Customizations.exportedUI`. In Word, import the customization file in to the ribbon: Right-Click on ribbon -> `Customize the Ribbon...`.  This simple import procedure overwrites other custom ribbons; however, most users have no custom ribbons (ACROBAT and other "professional" add-in ribbon-tabs seems to get preserved).

![LaTeXribbon](https://user-images.githubusercontent.com/113771815/192064971-68e64966-ec35-418c-a8e7-f0a8dcb03c0c.PNG)


## Features

The macros in the MS Word template let you (all functionality through buttons on the new `LaTeX` tab in the Word ribbon):

- Preview LaTeX-environments (in current version, all environements are previewed at once): `equation`, `align`, `table`, `figure`, etc. (see full list below)
- Format some common LaTeX-markup to gray or white, and format text inside `label`, `(eq)ref`, `cite(tp)` to blue (see full list below)
- Format LaTeX-section markups into MS Word headings (see full list below). With this, navigation and section-folding is possible by using MS Word's navigation pane, as well as re-ordering whole sections 
- Hide LaTeX-markup (by using the "hidden text" feature of MS Word). See full list of markup below.
- Run pdfLaTeX for the whole Word document and display the resulting .pdf
- Delete LaTeX-markup
- Unicode <-> TeX-conversion of a few characters/markup items (see full list further below): 
   


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

1. After you drafted your .tex in your preferred LaTeX-editor, open the .tex in MS Word
 * The `MS-DOS` import setting works for example (other import settings may work, too)
 * Optionally: You draft your LaTeX directly in MS Word
2. For a nice viewing in MS Word:
     * Set paragraph spacing to zero
	 * Align paragraphs to block
	 * LaTeX-comments must be written with two %% to be hidden after the special formatting commands (hiding). Hence, comments with single % are comments in MS Word visible for co-authors
3. Run the corresponding preview command (accessible through the new LaTeX-Ribbon) to insert tables, figures and equations as pictures
4. Prepare the file for revision in MS Word:  
  - Convert the preview-images into inline pictures (button on LaTeX-Ribbon), then hide the markup completely.
5. If some LaTeX clutter is not recognized by hiding, marke this text and set hidden yourself in MS Word

### Note: MS Word programming jargon
- A picture in MS Word can be in the so-called `canvas` layer, then this picture is called a `shape`
- Alternatively, a picture can be in the in the so-called `document` layer (inline with text), then it is called an `inline shape`.
- If the generated shapes are converted to inline shapes, then a control character is added by Word into the text neaerby.
	- Hence, if you want to copy-back your LaTeX-code to another editor, delete the inline shapes first (which deletes also the hidden control character)
	- Indeed, the LaTeX-Ribbon-command `RunLaTeX` deletes also all inline shapes in order to be able to to run `pdftex` on the document. 


### Bibliography

There is no preview of bibliography, because this can be done in a minute manually: 
1. Run the runLaTeX-macro, or use your preferred Latex-editor to produce a .pdf of your work (which has the reference section)
2. Then view the .pdf with word, and paste the reference section at the end


## Example: test.docx

### Original .tex inside Word

![Original](https://user-images.githubusercontent.com/113771815/193218750-05aeb71a-1396-4813-b470-b6da7d0b177e.PNG)


## Prettified .tex inside Word

This is the same, valid .tex as before, but just with some text formatted as hidden in MS Word, and title, section etc. formatted by Word's `Title`, `Heading 1`, ec., and the environments previewed as pictures (either in the canvas layer (as so-called shapes in MS Word), or in the document layer (as inline shapes).

![Prettified_1](https://user-images.githubusercontent.com/113771815/193218780-43bcc801-f94e-4846-af6f-27233c4735a8.PNG)

![Prettified_2](https://user-images.githubusercontent.com/113771815/193218822-ec1c9d33-f94b-4e68-b4c6-4ecec306df72.PNG)


### List of recognized LaTeX markup 

The following LaTeX markup is recognized for gray-out coloring or hiding (* = zero or more characters):

- \ref{*}, \eqref{3}, \cite[*], \label{*}
- \begin{equation*} * \end{equation*}
- \begin{table*}*\end\{table*}
- \begin{figure*}*\end\{figure*}
- \begin{align*}*\end\{align*}
- \begin{multline*}*\end\{multline*}
- \begin{gather*}*\end\{gather*}
- \begin{alignat*}*\end\{alignat*}
- \begin{flalign*}*\end\{flalign*}
- \begin{abstract}*\end{abstract}
- \begin{document}, \end{document}
- \begin{frontmatter}, \end{frontmatter}
- \begin{keyword}, \end{keyword}
- \section*{, \subsection*{, \subsubsection*{
- \usepackage*{*}
- \documentclass*{*}
- \maketitle
- \begin{itemize}, \end{itemize}
- \begin{enumerate}, \end{enumerate}
- \ead{, \corref{*}, \cortext*{, \address{, \journal{*}
- \medskip, \bigskip, \smallskip, \sep
- \bibliography{*}
- \title{
- \author{, \date\{
- %%*^13  (this is a comment line, starting with %%)


The following LaTeX environment are compiled and their picture is added (* = zero or more characters; the * inside braces is literal):


- \begin{equation}*\end{equation}
- \begin{table}*\end{table}
- \begin{figure}*\end{figure}
- \begin{align}*\end{align}
- \begin{multline}*\end{multline}
- \begin{gather}*\end{gather}
- \begin{alignat}*\end{alignat}
- \begin{flalign}*\end{flalign}
- \begin{equation*}*\end{equation*}
- \begin{table*}*\end{table*}
- \begin{figure*}*\end{figure*}
- \begin{align*}*\end{align*}
- \begin{multline*}*\end{multline*}
- \begin{gather*}*\end{gather*}
- \begin{alignat*}*\end{alignat*}
- \begin{flalign*}*\end{flalign*}


The text in the following markup gets a word-style font:

- \section -> Heading 1 (MS Word style)
- \subsection -> Heading 2
- \subsubsection -> Heading 3
- \subsubsubsction -> Heading 4
- \title -> Title
- \author, \date -> Subtitle


The following LaTeX-items can be interchanged with UTF characters (4-digit numbers):

- "2022" <-> "\item<space>" (UTF: bullet)
- "00A0" <->  "~"  (UTF: non-breaking white-space)
- "2014" <->  "---" (UTF:em-dash)
- "2013" <->  "--" (UTF: en-dash)
- "201C" <-> "``"
- "00AB" <-> "``" 
- "201D" <-> "''" 
- "00BB" <-> "''" 
- "2018" <-> "`" 
- "2019" <-> "'"  
- "2026" <-> "\textellipsis" (UTF: ...)


## Caveats

-  Make sure to use only `convert to gray' before previewing (that this, do not hide the LaTeX markup), such that the blue highlighting is not shifted
- Only LaTeX files which consists of ASCII characters only work surely (not UTF files).  


## Path to the ghostscript executable

You will be asked whether to search for the path yourself (_path_ includes filename), or the program tries to search in "C:\Program Files\gs\gs*\bin\gswin64c.exe" or "C:\Program Files (x86)\gs\gs*\bin\gswin32c.exe".



## Limitations/Bugs

- Currently, you can preview and delete all generated pictures only together at once, which may take some time
- Only most common LaTeX-markup is recognized (labels, refs, and cites get not fully deleted)
- Equation numbering is basic (e.g. `align` gets not multiple equation numbers)
- The snippets of the environments, when converted to inline-shapes, do not wrap properly on a line on their own 



## Auxiliary info

- To edit the Word-template `.dotm`, do not double-click in Windows Explorer (this opens a derived new document), but use the Open-dialog from Word
- For conversion of figures and tables, you need the software `ghostscript`. in fact, a ghostscript executable is also delivered with by the common LaTeX distributions `TeXLive` or `MikTeX`, but these are somehow truncated versions seems not to work: in Texlive: "C:\...\texlive\<year>\tlpkg\tlgs\bin\gswin64c.exe"; in MikTeX: "...\MikTex\...\mgs.exe".






