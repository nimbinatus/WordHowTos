====================================
Generating Table of Contents in Word
====================================

Fields in Word can be used to create automatically updating Table of Contents
(TOCs) for documents. While the graphical user interface (GUI) can be used to
generate automatic TOCs (``Reference > Table of Contents > Insert``), creating
the TOC with a field is more flexible and more powerful.

Syntax of the Table of Contents
===============================

The field code is ``TOC``. The field switches, or modifications to the field,
are generally straightfoward, but there are some that may be confusing:

* ``\a``: the identifier switch
* ``\b``: the bookmark switch
* ``\c``: the sequence switch
* ``\d``: the chapter seperator switch
* ``\f``: the entry identifier switch
* ``\h``: the hyperlink switch
* ``\l``: the levels switch
* ``\n``: the page number exclusion switch
* ``\o``: the headings switch
* ``\p``: the entry-number separator switch
* ``\s``: the sequence numbering switch
* ``\t``: the custom style switch
* ``\u``: the applied style switch
* ``\w``: the tab preservation switch
* ``\x``: the newline preservation switch
* ``\z``: the web layout view switch

Depending on how you intend to use the TOC, you may or may not use all the
switches.

Identifier Switch
-----------------
Use the identifer switch, ``\a``, to generate a TOC with the labels provided in
captions. This is the switch to use when generating a Table of Figures and
Tables. For example, ::

  { TOC \a tables }

generates a Table of Tables.

In practice, the ``\c`` switch is more common and customizable.

Bookmark Switch
---------------
Use the bookmark switch, ``\b``, in conjunction with a bookmark to generate TOCs
for small sections of a document, such as a section of a larger document. This
switch is especially useful for maintaining inner TOCs inside of sections or for
generating special formatting for different sections of a larger TOC. This
switch requires a bookmark name to be provided after it.::

  { TOC \b bookmark_name }

Note the lack of quotations around the bookmark name.

Sequence Switch
---------------
Use the sequence switch, ``\c``, to generate a TOC of content marked with a
``SEQ`` field. Sequences are most commonly found with captions, but the ``SEQ``
field can be used for any ordered content. For example,::

  { TOC \c "section" }

provides a TOC populated by any heading containing an ``SEQ`` field with a
``section`` identifier, which may be useful for high-level TOCs.

Note the quotation marks around the sequence identifer.

Chapter Separator Switch
------------------------
The chapter seperator switch (``\d``) must be used with the sequence numbering
switch (``\s``).::

  { TOC \s chapter \d "." }

This field produces the following TOC: ::

  Entry 1 in Chapter 1............1.1
  Entry 2 in Chapter 1............1.5
  Entry 3 in Chapter 2............2.9
  . . .

If the chapter seperator switch is not provided with the sequence numbering
switch, Word defaults to using a dash (``-``).

Entry Identifer Switch
----------------------
Use the entry identifer switch (``\f``) when ``TC`` fields are used in the
document. One situation where this switch may be helpful is in a TOC for an
appendix full of tables without captions. The tables may be tagged with a ``TC``
field with an entry field of type *t*, and then an appropriate TOC may be made
(with the appendix itself bookmarked as *appendix*)::

  { TOC \b appendix \f t }

Note the lack of quotation marks around the entry identifer.

Hyperlink Switch
----------------
The hyperlink switch (``\h``) just makes each TOC entry a clickable link. It is
highly recommended for any Word doc that is provided in digital form. Use is
very simple: ::

  { TOC \h }

Levels Switch
--------------
The levels switch, ``\l``, defines one way to indicate the TOC levels in the
document. This switch requires the use of ``TC`` fields to mark entries in a
document, and those ``TC`` fields identify the outline levels for each entry.
To show only those ``TC`` entries that are at levels 1 and 2, use::

  { TOC \l 1-2 }

Note that a range must be provided every time, so there must be a range of
``1-1`` for a TOC containing only level 1 entries.::

  { TOC \l 1-1 }

Note the lack of quotations around the range.

Page Number Exclusion Switch
----------------------------
Use the page number exclusion switch, ``\n``, to remove page numbers at certain
levels of a TOC. The field::

  { TOC \n 2-2 }

would generate a TOC without page numbers at level 2. To remove all page numbers
at all ranges, use the switch with no range: ::

  { TOC \n }

Note the lack of quotations around the range.

Headings Switch
---------------
The headings, or outline, switch, ``\o``, defines one way to indicate the TOC
levels in the document. This switch uses built-in heading styles to identify
levels for the TOC. To show only level 1 and 2 headings in the TOC, use::

  { TOC \o "1-2" }

Note that a range must be provided every time, so there must be a range of
``1-1`` for a TOC containing only heading 1 entries.::

  { TOC \o "1-1"}

Note the quotations around the level range.

Entry-Number Separator Switch
-----------------------------
Use the entry-number separator switch, ``\p`` to customize which character
separates the TOC entry and the page number. For example, to have a line of
dashes separate the entry and the page number, use ::

  { TOC \p "-" }

which provides the following TOC: ::

  Entry 1---------------1
  Entry 2---------------9
  . . .

Sequence Numbering Switch
-------------------------
Use the sequence numbering switch, ``\s``, to provide the sequence number along
with the page number. This switch requires the use of ``SEQ`` fields in the
document. For example, to provide the chapter number along with the page number,
the ``SEQ`` field must be used on each chapter (in this example, with an
identifer of *chapter*), and then that identifer used in the TOC field.::

  { TOC \s chapter }

This field produces the following TOC: ::

    Entry 1 in Chapter 1............1-1
    Entry 2 in Chapter 1............1-5
    Entry 3 in Chapter 2............2-9
    . . .

Note the lack of quotation marks around the identifer.

Custom Style Switch
-------------------
Use custom style switch, ``\t``, to specify which custom styles belong at which
level in the TOC. For example, if there are custom styles called *generated* and
*jarvis* in the document's style sheet, use the following field to place those
styles at levels 2 and 3, respectively: ::

  { TOC \t "generated,2, jarvis,3"}

To have the built-in Heading 1 style as level 1 and *generated* as level 2, use
a mixture of the ``\o`` and ``\t`` switches: ::

  { TOC \o "1-1" \t "generated,2" }

In practice, this may be useful if a user creates custom styles that are easier
to remember (e.g., styles named for section headings) rather than using built-in
styles.

Applied Outline Switch
----------------------
The applied outline switch, ``\u``, adds to the TOC paragraphs that are directly
formatted as outlined at different levels through the ``Paragraph`` dialog. It
is mainly used with the Document Map function of Word.::

  { TOC \u }

In practice, this is rarely, if ever, used. If random text has begun to appear
in a TOC, check the field for this switch, and remove this switch if it appears.

Tab Presentation Switch
-----------------------
The tab presentation switch, ``\w``, simply prevents Word from stripping tab
characters from the text of entries. For example, a heading of ::

  Heading     Tabbed

would need the field to be::

  { TOC \w }

to generate a TOC of::

  Heading     Tabbed............1

rather than::

  Heading Tabbed................1

In practice, this is rarely, if ever, used.

Newline Presentation Switch
---------------------------
The newline presentation switch, ``\x``, simply prevents Word from stripping tab
characters from the text of entries. For example, a heading of ::

  Heading
  New Line

would need the field to be::

  { TOC \x }

to generate a TOC of::

  Heading
  New Line........................1

rather than::

  Heading New Line................1

In practice, this is rarely, if ever, used.

Web Layout View Switch
----------------------
The Web Layout View switch, ``\z``, just hides the page numbers and the
seperators from the screen when the document is in Web Layout View.::

  { TOC \z }

In practice, this is rarely, if ever, used. Word often adds it to TOCs created
with the GUI.

Syntax of Marked Entries
========================
When explicitly marking entries for the TOC, use the ``TC`` field. The field
requires the use of a name, provided in quotation marks, that will appear as is
in the TOC itself: ::

  { TC "Name of Entry" }

which would appear in the TOC as::

  Name of Entry..........1

This field has very few switches:

* ``\f``: the type switch
* ``\l``: the level switch
* ``\n``: the numbering switch

Type Switch
-----------
The type switch, ``\f``, indicates which type of entry the field is marking,
often a consideration when building a list of figures, illustrations, or tables.
This switch requires a single, unique letter as a type identifer; the letter
itself does not necessarily have to match the exact type of entry (e.g., figures
may be identifed with a *g* if necessary).

For example, to indicate a table entry::

  { TC \f t }

The related TOC would then use an entry identifier switch to generate a list of
only tables.

If nothing is specified with this switch, the entry automatically goes into a
generic TOC.

Level Switch
------------
The level switch indicates which outline level the entry is.::

  { TC \l 2}

means the entry is a level 2 entry in the TOC outline.

Numbering Switch
----------------
Use numbering switch, ``\n``, to prevent the entry from displaying a page number
in the TOC.::

  { TC \n }

The default is to omit this switch, providing the page number in the TOC.
