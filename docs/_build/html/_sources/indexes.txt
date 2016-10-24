==========================
Generating Indexes in Word
==========================

Indexes are, frankly, time-eaters. A good index will take time to set up, and
there are many resources (and opinions) on the topic of just how to choose what
goes into an index, what a good one looks like, etc. Creating a basic one in
Word, however, is fairly straight-forward.

Fields in Word can be used to create automatically updating indexes for
documents. While the graphical user interface (GUI) can be used to generate
automatic indexes, using fields to mark the entries provides much more
flexibility and power.

.. _syntax:

Syntax of Marked Entries
========================
When marking entries, use the ``XE`` field code. The field requires text that
identifies what exactly the entry will be. For example, perhaps commas might be
a topic to tag for an index entry. The index entry for commas might simply be
*Commas*, so each index entry field code would be::

  { XE "Commas" }

no matter whether it would appear with the word *comma* or with a paragraph
mentioning when to use commas.

To get a subentry (i.e., a level 2 entry), use a colon in the text with a main
topic before the colon and the subtopic after the colon.::

  { XE "Entry:subentry" }

The field uses multiple switches, or modifiers.

* ``\b``: the bolding switch
* ``\f``: the type switch
* ``\i``: the italic switch
* ``\r``: the bookmark switch
* ``\t``: the text substitution switch
* ``\y``: the alternative character switch

Bolding Switch
--------------
Use the bolding switch, ``\b``, to bold that page number in the index for the
entry. This switch is often used to indicate where the reader may find the most
important information about an entry.::

  { XE "Entry" \b }

Type Switch
-----------
The type switch, ``\f``, indicates which type of entry the field is marking,
often a consideration when building a list of figures, illustrations, or tables.
This switch requires a single, unique letter as a type identifer; the letter
itself does not necessarily have to match the exact type of entry (e.g., figures
may be identifed with a *g* if necessary).

For example, to indicate a table entry::

  { XE "Entry" \f t }

The related index would then use an entry identifier switch to generate a list
of only tables.

If nothing is specified with this switch, the entry automatically goes into a
generic index.

Italic Switch
-------------
Use the italics switch, ``\i``, to italicize that page number in the index for
the entry.::

  { XE "Entry" \i }

Bookmark Switch
---------------
To provide page ranges in an index, use a bookmark to mark the entire section,
and then use the bookmark switch, ``\r``, to make the entry to span the entire
bookmark's page ranges. This example uses both the bookmark switch and the bold
switch: ::

    { XE "Entry:subentry" \r BookmarkName \b }

Text Substitution Switch
------------------------
To note a different entry for reference, use the ``\t`` switch.

    ``{ XE "Entry:subentry" \t "(``*see*`` Entry 2)" }``

To have an entry appear at a different spot than its alphabetical spot (e.g., a
"see also" note at the end of a list), use a semicolon to prove a separate
alphabetical index term with an empty text substitution switch:

\  { XE "Entry:*See also* Entry 2;zz" \t "" }

In this case, the entry would appear at the end of the list (as *zz* is at the
very end of an alphebetized list since very few, if any, words in most English
indexes would begin with *zz*) as follows:

\  Entry
    [...]
    See also *Entry 2*.

Note that the italics switched. This is a function of how the index is set up.

Alternative Character Switch
----------------------------
The alternative character switch, ``\y``, only is used when at least one East
Asian language is enabled in the document. It sorts the entry by a different
character than the character that begins the entry itself. This different
character, called a *yomi*, provides the ability to phonetically sort
non-English words. So, to get an entry to sort under the letter *D*::

  { XE "Entry" \y d }


Generation of an Index's Entries
================================

There are three methods to create an index: marking entries manually over time
with field codes, marking entries manually over time with the built-in GUI, or
creating a table to automark entries after the document is created.

Marking Entries with Field Codes
--------------------------------
The most powerful way to mark entries is to use the ``XE`` field code and any
related switches. For more information, see :ref:`syntax`.

Marking Entries with Word's Built-In GUI
----------------------------------------
To mark an entry manually, place the cursor where the entry should refer (i.e.,
after a specific word, at the beginning of a section, etc.). Next, go to
``References > Mark Entry`` (under ``Index``). A dialog box will appear. Enter
in the data, and then choose ``OK``. Looking under the paragraph marks will
reveal the field with the ``XE`` field code and the data just entered.

Marking Entries with AutoMark
-----------------------------
An alternate way of marking entries requires the document to be complete and
index entries chosen. By creating a simple Excel spreadsheet with one column for
the word that needs to be marked and one column for the index entry for that
word, the ``AutoMark`` feature can be used. To import the spreadsheet for
automatically marking the index entries, use ``References > Insert Index``
(under ``Index``). Choose ``AutoMark…`` at the bottom right. From the file
dialog box, choose the simple Excel spreadsheet and then choose ``OK``. Word
will mark every instance of the entry's appearance in the document. This can be
detrimental as, for example, every occurrence of the word *enter* will be tagged
"something" if this is used as an index entry. So the phrase "Enter your
password" will be tagged as "Enter ``[index entry field code]`` your password" along
with "the Enter ``[index entry field code]`` key is to the right." Also, if the
words *semicolon* and *semicolons* were in the index list to be tagged as
"Punctuation, semicolon", every instance of the word in a section on semicolons
would be tagged:

    Semicolons ``[index entry field code]`` can be used in place of a
    coordinating conjunction in certain cases. For example, the semicolon
    ``[index entry field code]`` in "Semicolons ``[index entry field code]`` are
    used here; semicolons ``[index entry field code]`` are not used here." is an
    accurate use, but the semicolon ``[index entry field code]`` in "Semicolons
    ``[index entry field code]`` are not like colons; Jimmy likes to place
    semicolons ``[index entry field code]`` everywhere." is not.
