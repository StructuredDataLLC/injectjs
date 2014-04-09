injectjs
========

Javascript macros for Excel via V8

Background
----------

Some of us have to use Excel a lot, and it is a good environment.  Scripting is a bit of a pain, though, 
as the default scripting language is not very friendly.

In the last version of Office, Microsoft added javascript-based "Apps". The [API sample app] (http://office.microsoft.com/en-us/store/api-tutorial-for-office-WA104077907.aspx?queryid=f7205dbe-1072-4b5c-a0da-4ddfa0be7043&css=excel&CTT=1) 
for Excel lets you write inline javascript (and it itself uses [codemirror](http://codemirror.net/) - very slick).
But then it turns out that the Javascript API doesn't let you do much.  For example, you can read values from
a spreadsheet, but not formulae.  

So OK, they designed it for a particular purpose (seems to be doing design and layout), and not as a general
scripting engine.  But since everything else is embedding v8, it seemed like the thing to do.

Is this a good idea?
--------------------

Possibly not.  Scripting in Excel has always been a security problem, and this just makes things worse.  At the 
moment, we're not supporting any commonjs-type extensions, so actual damage abililty is limited to what's built in
to Excel (which is a lot), but it won't run arbitrary extensions.  

Work in progress
----------------

This is very much a work in progress.  It's actually quite usable in 2013.  It kind of works in 2007 & 2010, but it
won't load or save scripts automatically so it's not more than a toy in those environments.

The script host side is a bit messy as I'm out of practice with COM/C++.  It can probably be sped up, and cleaned up,
and it would be good to stop leaking references (marked).

The COM objects are all auto-wrapped, so all methods and properties are available.  COM has a thing where 
property accessors can have arguments (sometimes required).  We treat those as functions so they should work as 
expected.  Events work, possibly not un-hooking correctly.  Indexed accessors are supported for integer indexing,
but not String (or any other) indexes.  Iterators aren't supported, but you can loop.

Missing
-------

Probably most important is autocompletion/tooltips in the editor to simplify writing.  Then automatically translating 
enum values, at the moment they're returned as ints.

Dependencies
------------

+ V8, built as shared libraries 
+ Scintilla.net, dot-net wrapper for Scintilla
+ Latest VSTO


