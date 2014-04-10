injectjs
========

Javascript macros for Excel via V8

Background
----------

Some of us use Excel a lot, and it is a good environment.  Scripting is a bit of a pain, though, 
as the default scripting language is not very friendly.

In the last version of Office, Microsoft added javascript-based "Apps". The [API sample app] (http://office.microsoft.com/en-us/store/api-tutorial-for-office-WA104077907.aspx?queryid=f7205dbe-1072-4b5c-a0da-4ddfa0be7043&css=excel&CTT=1) 
for Excel lets you write inline javascript (and it itself uses [codemirror](http://codemirror.net/) - very slick).
But then it turns out that the Javascript API doesn't let you do much.  For example, you can 
read values from a spreadsheet, but not formulas (not formulae, either).  

So OK, they designed it for a particular purpose (seems to be doing design and layout), and not as a general
scripting engine.  But once we got a taste for it, we decided that this was what we wanted.  And since everything 
else is embedding v8, it seemed like the thing to do. 

Is this a good idea?
--------------------

Possibly not.  Scripting in Excel has always been a security problem, and this just makes things worse.  At the 
moment, we're not supporting any commonjs-type extensions, so actual damage abililty is limited to what's built in
to Excel (which is a lot), but it won't run arbitrary extensions.  

Work in progress
----------------

This is very much a work in progress.  It has bugs and (likely) resource leaks.

It's actually quite usable in Excels 2007, 2010 and 2013. The COM objects are all auto-wrapped (late bound), 
so all methods and properties are available.  COM has a thing where property accessors can have arguments 
(sometimes required).  We treat those as functions so they should work as expected.  Events work, possibly 
not un-hooking correctly.  Indexed accessors are supported for integer indexing,
and you can use the Item property (method) for string indexes.  Iterators aren't supported, but you 
can loop.

Missing
-------

Immediate pain points:

+ better autocompletion/tooltips
+ resource utilization
+ file (and module) management

Dependencies
------------

+ [V8](https://github.com/v8/v8), built as shared libraries 
+ [Scintilla.net](https://scintillanet.codeplex.com/), dot-net wrapper for [Scintilla](http://www.scintilla.org/)
+ Latest [VSTO](http://www.microsoft.com/en-us/download/details.aspx?id=40791)

We can probably distribute all of this except for VSTO, let me know if you want a binary.



Using
=====

You need to know javascript and Excel's object model pretty well.  There is only one object exposed in the environment -
"Application", representing the Excel application object.  All other object descend from Application.  There are no 
"magic" fields like ActiveWorkbook (in VB that's just an alias for Application.ActiveWorkbook).  

Enums are exposed in the global namespace, so it's legal to say something like

Application.ActiveCell.Interior.Color = XlRgbColor.rgbCrimson;

Enums and the primary object are hinted in the editor, still working on hinting scoped fields.

Standard javascript methods alert and confirm are supported.  There's a console facility but (atm) it only supports 
log(string...)

Running Scripts
---------------

Scripts are not automatically executed nor are they "live".  Too unstable at the moment.  To excute a script, click Execute.

Loading and Saving
------------------

Scripts are stored in the document's XML.  It doesn't matter what the filetype is (xlsx, xlsm, &c).  Of course you
have to save the file if you make changes, but it will set the dirty flag.

Examples
--------

```javascript
var ws = Application.ActiveWorkbook.Sheets.Add();
ws.Name = "Update";
ws.SelectionChange = function(){
   console.log( Application.ActiveCell.Address());
};
```

```javascript
function search_highlight()
{
	console.log( "Search: " + 
		Application.Selection.Address().toString()
			.replace( /\$/g, "" )
			.replace( /\:/, " => "));

	var min = Application.Evaluate( "=MIN(" + 
		Application.Selection.Address() + ")" );
	console.log( "Min is " + min );
		
	Application.Selection.Font.Color = 0;
	var rng = Application.Selection.Find( min );
	var count = 0;

	if( typeof( rng ) != "undefined" ){ // need coffeescript
		var start = rng.Address();
		while( true ){
			count++;
			rng.Font.Color = XlRgbColor.rgbWhiteSmoke
			rng.Interior.Color = XlRgbColor.rgbDarkRed;
			rng = Application.Selection.FindNext(rng);
			if( rng.Address() == start ) break;
		}
	}

	console.log( "Done; " + count + " found" );
}

search_highlight();
```








