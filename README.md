butil
=====

This is a set of VBA utility modules you can use for your project

I always needed a bunch of utility modules that I make and I make them a lot of times. 'Where's the code reuse' I ask? So I created a project housing my utilities and so *butil* was born. The word *butil* is b + util which signifies basic utilities but it also means seed in Filipino, a nice occurence.

quick start
====

This is a <a href="https://github.com/FrancisMurillo/chip">chip</a> project, so you can download this via *Chip.ChipOnFromRepo "Butil"* or if you want to install it via importing module all utility modules. I suggest using the chip as it is more canonical to the project.

And include in your project references the following, not really needed but this is needed to run chip.

1. **Microsoft Visual Basic for Applications Extensibility 5.3** - Any version would do but it has been tested with version 5.3
2. **Microsoft Scripting Runtime**

You should see a bunch of *Util modules all over your project. Check out the list of available utilities on the next section

utilities
====

So far here are the utilities made.

1. **ArrayUtil** - A set of array utilities for handling arrays canonically, a must for most of my projects
2. **AssertUtil** - Nothing fancy here, just some methods to help with assertion in particular vase 
3. **SheetUtil** - A set of sheet manipulation methods which happen a lot.
4. **RangeUtil** - Likewise supporting SheetUtil is this bad boy handling ranges
