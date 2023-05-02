# Visio_macros
Microsoft Visio macros written in VBA

This VBA file contains four macros.  Their function is to move the pin location of a shape, *while keeping the shape in the same location*.
This is useful if you want to use the pin location to align several nonidentical shapes, etc.

Note the included macros only move the pin location to the four corners of the object.  Extension to the side midpoints should be straightforward - I didn't do it because I don't find that as useful.


Limitation:  This has not been tested on objects that have been flipped or rotated.

Further work:  It would be great to be able to assign these macros to ribbon buttons, but that requires code beyond VBA.
