# FIDE
Source code for Fluke 9010 IDE (FIDE) used to create and edit scripts 

## Additional Resources
See [QuarterArcade Tech Center](http://tech.quarterarcade.com/tech/) for additional details.

# History
This project is not longer maintained, but provided here for reference. The Fluke signature algorithm is in the source code. 
This is probably the most interesting part of the code, as the signature process can be used to validate ROMs. 

Note that the Fluke complier is needed to in order to convert the 9lc source code files into programs that can run
on the Fluke 9010A. The VB application shells out to a DOS window to invoke the compiler. There's no other 
way that I am aware of to compile scripts to binaries (Fluke published some technical details so it is possible
to write a compiler. It just wasn't something I was interested in doing.)
