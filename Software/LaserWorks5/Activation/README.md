Activating RDCAM
================

> Do you recognize this screenshot? It means that you're trying to start the program for the first time and you don't have a board or your system doesn't recognize it.

> ![Please link machine][1]

> RDCAM should activate automagically when detects the board. However, if you're willing to contribute on this project and you don't have such board, you can do this trick to activate the program. This recipe is valid for Windows and Unix/**Linux**.

1. Find the `%AppData%` path:

   Unix/**Linux**:

        wine cmd /c set 2>/dev/null | grep '^APPDATA'

   Windows:

        ECHO %AppData%

2. Copy the activation file:

   You can find the activation file on this folder. Simply copy it to your `%AppData%` folder. Now the program should open up correctly.


Why it works?
=============
It isn't a `dll` file. It's only a raw binary.

Here you're a python version of the activation checking routine:

     import os

     size_in_bytes = os.path.getsize(os.getenv('APPDATA') + 'r5.dll')

     if size_in_bytes is not 10:
         print("Please link machine")
     else:
         print("The activation file is a good one!")

So any file with a length of 10 bytes is a valid activation dll.

[1]: https://cloud.githubusercontent.com/assets/11387611/10869358/dbd3234a-80ad-11e5-9668-0386d3759eb6.png
