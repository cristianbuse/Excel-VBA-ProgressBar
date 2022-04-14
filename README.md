# Excel-VBA-ProgressBar
Flexible Progress Bar for Excel

Related [Code Review question](https://codereview.stackexchange.com/questions/273741/progress-bar-for-excel)

The Progress Bar from this project has the following features:
 - Works on both Windows and Mac
 - The user can cancel the displayed form via the X button (or the Esc key), if the ```AllowCancel``` property is set to ```True```
 - The form displayed can be Modal but also Modeless, as needed (see ```ShowType ``` property)
 - The progress bar calls a 'worker' routine which:
   - can return a value if it's a ```Function```
   - accepts a variable number of parameters and can change them ```ByRef``` if needed
   - can accept the progress bar instance at a specific position in the parameter list but not required
   - can be a macro in a workbook (see ```RunMacro``` ) or a method on an object (see ```RunObjMethod```)
 - Has the ability to show how much time has elapsed and an approximation of how much time is left if the ```ShowTime``` property is set to ```True```
 - The userform module has a minimum of code (just events that are going to get raised) and has no design time controls which makes it easily reproducible

## Installation
Just import the following code modules in your VBA Project:
* [ProgressBar.cls](https://github.com/cristianbuse/Excel-VBA-ProgressBar/blob/master/src/ProgressBar.cls)
* [ProgressForm.frm](https://github.com/cristianbuse/Excel-VBA-ProgressBar/blob/master/src/ProgressForm.frm) (you will also need the [ProgressForm.frx](https://github.com/cristianbuse/Excel-VBA-ProgressBar/blob/master/src/ProgressForm.frx) when you import) - Alternatively, this can be easily recreated from scratch in 3 easy steps:
  1. insert new form
  2. rename it to ```ProgressForm```
  3. add the following code:
      ```VBA
      Option Explicit

      Public Event Activate()
      Public Event QueryClose(Cancel As Integer, CloseMode As Integer)

      Private Sub UserForm_Activate()
          RaiseEvent Activate
      End Sub
      Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
          RaiseEvent QueryClose(Cancel, CloseMode)
      End Sub
      ```
      
 You will also need:
* **LibMemory** from the [submodules folder](https://github.com/cristianbuse/Excel-VBA-ProgressBar/tree/master/submodules) or you can try the latest version [here](https://github.com/cristianbuse/VBA-MemoryTools/blob/master/src/LibMemory.bas)

Note that ```LibMemory``` is not available in the Zip download. If cloning via GitHub Desktop the submodule will be pulled automatically by default. If cloning via Git Bash then use something like:
```
$ git clone https://github.com/cristianbuse/Excel-VBA-ProgressBar
$ git submodule init
$ git submodule update
```
or:
```
$ git clone --recurse-submodules https://github.com/cristianbuse/Excel-VBA-ProgressBar
```

## Demo

Import the following code modules:
* [Demo.bas](https://github.com/cristianbuse/Excel-VBA-ProgressBar/blob/master/src/Demo/Demo.bas) - run ```DemoMain```
* [DemoClass.cls](https://github.com/cristianbuse/Excel-VBA-ProgressBar/blob/master/src/Demo/DemoClass.cls)

There is also a Demo Workbook available for download.

## License
MIT License

Copyright (c) 2022 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.