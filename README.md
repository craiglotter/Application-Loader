Application Loader
==================

Application Loader is a simple program that checks the registry for a specified version of the .NET framework. If the .NET framwork appears to be installed, Application Loader will then launch the program specified in the program's input file. Application Loader is useful as a loader for .NET-enabled programs.

Created by Craig Lotter, September 2005

*********************************

Project Details:

Coded in Visual Basic 6 using Visual Studio 6
Implements concepts such as Registry manipulation and File handling.
Level of Complexity: Simple

*********************************

Update 20070921.06:

- Added new key checking. Now checks HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\.NETFramework\v*.*. This picks up .NET Framework 3
