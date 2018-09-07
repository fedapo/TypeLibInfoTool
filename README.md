# TypeLibInfoTool
Application to inspect a COM TypeLib and generate IDL and .manifest files. The
tool is made as an HTML application using Javascript for the logic.

This tool has been presented for the first time in a CodeProject article
[Generating IDL and Manifest Files with TypeLibInfoTool](http://www.codeproject.com/KB/COM/TypeLibInfoTool.aspx)
that I wrote in 2011.

## Introduction

A tool is presented to inspect the interface of any COM Type Library 
and generate its IDL file or a fragment of a .manifest file that can be 
used for deploying registration-free applications.

The functionality is similar to many of its predecessors, e.g. the 
Object Browser of the VB6 IDE, Microsoft OLE View, and many others. 
Compared to these it adds the simplicity and openness of a scripting 
language and some extra functionality that helps writing .manifest files
needed for Registration-Free COM (see [\[1\]](#ref1)).

## Background

Type Libraries are central to the COM technology. Each COM component 
itself is represented, in its interface, by a Type Library. These are 
implemented as files of several types.

* Dynamic Link Libraries (*.dll). This is probably the most common form taken by COM components.
* ActiveX Components (*.ocx). They are a special case of DLL with specific machinery to embed the objects it instantiates within a GUI.
* Executables (*.exe). Stand-alone executables can also expose a COM interface. This is often done to give the ability to programmatically drive an application from another one.
* Type Libraries (*.tlb). No code is contained within this kind of files, this is a binary form of IDL file and can be used to import interfaces or reference the objects in a Type Library when building new components.

This tools makes use of the _TypeLib Information Object Library_ which
presents an API for browsing type libraries. A sample usage of this library can
be found in this article ([\[2\]](#ref2)). The TLI component itself is
implemented as a type library, an interesting exercise is inspecting it with
the application presented here (the file is _TlbInf32.dll_ and is usually
located in _%SystemRoot%\system32_.)

The other main technology this tool relies on is Microsoft HTML 
Applications (HTA). This gives flexibiliy to users who decide to dig 
into the source code as it is immediately available as HTML and 
Javascript in the same file that gets executed.

## Usage

The TLI component is not a standard part of a Windows install. It gets shipped
with other pieces of software, such as Microsoft Visual Studio. First of all,
make sure you have _TlbInf32.dll_ registered on your system. If not, get a copy
of it and register it in the usual way: go to the folder where you have it
installed (again, usually it is _C:\WINDOWS\system32_) and type the following
in the command line.

`regsvr32 TlbInf32.dll`

After this step you are set to go. Double-click on the file
_TypeLibInfoTool.hta_ and the application will start. Here is a snapshot of it.

![screenshot](images/com_typelib_tool.png =417x450)

From this point on three operations are possible.
  
* Generating an IDL File
* Registration Info for .manifest Files
* Summary of Information

### Generating an IDL File

The IDL file is an essential part when building a COM component. It 
defines the set of all interfaces, coclasses, enums, etc. that make up 
the interface of the type library. They are commonly used in C/C++ 
projects for this purpose. It is often necessary to inspect the IDL file
that was originally used to make a COM component, while this is seldom 
available it can be reconstructed. Moreover, VB6 projects that result in
COM components do not make explicit use IDL files, this is often a 
problem as some essential information is kept hidden. Again, the ability
to reconstruct the IDL file is critical. An interesting web page on 
this topic can be found here [\[3\]](#ref3).

The TypeLibInfoTool allows this operation, by clicking on the 
Generate IDL button the IDL file can be reconstructed out of a binary 
file (dll, ocx, exe, tlb).

### Registration Info for .manifest Files

Applications traditionally use COM components by looking up in the 
registry for the place where it is located (and other information). This
is possible as every component goes through a registration phase where 
its information gets written to the registry. This is usually 
accomplished when installing a piece of software. This has always been 
the model used by OLE/COM. While perfectly all right for many purposes, 
sometimes a one-click deployment is more desirable. More recently, 
starting from Windows XP another approach can be taken (see [\[4\]](#ref4)).
A special file can be present in the same folder as the executable 
which contains all registration information, the file has the same name 
as the executable with a .manifest extention appended. The OS checks for
the presence of this file and uses it before looking up the Windows 
registry. This removes the need for the regitration phase.

A manifest file is written as XML format and has a format which can 
be hard to manually compose. The proposed application generates the 
fragment of a manifest file that contains the required registration 
information for a specific component. For instance, consider the 
following fragment.

```xml
<file name='C:\WINDOWS\system32\msscript.ocx'>
  <typelib tlbid='{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}' version='1.0' flags='' helpdir=''/>
  <comClass progid='MSScriptControl.Procedure' clsid='{0E59F1DA-1FBE-11D0-8FF2-00A0D10038BC}' threadingModel='Apartment'/>
  <comClass progid='MSScriptControl.Procedures' clsid='{0E59F1DB-1FBE-11D0-8FF2-00A0D10038BC}' threadingModel='Apartment'/>
  <comClass progid='MSScriptControl.Module' clsid='{0E59F1DC-1FBE-11D0-8FF2-00A0D10038BC}' threadingModel='Apartment'/>
  <comClass progid='MSScriptControl.Modules' clsid='{0E59F1DD-1FBE-11D0-8FF2-00A0D10038BC}' threadingModel='Apartment'/>
  <comClass progid='MSScriptControl.Error' clsid='{0E59F1DE-1FBE-11D0-8FF2-00A0D10038BC}' threadingModel='Apartment'/>
  <comClass progid='MSScriptControl.ScriptControl' clsid='{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}' threadingModel='Apartment'/>
</file>
<comInterfaceExternalProxyStub
  name='IScriptProcedure'
  iid='{70841C73-067D-11D0-95D8-00A02463AB28}'
  proxyStubClsid32='{00020424-0000-0000-C000-000000000046}'
  baseInterface='{00000000-0000-0000-C000-000000000046}'
  tlbid = '{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}'/>
<comInterfaceExternalProxyStub
  name='IScriptProcedureCollection'
  iid='{70841C71-067D-11D0-95D8-00A02463AB28}'
  proxyStubClsid32='{00020424-0000-0000-C000-000000000046}'
  baseInterface='{00000000-0000-0000-C000-000000000046}'
  tlbid = '{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}'/>
<comInterfaceExternalProxyStub
  name='IScriptModule'
  iid='{70841C70-067D-11D0-95D8-00A02463AB28}'
  proxyStubClsid32='{00020424-0000-0000-C000-000000000046}'
  baseInterface='{00000000-0000-0000-C000-000000000046}'
  tlbid = '{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}'/>
<comInterfaceExternalProxyStub
  name='IScriptModuleCollection'
  iid='{70841C6F-067D-11D0-95D8-00A02463AB28}'
  proxyStubClsid32='{00020424-0000-0000-C000-000000000046}'
  baseInterface='{00000000-0000-0000-C000-000000000046}'
  tlbid = '{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}'/>
<comInterfaceExternalProxyStub
  name='IScriptError'
  iid='{70841C78-067D-11D0-95D8-00A02463AB28}'
  proxyStubClsid32='{00020424-0000-0000-C000-000000000046}'
  baseInterface='{00000000-0000-0000-C000-000000000046}'
  tlbid = '{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}'/>
<comInterfaceExternalProxyStub
  name='IScriptControl'
  iid='{0E59F1D3-1FBE-11D0-8FF2-00A0D10038BC}'
  proxyStubClsid32='{00020424-0000-0000-C000-000000000046}'
  baseInterface='{00000000-0000-0000-C000-000000000046}'
  tlbid = '{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}'/>
<comInterfaceExternalProxyStub
  name='DScriptControlSource'
  iid='{8B167D60-8605-11D0-ABCB-00A0C90FFFC0}'
  proxyStubClsid32='{00020424-0000-0000-C000-000000000046}'
  baseInterface='{00000000-0000-0000-C000-000000000046}'
  tlbid = '{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}'/>
```

This is the output of the TypeLibInfoTool for the component _msscript.ocx_
(used to execute scripts within an application). Such a fragment can be
inserted within a manifest file with a structure such as the following (the
file in this example can be called _all_needed_components.manifest_).

```xml
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<assembly xmlns='urn:schemas-microsoft-com:asm.v1' manifestVersion='1.0'>
  <assemblyIdentity name='all_needed_components' type='win32' version='1.0'/>
  <description>All the needed COM components by our application</description>
  <file name='...'>
  </file>
  ...
</assembly>
```

When all the components used by the application are inserted in this manifest
file it can be used for our click-once application. Actually, another manifest
file with the same name as the application (e.g.
_MyRegistrationFreeApp.exe.manifest_ for an executable called
_MyRegistrationFreeApp.exe_) is still needed. It is this file that references
the manifest file with all the registration info. Here is a sample.

```xml
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<assembly xmlns='urn:schemas-microsoft-com:asm.v1' manifestVersion='1.0'>
  <assemblyIdentity name='MyRegistrationFreeApp' processorArchitecture='X86' type='win32' version='1.0' />
  <description>My Registration-Free Application</description>
  <dependency>
    <dependentAssembly>
      <assemblyIdentity name='all_needed_components' type='win32' version='1.0.0.0' />
    </dependentAssembly>
  </dependency>
</assembly>  
```

One last caveat on registration-free COM, when a component is 
implemented as an executable this all mechanism does not work. For some 
reason Microsoft engineers decided to leave executables out, therefore 
they still need registration. Remember also that Windows OS's before XP 
need all components to be registered in the traditional way.

### Summary of Information

The third operation available with the TypeLibInfoTool is a generating a
summary of the information of the Type Library interface. This simply operates
by enumerating each interface with its methods and properties, each coclass
with the interfaces it derives from, etc. The tool is quite rudimental, but can
be used for a quick inspection of a COM component before running one of the
operations above.

## Design

The application is written as an HTML Application. This means it is a
special kind of web page that gets displayed within an application 
frame instead of a web browser and is not subject to the usual 
restrictions when accessing system resource or executing code in COM 
components. All the logic is implemented as Javascript function within 
the web page.

The bulk of this code uses the TypeLib Information Object Library 
(TLI) to enumerate the information in a component and process it 
according to the task required (i.e. generating an IDL or a manifest 
file). As indicated in the to-do's section, this leaves room for 
extending the application for other operations such as automating 
documentation.

## To do

A number of desirable features have been left out, work on these will depend on
the interest this article can raise.

* Not all information contained in a Type Library are processed. For instance the ``importlib(...)`` directives are still missing in the IDL generation.
* Copy-and-Paste is the only way to get information out of the tool. A save button could be added.
* The function that enumerates all interfaces, coclasses, enums, etc. for generating the IDL file could easily be modified to automatically produce a documentation for a Type Library. The best approach could be generating an XML version of the IDL file and rely on an external XML Stylesheet
* This application is designed around the concept of retrieving information out of a _single_ COM component. This poses a limitation when one needs to generate a manifest file for an application, which in general requires gathering information from several files. For this reason, some manual copy-and-paste work is required. It would be nice have the TypeLibInfoTool help in automate this task to a greater degree.

## Author

&copy; Federico Aponte <<federico.aponte@gmail.com>> (2011-2018)

## License

[GPL v3 or later](http://www.gnu.org/copyleft/gpl.html)

## References

<a id="ref1"></a>[1] [Registration-Free Activation of COM Components: A Walkthrough](http://msdn.microsoft.com/en-us/library/ms973913.aspx)
<a id="ref2"></a>[2] [Visual Basic: Inspect COM Components Using the TypeLib Information Object Library](http://msdn.microsoft.com/en-us/magazine/bb985086.aspx)
<a id="ref3"></a>[3] [Creating Type Libraries Using IDL](http://edndoc.esri.com/arcobjects/9.1/ExtendingArcObjects/Ch02/TypeLibrariesAndIDL.htm)
<a id="ref4"></a>[4] [Simplify App Deployment with ClickOnce and Registration-Free COM](http://msdn.microsoft.com/en-us/magazine/cc188708.aspx)
