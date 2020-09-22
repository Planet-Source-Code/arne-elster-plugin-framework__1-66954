<div align="center">

## Plugin Framework


</div>

### Description

This class finds and loads plugins, you just specify the interface the plugins have to implement.

----

How it works: You create a new ActiveX DLL, add a class to it, and enter all the properties and functions your plugins have to implement (you design the plugin interface). After compiling, you create a plugin (again an ActiveX DLL) which implements ('Implements' keyword) the interface, and add some code to it. After compiling this DLL you tell the Plugin Framework which interface your plugins implement and let it search for it. The Framework will look in the directory you specified, register all DLLs and if one of the classes in a DLL implements the plugin interface, the Framework will return it to you. No need for CreateObject anymore. /// How to get this demo working: Compile the plugins in the "plugins" directory, and place the DLLs there. Then simply run "Projekt1.vbp", but make sure you referenced the typelib in the "typelib" directory.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2006-10-28 01:52:40
**By**             |[Arne Elster](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/arne-elster.md)
**Level**          |Advanced
**User Rating**    |5.0 (65 globes from 13 users)
**Compatibility**  |VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Plugin\_Fra2028191112006\.zip](https://github.com/Planet-Source-Code/arne-elster-plugin-framework__1-66954/archive/master.zip)








