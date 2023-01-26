# VbaDI - An IoC Container for VBA

VbaDI is an Windows Excel VBA add-in (.xlam) that provides an Inversion of Control (IoC) container for VBA projects.

#### VbaDI provides:
1. The means to statically declare class dependencies.
2. A fluent registration interface for configuring application objects.
3. The ability to Build/Resolve an application's object graph despite the absence of parameterized constructors in VBA.
4. A generalized and re-usable alternative to [_Pure DI_](https://blog.ploeh.dk/2014/06/10/pure-di/).

## Using VbaDI:

### Installing the VbaDI add-in: 
   1. Download VbaDI.xlam from this repository and copy it to your Excel add-in folder.
      - To find where add-ins are stored on your system, open any Excel workbook: Click _File -> Options -> Add-ins_.  Location of current add-ins will be visible.
   2. Open an existing Excel project (.xlsm) or open a new workbook and press Alt+F11 to open the Visual Basic Editor (VBE)   
   3. In the VBE, click on _Tools->References_ and find and check the box for 'VbaDIAddIn'.  It may be necessary to click _Browse..._ to find the add-in.
   4. If you opened a new workbook in step 2 - save the workbook as a macro-enabled .xlsm file.

### Recommended Tools and References
1. Get [Rubberduck](https://rubberduckvba.com/) - Make your VBA development efforts more efficient, productive, and correct.
2. See #1
3. "Dependency Injection in .NET" by Mark Seeman.
4. [CastleWinsor](http://www.castleproject.org/projects/windsor/) (CW) Inversion of Control (IoC) container for C# - A production OSS IoC container that was used as a model for _VbaDI_.

### IoC Container process flow

1. [Register](https://github.com/BZngr/SandboxMd/edit/main/README.md#register) classes, interfaces, and configuration values with the Container.
2. [Resolve](https://github.com/BZngr/SandboxMd/edit/main/README.md#resolve) the application's object graph.  Compose registered classes with their dependencies. 
3. [Release](https://github.com/BZngr/SandboxMd/edit/main/README.md#release) the Container and its object references.

This process flow is referred to as the [_Register -> Resolve -> Release_](https://blog.ploeh.dk/2010/09/29/TheRegisterResolveReleasepattern/) pattern:

### Register:

Configuration expressions begin by calling 1 of 2 functions exposed by the _VbaDI_ module: `VbaDI.Instance` or `VbaDI.ForInterface`.  
These functions are the entry points to a fluent builder API - `IVbaDIFluentRegistration` [^1].  

Some examples of registering elements with the Container:

_Registration of a Class instance_

```vba
container.Register VbaDI.Instance(New MyServiceImpl) 'defaults to Singleton lifestyle

'Or, declare Lifestyle explicitly:

'Singleton
container.Register VbaDI.Instance(New MyServiceImpl).AsSingleton()

'Transient
container.Register VbaDI.Instance(New MyServiceImpl).AsTransient()
```

_Interface(s) registration with an implementing class instance_
```vba
container.Register VbaDI.ForInterface(TypeName(New IMyService)).Use(New MyServiceImpl)

container.Register VbaDI _
    .ForInterface(TypeName(New ICopier), TypeName(New IFax), TypeName(New IShredder)) _
    .Use(New BizMachine)
```

_Registration of a component that has a Value dependency_
```vba
Dim cxnString As String
cxnString = "Provider= Microsoft.ACE.OLEDB.12.0; Data Source='" & filepath & "';"
	
container.Register VbaDI.ForInterface(TypeName(IRepo)).Use(New AppRepo) _
     .DependsOnValue("AdoConnectionString", cxnString)
```

##### Notes regarding IVbaDIFluentRegistration and the _Register_ phase

1. Classes, Interface assignments, and Value dependencies are identified by [RegistrationIDs](https://github.com/BZngr/SandboxMd/edit/main/README.md#registrationids).
2. The first registered RegistrationID/Instance pair is the only RegistrationID/Instance pair cached by the Container.
3. The first registration of an interface or an object value dependency _wins_.  Subsequent registrations are ignored.
4. If a class has no dependencies, but _is_ a dependency of another registered class, then it needs to be registered with the container.

#### RegistrationLoaders

The registration process can be accomplished by _Loaders_ which are `Class Modules` that implement the single method `IVbaDIRegistrationLoader` interface:

###### IVbaDIRegistrationLoader
```vba
Public Sub LoadToContainer(ByVal pContainer As IVbaDIContainer)
End Sub
```

An example of typical Loader object content:
```vba
Option Explicit

Implements IVbaDIRegistrationLoader

Private Sub IVbaDIRegistrationLoader_LoadToContainer(ByVal pContainer As IVbaDIContainer)
    
     pContainer.Register VbaDI.ForInterface(TypeName(New IMyService)) _
	     .Use(New MyServiceImpl)
    
     pContainer.Register VbaDI.ForInterface(TypeName(New ILogger)) _
	     .Use(New FileLoggerImpl)
     
     '... and so on
End Sub
```
###### Notes regarding Loaders/IVbaDIRegistrationLoader

1. Using a RegistrationLoader to configure the IoC container is recommended.
2. One or more RegistrationLoaders can be used to configure a container.
3. RegistrationLoaders are custom class modules of an application, not the add-in.  
4. RegistrationLoaders help organize/modularize IoC container configuration.

### Resolve:

Once all classes involved in DI are registered with the Container, the _Resolve_ method is invoked on the Container and the application's object graph is assembled.  The _Resolve_ process requires far less user code when compared to the  _Register_ process.  For each use of an IoC Container, there _should_ be a single call to `IVbaDIContainer.Resolve(<RegistrationID>)`. 

The _Resolve_ process relies upon the `IVbaDIQueryCompose` interface (declared by the add-in).  The `IVbaDIQueryCompose` interface provides the ability for class modules to statically declare their dependencies.  `IVbaDIQueryCompose` is also the mechanism by which the IoC Container injects dependencies.

#### IVbaDIQueryCompose
```
'@Description "Returns the dependency RegistrationIDs that an object requires
Public Property Get RegistrationIDs() As Collection
End Property

'@Description "Retrieve dependencies by RegistrationID using IVbaDIDependencyProvider"
Public Sub ComposeObject(ByVal pProvider As IVbaDIDependencyProvider)
End Sub
```

If a registered object implements the `IVbaDIQueryCompose` interface, then the Container will:
1. Call the `RegistrationIDs` property
    - The implementing object is responsible to return a `Collection` of required RegistrationIDs
2. Invoke `ComposeObject` delivering the required set of fully composed dependencies accessible by RegistrationID 	

For comparison: a parameterized C# instance constructor and its equivalent for a VBA object using `IVbaDIQueryCompose`:

```c#
public class ManipulateStuff
{
	private readonly DoStuff _doStuff;
	private readonly IDoOtherStuff _doOtherStuff;
	private readonly string _myStuffFilepath;
	
	//parameterized instance constructor
	public ManipulateStuff(DoStuff doStuff, IDoOtherStuff doOtherStuff, string stuffPath)
	{
            _doStuff= doStuff;
	    _doOtherStuff= doOtherStuff;
	    _myStuffFilepath = stuffPath;
	}
}
```

```vba
'VBA equivalent using VbaDI
'Class Module ManipulateStuff.cls
Option Explicit
Implements IVbaDIQueryCompose

Private Type TManipulateStuff
    DoStuff As DoStuff
    DoOtherStuff As IDoOtherStuff 
    MyStuffFilepath As String
End Type

Private this As TManipulateStuff

Private Property Get IVbaDIQueryCompose_RegistrationIDs() As Collection 
	
    Set IVbaDIQueryCompose_RegistrationIDs = New Collection

    With IVbaDIQueryCompose_RegistrationIDs
        .Add TypeName(New DoStuff)
        .Add TypeName(New IDoOtherStuff)
        .Add "ConfigFilepath"
    End With
End Property

Private Sub IVbaDIQueryCompose_ComposeObject(ByVal pProvider As IVbaDIDependencyProvider)
    With pProvider   
        Set this.DoStuff = .ObjectFor(TypeName(New DoStuff))
        Set this.DoOtherStuff = .ObjectFor(TypeName(New IDoOtherStuff))
        this.MyStuffFilepath = .ValueFor("ConfigFilepath")
    End With
End Sub
```
###### IVbaDIQueryCompose Notes: 

1. Class, Interface and Value dependencies are identified by [RegistrationIDs](https://github.com/BZngr/SandboxMd/edit/main/README.md#registrationids).
2. `IVbaDIQueryCompose.RegistrationIDs` is analogous to the C# example's constructor parameter list. 
3. `IVbaDIQueryCompose.ComposeObject` is analogous to the C# example's constructor function body.
4. If a class has no object, interface, or value dependencies, the developer can: 
    - Choose to _not_ implement `IVbaDIQueryCompose`, or 
    - Implement `IVbaDIQueryCompose` and return an empty RegistrationID `Collection` when `IVbaDIQueryCompose.RegistrationIDs` is called.

At completion of `ComposeObject`, an object is fully composed and initialized.

#### Lifestyles

Lifestyle settings determines how the lifetime of an injected dependency is managed.
The VbaDI container supports two object lifestyles: SINGLETON and TRANSIENT.

1. SINGLETON: 
    - Registration example:  `container.Register VbaDI.Instance(New MyServiceImpl).AsSingleton()`
    - If a Lifestyle is not specified, SINGLETON is the default.
    - The container provides the same, single, fully resolved, instance for each dependency request.
	- The SINGLETON instance cached by the Container is released when the Container is destroyed. 

2. TRANSIENT:
    - Registration example: `container.Register VbaDI.Instance(New MyServiceImpl).AsTransient()`
    - Objects with TRANSIENT lifestyle are created each time they are requested as a dependency.
	- The object's lifetime is controlled by the requesting object.  
	- Although the Container retains a reference to the first registered instance of a class, the cached instance is never provided in response to a transient dependency request. 
	- To support TRANSIENT lifestyle, a class must implement the `IVbaDIDefaultFactory` interface (declared by the add-in).  The Container will invoke `IVbaDIDefaultFactory.Create()` to create a new instance of the class.  The Container then provides the created (and resolved) instance when requested as a dependency. 

### RegistrationIDs

1. RegistrationIDs are unique `String` keys to identify classes, interfaces, and values.
2. RegistrationIDs are used/relevant only during the [Registration](https://github.com/BZngr/SandboxMd/edit/main/README.md#register) and [Resolve](https://github.com/BZngr/SandboxMd/edit/main/README.md#resolve) phases.
3. Object instances are registered using functions with an optional RegistrationID parameter.
    - If the optional RegistrationID parameter is not specified, the RegistrationID is set equal to `TypeName(<object instance>)`.
    - The optional parameter is provided to support user-defined RegistrationIDs. 
4. Interfaces and Value dependencies are registered using a RegistrationID (string) only.
5. Recommendation: Generate RegistrationIDs for classes and interfaces using `TypeName(<class object instance>)`/`TypeName(New <class module>)` or function(s) that leverage the `TypeName` function.  
    - Motivation: Enlists compiler support to keep RegistrationIDs consistent with class/interface module name changes over time.
6. Recommendation: Do not implement Lifecycle Handlers [^2] for classes registered with the _VbaDI_ container.
    - When using _VbaDI_, place code that _would_ have been in `Class_Initialize()` within the `IVbaDIQueryCompose.ComposeObject` implementation.
    - Motivation: Creating RegistrationIDs using `TypeName` can result in multiple, often temporary, object instantiations.
        - Lifecycle Handler implementations that 'do too much' can be a source of errors and/or slow performance (see #5 below)
7. If LifeCycle Handlers ___are___ implemented, a disciplined/best-practice approach to their implementations is necessary:
    - Do not include code that has side-effects, for example:
        - calls to an external resource (e.g., file and database operations)
    	- interactions with other modules and forms
    - Limit `Class_Initialize()` code to simple field initializations.

#### Putting it all together - CompositionRoot:

[CompositionRoot](https://medium.com/@cfryerdev/dependency-injection-composition-root-418a1bb19130) is where all object/dependencies are registered and the object graph is resolved.  CompositionRoot functionality is executed at application startup.  

An example of what implementing CompositionRoot in a `Standard Module` might look like:
```vba
'@Folder "CompositionRoot"
Option Explicit

Private Type TAppEntryPoint
    App As App
End Type

Private this As TAppEntryPoint

'@EntryPoint
Public Sub Main()
On Error GoTo ErrorExit
    
'Create the IoC Container
    Dim xContainer As IVbaDIContainer
    Set xContainer = VbaDI.CreateContainer()
    
'Register...Register using a RegistrationLoader
    xContainer.RegisterUsingLoader New AppIoCLoader
    
'Resolve...One call does it all
    Set this.App = xContainer.Resolve(TypeName(New App))
 
'Invoke the App...All object map instances are assembled and initialized  
    this.App.Main
    
'the Container is released when it goes out of scope  
    Exit Sub
    
ErrorExit:
    Debug.Print "Startup Error: " & Err.Description
    MsgBox "Error Encountered during Application Startup"
End Sub
```

### Release

As can be seen in the CompositionRoot example above, VbaDI does not expose a _Release_ method on the `IVbaDIContainer` interface.  VbaDI relies on garbage collection to implement the _Release_ phase.

[^1]: VbaDI's fluent registration API represent a small fraction of the options provided by full-featured IoC containers like [CastleWinsor](http://www.castleproject.org/projects/windsor/).
[^2]: VBA Lifecycle Handlers are `Class_Initialize()` and `Class_Terminate()`
