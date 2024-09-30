# Overview
**HeadlessWindowsAutomation** is a simple library designed for automating **Windows desktop applications in a headless environment**.  
It leverages the C# **AutomationElement** API to provide a simple yet powerful interface for interacting with UI elements.  
You can automate complex UI interactions and run your testing on **pipelines**.  

To run on a pipeline, your agent must run as **interactive** => started from a user and not a service.  
Consider using `Indirect display driver` (such as [Virtual Display Driver](https://github.com/itsmikethetech/Virtual-Display-Driver)) to avoid some headless limitations.

# Key features
- **AutomationElementWrapper**: Main class, provide automation functionalities.
  - find element(s) using the AutomationElement APIs or XPath.
  - interact with them (`Click`, `SetText`, `Keyboard` sub-class, etc.).  
    In the end, everything will be send as `message` using the `window handle`. 
    So it's better to interact with a component **with a handle**, if you don't have one then use the `ClickOnLocation`.  
    Note that the `Keyboard` might not be compatible with headless, depending on your environment.
- **WindowsAPIHelper**: low level utilities function around Windows APIs.  
  You most likely won't use it unless you need to find elements not under your application. Like an owned top-level window.
- **ProcessHelper**: utility methods for managing and interacting with processes.  
  For example, can be used to find your process when you started it from a command line. 

Refer to the XML documentation comments in the source code for more details.

# Sample Usage
## Setup
Inside your main, start your application and setup the automation:
```C#
WindowsAPIHelper.SetProcessDPIAware();  // Optional. Set DPI

// Start your app
Process myapp_cmd = Process.Start("cmd.exe", "/c myapp.exe"); // Dummy example when you don't directly start your app

// Get myapp and not just the terminal
ProcessHelper processHelper = new ProcessHelper();
processHelper.AdditionalProcessValidCheck = (Process proc) =>
{
    // we also force myapp in the name
    return proc.ProcessName.IndexOf("myapp", StringComparison.OrdinalIgnoreCase) >= 0;
};
int myapp_pid = processHelper.GetApplicationProcessIdFromParent(myapp_cmd);
Process myapp = Process.GetProcessById(myapp_pid);

// Set the root/main AutomationElementWrapper on myapp
var mainWindow = new AutomationElementWrapper(myapp.MainWindowHandle);
AutomationElementWrapper.MainWindowCallback = () => {
    return mainWindow;
};

// Automate your app
...
```

## Automate
After `Setup`, automate your application:
```C#
// Set the text in an Edit control
_ = mainWindow.FindElementByXPath("//Edit[@AutomationId='123']").SetText("foo");

// Use xpath to navigate the components tree of your application
// It follows the xpath spec, refer to the doc. // is descendant, / is children, . is relative, etc.
AutomationElementWrapper subPane = mainWindow.FindElementByXPath("./Pane/Pane[@Name='Foo' and @AutomationId='1234']");
// Print the children in the console for debug, you can also use tools like Accessibility Insights or Spy++
subPane.PrintAllChildren();

AutomationElementWrapper searchBtn = subPane.FindElementByXPath("./Button[@Name='Search']")
  .Click();
```