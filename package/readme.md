# Overview
**HeadlessWindowsAutomation** is a simple library designed for automating **Windows desktop applications in a headless environment**.  
It leverages the C# **AutomationElement** API to provide a simple yet powerful interface for interacting with UI elements.  
You can automate complex UI interactions and run your testing on **pipelines**.  

To run on a pipeline, your agent must run as **interactive** => started from a user and not a service.  
Consider using `Indirect display driver` (such as [Virtual Display Driver](https://github.com/itsmikethetech/Virtual-Display-Driver)) to avoid some headless limitations.

# Key features
- **AutomationElementWrapper**: Main class, provide automation functionalities.
  - find element(s) using the AutomationElement APIs or **XPath**.  
    The find **auto-waits** for the element! You don't need to force a wait.  
    That behavior is configurable using the `.Config` property.
  - interact with them (`Click`, `SetText`, `Keyboard` sub-class, etc.).  
    In the end, everything will be send as `message` using the `window handle`. 
    So it's better to interact with a component **with a handle**, if you don't have one then use the `ClickOnLocation`.  
    On Windows, you can interact with components using messages even if they are **not visible**.  
    Hence why **HeadlessWindowsAutomation** allows actual headless on a desktop application!  
    Note that the `Keyboard` might not be compatible with headless, depending on your environment.  
  - Take a screenshot to easily diagnose an error, especially useful for automated testing.    
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

### Take a screenshot
Whenever you encounter an error (like in a `try ... catch`) we recommend you to take a screenshot.  
When calling the method `TakeScreenshot`, you need to name the file properly.  
Example of implementation assuming you are using **NUnit**:
```C#
public void TakeScreenshot([System.Runtime.CompilerServices.CallerMemberName] string caller = "")
{
  if (mainWindow != null)
  {
    string workDir = NUnit.Framework.TestContext.CurrentContext.WorkDirectory;
    DateTime dateTime = DateTime.Now;
    string screenshotName = $"{caller}_{dateTime.Hour}h{dateTime.Minute}min{dateTime.Second}s{dateTime.Millisecond}ms.png";
    string screenshotPath = Path.Combine(workDir, screenshotName);
    mainWindow.TakeScreenshot(screenshotPath, ImageFormat.Png);
    NUnit.Framework.TestContext.AddTestAttachment(screenshotPath);
  }
}
```
