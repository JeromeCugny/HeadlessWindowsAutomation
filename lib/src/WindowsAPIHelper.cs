using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using System.Windows.Forms;

namespace HeadlessWindowsAutomation
{
    /// <summary>
    /// Represents a window in the Windows operating system and provides methods to interact with it.
    /// </summary>
    public class Window
    {
        /// <summary>
        /// Handle of the window.
        /// </summary>
        public IntPtr Handle { get; }
        private const uint GW_OWNER = 4;

        /// <summary>
        /// Initializes a new instance of the <see cref="Window"/> class with the specified window handle.
        /// </summary>
        /// <param name="hwnd">The handle of the window.</param>
        public Window(IntPtr hwnd) 
        {
            this.Handle = hwnd;
        }

        /// <summary>
        /// Retrieves the <see cref="AutomationElement"/> associated with this window.
        /// </summary>
        /// <returns>The <see cref="AutomationElement"/> associated with this window.</returns>
        public AutomationElement GetAutomationElement()
        {
            return AutomationElement.FromHandle(this.Handle);
        }

        /// <summary>
        /// Retrieves the handle of the owner window.
        /// </summary>
        /// <returns>The handle of the owner window.</returns>
        public IntPtr GetOwnerWindowHandle()
        {
            return WindowsAPIHelper.GetWindow(this.Handle, GW_OWNER);
        }

        /// <summary>
        /// Retrieves the handles of the windows owned by this window.
        /// </summary>
        /// <returns>A list of handles of the windows owned by this window.</returns>
        public List<IntPtr> GetOwnedWindows()
        {
            List<IntPtr> ownedWindows = new List<IntPtr>();
            WindowsAPIHelper.EnumWindows((hWnd, lParam) =>
            {
                if (WindowsAPIHelper.GetWindow(hWnd, GW_OWNER) == this.Handle)
                {
                    ownedWindows.Add(hWnd);
                }
                return true; // Continue enumeration
            }, IntPtr.Zero);

            return ownedWindows;
        }

        /// <summary>
        /// Retrieves the handles of the child windows of this window.
        /// </summary>
        /// <returns>A list of handles of the child windows of this window.</returns>
        public List<IntPtr> GetChildrenWindows()
        {
            List<IntPtr> children = new List<IntPtr>();

            WindowsAPIHelper.EnumChildWindows(this.Handle, (hWnd, lParam) =>
            {
                children.Add(hWnd);
                return true; // Continue enumeration
            }, IntPtr.Zero);

            return children;
        }
    }

    /// <summary>
    /// Provides helper methods for interacting with Windows API and UI Automation.
    /// </summary>
    public static class WindowsAPIHelper
    {
        /// <summary>
        /// Retrieves an AutomationElement from a window handle.
        /// </summary>
        /// <param name="hwnd">The handle of the window.</param>
        /// <returns>The AutomationElement associated with the window handle, or null if the handle is invalid.</returns>
        public static AutomationElement GetAutomationElement(IntPtr hwnd)
        {
            if (hwnd != IntPtr.Zero) return AutomationElement.FromHandle(hwnd);
            else
            {
                Console.Error.WriteLine("Cannot get AutomationElement from an unspecified window handle");
                return null;
            }
        }

        /// <summary>
        /// Retrieves the immediate parent of an AutomationElement.
        /// </summary>
        /// <param name="element">The AutomationElement whose parent is to be retrieved.</param>
        /// <returns>The immediate parent AutomationElement.</returns>
        public static AutomationElement GetImmediateParent(AutomationElement element)
        {
            // Create a TreeWalker instance
            TreeWalker walker = TreeWalker.ControlViewWalker;

            // Get the parent element
            AutomationElement parent = walker.GetParent(element);

            return parent;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        /// <summary>
        /// Visits the windows with the specified window as parent.
        /// </summary>
        /// <param name="hWndParent">The handle of the parent window.</param>
        /// <param name="immediateChild">If true, only immediate children are retrieved; otherwise, all descendants are retrieved.</param>
        /// <param name="visitor">Visitor function for each child. Return false to stop the visit.</param>
        public static void VisitWindows(IntPtr hWndParent, bool immediateChild, Func<AutomationElement, bool> visitor)
        {
            EnumChildWindows(hWndParent, (hWnd, lParam) =>
            {
                try
                {
                    AutomationElement element = WindowsAPIHelper.GetAutomationElement(hWnd);
                    if (element != null) return visitor(element);
                }
                catch (Exception ex)
                {
                    // Can return non-usable handle due to permission
                    Console.Error.WriteLine($"Exception in VisitWindows: {ex.Message}");
                }
                return true; // Continue enumeration
            }, IntPtr.Zero);
        }

        /// <summary>
        /// Visits the top level windows.
        /// </summary>
        /// <param name="visitor">Visitor function for each window. Return false to stop the visit.</param>
        public static void VisitTopWindows(Func<AutomationElement, bool> visitor)
        {
            EnumWindows((hWnd, lParam) =>
            {
                try
                {
                    AutomationElement element = WindowsAPIHelper.GetAutomationElement(hWnd);
                    if (element != null) return visitor(element);
                }
                catch (Exception ex)
                {
                    // Can return non-usable handle due to permission
                    Console.Error.WriteLine($"Exception in VisitTopWindows: {ex.Message}");
                }
                return true; // Continue enumeration
            }, IntPtr.Zero);
        }

        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();

        /// <summary>
        /// Combine two 16-bit values into a single 32-bit value.
        /// cf. MAKELPARAM macro in C++
        /// </summary>
        /// <param name="low">Value to combined</param>
        /// <param name="high">Value to combined</param>
        /// <returns>32-bit value</returns>
        public static IntPtr MakeLParam(int low, int high)
        {
            return (IntPtr)((high << 16) | (low & 0xFFFF));
        }

        /// <summary>
        /// Combine two 16-bit values into a single 32-bit value.
        /// cf. MAKEWPARAM macro in C++
        /// </summary>
        /// <param name="low">Value to combined</param>
        /// <param name="high">Value to combined</param>
        /// <returns>32-bit value</returns>
        public static IntPtr MakeWParam(int low, int high)
        {
            return MakeLParam(low, high);
        }

        /// <summary>
        /// Determines if the specified key is an Alt key.
        /// </summary>
        /// <param name="keyCode">The key code to check.</param>
        /// <returns>True if the key is an Alt key; otherwise, false.</returns>
        public static bool IsAltKey(Keys keyCode)
        {
            return keyCode == Keys.Alt || keyCode == Keys.Menu || keyCode == Keys.LMenu || keyCode == Keys.RMenu;
        }

        /// <summary>
        /// Determines if the specified key is a system key.
        /// </summary>
        /// <param name="keyCode">The key code to check.</param>
        /// <returns>True if the key is a system key; otherwise, false.</returns>
        public static bool IsSystemKey(Keys keyCode)
        {
            // F10, ALT, and anything alt +.
            return keyCode == Keys.F10 || IsAltKey(keyCode);
        }

        /// <summary>
        /// Determines if the specified key is an extended key.
        /// </summary>
        /// <param name="keyCode">The key code to check.</param>
        /// <returns>True if the key is an extended key; otherwise, false.</returns>
        public static bool IsExtendedKey(Keys keyCode)
        {
            // cf. https://learn.microsoft.com/en-us/windows/win32/inputdev/about-keyboard-input#extended-key-flag
            return keyCode == Keys.RMenu || keyCode == Keys.RControlKey || keyCode == Keys.Insert || keyCode == Keys.Delete
                || keyCode == Keys.Home || keyCode == Keys.End || keyCode == Keys.PageUp || keyCode == Keys.PageDown
                || keyCode == Keys.Up || keyCode == Keys.Down || keyCode == Keys.Left || keyCode == Keys.Right
                || keyCode == Keys.NumLock || keyCode == Keys.PrintScreen || keyCode == Keys.Divide;
                // missing ENTER key in the numeric keypad, it's the same value as the regular ENTER ...
                // missing BREAK (Keys.ControlKey + Keys.Pause) key, useless, and would need to check Keys[]
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowDC(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        public static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, int dwRop);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        public const int SRCCOPY = 0x00CC0020;

        [DllImport("gdi32.dll", SetLastError = true)]
        static extern uint GetBkColor(IntPtr hdc);

        /// <summary>
        /// Retrieves the background color of the specified window.
        /// </summary>
        /// <param name="hWnd">A handle to the window whose background color is to be retrieved.</param>
        /// <returns>
        /// A <see cref="System.Drawing.Color"/> structure that represents the background color of the specified window.
        /// Returns <see cref="System.Drawing.Color.Empty"/> if the window handle is invalid, 
        /// the device context cannot be retrieved, or the background color cannot be determined.
        /// </returns>
        /// <remarks>
        /// This method uses the GetDC and GetBkColor functions from the Windows API to retrieve the background color.
        /// If any step in the process fails, an error message is printed to <see cref="System.Console.Error"/> and 
        /// <see cref="System.Drawing.Color.Empty"/> is returned.
        /// </remarks>
        public static System.Drawing.Color GetBackgroundColor(IntPtr hWnd)
        {
            if (hWnd != IntPtr.Zero)    // We don't want to use the DC for the whole screen
            {
                IntPtr hdc = GetDC(hWnd);
                if (hdc != IntPtr.Zero)
                {
                    uint colorRef = GetBkColor(hdc);
                    ReleaseDC(hWnd, hdc);

                    if (colorRef != 0xFFFFFFFF) // CLR_INVALID
                    {
                        // Extract RGB values from colorRef
                        int r = (int)(colorRef & 0x000000FF);
                        int g = (int)((colorRef & 0x0000FF00) >> 8);
                        int b = (int)((colorRef & 0x00FF0000) >> 16);

                        return System.Drawing.Color.FromArgb(r, g, b);
                    }
                    else Console.Error.WriteLine("Failed to get background color.");
                }
                else Console.Error.WriteLine("Failed to get device context.");
            }
            else Console.Error.WriteLine("Window handle cannot be null.");

            return System.Drawing.Color.Empty;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        // Messages (non exhaustive)
        // Window Messages
        public const uint WM_SETFOCUS = 0x0007; // Sent to a window after it has gained the keyboard focus.
        public const uint WM_SETTEXT = 0x000C; // Sets the text of a window.
        public const uint WM_GETTEXT = 0x000D; // Copies the text of a window into a buffer.
        public const uint WM_GETTEXTLENGTH = 0x000E; // Gets the length of the text of a window.
        public const uint WM_CLOSE = 0x0010; // Closes a window.
        public const uint WM_QUIT = 0x0012; // Posts a quit message and exits the application.
        public const uint WM_SETCURSOR = 0x0020; // Sent to a window if the mouse causes the cursor to move within a window and mouse input is not captured.
        public const uint WM_NCHITTEST = 0x0084; // Sent to a window in order to determine what part of the window corresponds to a particular screen coordinate.
        public const uint WM_COMMAND = 0x0111; // Sent when the user selects a command item from a menu, or when a control sends a notification message to its parent window.
        public const uint WM_SYSCOMMAND = 0x0112; // Sent when the user selects a system command from the window menu.
        public const uint WM_KEYDOWN = 0x0100; // Posted to the window with the keyboard focus when a nonsystem key is pressed.
        public const uint WM_KEYUP = 0x0101; // Posted to the window with the keyboard focus when a nonsystem key is released.
        public const uint WM_CHAR = 0x0102; // Posted to the window with the keyboard focus when a character is typed.
        public const uint WM_DEADCHAR = 0x0103; // Posted to the window with the keyboard focus when a dead key is pressed.
        public const uint WM_SYSKEYDOWN = 0x0104; // Posted to the window with the keyboard focus when the user presses the F10 key (which activates the menu bar) or holds down the ALT key and then presses another key.
        public const uint WM_SYSKEYUP = 0x0105; // Posted to the window with the keyboard focus when the user releases a key that was pressed while the ALT key was held down.
        public const uint WM_SYSCHAR = 0x0106; // Posted to the window with the keyboard focus when a WM_SYSKEYDOWN message is translated by the TranslateMessage function. It specifies the character code of a system character key that is, a character key that is pressed while the ALT key is down.
        public const uint WM_SYSDEADCHAR = 0x0107; // Sent to the window with the keyboard focus when a WM_SYSKEYDOWN message is translated by the TranslateMessage function. WM_SYSDEADCHAR specifies the character code of a system dead key that is, a dead key that is pressed while holding down the ALT key.
        public const uint WM_LBUTTONDOWN = 0x0201; // Posted when the left mouse button is pressed.
        public const uint WM_LBUTTONUP = 0x0202; // Posted when the left mouse button is released.
        public const uint WM_LBUTTONDBLCLK = 0x0203; // Posted when the user double-clicks the left mouse button while the cursor is in the client area of a window.
        public const uint WM_RBUTTONDOWN = 0x0204; // Posted when the right mouse button is pressed.
        public const uint WM_RBUTTONUP = 0x0205; // Posted when the right mouse button is released.
        public const uint WM_MOUSEMOVE = 0x0200; // Posted when the mouse is moved.
        public const uint WM_MOUSEWHEEL = 0x020A; // Sent when the mouse wheel is rotated.
        // Control Messages
        public const uint BM_CLICK = 0x00F5; // Simulates a button click.
        public const uint EM_SETSEL = 0x00B1; // Selects a range of characters in an edit control.
        public const uint EM_GETSEL = 0x00B0; // Gets the starting and ending character positions of the current selection in an edit control.
        public const uint EM_REPLACESEL = 0x00C2; // Replaces the current selection in an edit control with the specified text.
        public const uint LB_ADDSTRING = 0x0180; // Adds a string to a list box.
        public const uint LB_GETCOUNT = 0x018B; // Gets the number of items in a list box.
        public const uint LB_GETTEXT = 0x0189; // Gets the string of a list box item.
        public const uint CB_ADDSTRING = 0x0143; // Adds a string to the list in a combo box.
        public const uint CB_GETCOUNT = 0x0146; // Gets the number of items in the list in a combo box.
        public const uint CB_GETCURSEL = 0x0147; // Gets the index of the currently selected item in a combo box.
        public const uint CB_SETCURSEL = 0x014E; // Selects a string in the list of a combo box.
        // Params
        public const uint MK_CONTROL = 0x0008; // The CTRL key is down.
        public const uint MK_LBUTTON = 0x0001; // The left mouse button is down.
        public const uint MK_MBUTTON = 0x0010; // The middle mouse button is down.
        public const uint MK_RBUTTON = 0x0002; // The right mouse button is down.
        public const uint MK_SHIFT = 0x0004; // The SHIFT key is down.
        public const int BN_CLICKED = 0; // Button notification codes
        public const int LBN_SELCHANGE = 1; // ListBox notification codes
        public const int CBN_SELCHANGE = 1; // ComboBox notification codes
        public const int EN_CHANGE = 0x0300; // Edit control notification codes
        public const int TCN_SELCHANGE = 0x1301; // Tab control notification codes
        public const int MN_SELECTITEM = 0x01E5; // Menu notification codes
        public const int TBM_SETPOS = 0x0405; // Slider notification codes
        public const int SBM_SETPOS = 0x00E0; // ScrollBar notification codes
        public const int TVN_SELCHANGED = 0x0200; // TreeView notification codes
        public const int STN_CLICKED = 0; // Static control notification codes
        public const int TTN_SHOW = 0x0401; // ToolTip notification codes
        public const int SB_SETTEXT = 0x0400; // StatusBar notification codes
        public const int PBM_SETPOS = 0x0402; // ProgressBar notification codes
        public const int UDN_DELTAPOS = 0x0403; // UpDown control notification codes
        public const int HDN_ITEMCHANGED = 0x0201; // Header control notification codes
        public const int TBN_DROPDOWN = 0x0404; // ToolBar notification codes
        public const int WM_USER = 0x0400; // User-defined message
    }
}
