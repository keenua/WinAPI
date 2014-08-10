WinAPI
======

C# wrapper for WinAPI functions for windows manipulation, registry and mouse/keyboard events

Here's a simple example

```C#

            // Find a notepad window
            var win = WindowSearch.GetWindowByText("notepad", "", false, true);

            // Set it foreground
            WindowManipulation.SetForegroundWindow(win);

            Thread.Sleep(500);

            //Type "Hi"
            MouseKeyboard.Type("Hi", 10);

            // Press enter
            MouseKeyboard.PressKey(KeyConstants.VK_RETURN);
            
```
