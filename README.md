# XLPolice

This application terminates Excel processes that remain even all Excel documents are closed.

These remaining Excel processes can cause a delay when opening Excel files (.xlsx/.xlsm) from Explorer, at least in my environment. When this tool is running, it detects and terminates these leftover Excel processes, even if no file is open and there is no application window. This helps to prevent the time-consuming issue of opening files in my environment.

While killing the process might have some adverse effects, it hasn't caused any problems in my environment so far.

## Get Started

1. Download the file
1. Extract the zip file and place `XLPolice.exe` in any local path.
1. Run XLPolice.exe

When you run `XLPolice.exe`, it will reside and continue to monitor the Excel process. 
It does not have a UI, so if you want to terminate it, you must terminate the process from the task manager.





