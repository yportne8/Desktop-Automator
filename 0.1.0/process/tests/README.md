### The subfolder tests is required and should not be moved.

### Running the package as an executable: python.exe -m process
Testing for the code within __main__, the Console App, cannot be fully
automated due to the nature of the interface. As such, all user related
components must be manually tested prior to release updates.

Users of the package do not need the Console interface to run their processes.
Once a process has been installed, if the .main function has been populated,
the operator can call on python to run their process directly on the command 
line as such:

```bash
START "python.exe -m process_name username password"
```