# 3M FX Book Sorter StoreAndForward.vbs

 This script attempts to make up for a gap in functionality with the aging 3M FX Book Sorter software.
 When the Library Management System (LMS) is unavailable the sorter software continues to accept items for return (offline mode). However, it makes no attempt to forward the associated transactions onto the LMS when it becomes available again.
 Instead it holds onto the transactions with the expectation that a Library staff member will visit the Admin
 interface and manually batch-forward the transactions (lots of clicking required).

 This is made slightly worse in that the software operation mode has to be temporarily changed from `Normal` to `Out of Service` in order to forward the transactions. This can cause unwanted disruption when done during hours the Library is open.

 3M at least provide an option to be emailed when stored transactions need forwarding, but still manual work is required.

 This script automates the forwarding process. It is designed to be set up as a scheduled task on the Induction PC(s), to be run once a week outside normal hours. However, it can also be run from any Windows PC/Server and at any frequency/time. Be aware, though, the 3M software does not support `TLS`, so the Admin interface password is sent plaintext. So if you are going to set this up on a remote PC/server, ensure that it and the Induction PC communicate over a network **that you trust**.

 The first thing you need to do is set the `IntelligentReturnSystemManagerPassword` in the environment with:

 `setx IntelligentReturnSystemManagerPassword password`

 Note: for Task Scheduler to "see" the new environment variable, you probably need to terminate `Taskeng.exe` (the one running as the user rather than `SYSTEM`). Alternatively reboot the system.

 To test the script from the command line:
 ```
 cscript.exe //nologo 3MStoreAndForward.vbs /inductionpcname:inductionpc.domain.ac.uk /testmode /forwardsleeptime:5
 ```

 `/testmode` runs through the operation mode changes and store-forward procedure irrespective of whether there are any items to process.

 `/forwardsleeptime:seconds` indicates the number of seconds to wait after forwarding transactions but before setting the Induction PC to `Normal` operation mode. Default is 2 minutes but for testing purposes you'll want to set this to something like 5 seconds. The amount of time required is determined by the number of transactions to process, so if you're running this script regularly and not allowing the numbers to build up, then it could be that 15 seconds is all that's required. In practice, if you're running the scheduled task in the small hours when few-to-no people are using the Library, then a longer wait period is likely perfectly acceptable.

The script provides a fair amount of output as it runs. This is designed to be logged to a file for later checking. The scheduled task guidance below demonstrates how this logging can be captured to a file.

To create a scheduled task, do so in the normal way:

 ```
 Program/script:   cmd.exe
 Add arguments: cscript.exe //nologo 3MStoreAndForward.vbs >>3MStoreAndForward.log 2>&1 & cmd.exe /C echo --->>3MStoreAndForward.log
 Start in:  the folder where the script is
 ```
  The `cmd.exe /C echo ---` creates some separation in the log file, helpful in delineating each scheduled task run. A scheduled task `3MStoreAndForward.xml` export file is provided. To simplify creating a scheduled task you can import this file into Task Scheduler.

### Limitations & Improvements

- The script makes no attempt to record the item IDs of the transactions that are stored or were forwarded. Though it is possible to scrape these from the relevant admin web page, it's felt that there's little value in doing so.
- It's probably possible to avoid the script's "sleep" period — that comes after the forwarding operation is started — by understanding better how the final `StoreAndForwardStart.aspx` HTTP response (the one that indicates it's finished) can be distinguished from intermediate responses. However, testing this script in a live environment was difficult, and relied upon there being some genuine transactions to forward (which is seldom the case).
- The script requires that the `IntelligentReturnSystemManagerPassword` be set as a plaintext user environment variable. This may be considered a security issue for some. In our case, though, we perceive it as a low risk.

### For interest...
A very basic `.cmd` implementation is also provided in this repo. It requires `curl` for Windows. See the REMarks in `3mStoreAndForward.cmd` for more information. This was put together as a rapid prototype, before the VBScript solution was developed, in order to understand the flow and build confidence that a solution of this nature would work.
