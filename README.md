# SimpleID Capture VB6 Sample
This is a Windows Forms application that shows how to integrate VB6 code with SimpleID Capture.

## Third-Party components ##

We use the third party components that are in the "dependencies" folder, both of which need to be registered in your Windows. To register the dlls you need to copy to the Windows / system32 and Windows / SysWOW64 folders and run the "RegisterHTX32.bat" and "RegisterHTX64.bat" files. Feel free to use other libraries.

## Running the application ##

Before running the application, you must have a valid installation of SimpleID Core and SimpleID Capture. You also need to set the API key and url in the Main.frm file. Change the following lines, create and run the project.

```frm
wsUrl = "URL_SIMPEID"
apiKey = "YOUR_API_KEY"
```

## Contribution ##

Suggestions and new features for this sample are more than welcome. Feel free to submit a PR.

## License ##

This sample is provided under [The MIT License].