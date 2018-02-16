# Licensing-API-VB6
A sample application in VB6 that uses the COM wrapper around the elm licensing API. 

# Usage
If you want to license your legacy VB6 application using the cloud-based [elm](https://elm.evoleap.com/) licensing system, 
you can use this sample application to guide your development. The key aspects of licensing calls are in the LicenseController class module.
You should be able to take this code with minimal changes and use it. The StatePersistence code must be modified to save the licensing
state in an encrypted format. The sample application uses plain text, which is not appropriate for production code. 
