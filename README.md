# An interactive code walk-through on how to authenticate and authorize to use the Azure IoT Central REST APIs
This is a small codebase built from create-react-app and uses MSAL.js. The code demonstrates the call patterns required to Auth'N/Auth'Z against MS Graph, IoT Central, Azure Resource Management (ARM) and working examples of their respective APIs. This should provide you with enough base code to start your own application.

The code demonstrates two call patterns.  One pattern is used for IoT Central data plane APIs and single sign-on.  The other is used for IoT Central control plane APIs.  

Data plane APIs are REST APIs to interact with a specific IoT Central application.  Click [here](https://docs.microsoft.com/en-us/rest/api/iotcentral/) to see the list of available APIs.  Single sign-on enables deep-linking from your application into an IoT Central application.

Control plane APIs are used to manage resources in you Azure subscription.  Use these APIs to create, delete or find and IoT Central application. 


## Install
```
npm i
```

## Run
```
npm start
````

## Usage
```
http://localhost:4001
````
