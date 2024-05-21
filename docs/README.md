# Documentation

Documentation for the Microsoft Security Updates CVRF API can be found in the form of Swagger documentation.

The documentation can be used to help developers who wish to implement their own API requests / response handling to get the correct responses and fields. For more information check out http://swagger.io. It is a great project and has great staff working on it. 

To view the API, you can open the [swagger-ui demo](http://petstore.swagger.io/) and put the link to the raw [swagger.json](swagger.json?raw=1) in the top bar then click explore. Swagger-ui can be downloaded locally and run if you need to integrate this into how you interact with the api.

As a note about the swagger, there are two response models defined for the CVRF endpoint. one for json and the other for xml. This is because swagger does not currently have a way to return two different schemas for XML and JSON responses. If you are looking at the swagger definitions and want the correct XML response, please change the response ref accordingly. To find the location to change, search the swagger.json or swagger.yaml for the text: `#/definitions/cvrfReturnTypes200` then simply add \_xml to the end (so it reads `#/definitions/cvrfReturnTypes200_xml`).

For more information on CVRF, please view the following links:

https://cve.mitre.org/cve/cvrf.html

http://www.icasi.org/the-common-vulnerability-reporting-framework-cvrf-v1-1/
