# Documentation

As described in the main `README` the above swagger definition documents have two return models for the CVRF endpoint. This is because swagger does not currently have a way to return two different schemas for XML and JSON responses. If you are looking at the swagger definitions and want the correct XML response, please change the response ref accordingly. To find the location to change, search the swagger.json or swagger.yaml for the text: `#/definitions/cvrfReturnTypes200` then simply add \_xml to the end (so it reads `#/definitions/cvrfReturnTypes200_xml`.

For more information on CVRF, please view the following links:

https://cve.mitre.org/cve/cvrf.html

http://www.icasi.org/the-common-vulnerability-reporting-framework-cvrf-v1-1/
