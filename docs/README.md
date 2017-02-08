# Documentation

As described in the main `README` the above swagger deffinition documents have two return models for the CVRF endpoint. This is because swagger does not currently have a way to return two different schemas for XML and JSON responses. If you are looking at the swagger deffinitions and want the correct XML response, please change the response ref accordingly at line 89 (for yaml) or 114 (jor json).

For more information on CVRF, please view the following links:

https://cve.mitre.org/cve/cvrf.html

http://www.icasi.org/the-common-vulnerability-reporting-framework-cvrf-v1-1/
