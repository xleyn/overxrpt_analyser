<h1>The Overxrpt Analyser Project</h1>
<h2>Scope of Work</h2>
Welcome to the Overxrpt Analyser Project. 
This project was developed for the Regional Radiation Protection Service and automates the analysis process for overexposure reports received from Landauer.
The process is fully automated, from receiving a report to structuring an email to be sent off to the client.
<h2>Requirements</h2>
All required resources are stored in this repository bar the following two documents (for privacy reasons):  
<ul>
  <li>The reference Landauer investigation levels spreadsheet</li>
  <li>The dosimetry formal investigation form</li>
</ul>

Both of these documents should be stored in a docs/ folder in the root directory for this software to correctly function.
Please create and activate a conda environment using the environ.yml file for dependency configuration.

<h2>Outline of Code Flow</h2>
The key logic behind the code's function is documented in the following steps:
<ul>
  <li>Report selection</li>
  <li>Dataframe loading</li>
  <li>Report reading</li>
  <li>Pulling inestigation levels</li>
  <li>Analysis and email generation</li>
</ul>

<h3>Report Selection</h3>
Prior to code running, the user saves a Landauer report in the correct folder on the RRPS R: drive. 
They must ensure it is saved in the folder for the year that the report relates to. 
During runtime, the user selects the report to be analysed.

<h3>Dataframe Loading</h3>
The investigation levels Excel spreadsheet is loaded into a pandas dataframe. 
Based on the report's account and subaccount code, the dataframe is cropped.

<h3>Report Reading</h3>
The contents of the table on the report are read. The information from each row is stored in a separate object, for ease of processing.

<h3>Pulling Investigation Levels</h3>
Relevant investigation levels relating to each Row object are pulled from the dataframe. 

<h3>Analysis and Email Generation</h3>
The doses for each row are compared to the pulled investigation levels and a response is formed based on comparison logic.
An email is automatically generated from this response and additional information in the Excel spreadsheet.
The user must check over the analysis before sending off to the client!

