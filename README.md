# CCG Starter Kits 
## Automatic Data in Brief article creation

This repository is used to automate the word document creation for country models.

There is one script for each region, containing the appropriate text and layout for that region. You should use (or download) the script for the correct region and then edit the file paths, replacing them with the corresponding file path on your PC – there are many file paths, so please make sure that you change them all. If you are working with documents in the cloud, please ensure that you ‘sync local files’ so that they are on your PC.   

Before running the script, make sure that the following files are uploaded for the country/countries that you are trying to run:  
<ul>
  <li>New Country Data Collection </li>
  <li>Country FF Results, Country LCv2 Results, Country NZv2 Results – once you are happy with them and ready to run. Even if you do not include Net Zero in the article, you should upload a ‘dummy’ excel results file with the same name and delete the figures from the article afterwards. </li>
</ul> 

You can indicate which country or countries you want to create articles for in the script around line 143. If you want to run multiple countries, you should leave these lines as shown below and add a column to the Starter Kits – List of Countries.xlsx, putting a 1 for countries that you want to include, and a 0 if you don’t want to include them. You should then add the name of that column in the square brackets after Country_Data in lines 143 and 144 (where it says Run_SA below). If you want to only run one country, you should # out lines 143 and 144, then remove the # from lines 146 and 147 and insert the country name and ISO3 code.   

<code>
countries_short = country_data[country_data['Run_SA'] == 1]['Country']

icountries = country_data[country_data['Run_SA'] == 1]['ISO3']
</code>

In the Starter Kits – List of Countries, make sure that you have indicated the domestic refinery capacity (from the McKinsey Refinery Reference Desk) in tb/d in column S. The author list and affiliations should also be added in columns L and M. You should also indicate the reference list to be used in column U. Reference lists should be named ‘Reference List – Name’ and stored in this folder, with the Name portion written in column U of the Starter Kits – List of Countries. Be sure not to activate URLs in your reference list, otherwise the script cannot pick them up. Ensure that you have the reference list downloaded onto your computer and that you update the file path in line 373 of the script (see below).  


<code>
 Variable Reference Imports
 
refl = country_data[country_data['Country']==country]['Reference List'].values[0]
  
input_doc = Document(r'C:/Users/username/.../Reference List {}.docx'.format(refl)

</code>

You should then be ready to run the script and create the articles. You should see figures created in a ‘figures' folder and articles in an 'outputs' folder, both in the file path you specified in the script.  

Once the articles are ready, there are a few manual things to do (listed below). You should also check that the references are correct e.g. if you used a country-specific unique source but a generic regional reference list you would need to adjust.  

Review Checklist:

<ul>
  <li>Add the Zenodo link for the country (available in the DiBs for review spreadsheet) to the Specifications Table in the ‘Data accessibility’ row (bottom row). Amend so the text says the following: 
    <ul> <li>‘With the article and in a repository. Repository name: Zenodo. Data identification number: v1.0.0. Direct URL to data: insert link’  </li>
      </ul> 
  </li>
  <li> Also add the Zenodo reference for the country to reference 32 or 31 in the reference list. You can reference it using the following:  
<ul> 
  <li>Cannone, C., Allington, L., Pappis, I., Cervantes Barron, K., Usher, W., et al. (2021). CCG Starter Data Kit: Country. (Version v1.0.0) [Data set]. Zenodo. DOI </li>
  </ul> </li>
  <li> Check that the top cells are merged for tables 1, 3, and 6 e.g. so Estimated Installed Capacity runs across 2015, 2016, 2017, 2018 in table 1 (this should be done but just a check)  </li>
  <li> Check that all the tables are filled in </li>
  <li> Check the first sentence on refinery capacity (section 1.4) makes sense </li>
  <li> In the first paragraph in section 2, check that only one of ‘the International Energy Agency (IEA)/UN Stats’ is written – if both are left, check which one is used for reference 26/27 and delete the other  </li>
  <li> Click on the URLs in the references section to make them hyperlinks  </li>
  <li> Give the results a sense check: 
    <ul>
     <li> Check that fossil future generation is nearly all fossil and that net zero is 100% renewable in 2050 </li> 
     <li> Check that the net zero scenario doesn’t seem ridiculously unrealistic – if it does, flag it and we can delete it for that country </li>
      <li> Make sure that electricity imports aren’t included in capacity graphs, only in generation graphs</li> 
      <li> Check that transport is all electric by 2050 in net zero</li>
      <li>Check that the 3 CO2 graphs start from roughly the same point in 2020 and that they are falling towards 0 in net zero </li> 
      </ul>
  </li>
  <li> Add ‘Writing - Review & Editing’ to your name in the credit statement and make sure you’re in the author list (with your affiliation) </li>
 
</ul> 






Please cite this CCG Starter Kits repository as:
K. Cervantes Barron, W. Usher, C. Cannone, and L. Allington, “CCG_Starter_Kits: Automatic Data in Brief article creation,” CCG_Starter_Kits: Automatic Data in Brief article creation, 2021. [Online]. Available: https://zenodo.org/record/6243119.
