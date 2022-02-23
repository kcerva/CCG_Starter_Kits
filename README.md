# CCG Starter Kits 
## Automatic Data in Brief article creation

This repository is used to automate the word document creation for country models.

There is one script for each region, containing the appropriate text and layout for that region. You should use the script for the correct region and then edit the file paths, replacing them with the corresponding file path on your PC – there are many file paths, so please make sure that you change them all. If you are working with documents in the cloud, please ensure that you ‘sync local files’ so that they are on your PC.   

You can indicate which country or countries you want to create articles for in the script around line 143. If you want to run multiple countries, you should leave these lines as shown below and add a column to the Starter Kits – List of Countries.xlsx, putting a 1 for countries that you want to include, and a 0 if you don’t want to include them. You should then add the name of that column in the square brackets after Country_Data in lines 143 and 144 (where it says Run_SA below). If you want to only run one country, you should # out lines 143 and 144, then remove the # from lines 146 and 147 and insert the country name and ISO3 code.   

<code>
countries_short = country_data[country_data['Run_SA'] == 1]['Country']

icountries = country_data[country_data['Run_SA'] == 1]['ISO3']
</code>





Please cite this repository as:
K. Cervantes Barron, W. Usher, C. Cannone, and L. Allington, “CCG_Starter_Kits: Automation of word document creation for country files,” CCG_Starter_Kits: Automation of word document creation for country files, 2021. [Online]. Available: https://zenodo.org/record/6243119.
