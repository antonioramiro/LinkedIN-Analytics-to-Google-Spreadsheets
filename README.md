# LinkedIN Analytics to Google Spreadsheet

This tool aims at **easing the analysis of a company's online marketing performance**. It was developed during a Summer Internship (2020) at [Value for Health CoLab](vohcolab.org) by [António Ramiro](https://antonioramiro.github.io/?utm_source=github&utm_medium=readme&utm_campaign=LinkedIN-Analytics-to-Google-Sheets), under the supervision of [Salomé Azevedo](https://pt.linkedin.com/in/salome-azevedo), Head of Digital Health at [Value for Health CoLab](vohcolab.org). 

The internship comprised the design and **creation of a proof of concept of a Digital Marketing Dashboard**, that allows the Communication Manager at VoH CoLab to have a quick overview over all of the organization's online marketing. This repository is the first step of the dashboard and encompasses the **importation of data from Linked-IN's company analytics page to a Google Spreadsheet, that it can subsequently update, when provided more data.**

The ultimate goal would be to provide an **automated dashboard that integrates updated analytics from several sources**, whether it be social networks, such as Linked-IN and Twitter, or Google Analytics all in one place. The dashboard, where KPIs are computed, is not automatically created and is not included in this repository.

Feel free to contact me via e-mail (antonio.ramiro@tecnico.pt) if you have any questions, doubts or suggestions. This project is subject to a MIT License (`LICENSE` file), so help yourself and fork it!

# Instructions of use

## Google Sheets API setup
*Tutorial based on https://developers.google.com/sheets/api/quickstart/python*

 1. Go to [Google Cloud Platform](https://console.developers.google.com/);
 2. Click in *Select a project* and then, in the top right corner, click *New Project*;
 3. Name the project and, if available, select an organization;
 4. Click in *Select a project* and then select the project you just created;
 5. Search for `Google Sheets API` or go [here](https://console.developers.google.com/apis/library/sheets.googleapis.com) and click *Enable*;
 6. In your project page, click *Create Credentials*;
 7. This should open a Credentials page, where you name the API you're using (*Google Sheets API*) in the first dropdown menu, choose *Other UI* in the second dropdown and choose *User data* in the last form input;
 8. Click *What credentials do I need?* and in the modal prompt, click *Set up consent screen*;
 9. Choose *External*, click *Create*.
 10. Name your application, choose the e-mail associated to the app and fill the last field, regarding the developer's contact. The remaining fields are not mandatory. *Save and continue*;
 11. *Save and continue* and *Save and continue*.
 12. Head to [Credentials](https://console.developers.google.com/apis/credentials) in the left-side panel and choose *Create*, *OAuth Client ID*;
 13. Choose *Desktop app* as the type of application, choose a name for the client and click *Create*;
 14. Click *Ok* in the confirmation menu and then download the credentials, by clicking in the *download symbol*, next to the client you just created.
 15. Rename the file to `credentials.json` and copy it into the same folder as `dmt.py`.

## Dependencies and needed files

 1. To run this python script, there are some dependencies needed beforehand. Install them for the script to work:

| Name  | Role  | Import |  Documentation |
|:---|:---|:---:|---:|
| os | Miscellaneous operating system interfaces | `pip install os`  | https://docs.python.org/3/library/os.html | 
| pandas | Fast, powerful, flexible and easy to use open source data analysis and manipulation tool  | `pip install pandas` |  https://pandas.pydata.org/docs/user_guide/index.html |
| pickle | Python object serialization | (bundled with Python) | https://docs.python.org/3/library/pickle.html |
|   google-api-python-client | Interacting with the API via Python  |  `pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib ` | https://github.com/googleapis/google-api-python-client/blob/master/docs/README.md |
|    google-auth-httplib2 | Library to help migrate users from oauth2client to google-auth  |  " | " |
|   google-auth-oauthlib | Login to your Google account  | " | " |



</br>

2. Next, export the excel files from your Linked-IN page, by going to your company Admin page on Linkedin - it can be found under https://www.linkedin.com/company/XXXXXX/admin/, where XXXXXX is your company Linked-ins's ID;
3. Click **Analytics**, choose **Visitors**, then click the **Export** button on the right, choose the desired time frame and click **Export**;
4. Do the previous step, but instead of **Visitors**, choose **Updates**;
5. Do the previous step, but instead of **Updates**, choose **Followers**;
6. Drag the three `.xls` files to the same folder as `dmt.py` and you can now run the script. 

## Running the script for the first time
1. Execute `dmt.py`;
2. You'll be asked to open a link to Google's Authentication page, open it and use the same account you used while creating the project;
3. You need to have the **three excel files** exported from Linked-in and **no .url file*, since that is what the script uses to detect whether it is the first time it's being run or not.
4. You can find a Google Spreadsheet created in your Google Drive. You can move it inside your drive.

## Tips
- You can add other users to your API, by going to https://console.developers.google.com/ and selecting the folder with the cogwheel at the top. Then you click in your projects name and on the right-side tab you can **add member**.
- After the first run of the script, the file `token.pickle` holds a token for your credentials, so you don't have to be always logging in.
- You should not share `token.pickle` nor `credentials.json` to people you don't trust.
- After the first run, you can share the `.url` shortcut with your team, so that they can also run the script to the same spreadsheet you're using. Their accounts should be added to the project and the spreadsheet file should have them as editors.
- Deleting the `.url` file will create a new spreadsheet. If accidentally deleted, just create a new shortcut to the old spreadsheet.
- The tool is highly dependant on the names of the `.xls` files exported from Linked-IN. They should not be changed.
- You can find old `.xls` files from your previous runs at `/archive`, labelled with the date of import.

# Developer notes

## Tasks Left Undone
- Convert .py file and its dependencies to one `.exe` or `.dmt` executable, using [pyinstaller](https://www.pyinstaller.org/);
- Add computed KPIs from the .xls files on script, instead of manually, in the sheet;
- Currently, only sheets with temporal data are added correctly, since it's a straightforward process of appending the data. However, there are two additional types of data, that are currently not being updated the right way:
    - **Updates Aggregated Data**: Update data is comprised of a list of updates (posts) and its engagement data. This type of data is continuously changing, meaning that a Linked-IN post from a year ago can still receive interactions and therefore change its data. In order to always have the most recent data in the DMD, the goal would be to read the the excel source file and match it with the existing data in the Google Spreadsheet data. A cycle would run the posts in the source excel and update the information uploaded to the GDrive sheet previously. (This applies to the second sheet of the updates.xls, Update engagement - non-aggregated)
    - **Demographics data**: Demographic data in the Excel files corresponds to sheets that only have two columns, one of them labelling the data and the second with the actual data. For instance, `vohcolab_followers_1600336927413.xls` > `Company size`, has in its first column the labels (0-3 people; 3-10 people; 10-50 people; etc...) and in the second column *how many people work in a company which fits each of the labels*. My idea to integrate this in the DMD would be to firstly create a column with all the labels and then append new columns, whose header would have the date in which the append happened. Note that to append new data, the left-most column has to be read, in order for the data placement to be correct (apples to apples). (Demographics data can be found in followers.xls in all but the first sheet and in visitors.xls in all but the first sheet);
- A routine that verifies that the data being added isn't already there and, if there is, to skip it until new data is found (so that there are no repeated days.);
- Add the remaining sheets as data sources (Update engagement, followers.xls all but the first sheet and in visitors.xls all but the first sheet);
- Edit date format to the european standard before sending it to the Spreadsheet;



## Good Ideas for the Future
- Instead of sourcing the data from Excel files downloaded from Linked-IN, source the date directly from [Linkedin's API](https://developer.linkedin.com/docs/v1/companies/company-analytics/company-statistics);
- Add more sources of data to overview: Google Analytics, twitter, etc;
- Create the Dashboard with Google's charts API for Sheets;
- Improve visual interface of the dashboard;
- A progress bar, when running the script;
- Remove the `.json` body from the `.py` file and place it elsewhere;
- The `.json` in the `.py` file is hardcoded with the header of the `.xls` to be pasted in the new sheet. This process should be automated with a `for` cycle, that reads the `.xls` and creates the header automatically (not hard-coded)
- Merge append-data process to a one command only operation, instead of one time for each sheet (see ['Writing multiple ranges'](https://developers.google.com/sheets/api/guides/values#writing_multiple_ranges));
- The import function is **extremely** sensitive to the excel's file name. Instead, it could determine which file is which by its content, in a relative manner, and not 'harcodedly' parsing the file name and looking for the categorizing world (visitors, followes or updates).

## Great Ideas
- Instead of using a Google Sheet as a database, use a tool specifically designed to serve as a database (MariaDB, mySQL, MongoDB, Firebase);
- Instead of serving content in a Google Spreadsheet-crafted dashboard, display data in a modern looking web page (example: mkt.vohcolab.org);


SQLite