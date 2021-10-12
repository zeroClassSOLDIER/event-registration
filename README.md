# E.R.A (Event Registration Application)
A tool for organizers to create and manage events for airforce classes and training in SharePoint Online.

# Build
1. Clone the source code
2. Install the libraries
```js
npm i
```
3. Link TypeScript
```js
npm link typescript
```
4. Build the Solution
```js
npm run all
```
# Installation and Deployment
1) Go to Site content -> Site Assets where the app shall live
2) Create a new folder and it _must_ be named "Event-Registration" </br> <span style="color: red;"> !!!IMPORTANT!!! The folder _must_ be named that or ERA will not work! </span>
![image](https://user-images.githubusercontent.com/84786032/135375839-04a622ef-1f03-467d-9da0-0fbc9bcab081.png)
3) Click inside of the folder and upload the following: <br/>
 ![image](https://user-images.githubusercontent.com/84786032/135656176-bba0656e-1ead-4f0d-94e8-b4599e576ebe.png)
5) Navigate to Site Pages <br/>
![image](https://user-images.githubusercontent.com/84786032/135382272-663b6ec0-4bbb-479f-af3c-a9fe4a56d6cc.png)
7) Click New -> Web Part Page
8) Type the name of the page before ".aspx" (era is recommended) and select "Full Page, Vertical" for layout template. <br/>
![image](https://user-images.githubusercontent.com/84786032/135384284-acbfc28c-74d1-4b5b-8350-b835e5db7a41.png)
10) Click "Add a Web Part" and set Categories to "Media and Content" and Parts to "Content Editor" <br/>
![image](https://user-images.githubusercontent.com/84786032/135382433-71365861-3324-4092-ba33-b8ab077bb93f.png)
12) Click the dropdown arrow in the upper right corner and choose Edit Web Part <br/>
![image](https://user-images.githubusercontent.com/84786032/135382472-74cde243-75bc-4a6b-979b-c265e2aa56d8.png)
14) Under "Content Link" copy-paste link to index.html file and under "Appearance" set "Chrome Type" to None <br/>
![image](https://user-images.githubusercontent.com/84786032/135657432-f3c7a841-04c3-4df8-8ff8-d09dba15c92a.png)
16) Click Install on the Installation Required screen <br/>
![image](https://user-images.githubusercontent.com/84786032/135657499-b92ad00a-3b73-41b1-9a31-e729b2ff28d9.png)
18) Click Refresh after installation is loaded <br/>
![image](https://user-images.githubusercontent.com/84786032/135657579-79bc691e-85e6-4c0e-967a-67cffbf16918.png)
20) Click Stop editing if present. 
![image](https://user-images.githubusercontent.com/84786032/135658055-b99a2850-267a-4237-9f48-f3491c4a99aa.png)
22) CONGRATULATIONS!!! ERA is now ready to go! It will appear similar to the screen-shots below. <br/> Administrators/Organizers view
![image](https://user-images.githubusercontent.com/84786032/135383112-b083be33-8360-4393-945a-0bdc3f0ed5db.png) <br/> Members/Attendees view
![image](https://user-images.githubusercontent.com/84786032/135658218-876f7abc-9e3d-40a1-bd04-f738d806d99d.png)



# User's Guide
## Organizer
1) Create an event with New Event button and view what security groups are managers and members with Manage Groups
1) View an event's details by clicking the View icon to the left of Title
2) Upload a document by clicking the Upload icon under Documents. View and delete options will be present
3) Edit, delete, and view an event's roster by clicking the Manage Events icon. View Roster will have a print option
3) Type title or date into Search to find a particular event
4) To view coures that have already started or ended click the filters button next to Search and click Show inactive events
### Change Members and Managers group
1) Navaigate to the "Event-Registration" folder and edit the "eventreg-config.json"
2) Enter the name of the new group for either the admin group or the members group
## Attendee
1) Type title or date into Search to find a particular event
2) View an event's details by clicking the View icon to the left of Title
3) Register or unregister by clicking the Registration button
4) If regitered for an event, use the Add event calendar button to add it to Outlook 
