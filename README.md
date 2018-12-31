# stockMarketPredictionTool
This is an Excel based automated tool developed by me which extensively uses D3.Js, R, Excel VBA and Phantom JS to provide an integrated appication that not only automatically fetches the data from the web but also develops models, creates interactive hovering data label charts on excel and also test the accuracy of the predictions just at the click of a button. Though the models are very basic, the automation mechanism is effective and can be used by people who still want to rely on basic excel apps for stock data analytics. 

# Demo of the tool available here: https://youtu.be/VBx7Ik6aw7c

# Instructions to be followed before running the application:

1	Setting up IIS localhost server on your system:								
	a) Go To Windows Start and type "Internet Information Services" in the search box								
	b) Click on "Internet Information Services (IIS) Manager"								
	c) Click on the arrow besides the system user name under "Connections" to expand the list								
	d) Right-click on "Sites" and click on "Add Web Site"								
	e) Put a Site Name of your choice and set the Physical path to the location of the "DissertationII" folder (You would have already extracted, unzipped and saved the DissertationII folder in your local drive)								
	f) Click on "Pass through Authentication"								
	g) Click on "Specific User:"								
	h) Click on "Set"								
	i) Enter your system user name (along with workgroup name, if any) and your system password								
	j) Click "OK"								
	k) Click "OK" again								
	l) Click on "Test Settings" and ensure the connection is successful before proceeding								
	m) Assign Port number as "81" and click on OK															
	n) Now go to your browser (preferably Chrome) and type "localhost:81" in the address bar								
	o) Ensure you do not get any error message & all sub-folders in the DissertationII folder are listed on the page								

2	Installing R for Windows 3.4.3:								
	a) Download the setup of "R for Windows 3.4.3" from the following link: https://cran.r-project.org/bin/windows/base/old/3.4.3/				
	b) Install R by opening the setup file and following the instructions								

3	Setting up PhantomJs:								
	a) Download the  "phantomjs.exe" file for windows from the following link: http://phantomjs.org/download.html								
	b) Extract the contents of the downloaded zip file								
	c) Create a folder in C:\ drive called "PhantomJs"								
	d) Copy the contents of the downloaded extracted folder in the C:\PhantomJs folder:								
												
4	Setting up Environment Variables:								
	a) Go To Windows Start and type "Environment Variables" in the search box								
	b) Click on "Edit environment variables for your account"								
	c) Click on the "New" button in the "Environment Variables" Dialog box								
	d) Mention "Variable Name" as: "PATH"								
	e) Mention "Variable Value" as: the path of the exe files for R and PhantomJS like this:								
	C:\Program Files\R\R-3.4.3\bin;C:\PhantomJs\bin								
													
5	Setting up Analysis Toolpak Excel Add-ins								
	a) Go to File -> Excel Options in the Excel Application.								
	b) Go to Add-ins.								
	c) In the Manage box, click Excel Add-ins, and then click Go.								
	d) The Add-Ins dialog box appears.								
	e) In Add-Ins available box, select "Analysis Toolpak" & "Analysis Toolpak-VBA" check box, & click OK.								
													
6	Important: Click on "Enable Content" before running the application.								
7	All Set!!! Now the application is ready to be used!								
