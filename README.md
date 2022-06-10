# PART A: SETTING UP YOUR PC TO RUN THE SCRIPT

1. Install Python 3.8.10
	* Go to Python Release [Python Release Python 3.8.10 | Python.org](https://www.python.org/downloads/release/python-3810/) – really, you can install any version of 3.8 but only some of the 3.8.x version have the helpful “installer option” at the bottom of the page. Move around the version list if you need slightly different version of 3.8.x but make sure it is 3.8.x – it doesn’t work with 3.9 or 3.10 or 3.6
	* Scroll to the bottom and click the appropriate “Windows Installer” hyperlink option (e.g. 64-bit option below)...
![python-download-selection](/readme-images/python-download-selection.png)
	* Open the file when download finished •  In the window that appears, select “Add Python 3.8 to PATH"
	![install-python-add-to-path](/readme-images/install-python-add-to-path.png)
  

	* Click "Install now"
	* When installation has complete (as below) move on to the next step.
	* ![successfully-installed-python](/readme-images/successfully-installed-python.png)

2.  Install pipenv
	* Open the Command Prompt by clicking windows icon and type “cmd” then select the option that appears
	* Type the command…	

			`pip install pipenv`
    
		...which will install pipenv.
	* Wait for it to complete
3. Create a host folder
	* First create a folder somewhere on your computer where you will store the python script  (the “main.py” file) , the invoices and where the environment will be run to create the report e.g. “C:\Users\P1234567\Documents\pdf_converter_environment_folder”
	* Place the “main.py” file into that folder – THIS IS AVAILABLE FROM THE AUTOMATION AND CAN BE DOWNLOADED DIRECTLY FROM GITLAB – Ask Jay Williams (jay.williams@optus.com.au)
	* Inside of that folder, create another folder named “invoices” – NOTE IT MUST BE NAMED INVOICES EXACTLY OR ELSE THE SCRIPT WILL NOT KNOW WHERE TO LOOK FOR THE INVOICES
	* Now your host folder should look like this…
	 
	 ![host-file-structure-after-invoices-and-main](/readme-images/host-file-structure-after-invoices-and-main.png)

4. Create your virtual environment setup with pipenv
	* Using the Command Prompt, navigate to your host folder which you created above e.g. “C:\Users\P1234567\Documents\pdf_converter_environment_folder” using the command…
	
	`cd “C:\Users\P1234567\Documents\pdf_converter_environment_folder”`
	* When inside that folder, instantiate the environment with the correct python version and required python packages using the command… 
	
	`pipenv install --python 3.8 pdfminer.six openpyxl`
	* If it runs successfully, you should see the following in the command prompt… 
	
		![command-prompt-successfully-created-virtual-environment](/readme-images/command-prompt-successfully-created-virtual-environment.png)
	
	 	… and the following two files in the host folder…

		![host-file-structure-after-generating-virtual-environment](/readme-images/host-file-structure-after-generating-virtual-environment.png)


5. Download the iLovePDF Desktop application
	* Click this link and install it - [iLovePDF Desktop App. PDF Editor & Reader](https://www.ilovepdf.com/desktop)

# PART B: USING THE SCRIPT
1. Remove Optus Permissions on the PDF/PDFs
	* If the PDF has an Optus permission, the file will not be accessible by the program. It may show a little briefcase icon on the top right corner if the file has Optus only permissions…
	![pdf-files-with-permissions](/readme-images/pdf-files-with-permissions.png)

	* To remove permissions, right-click the file/files and select “File Ownership” then click “Personal”. The briefcase icon will disappear.

2. Repair the PDFs with iLovePDF
	* In the iLovePDF application select the “Repair PDF” tool from the right-hand tool-list

	![repair-pdf-tool-selection-ilovepdf](/readme-images/repair-pdf-tool-selection-ilovepdf.png)

	* Click “Open File” and navigate to and select all of your desired invoice files. Select them all & click “Open”
	* Click “Repair PDF
	
	![repair-pdf-button](/readme-images/repair-pdf-button.png)

	* Click “Open Folder” to navigate to the folder where the repaired PDFs are located
  
	![your-pdfs-have-been-repaired-successfully](/readme-images/your-pdfs-have-been-repaired-successfully.png)

 3. Put the PDFs in in the “invoices” folder
	*	Navigate to the host folder that you created in Part A above (e.g. above “C:\Users\P1234567\Documents\pdf_converter_environment_folder”)
	*	Go to the “invoices” folder
	*	Delete (if any) the old PDF invoices or else they will have their data extracted
	*	Place all of the repaired, new PDF invoices into this folder
 4. Run the program in the virtual environment
	* Navigate to the folder where you previously created the pipenv virtual environment, and created the “invoices” directory.
	* Place all of the repaired PDF documents into the “invoices” directory.
	* Using the Command Prompt, navigate to that folder (e.g. above “C:\Users\P1234567\Documents\pdf_converter_environment_folder”)
	* Once you’re in that directory run the command…
	`pipenv run python main.py` 
	 You will see the output of that code in Command Prompt e.g. …
	 ![command-prompt-output](/readme-images/command-prompt-output.png)
	 
 5. Error Handling
	* If an error appears look in the Command Prompt’s output. You may see the following;
	
		**A.** If an error occurred on a specific PDF, the invoice name/number just above the error message.
		* Write that invoice name/number somewhere for later reference.
		* Remove that invoice from the “invoices” folder. It appears that this PDF isn’t compatible with the script (for whatever reason). The data on that invoice will need to be entered into the “invoice_report.xslx” file later manually.
		* TRY RUNNING THE SCRIPT AGAIN AND KEEP DELETING BAD PDFS UNTIL IT RUNS TO COMPLETION as in, it show’s the “++++++++++++END OF SCRIPT RUN+++++++++++++” MESSAGE SEEN IN THE SCREEN SHOT BELOW.
	![command-prompt-output-end-of-script-run](/readme-images/command-prompt-output-end-of-script-run.png)


		**B.**  If no specific PDF name/number appears to be the cause of the issue, reach out for support from the automation team (Sorry. I could have done my job a little better.)
		![apology-dog](/readme-images/apology-dog.png)

		**C.** If the PDF runs to completion, you need to open the xslx report and review the data for accuracy.
		* Look at the data briefly by eye to see if it makes sense. 
		* Then, search the sheet (Ctrl + F on Windows) for "Error - Please enter manually" within the page and make the ammendment. 
		* Then, search the sheet (Ctrl + F on Windows) for "Line item error" within the page and makethe ammendment.


