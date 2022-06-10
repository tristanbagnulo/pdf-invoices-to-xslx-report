# These are all the packages ad modules used in this
# script.
import os
from pdfminer import high_level
from openpyxl import Workbook

print("\n +++++++++BEGINNING OF SCRIPT RUN++++++++++\n")

# Declare an excel workbook
book = Workbook()

# Declare a sheet within that workbook 
sheet = book.active

# Add the column titles to the first row of the excel sheet
# which will later be printed out. This function is enabled
# by the openpyxl package which supports all the excel doc
# commands in this script.
sheet.append(["Invoice Number","PO Ref #","Invoice Date","OB SNOW Ref#","Customer Ref","SLA","Location","Description (item 1)","Quantity (item 1)","Price (item 1)","Total (item 1)","Description (item 2)","Quantity (item 2)","Price (item 2)","Total (item 2)","Description (item 3)","Quantity (item 3)","Price (item 3)","Total (item 3)","Description (item 4)","Quantity (item 4)","Price (item 4)","Total (item 4)", "Total Excluding GST", "Total GST", "Total Including GST", "Error Detection"])

# Access the invoice pdf files and extract the data from each into the excel sheet.
# This is done ONCE for EACH PDF file within the "/invoices" directory. 
# "os.listdir" is a list of the contents of the "invoices" directory.
# E.g. if you places two invoices in the invoices folder, e.g. INV10001 and 
# INV10002 it would hold them in a list as [INV10001,INV10002]. The below
# for-loop takes one of those names at a time and runs the script against it.
# For each loop, the variable "file_name" will refer to the name of the PDF
for file_name in os.listdir('invoices'):

  #Below prints the name of the file being processes.
  print("Start of "+file_name)


  #Extract all text from each invoice document. 
  # First we need to provide the path to the file by adding
  #"inoices/" behind the file name.
  pdf_filename = "invoices/"+file_name
  # the "pages" variable, refers to which page to extract
  # the text from in the PDF. In thsi case, it's only the 1st page
  # if you need to change up the data extraction at all consult the 
  # pdfminer package instructions which you can simply google.
  pages = [0]
  # The below "extracted_text" variable now refers to one continuous string
  # that holds all of the information from the PDF that was scraped.
  # It will only hold the information from 1 PDF as a time and will change
  # with each of the above for loops.
  extracted_text = high_level.extract_text(pdf_filename,'',pages)

  #Data extraction 1 - Find the high-level invoice data
  # This splits the entire string into a list containing two strings:
  # 1. The string before the "Invoice #\n\nDate\n\nYour Ref / PO#\n\nDue Date\n\n"
  # text which is used to determine WHERE to split the string.
  # Its in position 0 of the resulting list; 
  # 2. The string after the text which starts with the data that we want to extract.
  # It is in position 1 (see the "[1]" below which refers to it)
  
  # NOTE: This type of .split() funciton is used MANY MANY times throughout this 
  # script to locate and extract a SPECIFIC piece of information or pieces of
  # information. For more infor on how to use it google "split function python".
  meta_data_containing_string = extracted_text.split("Invoice #\n\nDate\n\nYour Ref / PO#\n\nDue Date\n\n")[1]
  #Store information in variables

  # The string containing the extracted text from a pdf separates many lines of
  # data using two new lines aka "\n\n" (with "\n" denoting a single new line).
  # In the below line of code, I split a single string 3 times on a "\n\n" delimiter
  # i.e. every time it sees "\n\n" it will split the string and place all the characters
  # behind that "\n\n" into a separate string within a list and continue doign that until it
  # splits the string 3 times.
  # I'm doing this because the first 3 lines in the "meta_data_containing_string" string
  # store information that I want to extract and assign to their own identifying variables
  meta_data_strings_list = meta_data_containing_string.split("\n\n",3)

  # After I have a nice list of stinrgs I format them appropriately (e.g. by using
  # "int(string)" to turn a string-type number into an integer data type) and
  # assign them to a variable which I will later call upon to push that data onto the excel
  # report :).
  invoice_no = int(meta_data_strings_list[0])
  invoice_date = meta_data_strings_list[1]
  po_ref_number = int(meta_data_strings_list[2])

  # The below print statements are for testing when needed. Simple uncomment them when
  # you want to see their values at runtime...
  # print ("Invoice #: "+str(invoice_no))
  # print("Invoice date: "+str(invoice_date))
  # print("PO Ref #: "+str(po_ref_number))

  #Data extraction 2 - Get the invoice line items
  
  # The "list(filter(None,list))" takes the list of separate strings
  # and removes every string on it that has a "None" value i.e. ''.
  # This improves the list for later processing.
  extracted_text_list_double_new_line_delimiter = list(filter(None,extracted_text.split("\n\n")))  

  # The below for-loop takes in the "extracted_text_list_double_new_line_delimiter" list above
  # and makes a new list called "line_items_list" which contains only strings that have
  # "Optus Business" within them. These strings only occur for line items on each invoice.
  # These line items contain important information that will go in to our report.
  #First we declare the new list and make it empty 
  line_items_list = []  
  number_of_line_item_errors = 0
  # Then we loop through each string in the list
  for item in extracted_text_list_double_new_line_delimiter:
    # If "Optus Busines" is in it...
    if "Optus Business" in item:
      # ... we add the new string to the the list with ".append(string)""
      line_items_list.append(item)


  # Process each line item-derived string

  # Counter variable for the for-loop
  line_item_count = 0
  # Add line item descriptions to a list for later reference 
  description_list = []
  for item in  line_items_list:
    #Below print statement is for testing and debugging. Remove comments to see at runtime.
    # print("Line item in for loop: "+ item)

    #Add one to counter for each loop undergone (useful later)
    line_item_count+=1

    #Below print statement is for testing and debugging. Remove comments to see at runtime.
    # print("Ongoing Line Item Count: "+ str(line_item_count))

    # Data extracton 2.1 - For the "Meta" Line item
    # On an invoice, there is a line item which contains ">" characters and the word
    # "Metro" or "Regional" - this string is unique and contains vital information
    # so it must be processes uniquely to extract that information. 

    # We determing if the line item is a Meta line item...
    if (">" in item and "Regional" in item) or (">" in item and "Metro" in item) :
      # We mark its position (by index) on the line_items_list list.
      indexOfMetaLineItem = line_items_list.index(item);
      # Clean the item and prepare it for extraction
      # Create another list with all the strings by splitting the meta string at each ">" character
      meta_item_list = item.split(">")

      # Another counter for another for look
      counter_for_meta_item_list_cleaning = 0
      
      for item in meta_item_list:
        #Remove new line characters and space characters from edges of each item....
        # ".strip() removes spaces ' ' from the front and end of the string"
        # ".replace("\n,"")" replaces new lines with nothing
        meta_item_list[counter_for_meta_item_list_cleaning] = meta_item_list[counter_for_meta_item_list_cleaning].replace("\n","").strip()
        
        #Remove the "Optus Business" string within this list. It is no longer useful
        # as it was only used as an indicator to find a string previously.
        if "Optus Business" in item:
          del meta_item_list[counter_for_meta_item_list_cleaning]
      # Cut irrelevant data away from list by retaining only the first 3 items
      meta_item_list = meta_item_list[:3] 

      #The below print statements are for testing and debugging. Remove comments to see the output at runtime.
      # print("Total Meta Item List:")
      # print(meta_item_list) 

      # Data extracton 2.1.1 - Find the the Location
      for item in meta_item_list:
        # If "Metro" is within any string  of the meta_item_list, mark the location as "Metro"
        if "Metro" in item:
          location = "Metro"
          # Mark its location within the list
          index_of_location_item = meta_item_list.index(item)
          # Then delete the item from that list. It is no longer useful.
          del meta_item_list[index_of_location_item]
          break
          # Else, if "Regional" is within any string  of the meta_item_list, mark the location as "Regional"
        elif "Regional" in item:
          location = "Regional"
          # Mark its location within the list
          index_of_location_item = meta_item_list.index(item)
          # Then delete the item from that list. It is no longer useful.
          del meta_item_list[index_of_location_item]
          break


      # Data extracton 2.1.2 - Find the OB SNOW Ref #
      for item in meta_item_list:
        # Check if it's the OB SNOW Ref # or the Customer Ref (name) count string
        # This is done by counting the number of digits and character within each string
        # within the meta_line_item list. This is important to do because the OB SNOW Ref #
        # has a very specific character and number count...
        digit_count = 0
        letter_count = 0
        for ch in item:
          if ch.isdigit():
            digit_count = digit_count+1
          elif ch.isalpha():
            letter_count = letter_count+1
          else:
            pass
        # If it is the OB SNOW Ref #, declare that. If not, next item.
        # Note, all the OB SNOW Ref # values that  I saw had 7 numbers and 
        # 4 or 5 letters so that is the qualifying indicator to mark it as or as not that character.
        if digit_count == 7 and letter_count >=4 and letter_count<=5:
          # Clean away the new lines and spaces on either side of the number 
          # (as above) with strip() and replace().
          ob_snow_ref_number = item.replace("\n","").strip()
          # Declare where the OB SNOW Ref # was within the string.
          index_of_ob_snow_ref_number_item = meta_item_list.index(item)
          
          #Delete that line item as it is no longer useful.
          del meta_item_list[index_of_ob_snow_ref_number_item]
          break

      # Data extraction 2.1.3 - Find the Customer Reference (of customer name)
      # The customer_ref or name is almost always the first string after the "Optus Business"
      # one is removed.
      customer_ref = meta_item_list[0]
        
      # Print the 3 extracted data points (useful for debugging and summary).
      print("Customer Ref: "+customer_ref)
      print("Location: "+location)
      print("OB SNOW Ref#: "+ob_snow_ref_number)
      
    # Data extraction 2.2 - For the SLA-containing line item
    # Now, we are going to extract the data from the line item that I call the "SLA" line item.
    # It is detected by having in it any of the below 6 different SLA codes.
    elif "14x7x2"in item or "24x7x2" in item or"24x7x4" in item or"8x5x4"in item or"8x5xNBD"in item:
      if "14x7x2" in item:
        # If it is detected, the item code is declared with the
        # variable "sla"
        sla="14x7x2"
        # Then, the remaining string is split, has unecessary items removed
        # and is assigned to the variable, "sla_line_item" which is a description
        # of that line item, it is normally "callout" but it can vary.
        sla_line_item = item.split("14x7x2")[-1].replace("-","").strip()
      elif "24x7x2" in item:
        sla="24x7x2"
        sla_line_item = item.split("24x7x2")[-1].replace("-","").strip()
      elif "24x7x4" in item:
        sla="24x7x4"
        sla_line_item = item.split("24x7x4")[-1].replace("-","").strip()
      elif "8x5x4" in item:
        sla="8x5x4"
        sla_line_item = item.split("8x5x4")[-1].replace("-","").strip()
      elif "8x5xNBD" in item:
        sla="8x5xNBD"
        sla_line_item = item.split("8x5xNBD")[-1].replace("-","").strip()
      print("SLA: "+sla)
      # The sla_line_item is added to the "description_list" list for later use.
      description_list.append(sla_line_item)
    
    # Data extraction 2.3 - For "Normal" line items which contain neither Meta nor SLA information
    # These types of line item countain information about the location but no "<" characters as in
    # the Meta item.
    elif "Metro" in item and "<" not in item:
      # For each of these items, useful descriptive informion is extracted by first removing 
      # unecessary data. This improves the clarity of the extracted data.
      normalLineItem = item.replace("- Metro -","").replace("Optus Business","").strip()
    elif "Regional" in item and "<" not in item:
      normalLineItem = item.replace("- Regional - ","").replace("Optus Business","").strip()
      # That description is then added to the description list.
      description_list.append(normalLineItem)
    else:
      # In the rare case where a line item cannot be categorised as Meta, SLA or Normal, an error
      # message is added in the excel report, informing the user that it must be processed manually.
      number_of_line_item_errors += 1
      error_message = "Line item error/s - " +str(number_of_line_item_errors) +" line item errors detected. Please check for missing data & incorrect prices, totals and quantities in the spreadsheet and enter data manually"
      

  # Data extraction 3 - Line item quantities, prices, totals

  # This integer, 'countNonMetaLineItems' is used later for cleaning the data
  # because sometimes an invoice will have a Meta line item and will or will not
  # have a quantity specified for that line item. This integer is used to determine
  # whether there is a quantity that shouldn't be there and remove it if there is.
  # It does this by compating the "countNonMetaLineItems" int with the count of 
  # quantity values.
  countNonMetaLineItems = line_item_count-1

  # The below print statement is for debugging. Uncomment it to see the output.
  # print("Total non-meta line items: "+str(countNonMetaLineItems))
  
  # Again, this separates all the extracted text into separate strings
  # delimited by the "Optus Business" delimiter and takes only the last string
  # which is where the data that we need is located. It then splits THAT string by each
  # new line character "\n". It then removes all the strings that have a None or '' value.
  stringsList = list(filter(None,extracted_text.split("Optus Business")[-1].split("\n")))
  # Then the first one is deleted.
  del stringsList[0]
  
  # declare an arrayList to store the quantity, price and total data for later reference
  quantitiesArray = []
  pricesArray = []
  totalsArray = []

  # Put quantities, prices & totals into lists
  dollarSignIterations = 0
  # print(stringsList)
  for item in stringsList:
    if "Project" in item:
      # When "Project" appears, we have alreads passed the quantity, 
      #price and line-item total data that we need, so it ends the for-loop.
      break
    #if it is a nubmer, it's a quantity
    elif item.isnumeric():
      itemInt = int(item)
      # store into list
      quantitiesArray.append(itemInt)
    # if it has a "$" in it, it's a price OR line item total
    elif "$" in item:\
      # Below print statement is for debugging. Uncomment it to see output at runtime.
      # print("$ item #" + str(dollarSignIterations) + ": " +str(item))

      # Clean the price/item total item by removing $ and , characters.
      cleanedItem = item.replace("$","").replace(",","")
      # Also, cast them to the "flaot" data type (which has decimals unlike an int).
      itemFloat = float(cleanedItem)

      # Store the prices/item-totals into either the prices or totals array.
      # Start with prices and when the number of items stored is the number of
      # Non-Meta line items, start storing prices. The reason it's done this way is because 
      # The price and total data occur as price first then totals after on the extracted text
      # string.
      if dollarSignIterations>=countNonMetaLineItems:
        # store into list
        totalsArray.append(itemFloat)
      else:    
        # store in to list
        pricesArray.append(itemFloat)
      dollarSignIterations += 1

  # Below two print statements are for debugging. Uncomment them to see output at runtime.
  # print("Quantities Array PRE-cleaning:")
  # print(quantitiesArray)

  # Clean up quantities array my deleting the erronious quantity that sometimes appears
  # with the Meta line item.

  # Below two pring statements are debugging. Uncomment them to see output at runtime. 
  # print("Quantities Array PRE-cleaning:")
  # print(quantitiesArray)

  if len(quantitiesArray) > countNonMetaLineItems:
    del quantitiesArray[indexOfMetaLineItem]
  
  # Below two string statements are for debugging. Uncomment them to see output at runtime. 
  # print("Quantities Array POST-cleaning:")
  # print(quantitiesArray)

  # Data extraction 4 - Total invoice amount excluding gst, including gst and gst

  # First, find the location of the data in the extracted text string that we are 
  # looking for. It always comes after the string "TOTAL EX\n\nGST" so split
  # the extracted text on that delimiter and take the 2nd half (at index "[1]").
  string_containing_total_invoice_cost = extracted_text.split("TOTAL EX\n\nGST")[1]

  # As before, split the big string into a list of strings by separating at each "\n\n".
  list_total_invoices_dollar_amounts = string_containing_total_invoice_cost.split("\n\n")
  
  # Add total invoice price items to a list for later processing
  dollar_ammounts_list = []
  for item in list_total_invoices_dollar_amounts:
    # The prices are in the strings with a "$" in them :)
    if "$" in item:
      # Clean up the resulting data and append it to the dollar_amounts_list.
      # Do so by removing $ and , and cast the type from String to float 
      # so that you can store the decimal value.
      dollar_ammounts_list.append(float(item.replace("$","").replace(",","")))
  
  # Below string statement is for debugging. Uncomment it to see output at runtime. 
  # print(dollar_ammounts_list)

  # Error check the items so that they are entered into the right fields.
  # The below formula makes sure that GST is 10% of TOTAL EXCLUDING GST 
  # and that TOTAL INCLUDING GST IS equal to GST + TOTAL EXCLUDING GST.
  if dollar_ammounts_list[0]*.1==dollar_ammounts_list[1] and dollar_ammounts_list[0]+dollar_ammounts_list[1]==dollar_ammounts_list[2]:
    total_excluding_gst = dollar_ammounts_list[0]
    total_gst = dollar_ammounts_list[1]
    total_including_gst = dollar_ammounts_list[2]
    # If it's not correct, these data fields are populated with an error message to tell
    # the user to enter them manually.
  else: 
    total_excluding_gst = "Error - Please enter manually"
    total_gst = "Error - Please enter manually"
    total_including_gst = "Error - Please enter manually"

  # The below 6 string statements are for debugging. Uncomment them to see output at runtime. 
  # print("Quantities array: ")
  # print(quantitiesArray)
  # print("Prices array: ")
  # print(pricesArray)
  # print("Totals array: ")
  # print(totalsArray)

  # Create a list that stores all data to print to excel.
  total_data_list = [invoice_no, po_ref_number, invoice_date, ob_snow_ref_number, customer_ref, sla, location]
  #Add the item description, quantity, price and total information to the main list above...
  loop_counter = 0
  for description in description_list:
    # The below if statement checks the price, quantity and item total data for correctness
    # and adds an error message to the .xslx report document if it doesn't add up correctly.
    if pricesArray[loop_counter]*quantitiesArray[loop_counter]!=totalsArray[loop_counter]:
      quantitiesArray[loop_counter]="Error - Please enter manually"
      pricesArray[loop_counter]="Error - Please enter manually"
      totalsArray[loop_counter]="Error - Please enter manually" 
      
    # Actually add that data to the large total data list in sqeuence so that a description is 
    # followed by its quantity, price and sub-total. This repeats for EVERY line item.
    total_data_list.append(description)
    total_data_list.append(quantitiesArray[loop_counter])
    total_data_list.append(pricesArray[loop_counter])
    total_data_list.append(totalsArray[loop_counter])
    loop_counter += 1
  
    # Fill up empty spaces in the main list with None values.
  length_of_total_list = len(total_data_list)
  if length_of_total_list < 23:
    number_of_spaces_to_fill = 23 - length_of_total_list
    empty_spaces_list = [None]*number_of_spaces_to_fill
    total_data_list.extend(empty_spaces_list)

  # Finish the total data list by appending the final total excl, gst, and total incl cost items.
  total_data_list.extend([total_excluding_gst,total_gst,total_including_gst])
    
  # Print the entire resulting data for the invoice being processed.
  print("Total items list: ")
  print(total_data_list)
  
  # Append all the information in the total_data_list list to the next row in the "invoice_report.xslx" file.
  sheet.append(total_data_list)
  # End of script for THIS invoice document is denoted with the below 
  # "=================================" separator. 
  print("=================================")

# After every invoice in the invoices folder is processed, the xslx document is SAVED
# and is now available in the project folder (the folder in which the "invoices" folder
# is located).
book.save('invoice_report.xlsx')

# The below is printed to indicate the successful processing of ALL invoices
# in the invoices folder.
print("+++++++++END OF SCRIPT RUN++++++++++")