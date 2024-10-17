import requests
from bs4 import BeautifulSoup
import pandas as pd
import logging
from datetime import datetime

date = datetime.now().strftime("%m%d%y")

#Sets up logging
logging.basicConfig(format='%(asctime)s %(levelname)-8s %(message)s', level=logging.INFO, filename="update_item_notes.log", datefmt='%Y-%m-%d %H:%M:%S')

#Set up default things for base URL and bib API key
alma_base = 'https://api-na.hosted.exlibrisgroup.com/almaws/v1'

#Insert Bib R/W API key:
bibapi = 'ENTER-BIBS-RW-API-KEY-HERE'


#Required headers for requests to work
headers = {"Accept": "application/xml", "Content-Type": "application/xml"}

#Update log file
logging.info("Script started-" + date)

#Take input file
inputfile = input('Enter Input Filename without extension (Note: must be in .xlsx format and in same folder as this script): ')

#Keep_default_na=False so that it doesn't add "nan" to blank cells from input file
source = pd.read_excel(f"{inputfile}.xlsx", keep_default_na=False)

#Reads columns from input spreadsheet, all as strings
lookupitembarcode_raw = source['Barcode'].astype(str)

lookup_update_fulfillment_note = source['Fulfillment Note'].astype(str)
lookup_update_public_note= source['Public Note'].astype(str)
lookup_update_internal_note_1 = source['Internal Note 1'].astype(str)
lookup_update_internal_note_2 = source['Internal Note 2'].astype(str)
lookup_update_internal_note_3 = source['Internal Note 3'].astype(str)
lookup_update_statistics_note_1 = source['Statistics Note 1'].astype(str)
lookup_update_statistics_note_2 = source['Statistics Note 2'].astype(str)
lookup_update_statistics_note_3 = source['Statistics Note 3'].astype(str)


#Removes precautionary ""s from long ID numbers
lookupitembarcode = ([s.replace('"', '') for s in lookupitembarcode_raw])



#The actual function, loop through input file data, pull record from Alma, update the enum/chron fields based on input file data, send data back to Alma
for i, itemid in enumerate(lookupitembarcode, 0):
   barcode = lookupitembarcode[i]
   update_fulfillment_note = lookup_update_fulfillment_note[i]
   update_public_note = lookup_update_public_note[i]
   update_internal_note_1 = lookup_update_internal_note_1[i]
   update_internal_note_2 = lookup_update_internal_note_2[i]
   update_internal_note_3 = lookup_update_internal_note_3[i]
   update_statistics_note_1 = lookup_update_statistics_note_1[i]
   update_statistics_note_2 = lookup_update_statistics_note_2[i]
   update_statistics_note_3 = lookup_update_statistics_note_3[i]


   #Let user know what's going on and also add to log file
   print(f"Processing {barcode}, entry # {str(i+1)} of {str(len(lookupitembarcode))}")
   logging.info(f"Processing {barcode}, entry # {str(i+1)} of {str(len(lookupitembarcode))}")


   r = requests.get(f"{alma_base}/items?view=label&item_barcode={barcode}&apikey={bibapi}", headers=headers)
   
   #Check that response is 200 (record found, no errors), log any errors, if record is found then proceed 
   if r.status_code != 200:
      logging.error(f"Error on {barcode}, record not found?")
   else:
      soup = BeautifulSoup(r.content, "xml") 
     
      mms_id = soup.mms_id.string
      holding_id = soup.holding_id.string
      item_pid = soup.pid.string

      #Find enum and chron fields in item record data, replace their contents with the updated data from spreadsheet:
      fulfillment_note = soup.fulfillment_note
      fulfillment_note.string = update_fulfillment_note

      public_note = soup.public_note
      public_note.string = update_public_note

      internal_note_1 = soup.internal_note_1
      internal_note_1.string = update_internal_note_1

      internal_note_2 = soup.internal_note_2
      internal_note_2.string = update_internal_note_2

      internal_note_3 = soup.internal_note_3
      internal_note_3.string = update_internal_note_3

      statistics_note_1 = soup.statistics_note_1
      statistics_note_1.string = update_statistics_note_1

      statistics_note_2 = soup.statistics_note_2
      statistics_note_2.string = update_statistics_note_2

      statistics_note_3 = soup.statistics_note_3
      statistics_note_3.string = update_statistics_note_3

      #Takes the Soup item...
      newitemdata = soup.item

      #Makes it a string instead of an object??? (So it can be sent off in the request)
      newitemdata_str = str(newitemdata)


      #Sends updated item record back to Alma as a PUT request
      updatepush = requests.put(f"{alma_base}/bibs/{mms_id}/holdings/{holding_id}/items/{item_pid}?generate_description=false&apikey={bibapi}", data=newitemdata.encode('utf-8'),headers=headers)

      #Checks that update was successful and logs result
      if updatepush.status_code != 200:
         logging.error(f"Record not updated, Item ID {barcode}")
      else:
         logging.info(f"Item barcode {barcode} updated successfully")