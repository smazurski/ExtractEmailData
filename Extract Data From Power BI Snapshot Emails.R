library(tidyverse)
library(magick)
library(tesseract)
library(reshape2)
library(devtools)
library(RDCOMClient)
library(remotes)
library(RDCOMOutlook)


# Open Outlook within R
outlook_app <- COMCreate("Outlook.Application")

# Create a search object to search the mail box by given criteria (e.g. subject)
search <- outlook_app$AdvancedSearch(
  "Inbox",
  "urn:schemas:httpmail:subject = 'Inventory Status Report'"
)


# Saves search results into results object
# This part might take a moment. Make sure this finishes before running the next line.
results <- search$Results() 

# Count the number of emails with this subject line - I happened to have 5870 at this moment
results$Count()
  
# Save Email Results
email5870 <- results$Item(5870)


# Create temporary file 
target <- tempfile()

# Save email image attachments to target tempfile
testfilename <- email5870$Attachments(1)$SaveAsFile(target)


# read image into object
image_I_S5870 <- image_read(target)


### Loading File From PC for purposes of this tutorial
image_I_S5870 <- image_read("C:/users/steve/documents/Inventory Status - (Status Report).png")
###


# Crop image
image_I_S5870 <- image_crop(image_I_S5870, geometry = "300x1000")


# OCR values from image
img_rows <- ocr_data(image_I_S5870)


# View Data Frame
view(img_rows)


# Extract relevant values from image
img_rows <- img_rows[c(2,8,12,16,20,24,28,30:31),1]


# Convert to numeric
img_rows$word <- as.numeric(img_rows$word)


# Bind the yard names
img_rows$Yard <- rbind("Inventory", "EastClass", "EastDeparture", "EastReceiving", 
                       "WestClass", "WestDeparture", "WestReceiving", "EastRehumps","WestRehumps")


# Spread dataframe
# Inventory <- spread(img_rows,Yard, word)


# include DateTime Stamp
# Inventory <- cbind(Inventory, as.POSIXct(results$Item(5870)$ReceivedTime()*(60*60*24),origin="1899-12-30", tz="GMT"))

# Change Timestamp Name
# colnames(Inventory)[10] <- "TimeStamp"


# create excel file and copy values
# write.csv(Inventory, file = "c:/users/steve/documents/inventory_history.csv")
