#' Title: "Producing AI summary of emails"

message("Loading packages")

library(remotes)
library(devtools)
# RDCOMClient can be installed by running pak::pak('omegahat/RDCOMClient')
library(RDCOMClient)
library(tidyverse)
library(lubridate)
library(curl)
library(ellmer)

run_email_summary_once <- function(sharedEmail,
                                   folderType,
                                   start_date,
                                   end_date
                                   ){

message("Starting Outlook application, getting the MAPI namespace, and accessing inbox and archive messages")
OutApp <- COMCreate("Outlook.Application")
outlookNameSpace <- OutApp$GetNameSpace("MAPI")

# Delete the next 2 lines whenever you want to summarize emails from your own Outlook folder.
recipient <- outlookNameSpace$CreateRecipient(sharedEmail)
recipient$Resolve()

# Set the inbox variable to "outlookNameSpace$GetDefaultFolder(6)" whenever you want to summarize emails from your own inbox.
inbox <- outlookNameSpace$GetSharedDefaultFolder(recipient, 6)
if(folderType == "Inbox"){
  messages <- inbox$Items()
}
else if(folderType == "Archives"){
  archive <- inbox$Parent()$Folders("Archive")
  messages <- archive$Items()
}

todays_date <- format(Sys.Date(), "%m%d%Y")

total_summary <- ""

filter_string_total <- paste0("[ReceivedTime] >= '", start_date, " 12:00AM' AND [ReceivedTime] < '", end_date, "11:59PM'")
number_of_emails <- messages$Restrict(filter_string_total)

days_in_between <- as.numeric(difftime(end_date, start_date, units = "days"))

# Any select emails received on or after "start_date" below will be flagged.
if (number_of_emails$Count() > 0) {
for (j in 0:days_in_between) {
  day_beginning <- format(Sys.time() - (j*86400), "%m/%d/%Y 12:00 AM")
  day_end <- format(Sys.time() - (j*86400), "%m/%d/%Y 11:59 PM")

  filter_string <- paste0("[ReceivedTime] >= '", day_beginning, "' AND [ReceivedTime] < '", day_end, "'")
  this_days_emails <- messages$Restrict(filter_string)

  combined_emails <- ""
  summary_chunk <- ""

  # Flagging emails with keywords found in filter
  if (this_days_emails$Count() > 0) {
    for (i in 1:this_days_emails$Count()) {
      item <- this_days_emails$Item(i)
      combined_emails <- paste(combined_emails, "### START OF EMAIL (File: ", item$Subject(), ")\n", item$Body(), "\n### END OF EMAIL\n---\n", collapse = "\n")
    }
    message("Sending email batch to Groq")
    # chat_github was discontinued on July 30, 2026
    # Set up GROQ_API_KEY in .Renviron and sign up on the Groq website using your email address before proceeding
    chat <- chat_groq(model = "llama-3.3-70b-versatile")
    summary_chunk <- chat$chat(paste0("Summarize the following emails:\n\n", combined_emails))
    total_summary <- c(total_summary, summary_chunk)
  } else {
    message("No emails to summarize this day.")
  }
}
    return(total_summary)
  } else {
    message("No emails to summarize.")
  }
}
