#' Title: "Producing AI summary of emails"

message("Loading packages")

library(remotes)
library(devtools)
# RDCOMClient can be installed by running pak::pak('omegahat/RDCOMClient')
pak::pak('omegahat/RDCOMClient')
library(RDCOMClient)
library(tidyverse)
library(lubridate)
library(curl)
library(ellmer)
library(xml2)

if(file.exists("scheduler_config.rds")){
  message("Loading scheduler_config.rds")
  scheduler_config <- readRDS("scheduler_config.rds")
  sharedEmail <- scheduler_config$sharedEmail
  folderType <- scheduler_config$folderType
  freq <- scheduler_config$freq
  start_taskscheduler_date <- scheduler_config$start_taskscheduler_date
  folder_dir <- scheduler_config$folder_dir
} else {
  message("scheduler_config.rds not found. Please run the hackathon_shiny_app.R script to create the configuration file.")
}

message("Starting Outlook application, getting the MAPI namespace, and accessing inbox messages")
OutApp_scheduler <- COMCreate("Outlook.Application")
outlookNameSpace_scheduler <- OutApp_scheduler$GetNameSpace("MAPI")

# Delete the next 2 lines whenever you want to summarize emails from your own Outlook folder.
recipient_scheduler <- outlookNameSpace_scheduler$CreateRecipient(sharedEmail)
recipient_scheduler$Resolve()

# Set the inbox_scheduler variable to "outlookNameSpace_scheduler$GetDefaultFolder(6)" whenever you want to summarize emails from your own inbox.
inbox_scheduler <- outlookNameSpace_scheduler$GetSharedDefaultFolder(recipient_scheduler, folderType)
if(folderType == "Inbox"){
messages_scheduler <- inbox_scheduler$Items()
} else if(folderType == "Archives"){
    archive_scheduler <- inbox_scheduler$Parent()$Folders("Archive")
    messages_scheduler <- archive_scheduler$Items()
}

total_summary_scheduler <- ""

if(freq == "DAILY"){
    filter_string_total_scheduler <- paste0("[ReceivedTime] >= '", format(Sys.Date(), "%m/%d/%Y"), " 12:00AM' AND [ReceivedTime] < '", format(Sys.Date(), "%m/%d/%Y"), "11:59PM'")
    days_in_between_scheduler <- 0
} else if(freq == "WEEKLY") {
    filter_string_total_scheduler <- paste0("[ReceivedTime] >= '", format(Sys.Date() - 7, "%m/%d/%Y"), " 12:00AM' AND [ReceivedTime] < '", format(Sys.Date(), "%m/%d/%Y"), "11:59PM'")
    days_in_between_scheduler <- 7
} else if(freq == "MONTHLY") {
    filter_string_total_scheduler <- paste0("[ReceivedTime] >= '", format(Sys.Date() %m-% months(1), "%m/%d/%Y"), " 12:00AM' AND [ReceivedTime] < '", format(Sys.Date(), "%m/%d/%Y"), "11:59PM'")
    days_in_between_scheduler <- as.numeric(difftime(start_taskscheduler_date, format(start_taskscheduler_date %m-% months(1)), units = "days"))
}

number_of_emails_scheduler <- messages_scheduler$Restrict(filter_string_total_scheduler)

# Any select emails received on or after "start_date" below will be flagged.
if (number_of_emails_scheduler$Count() > 0) {
for (j in 0:days_in_between_scheduler) {
  day_beginning_scheduler <- format(Sys.Date() - j, "%m/%d/%Y 12:00 AM")
  day_end_scheduler <- format(Sys.Date() - j, "%m/%d/%Y 11:59 PM")

  filter_string_scheduler <- paste0("[ReceivedTime] >= '", day_beginning_scheduler, "' AND [ReceivedTime] < '", day_end_scheduler, "'")
  this_days_emails_scheduler <- messages_scheduler$Restrict(filter_string_scheduler)

  combined_emails_scheduler <- ""
  summary_chunk_scheduler <- ""

  # Flagging emails with keywords found in filter
  if (this_days_emails_scheduler$Count() > 0) {
    for (i in 1:this_days_emails_scheduler$Count()) {
      item <- this_days_emails_scheduler$Item(i)
      combined_emails_scheduler <- paste(combined_emails_scheduler, "### START OF EMAIL (File: ", item$Subject(), ")\n", item$Body(), "\n### END OF EMAIL\n---\n", collapse = "\n")
    }
    message("Sending email batch to Groq")
    # chat_github was discontinued on July 30, 2026
    # Set up GROQ_API_KEY in .Renviron and sign up on the Groq website using your email address before proceeding
    chat_scheduler <- chat_groq(model = "llama-3.3-70b-versatile")
    summary_chunk_scheduler <- chat_scheduler$chat(paste0("Summarize the following emails and explicitly tell me if any actions need to be taken:\n\n", combined_emails_scheduler))
    total_summary_scheduler <- c(total_summary_scheduler, summary_chunk_scheduler)
  } else {
    message("No emails to summarize this day.")
  }
}
    writeLines(total_summary_scheduler, paste0(folder_dir, "/Groq_email_summary", todays_date, ".txt"))
    return(total_summary_scheduler)
  } else {
    message("No emails to summarize.")
  }
