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
library(rvest)

if(file.exists(paste0(dirname(Sys.getenv("HACKATHON_R_SCRIPT_PATH")), "/scheduler_config.rds"))){
  message("Loading scheduler_config.rds")
  scheduler_config <- readRDS(paste0(dirname(Sys.getenv("HACKATHON_R_SCRIPT_PATH")), "/scheduler_config.rds"))
  sharedEmail <- scheduler_config$sharedEmail
  folderType <- scheduler_config[[2]]
  freq <- scheduler_config$freq
  start_taskscheduler_date <- scheduler_config$start_taskscheduler_date
  folder_dir <- scheduler_config[[3]]
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
inbox_scheduler <- outlookNameSpace_scheduler$GetSharedDefaultFolder(recipient_scheduler, 6)
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
      item_scheduler <- this_days_emails_scheduler$Item(i)
      
      get_message_text_scheduler <- function(item_scheduler){
        body_format_scheduler <- tryCatch(item_scheduler$BodyFormat(), error = function(e) NULL)
        if(body_format_scheduler == 0 | body_format_scheduler == 3) {
          # Force Outlook to convert unidentified and RTF formats to HTML
          item_scheduler[["BodyFormat"]] <- 2
          body_format_scheduler <- 2
          
        }
        if(body_format_scheduler == 2) {
          email_text_scheduler <- tryCatch(item_scheduler$HTMLBody(), error = function(e) "")
          if (!is.null(email_text_scheduler) && nzchar(trimws(email_text_scheduler))) {
            clean_text_scheduler <- tryCatch({
              email_text_scheduler %>%
                read_html() %>%
                html_text2()
            }, error = function(e) "")
          } else if (!nzchar(trimws(email_text_scheduler))) {
            clean_text_scheduler <- tryCatch({
              inspector_scheduler <- item_scheduler$GetInspector()
              word_doc_scheduler <- inspector_scheduler$WordEditor()
              word_doc_scheduler[["Content"]][["Text"]]
            }, error = function(e) "")
            if (clean_text_scheduler == "") {
              clean_text_scheduler <- tryCatch({
                pa_scheduler <- item_scheduler$PropertyAccessor()
                pa_scheduler$GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
              }, error = function(e) "")
            }
          }
        } else if (body_format_scheduler == 1) {
          # Plain text format
          clean_text_scheduler <- tryCatch({
            item_scheduler$Body()
          }, error = function(e) "")
        } else {
          return("No readable text content found in email body")
        }
        return(clean_text_scheduler)
      }
      complete_email_scheduler <- paste("### START OF EMAIL (File: ", item_scheduler$Subject(), ")\n", get_message_text_scheduler(item_scheduler), "\n### END OF EMAIL\n---\n", collapse = "\n")
    }
    message("Sending complete email to Groq")
    # chat_github was discontinued on July 30, 2026
    # Set up GROQ_API_KEY in .Renviron and sign up on the Groq website using your email address before proceeding
    chat_scheduler <- chat_groq(model = "llama-3.3-70b-versatile", params = params(max_tokens = 2500))
    summary_chunk_scheduler <- chat_scheduler$chat(paste0("Summarize the following emails, placing the date ", format(Sys.Date() - j, "%m/%d/%Y"), " on top of the summary, and explicitly tell me if any actions need to be taken:\n\n", complete_email_scheduler))
  } else {
    message("No emails to summarize this day.")
  }
  total_summary_scheduler <- c(total_summary_scheduler, summary_chunk_scheduler)
}
    writeLines(total_summary_scheduler, paste0(folder_dir, "/Groq_email_summary", todays_date, ".txt"))
    return(total_summary_scheduler)
  } else {
    message("No emails to summarize.")
  }
