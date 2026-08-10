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
library(xml2)

run_email_summary_once <- function(sharedEmail,
                                   folder_dir,
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
  day_beginning <- format(end_date - j, "%m/%d/%Y 12:00 AM")
  day_end <- format(end_date - j, "%m/%d/%Y 11:59 PM")

  filter_string <- paste0("[ReceivedTime] >= '", day_beginning, "' AND [ReceivedTime] < '", day_end, "'")
  this_days_emails <- messages$Restrict(filter_string)

  combined_emails <- ""
  summary_chunk <- ""
  email_text <- ""
  clean_text <- ""
  
  # Flagging emails with keywords found in filter
  if (this_days_emails$Count() > 0) {
    for (i in 1:this_days_emails$Count()) {
      item <- this_days_emails$Item(i)
      
      get_message_text <- function(item){
        body_format <- tryCatch(item$BodyFormat(), error = function(e) NULL)
        if(body_format == 0 | body_format == 3) {
          # Force Outlook to convert unidentified and RTF formats to HTML
          item[["BodyFormat"]] <- 2
          body_format <- 2
          
        }
        if(body_format == 2) {
          email_text <- tryCatch(item$HTMLBody(), error = function(e) "")
          if (!is.null(email_text) && nzchar(trimws(email_text))) {
            clean_text <- tryCatch({
              email_text %>%
                read_html() %>%
                html_text2()
            }, error = function(e) "")
        } else if (!nzchar(trimws(email_text))) {
              clean_text <- tryCatch({
                item$Display()
                inspector <- item$GetInspector()
                word_doc <- inspector$WordEditor()
                word_doc$Range()$Select()
                word_doc$Range()$Copy()
                paste(readClipboard(), collapse="\n")
              }, error = function(e) "")
              if (clean_text == "") {
                clean_text <- tryCatch({
                pa <- item$PropertyAccessor()
                pa$GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
              }, error = function(e) "")
              }
        }
          item$Close(0)
      } else if (body_format == 1) {
        # Plain text format
        clean_text <- tryCatch({
          item$Body()
        }, error = function(e) "")
      } else {
          return("No readable text content found in email body")
      }
    return(clean_text)
      }
  complete_email <- paste("### START OF EMAIL (File: ", item$Subject(), ")\n", get_message_text(item), "\n### END OF EMAIL\n---\n", collapse = "\n")
  message("Sending complete email to Groq")
  # chat_github was discontinued on July 30, 2026
  # Set up GROQ_API_KEY in .Renviron and sign up on the Groq website using your email address before proceeding
  chat <- chat_groq(model = "llama-3.3-70b-versatile", params = params(max_tokens = 2500))
  summary_chunk <- chat$chat(paste0("Summarize the following emails, placing the date ", format(end_date - j, "%m/%d/%Y"), " on top of the summary, and explicitly tell me if any actions need to be taken:\n\n", complete_email))
    } 
    Sys.sleep(10)
  } else {
    message("No emails to summarize this day.")
  }
    total_summary <- c(total_summary, summary_chunk)
}
    message("Outputting and saving summary")
    writeLines(total_summary, paste0(folder_dir, "/Groq_email_summary", todays_date, ".txt"))
    return(total_summary)
  } else {
    message("No emails to summarize.")
  }
}
