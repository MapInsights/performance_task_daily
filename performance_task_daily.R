# This is the Waiting Time Task R file 
#======================================
library(dplyr)
library(flexdashboard)
pacman::p_load(rmarkdown, knitr)

## This will need to be changed to your local settings

home_dir <- "C:/Programs/gtc_tasks/performance_task_daily"
setwd(home_dir)

# Set the directory path for later
reports_dir <- paste(home_dir,"reports",sep = "/")
spreadsheets_dir <- paste(home_dir,"spreadsheets",sep = "/")

# Make sure you delete the folders and files from last week 
unlink("reports", recursive = TRUE, force = FALSE)
unlink("spreadsheets", recursive = TRUE, force = FALSE)

# Create a directory for spreadsheets otherwise R having a heart attack 
dir.create("spreadsheets",showWarnings = F)


# Go baby run the report
rmarkdown::render("performance_daily.Rmd", output_dir = "reports" )

# zip dashboard file
#setwd(reports_dir)
#zip(zipfile = 'performance_daily', files = 'performance_daily.html')


# Get the list of the generated files 
html_files <- list.files(reports_dir, recursive=TRUE)
xlsx_files <- list.files(spreadsheets_dir, recursive=TRUE)

# For each file set create the full path 
#html_file_paths <- paste(reports_dir, html_files, sep="/")
#xlsx_files_paths <- paste(spreadsheets_dir, xlsx_files , sep="/")
xlsx_name <- paste(as.Date(Sys.time() - (60*60*24)), "_", "results.xlsx", sep = '')
  
# Create mail lists
base_list <- "McGarry.Con@greentomatocars.com;Daria.Alekseeva@greentomatocars.com;Julia.Thomas@transdevplc.co.uk);James.Rowe@greentomatocars.com;sean.sauter@transdevplc.co.uk;Christina.Stone@greentomatocars.com;Ian.Bates@greentomatocars.com;Lee.Holland@greentomatocars.com;controldesk@greentomatocars.com;Duncan.Fendom@greentomatocars.com;maxim.starostin@magenta-technology.com;petr.popov@magenta-technology.com;Paul.Jobling@greentomatocars.com;Tyrone.Hunte@greentomatocars.com;antony.carolan@greentomatocars.com;Haider.Variava@greentomatocars.com;Sophie.Jacobsen@greentomatocars.com;Moses.Adegoroye@greentomatocars.com;Tim.Stone@greentomatocars.com;Andrew.Middleton@greentomatocars.com;Tahir.Nazir@greentomatocars.com;Sales@greentomatocars.com"
#daily_list <- "Daria.Alekseeva@greentomatocars.com"

# Send mail for each client and attach html report + spreadsheet
library(RDCOMClient)

# Send mail for 3D
OutApp <- COMCreate("Outlook.Application")
outMail = OutApp$CreateItem(0)
outMail[["subject"]] = 'Daily Performance GTC'
outMail[["To"]] = base_list
outMail[["body"]] = "Good day. This is an automated e-mail. Daily performance report is attached. Daria"
outMail[["Attachments"]]$Add(paste(reports_dir,'performance_daily.html', sep='/'))
outMail[["Attachments"]]$Add(paste(spreadsheets_dir,xlsx_name, sep='/'))
outMail$Send()
rm(list = c("OutApp","outMail"))