
# This is the server logic for a Shiny web application.
# You can find out more about building applications with Shiny here:
# 
# http://www.rstudio.com/shiny/
#

library(formattable)
library(shiny)
library(datasets)
library(dplyr)
library(tidyr)
#library(readr)

# GLOBAL DATA
# read_csv bug sur le origin.1.1
# us17 = read.csv("data/Fabric adjusted residual_J5_US17 et 18_juin 2015 - Macro_protov2.csv", stringsAsFactors = FALSE)

# Define server logic required to summarize and view the selected
# dataset
shinyServer(function(input, output, session) {
  
  baseData <- reactiveValues()

  observe({
    req(input$fileInput)
    baseData$csv <- read.csv(file=input$fileInput$datapath, header=input$header, sep=input$sep, quote=input$quote, dec=input$dec, stringsAsFactor=FALSE)
    })
  
  #updateselectedinput
  observe({
    req(input$fileInput)
    updateSelectInput(session = session,
                      inputId = "columns",
                      choices = names(baseData$csv))
  })
  
  
  observe({
    req(input$columns)
    updateSelectInput(session = session,
                      inputId = "criteria",
                      choices = unique(baseData$csv %>% select_(input$columns)))
  })
  
  observe({
    req(input$fileInput)
    updateSelectInput(session = session,
                      inputId = "valeurs",
                      choices = names(baseData$csv))
  })
  
  
  dataInput <- reactive({
    req(input$columns)
    req(input$criteria)
    req(input$valeurs)
    
    expr <- lazyeval::interp(quote(x == y & z != "NA"), x = as.name(input$columns), y = input$criteria, z = as.name(input$valeurs) )
    #print(file=stderr(), expr)
    filteredUS <- baseData$csv %>% filter_(expr) 
    filteredUS
  })
  
  output$data <- renderDataTable(dataInput())
    
  
  finalTable <- reactive({
    
    req(input$columns)
    req(input$criteria)
    req(input$valeurs)
    
    data <- dataInput()
    
    symetryFilter <- function(X) {
      if (X > input$orientation) {
        return(X - input$orientation)} else return(X )
    }
    
    breaksVal <- seq(0, input$orientation, by = as.numeric(input$classes))
    data$modulorient <- as.numeric(Map(symetryFilter,data[, input$valeurs]))
    
    data$levels <- cut(data$modulorient, breaks = breaksVal, right = FALSE )
    data$levels[is.na(data$levels)] <- paste0("[0,",input$classes,")")
    
    statByFactor <- data %>% group_by(levels) %>%  summarize(nb = n()) %>% complete(levels, fill = list("nb" = 0))
    
    output$detailsclassification <- renderTable(data %>% select_("modulorient", "levels", input$valeurs))
    output$detailsnb <- renderTable(statByFactor %>% select(levels,nb))
    
    #print(file=stderr(), statByFactor)
    
    # mean objects by class group (observed objects / number of classes )
    statByFactorFiltered <- statByFactor %>% filter(!is.na(levels))
    
    sumOfObject <- sum(statByFactorFiltered$nb)
    meanObjectByClass <- sumOfObject / (input$orientation/as.numeric(input$classes)) 
    
    output$detailsmean <- renderUI( {
      str1 <- paste("sum of object = ", sumOfObject) 
      str2 <- paste("orientation/classes = ", input$orientation/as.numeric(input$classes)) 
      str3 <- paste("object by class = ", meanObjectByClass)
      HTML(paste(str1,str2,str3, sep = "<br/>"))})
    
    statByClass <- statByFactorFiltered %>% 
      mutate(diff_obs_exp = statByFactorFiltered$nb - meanObjectByClass) %>%
      mutate(sqrt_exp = sqrt(meanObjectByClass)) %>%
      mutate(eij = diff_obs_exp / sqrt_exp)  %>%
      mutate(vij = 1 - (nb / sumOfObject)) %>%
      mutate(sqrt_vij = sqrt(vij)) %>%
      mutate(dij = eij / sqrt_vij) %>%
      mutate(P_calcule = pnorm(dij, 0, 1, TRUE)) %>%
      mutate(P_reduced = ifelse(P_calcule < 0, P_calcule, 1 - P_calcule)) %>%
      mutate(P_final = round(ifelse(dij >= 0, P_reduced, P_calcule),4))
    
     return (statByClass)
  })
  
  output$result <- formattable::renderFormattable({
    
    chooseColor <- function(x) {
      finalColor <- c()
      colorSelected <- "black"
      
      for (val in x)
      { 
        if (val <= 0.001) {
          colorSelected <- "green"}
        else if(val <= 0.01) {
          colorSelected <- "darkblue"}
        else if (val<= 0.05) {
          colorSelected <- "softblue"}
        else{
          if (val > 0.1) {
            colorSelected <- "red"}
          else {
            colorSelected <-  "orange"}
        }
        
        finalColor <- append(finalColor, colorSelected) 
        
      }
      return(finalColor) 
    }
    
    formattable(finalTable() , list(
       P_final = formatter("span", style = x ~ style(color = chooseColor(x), font.weight = "bold") ))
      )
    
  })
  
  output$downloadTable <- downloadHandler(
    filename = function() {paste('myoutput-', Sys.Date(), '.csv', sep='')},
    content = function(file) {
      write.csv(finalTable(),file)
    }
  )




})

