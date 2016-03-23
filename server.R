
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
us17 = read.csv("data/Fabric adjusted residual_J5_US17 et 18_juin 2015 - Macro_protov2.csv", stringsAsFactors = FALSE)

# Define server logic required to summarize and view the selected
# dataset
shinyServer(function(input, output, session) {
  
  #updateselectedinput
  observe({
    updateSelectInput(session = session,
                      inputId = "columns",
                      choices = names(us17))
  })
  
  
  observe({
    req(input$columns)
    updateSelectInput(session = session,
                      inputId = "criteria",
                      choices = unique(us17 %>% select_(input$columns)))
  })

  
  
  dataInput <- reactive({
    req(input$columns)
    req(input$criteria)
    
    expr <- lazyeval::interp(quote(x == y & z != "NA"), x = as.name(input$columns), y = input$criteria, z = as.name("Orientation.1"))
    print(file=stderr(), expr)
    filteredUS <- us17 %>% filter_(expr) 
    filteredUS
  })
  
  output$data <- renderDataTable(dataInput())
  
  output$result <- formattable::renderFormattable({
    req(input$columns)
    req(input$criteria)
    
    data <- dataInput()
    
    breaksVal <- seq(0, input$orientation, by = as.numeric(input$classes))
    data$levels <- cut(data$Orientation.1, breaks = breaksVal, right = FALSE )
    statByFactor <- data %>% group_by(levels) %>%  summarize(nb = n()) %>% complete(levels, fill = list("nb" = 0))
    
    print(file=stderr(), statByFactor)

    
     # mean objects by class group (observed objects / number of classes )
    statByFactorFiltered <- statByFactor %>% filter(!is.na(levels))

    sumOfObject <- sum(statByFactorFiltered$nb)
    meanObjectByClass <- sumOfObject / (input$orientation/as.numeric(input$classes)) 
    
    statByClass <- statByFactorFiltered %>% 
      mutate(diff_obs_exp = statByFactorFiltered$nb - meanObjectByClass) %>%
      mutate(sqrt_exp = sqrt(meanObjectByClass)) %>%
      mutate(eij = diff_obs_exp / sqrt_exp)  %>%
      mutate(vij = 1 - (nb / sumOfObject)) %>%
      mutate(sqrt_vij = sqrt(vij)) %>%
      mutate(dij = eij / sqrt_vij) %>%
      mutate(P_calcule = pnorm(dij, 0, 1, TRUE)) %>%
      mutate(P_reduced = ifelse(P_calcule < 0, P_calcule, round(1 - P_calcule,3)))
    
   chooseColor <- function(x) {
     finalColor <- "black"
     
     print(file=stderr(), x)
     
     if (x <= 0.001) {
       finalColor <- "green"}
     else if(x <= 0.01) {
       finalColor <- "darkblue"}
     else if (x<= 0.05) {
       finalColor <- "softblue"}
     else{
        if (x > 0.1) {
          finalColor <- "red"}
       else {
         finalColor <-  "orange"}
     }
       return(finalColor) 
   }
   
    formattable(statByClass, list(
      P_reduced = formatter("span", style = ~ style(color = chooseColor(P_reduced), font.weight = "bold") ))
    )
  })  




})

