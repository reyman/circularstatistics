us17 = read.csv("data/Fabric adjusted residual_J5_US17 et 18_juin 2015 - Macro_protov2.csv")
library(tidyr)
library(dplyr)

selection = "Orientation 1"
classe = 20
orientation = 360
filtre = "US"
critere = "US17"

breaksVal <- seq(0, orientation, by = classe)
filteredUS <- filter(us17, US=="US17" & Orientation.1 !="NA") %>% select(Orientation.1)
# http://www.r-bloggers.com/r-function-of-the-day-cut/
filteredUS$levels <- cut(filteredUS$Orientation.1, breaks = breaksVal, right = FALSE )

statByFactor <- filteredUS %>% group_by(levels) %>%  summarize(nb = n()) %>% complete(levels, fill = list("nb" = 0))

# mean objects by class group (observed objects / number of classes )
sumOfObject <- sum(statByFactor$nb)
meanObjectByClass <- sumOfObject / (orientation/classe) 

statByClass <- statByFactor %>% 
  mutate(diff_obs_exp = statByFactor$nb - meanObjectByClass) %>%
  mutate(sqrt_exp = sqrt(meanObjectByClass)) %>%
  mutate(eij = diff_obs_exp / sqrt_exp)  %>%
  mutate(vij = 1 - (nb / sumOfObject)) %>%
  mutate(sqrt_vij = sqrt(vij)) %>%
  mutate(dij = eij / sqrt_vij) %>%
  mutate(P_calcule = pnorm(dij, 0, 1, TRUE)) %>%
  mutate(P_reduced = ifelse(P_calcule < 0, P_calcule, round(1 - P_calcule,4)))
  


